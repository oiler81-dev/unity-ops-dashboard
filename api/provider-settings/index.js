const { getUserFromRequest } = require("../shared/auth");
const {
  resolveAccess,
  requireAccess,
  scopeEntitiesToAccess
} = require("../shared/permissions");
const { getTableClient } = require("../shared/table");

const SETTINGS_TABLE = "ProviderSettingsData";
const ENTITIES = ["LAOSS", "NES", "SpineOne", "MRO"];

// Tessa-confirmed defaults (pods looped into MD count)
const DEFAULTS = {
  LAOSS:    { mdCount: 21, paCount: 14, ptCount: 0  },
  NES:      { mdCount: 12, paCount: 1,  ptCount: 4  },
  SpineOne: { mdCount: 2,  paCount: 3,  ptCount: 2  },
  MRO:      { mdCount: 8,  paCount: 4,  ptCount: 3  }
};

function toNumber(value, fallback = 0) {
  if (value == null || value === "") return fallback;
  const n = Number(value);
  return Number.isFinite(n) ? n : fallback;
}

async function ensureTable(table) {
  try { await table.createTable(); } catch { /* already exists */ }
}

async function getEntitySettings(table, entity) {
  try {
    return await table.getEntity(entity, "settings");
  } catch (err) {
    if (err?.statusCode === 404 || err?.code === "ResourceNotFound" || err?.code === "TableNotFound") return null;
    throw err;
  }
}

module.exports = async function (context, req) {
  const respond = (status, body) => {
    context.res = {
      status,
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body)
    };
  };

  try {
    const user = getUserFromRequest(req);
    const access = resolveAccess(user);

    const authError = requireAccess(access);
    if (authError) return respond(authError.status, authError.body);

    const table = getTableClient(SETTINGS_TABLE);
    await ensureTable(table);

    // ── POST: save provider settings for one entity ──────────────────────────
    if (req.method === "POST") {
      if (!access.isAdmin) {
        return respond(403, { ok: false, error: "Admin only" });
      }

      const body = req.body || {};
      const entity = String(body.entity || "").trim();

      if (!ENTITIES.includes(entity)) {
        return respond(400, { ok: false, error: "Invalid entity" });
      }

      const mdCount = Math.max(0, toNumber(body.mdCount, 0));
      const paCount = Math.max(0, toNumber(body.paCount, 0));
      const ptCount = Math.max(0, toNumber(body.ptCount, 0));

      await table.upsertEntity({
        partitionKey: entity,
        rowKey: "settings",
        entity,
        mdCount,
        paCount,
        ptCount,
        updatedBy: access.email || user?.userDetails || "admin",
        updatedAt: new Date().toISOString()
      });

      return respond(200, {
        ok: true,
        message: "Provider settings saved",
        entity,
        mdCount,
        paCount,
        ptCount
      });
    }

    // ── GET: return entity settings scoped to caller's access ──────────────
    const visibleEntities = scopeEntitiesToAccess(access, ENTITIES);
    const results = await Promise.all(
      visibleEntities.map(async (entity) => {
        try {
          const saved = await getEntitySettings(table, entity);
          const defaults = DEFAULTS[entity];

          return {
            entity,
            mdCount: saved ? toNumber(saved.mdCount, defaults.mdCount) : defaults.mdCount,
            paCount: saved ? toNumber(saved.paCount, defaults.paCount) : defaults.paCount,
            ptCount: saved ? toNumber(saved.ptCount, defaults.ptCount) : defaults.ptCount,
            updatedBy: saved?.updatedBy || null,
            updatedAt: saved?.updatedAt || null,
            isDefault: !saved
          };
        } catch {
          const defaults = DEFAULTS[entity];
          return { entity, ...defaults, updatedBy: null, updatedAt: null, isDefault: true };
        }
      })
    );

    return respond(200, { ok: true, entities: results });

  } catch (error) {
    try { context?.log?.error?.("provider-settings failed", error); } catch (_) {}
    return respond(500, { ok: false, error: "Failed to process provider settings" });
  }
};
