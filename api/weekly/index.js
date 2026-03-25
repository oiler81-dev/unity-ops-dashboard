const { getUserFromRequest } = require("../shared/auth");
const { resolveAccess } = require("../shared/permissions");
const { getTableClient } = require("../shared/table");

const TABLE_NAME = "WeeklyRegionData";

function parseJsonSafely(value, fallback = {}) {
  if (!value) return fallback;

  try {
    return typeof value === "string" ? JSON.parse(value) : value;
  } catch {
    return fallback;
  }
}

module.exports = async function (context, req) {
  try {
    const user = getUserFromRequest(req);
    const access = resolveAccess(user);

    const entity = String(req.query.entity || "").trim();
    const weekEnding = String(req.query.weekEnding || "").trim();

    if (!entity || !weekEnding) {
      return {
        status: 400,
        body: {
          ok: false,
          error: "Missing entity or weekEnding"
        }
      };
    }

    if (!access.isAdmin && access.entity !== entity) {
      return {
        status: 403,
        body: {
          ok: false,
          error: "Forbidden"
        }
      };
    }

    const table = getTableClient(TABLE_NAME);

    let record = null;
    try {
      record = await table.getEntity(entity, weekEnding);
    } catch (err) {
      if (err.statusCode !== 404) {
        throw err;
      }
    }

    const values = record ? parseJsonSafely(record.valuesJson, {}) : {};

    return {
      status: 200,
      body: {
        ok: true,
        found: !!record,
        entity,
        weekEnding,
        values,
        raw: record || null,
        source: record?.source || null,
        status: record?.status || null,
        importedAt: record?.importedAt || null,
        updatedAt: record?.updatedAt || null
      }
    };
  } catch (error) {
    context.log.error("weekly GET failed", error);

    return {
      status: 500,
      body: {
        ok: false,
        error: "Failed to load weekly region data",
        details: error.message
      }
    };
  }
};
