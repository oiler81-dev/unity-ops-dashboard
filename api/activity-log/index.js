const { getUserFromRequest } = require("../shared/auth");
const {
  resolveAccess,
  requireAccess,
  scopeEntitiesToAccess,
  canViewEntity,
  safeErrorResponse
} = require("../shared/permissions");
const { getTableClient } = require("../shared/table");
const { parseJsonSafe } = require("../shared/audit");

const TABLE_NAME = "WeeklyAuditLog";

function safeString(value) {
  return value == null ? "" : String(value);
}

module.exports = async function (context, req) {
  try {
    const user = getUserFromRequest(req);
    const access = resolveAccess(user);

    const authError = requireAccess(access);
    if (authError) return authError;

    const entity = safeString(req.query.entity).trim();
    const weekEnding = safeString(req.query.weekEnding).trim();
    const limit = Math.max(1, Math.min(200, Number(req.query.limit) || 50));

    // If the caller named a specific entity, enforce that they can see it.
    // If they didn't, scope the set of partitions to what they're allowed to read.
    const allEntities = ["LAOSS", "NES", "SpineOne", "MRO"];
    let entitiesToQuery;
    if (entity) {
      if (!canViewEntity(access, entity)) {
        return { status: 404, body: { ok: false, error: "Not found" } };
      }
      entitiesToQuery = [entity];
    } else {
      entitiesToQuery = scopeEntitiesToAccess(access, allEntities);
      if (entitiesToQuery.length === 0) {
        return { status: 200, body: { ok: true, count: 0, items: [] } };
      }
    }

    const client = getTableClient(TABLE_NAME);

    const results = [];

    for (const partition of entitiesToQuery) {
      const filter = `PartitionKey eq '${partition}'`;
      const iter = client.listEntities({
        queryOptions: { filter }
      });

      for await (const row of iter) {
        if (weekEnding && safeString(row.weekEnding) !== weekEnding) {
          continue;
        }

        results.push({
          entity: safeString(row.entity || row.partitionKey),
          weekEnding: safeString(row.weekEnding),
          eventType: safeString(row.eventType),
          actorEmail: safeString(row.actorEmail),
          actorRole: safeString(row.actorRole),
          timestamp: safeString(row.timestamp),
          summary: safeString(row.summary),
          before: parseJsonSafe(row.beforeJson, {}),
          after: parseJsonSafe(row.afterJson, {}),
          metadata: parseJsonSafe(row.metadataJson, {})
        });
      }
    }

    results.sort((a, b) => {
      const aTime = new Date(a.timestamp || 0).getTime();
      const bTime = new Date(b.timestamp || 0).getTime();
      return bTime - aTime;
    });

    const items = results.slice(0, limit);

    return {
      status: 200,
      body: {
        ok: true,
        count: items.length,
        items
      }
    };
  } catch (error) {
    return safeErrorResponse(context, error, "Failed to load activity log");
  }
};