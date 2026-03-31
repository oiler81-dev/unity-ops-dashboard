const { getUserFromRequest } = require("../shared/auth");
const { resolveAccess } = require("../shared/permissions");
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

    if (!access.allowed) {
      return {
        status: 403,
        body: {
          ok: false,
          error: "Forbidden"
        }
      };
    }

    const entity = safeString(req.query.entity).trim();
    const weekEnding = safeString(req.query.weekEnding).trim();
    const limit = Math.max(1, Math.min(200, Number(req.query.limit) || 50));

    const client = getTableClient(TABLE_NAME);

    const results = [];
    const entitiesToQuery = entity ? [entity] : ["LAOSS", "NES", "SpineOne", "MRO"];

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
    context.log.error("activity-log failed", error);

    return {
      status: 500,
      body: {
        ok: false,
        error: "Failed to load activity log",
        details: error.message
      }
    };
  }
};