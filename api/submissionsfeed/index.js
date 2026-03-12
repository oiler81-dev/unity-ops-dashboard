const { getTableClient } = require("../shared/table");

function normalizeString(value) {
  if (value === undefined || value === null) return "";
  return String(value).trim();
}

function toIsoOrBlank(value) {
  if (!value) return "";
  const d = new Date(value);
  if (Number.isNaN(d.getTime())) return "";
  return d.toISOString();
}

module.exports = async function (context, req) {
  try {
    const table = getTableClient("DashboardSubmissions");
    const items = [];

    for await (const entity of table.listEntities()) {
      items.push({
        partitionKey: entity.partitionKey || "",
        rowKey: entity.rowKey || "",
        entityId: entity.entityId || entity.partitionKey || "",
        weekEnding: normalizeString(entity.weekEnding),
        market: normalizeString(entity.market),
        location: normalizeString(entity.location),
        submittedBy: normalizeString(entity.submittedBy),
        submittedByEmail: normalizeString(entity.submittedByEmail),
        submittedAt: toIsoOrBlank(entity.submittedAt || entity.createdAt || entity.timestamp),
        updatedAt: toIsoOrBlank(entity.updatedAt || entity.timestamp),
        status: normalizeString(entity.status),
        payload: entity.payload || null
      });
    }

    items.sort((a, b) => {
      const aTime = a.submittedAt ? new Date(a.submittedAt).getTime() : 0;
      const bTime = b.submittedAt ? new Date(b.submittedAt).getTime() : 0;
      return bTime - aTime;
    });

    context.res = {
      status: 200,
      headers: {
        "Content-Type": "application/json"
      },
      body: {
        ok: true,
        count: items.length,
        items
      }
    };
  } catch (err) {
    context.log.error("api/submissionsfeed failed", err);

    context.res = {
      status: 500,
      headers: {
        "Content-Type": "application/json"
      },
      body: {
        ok: false,
        error: "submissionsfeed_failed",
        details: err.message
      }
    };
  }
};
