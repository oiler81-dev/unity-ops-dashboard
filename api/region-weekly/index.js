const { getTableClient } = require("../shared/table");

const TABLE_NAME = "WeeklyRegionData";

function safeJsonParse(value, fallback = {}) {
  if (!value) return fallback;
  if (typeof value === "object") return value;

  try {
    return JSON.parse(value);
  } catch {
    return fallback;
  }
}

function normalizeDateString(value) {
  return String(value || "").trim().slice(0, 10);
}

function normalizeEntity(value) {
  return String(value || "").trim();
}

function buildResponse(entity, weekEnding, record) {
  return {
    entity,
    weekEnding,
    values: record?.valuesJson ? safeJsonParse(record.valuesJson, {}) : {},
    source: record?.source || null,
    importedAt: record?.importedAt || null,
    updatedAt: record?.updatedAt || null,
    found: !!record
  };
}

module.exports = async function (context, req) {
  try {
    const entity = normalizeEntity(req.query.entity);
    const weekEnding = normalizeDateString(req.query.weekEnding);

    if (!entity || !weekEnding) {
      context.res = {
        status: 400,
        headers: { "Content-Type": "application/json" },
        body: { error: "Missing entity or weekEnding." }
      };
      return;
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

    if (!record) {
      for await (const item of table.listEntities()) {
        const itemEntity =
          normalizeEntity(item.entity) ||
          normalizeEntity(item.partitionKey);

        const itemWeekEnding =
          normalizeDateString(item.weekEnding) ||
          normalizeDateString(item.rowKey);

        if (itemEntity === entity && itemWeekEnding === weekEnding) {
          record = item;
          break;
        }
      }
    }

    context.res = {
      status: 200,
      headers: { "Content-Type": "application/json" },
      body: buildResponse(entity, weekEnding, record)
    };
  } catch (error) {
    context.log.error("region-weekly GET failed", error);

    context.res = {
      status: 500,
      headers: { "Content-Type": "application/json" },
      body: {
        error: "Failed to load weekly region data.",
        details: error && error.message ? error.message : String(error)
      }
    };
  }
};
