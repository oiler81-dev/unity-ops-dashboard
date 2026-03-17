const { getTableClient } = require("../shared/table");

const TABLE_NAME = "WeeklyRegionData";

module.exports = async function (context, req) {
  try {
    const entity = String(req.query.entity || "").trim();
    const weekEnding = String(req.query.weekEnding || "").trim();

    if (!entity || !weekEnding) {
      context.res = {
        status: 400,
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

    context.res = {
      status: 200,
      headers: { "Content-Type": "application/json" },
      body: {
        entity,
        weekEnding,
        values: record?.valuesJson ? JSON.parse(record.valuesJson) : {},
        source: record?.source || "app",
        importedAt: record?.importedAt || null,
        updatedAt: record?.updatedAt || null
      }
    };
  } catch (error) {
    context.log.error("weekly GET failed", error);
    context.res = {
      status: 500,
      body: {
        error: "Failed to load weekly region data.",
        details: error.message
      }
    };
  }
};
