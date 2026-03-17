const { getTableClient } = require("../shared/table");

const TABLE_NAME = "WeeklyRegionData";

module.exports = async function (context, req) {
  try {
    const body = req.body || {};
    const entity = String(body.entity || "").trim();
    const weekEnding = String(body.weekEnding || "").trim();
    const values = body.values && typeof body.values === "object" ? body.values : {};

    if (!entity || !weekEnding) {
      context.res = {
        status: 400,
        body: { error: "Missing entity or weekEnding." }
      };
      return;
    }

    const table = getTableClient(TABLE_NAME);

    await table.upsertEntity({
      partitionKey: entity,
      rowKey: weekEnding,
      entity,
      weekEnding,
      valuesJson: JSON.stringify(values),
      source: "app",
      updatedAt: new Date().toISOString()
    });

    context.res = {
      status: 200,
      headers: { "Content-Type": "application/json" },
      body: {
        ok: true,
        status: "Draft",
        entity,
        weekEnding
      }
    };
  } catch (error) {
    context.log.error("weekly-save failed", error);
    context.res = {
      status: 500,
      body: {
        error: "Failed to save weekly region data.",
        details: error.message
      }
    };
  }
};
