const { getTableClient } = require("../shared/table");

const TABLE_NAME = "SharedPageData";

module.exports = async function (context, req) {
  try {
    const body = req.body || {};
    const page = String(body.page || "").trim();
    const weekEnding = String(body.weekEnding || "").trim();
    const values = body.values && typeof body.values === "object" ? body.values : {};

    if (!page || !weekEnding) {
      context.res = {
        status: 400,
        body: { error: "Missing page or weekEnding." }
      };
      return;
    }

    const table = getTableClient(TABLE_NAME);

    await table.upsertEntity({
      partitionKey: page,
      rowKey: weekEnding,
      page,
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
        page,
        weekEnding
      }
    };
  } catch (error) {
    context.log.error("shared-save failed", error);
    context.res = {
      status: 500,
      body: {
        error: "Failed to save shared page data.",
        details: error.message
      }
    };
  }
};
