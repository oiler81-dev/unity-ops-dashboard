const { getTableClient } = require("../shared/table");

const TABLE_NAME = "SharedPageData";

module.exports = async function (context, req) {
  try {
    const page = String(req.query.page || "").trim();
    const weekEnding = String(req.query.weekEnding || "").trim();

    if (!page || !weekEnding) {
      context.res = {
        status: 400,
        body: { error: "Missing page or weekEnding." }
      };
      return;
    }

    const table = getTableClient(TABLE_NAME);

    let record = null;
    try {
      record = await table.getEntity(page, weekEnding);
    } catch (err) {
      if (err.statusCode !== 404) {
        throw err;
      }
    }

    context.res = {
      status: 200,
      headers: { "Content-Type": "application/json" },
      body: {
        page,
        weekEnding,
        values: record?.valuesJson ? JSON.parse(record.valuesJson) : {},
        source: record?.source || "app",
        importedAt: record?.importedAt || null,
        updatedAt: record?.updatedAt || null
      }
    };
  } catch (error) {
    context.log.error("shared-data failed", error);
    context.res = {
      status: 500,
      body: {
        error: "Failed to load shared page data.",
        details: error.message
      }
    };
  }
};
