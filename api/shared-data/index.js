const { requireAuthorizedUser } = require("../shared");
const { getTableClient } = require("../shared/table");
const { buildSharedKpis } = require("../shared/sharedPageLogic");

module.exports = async function (context, req) {
  try {
    const auth = await requireAuthorizedUser(req);
    if (!auth.ok) {
      context.res = { status: auth.status, body: auth.body };
      return;
    }

    const page = req.query.page;
    const weekEnding = req.query.weekEnding || new Date().toISOString().slice(0, 10);

    if (!page) {
      context.res = { status: 400, body: { error: "Missing page" } };
      return;
    }

    const partitionKey = `${page}|${weekEnding}`;
    const inputsTable = getTableClient("WeeklyInputs");

    const rows = await inputsTable.listByPartition(partitionKey);
    const inputs = {};

    for (const row of rows) {
      if (row.metricKey) inputs[row.metricKey] = row.value;
    }

    context.res = {
      status: 200,
      body: {
        page,
        weekEnding,
        status: "Draft",
        inputs,
        kpis: buildSharedKpis(page, inputs)
      }
    };
  } catch (err) {
    context.log.error(err);
    context.res = { status: 500, body: { error: err.message } };
  }
};
