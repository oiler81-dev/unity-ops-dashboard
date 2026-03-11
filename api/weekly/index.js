const { getUserFromRequest, getUserEmail } = require("../shared/auth");
const { getPermissionByEmail } = require("../shared/permissions");
const { getTableClient } = require("../shared/table");
const { buildRegionKpis } = require("../shared/workbookLogic");

module.exports = async function (context, req) {
  try {
    const user = getUserFromRequest(req);
    if (!user) {
      context.res = { status: 401, body: { error: "Not authenticated" } };
      return;
    }

    const email = getUserEmail(user);
    const permission = await getPermissionByEmail(email);
    if (!permission) {
      context.res = { status: 403, body: { error: "Not authorized" } };
      return;
    }

    const entity = req.query.entity;
    const weekEnding = req.query.weekEnding || new Date().toISOString().slice(0, 10);

    if (!entity) {
      context.res = { status: 400, body: { error: "Missing entity" } };
      return;
    }

    const partitionKey = `${entity}|${weekEnding}`;
    const inputsTable = getTableClient("WeeklyInputs");
    const narrativeTable = getTableClient("WeeklyNarratives");
    const statusTable = getTableClient("SubmissionStatus");

    const rows = await inputsTable.listByPartition(partitionKey);

    const inputs = {};
    for (const row of rows) {
      inputs[row.metricKey] = row.value;
    }

    const narrativeRow = await narrativeTable.getEntity(partitionKey, "NARRATIVE");
    const statusRow = await statusTable.getEntity(partitionKey, "STATUS");

    context.res = {
      status: 200,
      body: {
        entity,
        weekEnding,
        status: statusRow?.status || "Draft",
        inputs,
        narrative: narrativeRow
          ? {
              commentary: narrativeRow.commentary || "",
              blockers: narrativeRow.blockers || "",
              opportunities: narrativeRow.opportunities || "",
              executiveNotes: narrativeRow.executiveNotes || ""
            }
          : {},
        kpis: buildRegionKpis(inputs)
      }
    };
  } catch (err) {
    context.log.error(err);
    context.res = {
      status: 500,
      body: { error: err.message }
    };
  }
};
