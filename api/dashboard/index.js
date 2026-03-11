const { getUserFromRequest, getUserEmail } = require("../shared/auth");
const { getPermissionByEmail } = require("../shared/permissions");
const { getTableClient } = require("../shared/table");
const { ENTITIES, buildExecutiveFromRows } = require("../shared/workbookLogic");
const { getReferenceMapForEntity } = require("../shared/reference");

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

    const weekEnding = req.query.weekEnding || new Date().toISOString().slice(0, 10);

    const inputsTable = getTableClient("WeeklyInputs");
    const statusTable = getTableClient("SubmissionStatus");

    const entityRows = [];
    const entityReferenceMaps = {};

    for (const entity of ENTITIES) {
      const partitionKey = `${entity}|${weekEnding}`;
      const [rows, statusRow, referenceMap] = await Promise.all([
        inputsTable.listByPartition(partitionKey),
        statusTable.getEntity(partitionKey, "STATUS"),
        getReferenceMapForEntity(entity)
      ]);

      const raw = {};
      for (const row of rows) {
        if (row.metricKey) raw[row.metricKey] = row.value;
      }

      entityRows.push({
        entity,
        raw,
        status: statusRow?.status || "Draft"
      });

      entityReferenceMaps[entity] = referenceMap;
    }

    context.res = {
      status: 200,
      body: buildExecutiveFromRows(weekEnding, entityRows, entityReferenceMaps)
    };
  } catch (err) {
    context.log.error(err);
    context.res = {
      status: 500,
      body: { error: err.message }
    };
  }
};
