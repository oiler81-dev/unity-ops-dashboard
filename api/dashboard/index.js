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
    const previousWeekEnding = (() => {
      const d = new Date(`${weekEnding}T12:00:00`);
      d.setDate(d.getDate() - 7);
      return d.toISOString().slice(0, 10);
    })();

    const inputsTable = getTableClient("WeeklyInputs");
    const statusTable = getTableClient("SubmissionStatus");
    const narrativeTable = getTableClient("WeeklyNarratives");

    const entityRows = [];
    const previousEntityRows = [];
    const entityReferenceMaps = {};
    const narratives = {};

    for (const entity of ENTITIES) {
      const currentPartitionKey = `${entity}|${weekEnding}`;
      const previousPartitionKey = `${entity}|${previousWeekEnding}`;

      const [currentRows, previousRows, statusRow, referenceMap, narrativeRow] = await Promise.all([
        inputsTable.listByPartition(currentPartitionKey),
        inputsTable.listByPartition(previousPartitionKey),
        statusTable.getEntity(currentPartitionKey, "STATUS"),
        getReferenceMapForEntity(entity),
        narrativeTable.getEntity(currentPartitionKey, "NARRATIVE")
      ]);

      const raw = {};
      for (const row of currentRows) {
        if (row.metricKey) raw[row.metricKey] = row.value;
      }

      const previousRaw = {};
      for (const row of previousRows) {
        if (row.metricKey) previousRaw[row.metricKey] = row.value;
      }

      entityRows.push({
        entity,
        raw,
        status: statusRow?.status || "Draft"
      });

      previousEntityRows.push({
        entity,
        raw: previousRaw,
        status: "Historical"
      });

      entityReferenceMaps[entity] = referenceMap;
      narratives[entity] = {
        commentary: narrativeRow?.commentary || "",
        blockers: narrativeRow?.blockers || "",
        opportunities: narrativeRow?.opportunities || ""
      };
    }

    context.res = {
      status: 200,
      body: buildExecutiveFromRows(weekEnding, entityRows, entityReferenceMaps, previousEntityRows, narratives)
    };
  } catch (err) {
    context.log.error(err);
    context.res = {
      status: 500,
      body: { error: err.message }
    };
  }
};
