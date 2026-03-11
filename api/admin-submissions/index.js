const { getUserFromRequest, getUserEmail } = require("../shared/auth");
const { getPermissionByEmail } = require("../shared/permissions");
const { getTableClient } = require("../shared/table");
const { ENTITIES } = require("../shared/workbookLogic");

module.exports = async function (context, req) {
  try {
    const user = getUserFromRequest(req);
    if (!user) {
      context.res = { status: 401, body: { error: "Not authenticated" } };
      return;
    }

    const email = getUserEmail(user);
    const permission = await getPermissionByEmail(email);
    if (!permission?.isAdmin) {
      context.res = { status: 403, body: { error: "Admin access required" } };
      return;
    }

    const weekEnding = req.query.weekEnding || new Date().toISOString().slice(0, 10);

    const statusTable = getTableClient("SubmissionStatus");
    const inputsTable = getTableClient("WeeklyInputs");
    const narrativeTable = getTableClient("WeeklyNarratives");

    const rows = [];

    for (const entity of ENTITIES) {
      const partitionKey = `${entity}|${weekEnding}`;

      const [statusRow, inputRows, narrativeRow] = await Promise.all([
        statusTable.getEntity(partitionKey, "STATUS"),
        inputsTable.listByPartition(partitionKey),
        narrativeTable.getEntity(partitionKey, "NARRATIVE")
      ]);

      const lastInputUpdate = [...inputRows]
        .map((r) => r.updatedAt)
        .filter(Boolean)
        .sort()
        .pop() || null;

      const lastInputUser = [...inputRows]
        .filter((r) => r.updatedAt === lastInputUpdate)
        .map((r) => r.updatedBy)
        .filter(Boolean)
        .pop() || null;

      const lastNarrativeUpdate = narrativeRow?.updatedAt || null;
      const lastNarrativeUser = narrativeRow?.updatedBy || null;

      const latestUpdatedAt = [lastInputUpdate, lastNarrativeUpdate].filter(Boolean).sort().pop() || null;

      const latestUpdatedBy =
        latestUpdatedAt === lastNarrativeUpdate
          ? lastNarrativeUser
          : lastInputUser;

      rows.push({
        entity,
        weekEnding,
        status: statusRow?.status || "Draft",
        submittedBy: statusRow?.submittedBy || "",
        submittedAt: statusRow?.submittedAt || "",
        approvedBy: statusRow?.approvedBy || "",
        approvedAt: statusRow?.approvedAt || "",
        updatedBy: latestUpdatedBy || "",
        updatedAt: latestUpdatedAt || "",
        inputCount: inputRows.length,
        hasNarrative: Boolean(
          narrativeRow &&
          (narrativeRow.commentary || narrativeRow.blockers || narrativeRow.opportunities)
        )
      });
    }

    const missing = rows.filter((r) => r.status !== "Submitted" && r.status !== "Approved");

    context.res = {
      status: 200,
      body: {
        weekEnding,
        rows,
        summary: {
          totalEntities: ENTITIES.length,
          submittedCount: rows.filter((r) => r.status === "Submitted" || r.status === "Approved").length,
          missingCount: missing.length,
          missingEntities: missing.map((r) => r.entity)
        }
      }
    };
  } catch (err) {
    context.log.error(err);
    context.res = { status: 500, body: { error: err.message } };
  }
};