const { getUserFromRequest, getUserEmail } = require("../shared/auth");
const { getPermissionByEmail } = require("../shared/permissions");
const { getTableClient } = require("../shared/table");
const { ENTITIES } = require("../shared/workbookLogic");

module.exports = async function (context, req) {
  try {
    const user = getUserFromRequest(req);
    if (!user) {
      context.res = {
        status: 401,
        body: { error: "Not authenticated" }
      };
      return;
    }

    const email = getUserEmail(user);
    const permission = await getPermissionByEmail(email);

    if (!permission || !permission.isAdmin) {
      context.res = {
        status: 403,
        body: { error: "Admin access required" }
      };
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

      const latestInputRow = [...inputRows]
        .filter((r) => r.updatedAt)
        .sort((a, b) => String(b.updatedAt).localeCompare(String(a.updatedAt)))[0] || null;

      const latestNarrativeAt = narrativeRow?.updatedAt || "";
      const latestInputAt = latestInputRow?.updatedAt || "";

      const useNarrative = latestNarrativeAt && String(latestNarrativeAt) > String(latestInputAt);

      rows.push({
        entity,
        weekEnding,
        status: statusRow?.status || "Draft",
        submittedBy: statusRow?.submittedBy || "",
        submittedAt: statusRow?.submittedAt || "",
        approvedBy: statusRow?.approvedBy || "",
        approvedAt: statusRow?.approvedAt || "",
        updatedBy: useNarrative
          ? (narrativeRow?.updatedBy || "")
          : (latestInputRow?.updatedBy || narrativeRow?.updatedBy || ""),
        updatedAt: useNarrative
          ? latestNarrativeAt
          : (latestInputAt || latestNarrativeAt || ""),
        inputCount: inputRows.length,
        hasNarrative: Boolean(
          narrativeRow &&
          (
            narrativeRow.commentary ||
            narrativeRow.blockers ||
            narrativeRow.opportunities ||
            narrativeRow.executiveNotes
          )
        )
      });
    }

    const submittedCount = rows.filter(
      (r) => r.status === "Submitted" || r.status === "Approved"
    ).length;

    const missingEntities = rows
      .filter((r) => r.status !== "Submitted" && r.status !== "Approved")
      .map((r) => r.entity);

    context.res = {
      status: 200,
      body: {
        weekEnding,
        rows,
        summary: {
          totalEntities: ENTITIES.length,
          submittedCount,
          missingCount: missingEntities.length,
          missingEntities
        }
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