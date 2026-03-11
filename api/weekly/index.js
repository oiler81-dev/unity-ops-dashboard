const crypto = require("crypto");
const { getUserFromRequest, getUserEmail, getDisplayName } = require("../shared/auth");
const { getPermissionByEmail, canEditEntity } = require("../shared/permissions");
const { getTableClient } = require("../shared/table");

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

    const body = req.body || {};
    const entity = body.entity;
    const weekEnding = body.weekEnding;
    const inputs = body.inputs || {};
    const narrative = body.narrative || {};

    if (!entity || !weekEnding) {
      context.res = { status: 400, body: { error: "entity and weekEnding are required" } };
      return;
    }

    if (!canEditEntity(permission, entity)) {
      context.res = { status: 403, body: { error: "You do not have edit access for this entity" } };
      return;
    }

    const partitionKey = `${entity}|${weekEnding}`;
    const inputsTable = getTableClient("WeeklyInputs");
    const narrativeTable = getTableClient("WeeklyNarratives");
    const auditTable = getTableClient("AuditLog");
    const statusTable = getTableClient("SubmissionStatus");

    const nowIso = new Date().toISOString();
    const displayName = permission.displayName || getDisplayName(user);

    for (const [metricKey, newValue] of Object.entries(inputs)) {
      const rowKey = `INPUT|${metricKey}`;
      const existing = await inputsTable.getEntity(partitionKey, rowKey);
      const oldValue = existing?.value ?? null;

      await inputsTable.upsertEntity({
        partitionKey,
        rowKey,
        entity,
        weekEnding,
        section: "DynamicRegionInput",
        metricKey,
        label: metricKey,
        value: newValue,
        valueType: typeof newValue,
        updatedBy: displayName,
        updatedAt: nowIso
      });

      if (String(oldValue ?? "") !== String(newValue ?? "")) {
        await auditTable.createEntity({
          partitionKey,
          rowKey: `${Date.now()}-${crypto.randomBytes(4).toString("hex")}`,
          entity,
          weekEnding,
          section: "DynamicRegionInput",
          metricKey,
          oldValue: oldValue ?? "",
          newValue: newValue ?? "",
          changedBy: displayName,
          changedAt: nowIso,
          changeType: "input_update"
        });
      }
    }

    await narrativeTable.upsertEntity({
      partitionKey,
      rowKey: "NARRATIVE",
      entity,
      weekEnding,
      commentary: narrative.commentary || "",
      blockers: narrative.blockers || "",
      opportunities: narrative.opportunities || "",
      executiveNotes: narrative.executiveNotes || "",
      updatedBy: displayName,
      updatedAt: nowIso
    });

    const existingStatus = await statusTable.getEntity(partitionKey, "STATUS");
    await statusTable.upsertEntity({
      partitionKey,
      rowKey: "STATUS",
      entity,
      weekEnding,
      status: existingStatus?.status || "Draft",
      submittedBy: existingStatus?.submittedBy || "",
      submittedAt: existingStatus?.submittedAt || "",
      approvedBy: existingStatus?.approvedBy || "",
      approvedAt: existingStatus?.approvedAt || "",
      updatedBy: displayName,
      updatedAt: nowIso
    });

    context.res = {
      status: 200,
      body: {
        ok: true,
        status: existingStatus?.status || "Draft",
        message: "Weekly data saved successfully"
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
