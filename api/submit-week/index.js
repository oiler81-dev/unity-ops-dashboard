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

    if (!entity || !weekEnding) {
      context.res = { status: 400, body: { error: "entity and weekEnding are required" } };
      return;
    }

    if (!canEditEntity(permission, entity)) {
      context.res = { status: 403, body: { error: "You do not have edit access for this entity" } };
      return;
    }

    const partitionKey = `${entity}|${weekEnding}`;
    const statusTable = getTableClient("SubmissionStatus");
    const auditTable = getTableClient("AuditLog");
    const nowIso = new Date().toISOString();
    const displayName = permission.displayName || getDisplayName(user);

    const existing = await statusTable.getEntity(partitionKey, "STATUS");
    const oldStatus = existing?.status || "Draft";

    await statusTable.upsertEntity({
      partitionKey,
      rowKey: "STATUS",
      entity,
      weekEnding,
      status: "Submitted",
      submittedBy: displayName,
      submittedAt: nowIso,
      approvedBy: existing?.approvedBy || "",
      approvedAt: existing?.approvedAt || "",
      updatedBy: displayName,
      updatedAt: nowIso
    });

    await auditTable.createEntity({
      partitionKey,
      rowKey: `${Date.now()}-${crypto.randomBytes(4).toString("hex")}`,
      entity,
      weekEnding,
      section: "Workflow",
      metricKey: "status",
      oldValue: oldStatus,
      newValue: "Submitted",
      changedBy: displayName,
      changedAt: nowIso,
      changeType: "status_change"
    });

    context.res = {
      status: 200,
      body: { ok: true, status: "Submitted" }
    };
  } catch (err) {
    context.log.error(err);
    context.res = { status: 500, body: { error: err.message } };
  }
};
