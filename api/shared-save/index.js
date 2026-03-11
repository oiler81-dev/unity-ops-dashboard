const crypto = require("crypto");
const { requireAuthorizedUser } = require("../shared");
const { getTableClient } = require("../shared/table");
const { getDisplayName } = require("../shared/auth");

module.exports = async function (context, req) {
  try {
    const auth = await requireAuthorizedUser(req);
    if (!auth.ok) {
      context.res = { status: auth.status, body: auth.body };
      return;
    }

    const { permission, user } = auth;
    if (!permission.isAdmin) {
      context.res = { status: 403, body: { error: "Only admins can edit shared pages right now" } };
      return;
    }

    const body = req.body || {};
    const page = body.page;
    const weekEnding = body.weekEnding;
    const inputs = body.inputs || {};

    if (!page || !weekEnding) {
      context.res = { status: 400, body: { error: "page and weekEnding are required" } };
      return;
    }

    const partitionKey = `${page}|${weekEnding}`;
    const inputsTable = getTableClient("WeeklyInputs");
    const auditTable = getTableClient("AuditLog");

    const nowIso = new Date().toISOString();
    const displayName = permission.displayName || getDisplayName(user);

    for (const [metricKey, newValue] of Object.entries(inputs)) {
      const rowKey = `SHARED|${metricKey}`;
      const existing = await inputsTable.getEntity(partitionKey, rowKey);
      const oldValue = existing?.value ?? null;

      await inputsTable.upsertEntity({
        partitionKey,
        rowKey,
        entity: page,
        weekEnding,
        section: "SharedPageInput",
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
          entity: page,
          weekEnding,
          section: "SharedPageInput",
          metricKey,
          oldValue: oldValue ?? "",
          newValue: newValue ?? "",
          changedBy: displayName,
          changedAt: nowIso,
          changeType: "shared_input_update"
        });
      }
    }

    context.res = {
      status: 200,
      body: {
        ok: true,
        status: "Draft",
        message: "Shared page saved successfully"
      }
    };
  } catch (err) {
    context.log.error(err);
    context.res = { status: 500, body: { error: err.message } };
  }
};
