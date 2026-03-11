const { getUserFromRequest, getUserEmail } = require("../shared/auth");
const { getPermissionByEmail } = require("../shared/permissions");
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

    const entity = req.query.entity;
    const weekEnding = req.query.weekEnding;

    if (!entity || !weekEnding) {
      context.res = { status: 400, body: { error: "entity and weekEnding are required" } };
      return;
    }

    const partitionKey = `${entity}|${weekEnding}`;
    const auditTable = getTableClient("AuditLog");
    const rows = await auditTable.listByPartition(partitionKey);

    rows.sort((a, b) => String(b.changedAt || "").localeCompare(String(a.changedAt || "")));

    context.res = {
      status: 200,
      body: { entity, weekEnding, rows }
    };
  } catch (err) {
    context.log.error(err);
    context.res = { status: 500, body: { error: err.message } };
  }
};
