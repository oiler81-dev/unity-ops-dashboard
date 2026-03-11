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
    if (!permission?.isAdmin) {
      context.res = { status: 403, body: { error: "Admin access required" } };
      return;
    }

    const table = getTableClient("UserPermissions");
    const rows = await table.listByPartition("USER");

    const users = rows.map((r) => ({
      displayName: r.displayName,
      email: r.email,
      role: r.role,
      entity: r.entity,
      canEdit: Boolean(r.canEdit),
      isAdmin: Boolean(r.isAdmin),
      isActive: r.isActive !== false
    }));

    context.res = {
      status: 200,
      body: { users }
    };
  } catch (err) {
    context.log.error(err);
    context.res = { status: 500, body: { error: err.message } };
  }
};
