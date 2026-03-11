const { getUserFromRequest, getUserEmail, getDisplayName } = require("../shared/auth");
const { getPermissionByEmail } = require("../shared/permissions");

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

    if (!permission) {
      context.res = {
        status: 403,
        body: { error: "No dashboard permission found for this user" }
      };
      return;
    }

    context.res = {
      status: 200,
      body: {
        displayName: permission.displayName || getDisplayName(user),
        email: permission.email || email,
        entity: permission.entity,
        role: permission.role,
        isAdmin: permission.isAdmin,
        canEdit: permission.canEdit,
        isActive: permission.isActive
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
