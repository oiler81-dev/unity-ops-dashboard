const { getUserFromRequest, getUserEmail } = require("./auth");
const { getPermissionByEmail } = require("./permissions");

async function requireAuthorizedUser(req) {
  const user = getUserFromRequest(req);
  if (!user) {
    return { ok: false, status: 401, body: { error: "Not authenticated" } };
  }

  const email = getUserEmail(user);
  const permission = await getPermissionByEmail(email);

  if (!permission) {
    return { ok: false, status: 403, body: { error: "Not authorized" } };
  }

  return { ok: true, user, permission };
}

module.exports = { requireAuthorizedUser };
