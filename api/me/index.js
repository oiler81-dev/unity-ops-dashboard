const { getUserFromRequest } = require("../shared/auth");
const { resolveAccess } = require("../shared/permissions");
const { ok } = require("../shared/response");

module.exports = async function (context, req) {
  const user = getUserFromRequest(req);
  const access = resolveAccess(user);
  return ok({ user, access });
};