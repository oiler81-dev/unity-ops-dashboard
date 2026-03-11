const { getUserFromRequest, getUserEmail } = require("../shared/auth");
const { getPermissionByEmail } = require("../shared/permissions");
const { buildExecutiveSample } = require("../shared/workbookLogic");

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

    const weekEnding = req.query.weekEnding || new Date().toISOString().slice(0, 10);

    context.res = {
      status: 200,
      body: buildExecutiveSample(weekEnding)
    };
  } catch (err) {
    context.log.error(err);
    context.res = {
      status: 500,
      body: { error: err.message }
    };
  }
};
