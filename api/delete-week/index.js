const { getUserFromRequest } = require("../shared/auth");
const { resolveAccess } = require("../shared/permissions");
const { ok, badRequest, forbidden, serverError } = require("../shared/response");
const { ensureTable } = require("../shared/table");
const { WEEKLY_TABLE } = require("../shared/constants");

module.exports = async function (context, req) {
  try {
    const user = getUserFromRequest(req);
    const access = resolveAccess(user);

    if (!access.isAdmin) {
      return forbidden("Admin only");
    }

    const body = req.body || {};
    const weekEnding = body.weekEnding;
    const entity = body.entity;

    if (!weekEnding || !entity) {
      return badRequest("Missing entity or weekEnding");
    }

    const client = await ensureTable(WEEKLY_TABLE);

    try {
      await client.deleteEntity(entity, weekEnding);
    } catch (error) {
      if (error?.statusCode === 404) {
        return badRequest("No record found to delete");
      }
      throw error;
    }

    return ok({
      ok: true,
      message: "Week deleted successfully",
      entity,
      weekEnding
    });
  } catch (error) {
    context.log.error("delete-week failed", error);
    return serverError(error, "Failed to delete week");
  }
};
