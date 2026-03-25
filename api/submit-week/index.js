const { getUserFromRequest } = require("../shared/auth");
const { resolveAccess } = require("../shared/permissions");
const { ok, badRequest, forbidden, serverError } = require("../shared/response");
const { ensureTable } = require("../shared/table");
const { WEEKLY_TABLE } = require("../shared/constants");

module.exports = async function (context, req) {
  try {
    const user = getUserFromRequest(req);
    const access = resolveAccess(user);

    if (!access.isAdmin) return forbidden("Admin only");

    const { weekEnding, entity } = req.body || {};

    if (!weekEnding || !entity) {
      return badRequest("Missing fields");
    }

    const client = await ensureTable(WEEKLY_TABLE);

    let record;
    try {
      record = await client.getEntity(entity, weekEnding);
    } catch {
      return badRequest("No data found");
    }

    record.status = "approved";
    record.approvedBy = access.email;
    record.approvedAt = new Date().toISOString();

    await client.upsertEntity(record, "Replace");

    return ok({
      ok: true,
      message: "Approved successfully"
    });

  } catch (error) {
    context.log.error(error);
    return serverError(error);
  }
};
