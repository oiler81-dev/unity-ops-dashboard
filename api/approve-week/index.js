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
      return badRequest("Missing fields");
    }

    const client = await ensureTable(WEEKLY_TABLE);

    let record;
    try {
      record = await client.getEntity(entity, weekEnding);
    } catch (error) {
      return badRequest("No data found");
    }

    const now = new Date().toISOString();

    record.status = "approved";
    record.approvedBy = access.email;
    record.approvedAt = now;
    record.updatedBy = access.email;
    record.updatedAt = now;

    await client.upsertEntity(record, "Replace");

    return ok({
      ok: true,
      message: "Approved successfully",
      entity,
      weekEnding,
      status: "approved"
    });
  } catch (error) {
    context.log.error("approve-week failed", error);
    return serverError(error, "Failed to approve week");
  }
};