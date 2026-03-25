const { getUserFromRequest } = require("../shared/auth");
const { resolveAccess, canAccessEntity } = require("../shared/permissions");
const { ok, badRequest, forbidden, serverError } = require("../shared/response");
const { ensureTable } = require("../shared/table");
const { WEEKLY_TABLE } = require("../shared/constants");

module.exports = async function (context, req) {
  try {
    const user = getUserFromRequest(req);
    const access = resolveAccess(user);

    if (!access.allowed) {
      return forbidden();
    }

    const body = req.body || {};
    const weekEnding = body.weekEnding;
    const entity = body.entity;

    if (!weekEnding || !entity) {
      return badRequest("Missing fields");
    }

    if (!canAccessEntity(access, entity)) {
      return forbidden("Cannot submit for this entity");
    }

    const client = await ensureTable(WEEKLY_TABLE);

    let record;
    try {
      record = await client.getEntity(entity, weekEnding);
    } catch (error) {
      return badRequest("No data found to submit");
    }

    record.status = "submitted";
    record.submittedBy = access.email;
    record.submittedAt = new Date().toISOString();

    await client.upsertEntity(record, "Replace");

    return ok({
      ok: true,
      message: "Submitted successfully",
      entity,
      weekEnding,
      status: "submitted"
    });
  } catch (error) {
    context.log.error("submit-week failed", error);
    return serverError(error, "Failed to submit week");
  }
};
