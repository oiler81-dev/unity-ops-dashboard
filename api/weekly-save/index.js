const { getUserFromRequest } = require("../shared/auth");
const { resolveAccess, canAccessEntity } = require("../shared/permissions");
const { ok, badRequest, forbidden, serverError } = require("../shared/response");
const { ensureTable } = require("../shared/table");
const { WEEKLY_TABLE, KPI_FIELDS } = require("../shared/constants");

module.exports = async function (context, req) {
  try {
    const user = getUserFromRequest(req);
    const access = resolveAccess(user);

    if (!access.allowed) {
      return forbidden();
    }

    const body = req.body;

    if (!body || !body.weekEnding || !body.entity || !body.data) {
      return badRequest("Missing required fields");
    }

    const { weekEnding, entity, data } = body;

    if (!canAccessEntity(access, entity)) {
      return forbidden("You cannot edit this entity");
    }

    const client = await ensureTable(WEEKLY_TABLE);

    const sanitizedData = {};
    for (const field of KPI_FIELDS) {
      const raw = data[field.key];
      sanitizedData[field.key] =
        raw === null || raw === undefined || raw === ""
          ? null
          : Number(raw);
    }

    const entityRecord = {
      partitionKey: entity,
      rowKey: weekEnding,
      ...sanitizedData,
      status: "draft",
      updatedBy: access.email,
      updatedAt: new Date().toISOString()
    };

    await client.upsertEntity(entityRecord, "Replace");

    return ok({
      ok: true,
      message: "Saved successfully",
      entity,
      weekEnding,
      data: sanitizedData
    });
  } catch (error) {
    context.log.error("weekly-save failed", error);
    return serverError(error, "Failed to save weekly data");
  }
};
