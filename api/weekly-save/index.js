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

    const body = req.body || {};
    const weekEnding = body.weekEnding;
    const entity = body.entity;
    const data = body.data || {};

    if (!weekEnding || !entity || !data) {
      return badRequest("Missing required fields");
    }

    if (!canAccessEntity(access, entity)) {
      return forbidden("You cannot edit this entity");
    }

    const client = await ensureTable(WEEKLY_TABLE);

    let existing = null;
    try {
      existing = await client.getEntity(entity, weekEnding);
    } catch (error) {
      if (error?.statusCode !== 404) {
        throw error;
      }
    }

    if (
      existing &&
      (existing.status === "submitted" || existing.status === "approved") &&
      !access.isAdmin
    ) {
      return forbidden("This week is locked and cannot be edited");
    }

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
      status: existing?.status || "draft",
      updatedBy: access.email,
      updatedAt: new Date().toISOString()
    };

    if (existing?.submittedBy) entityRecord.submittedBy = existing.submittedBy;
    if (existing?.submittedAt) entityRecord.submittedAt = existing.submittedAt;
    if (existing?.approvedBy) entityRecord.approvedBy = existing.approvedBy;
    if (existing?.approvedAt) entityRecord.approvedAt = existing.approvedAt;

    if (
      access.isAdmin &&
      existing &&
      (existing.status === "submitted" || existing.status === "approved")
    ) {
      entityRecord.overrideBy = access.email;
      entityRecord.overrideAt = new Date().toISOString();
    }

    await client.upsertEntity(entityRecord, "Replace");

    return ok({
      ok: true,
      message: access.isAdmin && existing && (existing.status === "submitted" || existing.status === "approved")
        ? "Override saved successfully"
        : "Saved successfully",
      entity,
      weekEnding,
      data: sanitizedData,
      status: entityRecord.status
    });
  } catch (error) {
    context.log.error("weekly-save failed", error);
    return serverError(error, "Failed to save weekly data");
  }
};
