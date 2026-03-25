const { getUserFromRequest } = require("../shared/auth");
const { resolveAccess, canAccessEntity } = require("../shared/permissions");
const { ok, badRequest, forbidden, serverError } = require("../shared/response");
const { ensureTable } = require("../shared/table");
const { WEEKLY_TABLE, KPI_FIELDS } = require("../shared/constants");

function normalizeNumber(value) {
  if (value === null || value === undefined || value === "") return null;
  const n = Number(value);
  return Number.isFinite(n) ? n : null;
}

function legacyToValues(data) {
  return {
    totalVisits: normalizeNumber(data.visitVolume),
    totalCalls: normalizeNumber(data.callVolume),
    npActual: normalizeNumber(data.newPatients),
    noShowRate: normalizeNumber(data.noShowRate),
    cancellationRate: normalizeNumber(data.cancellationRate),
    abandonmentRate: normalizeNumber(data.abandonedCallRate)
  };
}

function parseValuesJson(valuesJson) {
  if (!valuesJson) return {};
  try {
    return typeof valuesJson === "string" ? JSON.parse(valuesJson) : valuesJson;
  } catch {
    return {};
  }
}

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

    const existingStatus = String(existing?.status || "").toLowerCase();

    if (
      existing &&
      (existingStatus === "submitted" || existingStatus === "approved") &&
      !access.isAdmin
    ) {
      return forbidden("This week is locked and cannot be edited");
    }

    const sanitizedData = {};
    for (const field of KPI_FIELDS) {
      sanitizedData[field.key] = normalizeNumber(data[field.key]);
    }

    const existingValues = parseValuesJson(existing?.valuesJson);
    const newValues = {
      ...existingValues,
      ...legacyToValues(sanitizedData)
    };

    const now = new Date().toISOString();

    const entityRecord = {
      partitionKey: entity,
      rowKey: weekEnding,

      // keep flat fields for backward compatibility
      ...sanitizedData,

      // canonical rebuilt shape
      entity,
      weekEnding,
      valuesJson: JSON.stringify(newValues),

      status: existing?.status || "draft",
      source: existing?.source || "manual-entry",
      updatedBy: access.email,
      updatedAt: now
    };

    if (existing?.submittedBy) entityRecord.submittedBy = existing.submittedBy;
    if (existing?.submittedAt) entityRecord.submittedAt = existing.submittedAt;
    if (existing?.approvedBy) entityRecord.approvedBy = existing.approvedBy;
    if (existing?.approvedAt) entityRecord.approvedAt = existing.approvedAt;
    if (existing?.importedAt) entityRecord.importedAt = existing.importedAt;
    if (existing?.source && existing.source !== "manual-entry") {
      entityRecord.source = existing.source;
    }

    if (
      access.isAdmin &&
      existing &&
      (existingStatus === "submitted" || existingStatus === "approved")
    ) {
      entityRecord.overrideBy = access.email;
      entityRecord.overrideAt = now;
    }

    await client.upsertEntity(entityRecord, "Replace");

    return ok({
      ok: true,
      message:
        access.isAdmin &&
        existing &&
        (existingStatus === "submitted" || existingStatus === "approved")
          ? "Override saved successfully"
          : "Saved successfully",
      entity,
      weekEnding,
      data: sanitizedData,
      values: newValues,
      status: entityRecord.status
    });
  } catch (error) {
    context.log.error("weekly-save failed", error);
    return serverError(error, "Failed to save weekly data");
  }
};
