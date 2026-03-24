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

    const weekEnding = req.query && req.query.weekEnding ? req.query.weekEnding : null;
    let entity = req.query && req.query.entity ? req.query.entity : null;

    if (!weekEnding) {
      return badRequest("Missing weekEnding");
    }

    if (!entity) {
      entity = access.isAdmin ? "LAOSS" : access.entity;
    }

    if (!canAccessEntity(access, entity)) {
      return forbidden("You cannot view this entity");
    }

    const client = await ensureTable(WEEKLY_TABLE);

    let record = null;

    try {
      record = await client.getEntity(entity, weekEnding);
    } catch (error) {
      if (error?.statusCode !== 404) {
        throw error;
      }
    }

    if (!record) {
      const emptyData = {};
      for (const field of KPI_FIELDS) {
        emptyData[field.key] = null;
      }

      return ok({
        ok: true,
        found: false,
        entity,
        weekEnding,
        data: emptyData,
        status: "draft"
      });
    }

    const data = {};
    for (const field of KPI_FIELDS) {
      data[field.key] = record[field.key] ?? null;
    }

    return ok({
      ok: true,
      found: true,
      entity,
      weekEnding,
      data,
      status: record.status || "draft",
      updatedBy: record.updatedBy || null,
      updatedAt: record.updatedAt || null
    });
  } catch (error) {
    context.log.error("weekly failed", error);
    return serverError(error, "Failed to load weekly data");
  }
};
