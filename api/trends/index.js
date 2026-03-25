const { getUserFromRequest } = require("../shared/auth");
const { resolveAccess, canAccessEntity } = require("../shared/permissions");
const { ok, badRequest, forbidden, serverError } = require("../shared/response");
const { ensureTable } = require("../shared/table");
const { WEEKLY_TABLE } = require("../shared/constants");

function isValidDateString(value) {
  return typeof value === "string" && /^\d{4}-\d{2}-\d{2}$/.test(value);
}

module.exports = async function (context, req) {
  try {
    const user = getUserFromRequest(req);
    const access = resolveAccess(user);

    if (!access.allowed) {
      return forbidden();
    }

    const entity = req.query && req.query.entity ? req.query.entity : null;
    const limitRaw = req.query && req.query.limit ? Number(req.query.limit) : 12;
    const startDate = req.query && req.query.startDate ? req.query.startDate : null;
    const endDate = req.query && req.query.endDate ? req.query.endDate : null;

    if (!entity) {
      return badRequest("Missing entity");
    }

    if (!canAccessEntity(access, entity)) {
      return forbidden("You cannot view this entity");
    }

    if (startDate && !isValidDateString(startDate)) {
      return badRequest("Invalid startDate");
    }

    if (endDate && !isValidDateString(endDate)) {
      return badRequest("Invalid endDate");
    }

    const limit = Number.isFinite(limitRaw) && limitRaw > 0 ? Math.min(limitRaw, 52) : 12;

    const client = await ensureTable(WEEKLY_TABLE);

    const rows = [];
    const filter = `PartitionKey eq '${entity}'`;

    for await (const record of client.listEntities({
      queryOptions: { filter }
    })) {
      rows.push({
        entity,
        weekEnding: record.rowKey,
        status: record.status || "draft",
        visitVolume: Number(record.visitVolume || 0),
        callVolume: Number(record.callVolume || 0),
        newPatients: Number(record.newPatients || 0),
        noShowRate: Number(record.noShowRate || 0),
        cancellationRate: Number(record.cancellationRate || 0),
        abandonedCallRate: Number(record.abandonedCallRate || 0),
        updatedBy: record.updatedBy || null,
        updatedAt: record.updatedAt || null
      });
    }

    rows.sort((a, b) => a.weekEnding.localeCompare(b.weekEnding));

    let filtered = rows;

    if (startDate || endDate) {
      filtered = rows.filter((row) => {
        if (startDate && row.weekEnding < startDate) return false;
        if (endDate && row.weekEnding > endDate) return false;
        return true;
      });
    } else {
      filtered = rows.slice(-limit);
    }

    return ok({
      ok: true,
      entity,
      count: filtered.length,
      appliedFilter: startDate || endDate
        ? { mode: "dateRange", startDate: startDate || null, endDate: endDate || null }
        : { mode: "recentWeeks", limit },
      items: filtered
    });
  } catch (error) {
    context.log.error("trends failed", error);
    return serverError(error, "Failed to load trends");
  }
};
