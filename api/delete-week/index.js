const { getUserFromRequest } = require("../shared/auth");
const { resolveAccess } = require("../shared/permissions");
const { ok, badRequest, forbidden, serverError } = require("../shared/response");
const { ensureTable } = require("../shared/table");
const { WEEKLY_TABLE } = require("../shared/constants");
const { writeAuditEvent, parseJsonSafe } = require("../shared/audit");

function toNumber(value, fallback = 0) {
  if (value == null || value === "") return fallback;
  const n = Number(value);
  return Number.isFinite(n) ? n : fallback;
}

function buildAuditShape(entity, weekEnding, values, record = {}) {
  return {
    entity,
    weekEnding,

    newPatients: toNumber(values.newPatients ?? record.newPatients, 0),
    surgeries: toNumber(values.surgeries ?? record.surgeries, 0),
    established: toNumber(values.established ?? record.established, 0),
    noShows: toNumber(values.noShows ?? record.noShows, 0),
    cancelled: toNumber(values.cancelled ?? record.cancelled, 0),
    totalCalls: toNumber(values.totalCalls ?? record.totalCalls ?? record.callVolume, 0),
    abandonedCalls: toNumber(values.abandonedCalls ?? record.abandonedCalls, 0),
    visitVolume: toNumber(values.visitVolume ?? record.visitVolume, 0),
    callVolume: toNumber(values.callVolume ?? record.callVolume, 0),
    noShowRate: toNumber(values.noShowRate ?? record.noShowRate, 0),
    cancellationRate: toNumber(values.cancellationRate ?? record.cancellationRate, 0),
    abandonedCallRate: toNumber(values.abandonedCallRate ?? record.abandonedCallRate, 0),

    ptScheduledVisits: toNumber(values.ptScheduledVisits ?? record.ptScheduledVisits, 0),
    ptCancellations: toNumber(values.ptCancellations ?? record.ptCancellations, 0),
    ptNoShows: toNumber(values.ptNoShows ?? record.ptNoShows, 0),
    ptReschedules: toNumber(values.ptReschedules ?? record.ptReschedules, 0),
    ptTotalUnitsBilled: toNumber(values.ptTotalUnitsBilled ?? record.ptTotalUnitsBilled, 0),
    ptVisitsSeen: toNumber(values.ptVisitsSeen ?? record.ptVisitsSeen, 0),
    ptWorkingDays: toNumber(values.ptWorkingDays ?? record.ptWorkingDays, 5),
    ptUnitsPerVisit: toNumber(values.ptUnitsPerVisit ?? record.ptUnitsPerVisit, 0),
    ptVisitsPerDay: toNumber(values.ptVisitsPerDay ?? record.ptVisitsPerDay, 0)
  };
}

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

    let record;
    try {
      record = await client.getEntity(entity, weekEnding);
    } catch (error) {
      if (error?.statusCode === 404) {
        return badRequest("No record found to delete");
      }
      throw error;
    }

    const rawValues = parseJsonSafe(record.valuesJson, {});
    const beforeAudit = buildAuditShape(entity, weekEnding, rawValues, record);

    await client.deleteEntity(entity, weekEnding);

    await writeAuditEvent({
      eventType: "delete",
      entity,
      weekEnding,
      actorEmail: access.email,
      actorRole: access.role || "admin",
      before: beforeAudit,
      after: {},
      metadata: {
        source: "delete-week"
      }
    });

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