const { getUserFromRequest } = require("../shared/auth");
const { resolveAccess } = require("../shared/permissions");
const { getTableClient } = require("../shared/table");
const { writeAuditEvent, parseJsonSafe } = require("../shared/audit");

const TABLE_NAME = "WeeklyRegionData";

function toNumber(value, fallback = 0) {
  if (value == null || value === "") return fallback;
  const n = Number(value);
  return Number.isFinite(n) ? n : fallback;
}

function calculateDerived(values = {}) {
  const newPatients = toNumber(values.newPatients, 0);
  const surgeries = toNumber(values.surgeries, 0);
  const established = toNumber(values.established, 0);
  const noShows = toNumber(values.noShows, 0);
  const cancelled = toNumber(values.cancelled, 0);
  const totalCalls = toNumber(values.totalCalls ?? values.callVolume, 0);
  const abandonedCalls = toNumber(values.abandonedCalls, 0);

  const visitVolume = newPatients + surgeries + established;
  const callVolume = totalCalls;
  const scheduledAppointments = visitVolume + noShows + cancelled;

  const noShowRate =
    scheduledAppointments > 0 ? (noShows / scheduledAppointments) * 100 : 0;

  const cancellationRate =
    scheduledAppointments > 0 ? (cancelled / scheduledAppointments) * 100 : 0;

  const abandonedCallRate =
    totalCalls > 0 ? (abandonedCalls / totalCalls) * 100 : 0;

  const ptScheduledVisits = toNumber(values.ptScheduledVisits, 0);
  const ptCancellations = toNumber(values.ptCancellations, 0);
  const ptNoShows = toNumber(values.ptNoShows, 0);
  const ptReschedules = toNumber(values.ptReschedules, 0);
  const ptTotalUnitsBilled = toNumber(values.ptTotalUnitsBilled, 0);
  const ptVisitsSeen = toNumber(values.ptVisitsSeen, 0);
  const ptWorkingDays = Math.max(1, toNumber(values.ptWorkingDays, 5));

  const ptUnitsPerVisit =
    ptVisitsSeen > 0 ? ptTotalUnitsBilled / ptVisitsSeen : 0;

  const ptVisitsPerDay =
    ptWorkingDays > 0 ? ptVisitsSeen / ptWorkingDays : 0;

  return {
    newPatients,
    surgeries,
    established,
    noShows,
    cancelled,
    totalCalls,
    abandonedCalls,
    visitVolume,
    callVolume,
    noShowRate: Number(noShowRate.toFixed(2)),
    cancellationRate: Number(cancellationRate.toFixed(2)),
    abandonedCallRate: Number(abandonedCallRate.toFixed(2)),

    ptScheduledVisits,
    ptCancellations,
    ptNoShows,
    ptReschedules,
    ptTotalUnitsBilled,
    ptVisitsSeen,
    ptWorkingDays,
    ptUnitsPerVisit: Number(ptUnitsPerVisit.toFixed(2)),
    ptVisitsPerDay: Number(ptVisitsPerDay.toFixed(2))
  };
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

    if (!access.allowed) {
      return {
        status: 403,
        body: {
          ok: false,
          error: "Forbidden"
        }
      };
    }

    const entity = String(req.body?.entity || "").trim();
    const weekEnding = String(req.body?.weekEnding || "").trim();
    const input =
      req.body?.data && typeof req.body.data === "object"
        ? req.body.data
        : req.body?.values && typeof req.body.values === "object"
          ? req.body.values
          : {};

    if (!entity || !weekEnding) {
      return {
        status: 400,
        body: {
          ok: false,
          error: "Missing entity or weekEnding"
        }
      };
    }

    const table = getTableClient(TABLE_NAME);

    let existing = null;
    try {
      existing = await table.getEntity(entity, weekEnding);
    } catch (err) {
      if (err.statusCode !== 404) {
        throw err;
      }
    }

    const values = calculateDerived(input);
    const now = new Date().toISOString();
    const actorEmail = access.email || user?.userDetails || null;
    const actorRole = access.role || "user";

    const existingValues = existing ? parseJsonSafe(existing.valuesJson, {}) : {};
    const beforeAudit = existing
      ? buildAuditShape(entity, weekEnding, existingValues, existing)
      : null;

    const createdBy = existing?.createdBy || actorEmail;
    const createdAt = existing?.createdAt || now;

    const nextRecord = {
      partitionKey: entity,
      rowKey: weekEnding,
      entity,
      weekEnding,
      status: "saved",
      valuesJson: JSON.stringify(values),

      visitVolume: values.visitVolume,
      callVolume: values.callVolume,
      newPatients: values.newPatients,
      surgeries: values.surgeries,
      established: values.established,
      noShows: values.noShows,
      cancelled: values.cancelled,
      totalCalls: values.totalCalls,
      abandonedCalls: values.abandonedCalls,
      noShowRate: values.noShowRate,
      cancellationRate: values.cancellationRate,
      abandonedCallRate: values.abandonedCallRate,

      ptScheduledVisits: values.ptScheduledVisits,
      ptCancellations: values.ptCancellations,
      ptNoShows: values.ptNoShows,
      ptReschedules: values.ptReschedules,
      ptTotalUnitsBilled: values.ptTotalUnitsBilled,
      ptVisitsSeen: values.ptVisitsSeen,
      ptWorkingDays: values.ptWorkingDays,
      ptUnitsPerVisit: values.ptUnitsPerVisit,
      ptVisitsPerDay: values.ptVisitsPerDay,

      source: "app",
      createdBy,
      createdAt,
      updatedBy: actorEmail,
      updatedAt: now
    };

    await table.upsertEntity(nextRecord);

    const afterAudit = buildAuditShape(entity, weekEnding, values, nextRecord);

    await writeAuditEvent({
      eventType: existing ? "update" : "create",
      entity,
      weekEnding,
      actorEmail,
      actorRole,
      before: beforeAudit || {},
      after: afterAudit,
      metadata: {
        source: "weekly-save"
      }
    });

    return {
      status: 200,
      body: {
        ok: true,
        message: "Saved successfully",
        status: "saved",
        entity,
        weekEnding,
        values,
        createdBy,
        createdAt,
        updatedBy: actorEmail,
        updatedAt: now
      }
    };
  } catch (error) {
    context.log.error("weekly-save failed", error);

    return {
      status: 500,
      body: {
        ok: false,
        error: "Failed to save weekly region data.",
        details: error.message
      }
    };
  }
};