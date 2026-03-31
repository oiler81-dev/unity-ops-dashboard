const crypto = require("crypto");
const { getUserFromRequest } = require("../shared/auth");
const { resolveAccess } = require("../shared/permissions");
const { getTableClient } = require("../shared/table");

const TABLE_NAME = "WeeklyRegionData";
const AUDIT_TABLE_NAME = "WeeklyAuditLog";

function toNumber(value, fallback = 0) {
  if (value == null || value === "") return fallback;
  const n = Number(value);
  return Number.isFinite(n) ? n : fallback;
}

function safeText(value) {
  return value == null ? "" : String(value).trim();
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

function summarizeExistingRecord(record) {
  if (!record) return null;

  return {
    entity: safeText(record.entity || record.partitionKey || record.PartitionKey),
    weekEnding: safeText(record.weekEnding || record.rowKey || record.RowKey),
    status: safeText(record.status),
    valuesJson: safeText(record.valuesJson),
    visitVolume: toNumber(record.visitVolume),
    callVolume: toNumber(record.callVolume),
    newPatients: toNumber(record.newPatients),
    surgeries: toNumber(record.surgeries),
    established: toNumber(record.established),
    noShows: toNumber(record.noShows),
    cancelled: toNumber(record.cancelled),
    totalCalls: toNumber(record.totalCalls),
    abandonedCalls: toNumber(record.abandonedCalls),
    noShowRate: toNumber(record.noShowRate),
    cancellationRate: toNumber(record.cancellationRate),
    abandonedCallRate: toNumber(record.abandonedCallRate),

    ptScheduledVisits: toNumber(record.ptScheduledVisits),
    ptCancellations: toNumber(record.ptCancellations),
    ptNoShows: toNumber(record.ptNoShows),
    ptReschedules: toNumber(record.ptReschedules),
    ptTotalUnitsBilled: toNumber(record.ptTotalUnitsBilled),
    ptVisitsSeen: toNumber(record.ptVisitsSeen),
    ptWorkingDays: toNumber(record.ptWorkingDays),
    ptUnitsPerVisit: toNumber(record.ptUnitsPerVisit),
    ptVisitsPerDay: toNumber(record.ptVisitsPerDay),

    createdAt: safeText(record.createdAt),
    createdBy: safeText(record.createdBy),
    updatedAt: safeText(record.updatedAt),
    updatedBy: safeText(record.updatedBy)
  };
}

async function writeAuditLog(auditTable, payload) {
  const now = new Date().toISOString();
  const rowKey = `${now}__${crypto.randomUUID()}`;

  await auditTable.upsertEntity({
    partitionKey: safeText(payload.entity || "unknown"),
    rowKey,
    eventType: safeText(payload.eventType),
    entity: safeText(payload.entity),
    weekEnding: safeText(payload.weekEnding),
    actorEmail: safeText(payload.actorEmail),
    actorRole: safeText(payload.actorRole),
    actionSource: safeText(payload.actionSource || "app"),
    isAdminAction: payload.isAdminAction ? true : false,
    timestamp: now,
    summary: safeText(payload.summary),
    beforeJson: JSON.stringify(payload.before ?? null),
    afterJson: JSON.stringify(payload.after ?? null)
  });
}

module.exports = async function (context, req) {
  try {
    const user = getUserFromRequest(req);
    const access = resolveAccess(user);

    if (!access.authenticated) {
      return {
        status: 401,
        body: {
          ok: false,
          error: "Authentication required"
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
    const auditTable = getTableClient(AUDIT_TABLE_NAME);

    let existing = null;
    try {
      existing = await table.getEntity(entity, weekEnding);
    } catch (err) {
      if (err.statusCode !== 404) {
        throw err;
      }
    }

    if (existing && !access.isAdmin) {
      return {
        status: 403,
        body: {
          ok: false,
          error: "This entry already exists. Only admins can edit existing records."
        }
      };
    }

    const now = new Date().toISOString();
    const values = calculateDerived(input);
    const before = summarizeExistingRecord(existing);

    const createdAt = existing?.createdAt || now;
    const createdBy = existing?.createdBy || access.email || user?.userDetails || null;

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
      createdAt,
      createdBy,
      updatedBy: access.email || user?.userDetails || null,
      updatedAt: now
    };

    await table.upsertEntity(nextRecord);

    const after = summarizeExistingRecord(nextRecord);

    await writeAuditLog(auditTable, {
      eventType: existing ? "update" : "create",
      entity,
      weekEnding,
      actorEmail: access.email || user?.userDetails || "",
      actorRole: access.role || "user",
      isAdminAction: !!access.isAdmin,
      summary: existing
        ? `Updated ${entity} ${weekEnding}`
        : `Created ${entity} ${weekEnding}`,
      before,
      after
    });

    return {
      status: 200,
      body: {
        ok: true,
        message: existing ? "Updated successfully" : "Saved successfully",
        status: "saved",
        entity,
        weekEnding,
        isNew: !existing,
        adminEdited: !!existing && !!access.isAdmin,
        values,
        audit: {
          eventType: existing ? "update" : "create",
          actorEmail: access.email || user?.userDetails || null,
          timestamp: now
        }
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