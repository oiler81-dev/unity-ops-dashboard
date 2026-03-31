const crypto = require("crypto");
const { getUserFromRequest } = require("../shared/auth");
const { resolveAccess } = require("../shared/permissions");
const { ok, badRequest, forbidden, serverError } = require("../shared/response");
const { getTableClient, ensureTable } = require("../shared/table");
const { WEEKLY_TABLE } = require("../shared/constants");

const AUDIT_TABLE_NAME = "WeeklyAuditLog";

function safeText(value) {
  return value == null ? "" : String(value).trim();
}

function toNumber(value, fallback = 0) {
  if (value == null || value === "") return fallback;
  const n = Number(value);
  return Number.isFinite(n) ? n : fallback;
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
    const auditTable = getTableClient(AUDIT_TABLE_NAME);

    let record;
    try {
      record = await client.getEntity(entity, weekEnding);
    } catch (error) {
      if (error?.statusCode === 404) {
        return badRequest("No record found to delete");
      }
      throw error;
    }

    const before = summarizeExistingRecord(record);

    await client.deleteEntity(entity, weekEnding);

    await writeAuditLog(auditTable, {
      eventType: "delete",
      entity,
      weekEnding,
      actorEmail: access.email || user?.userDetails || "",
      actorRole: access.role || "admin",
      isAdminAction: true,
      summary: `Deleted ${entity} ${weekEnding}`,
      before,
      after: null
    });

    return ok({
      ok: true,
      message: "Week deleted successfully",
      entity,
      weekEnding,
      audit: {
        eventType: "delete",
        actorEmail: access.email || user?.userDetails || null
      }
    });
  } catch (error) {
    context.log.error("delete-week failed", error);
    return serverError(error, "Failed to delete week");
  }
};