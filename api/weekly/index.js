const { getUserFromRequest } = require("../shared/auth");
const {
  resolveAccess,
  requireAccess,
  requireEntityAccess,
  safeErrorResponse
} = require("../shared/permissions");
const { getTableClient } = require("../shared/table");

const TABLE_NAME = "WeeklyRegionData";

function parseJsonSafely(value, fallback = {}) {
  if (!value) return fallback;

  try {
    return typeof value === "string" ? JSON.parse(value) : value;
  } catch {
    return fallback;
  }
}

function toNumber(value, fallback = 0) {
  if (value == null || value === "") return fallback;
  const n = Number(value);
  return Number.isFinite(n) ? n : fallback;
}

function toText(value, fallback = "") {
  if (value == null) return fallback;
  return String(value).trim();
}

function normalizeWeeklyValues(values = {}, record = null) {
  return {
    newPatients: toNumber(values.newPatients ?? values.npActual ?? record?.newPatients, 0),
    surgeries: toNumber(values.surgeries ?? values.surgeryActual ?? record?.surgeries, 0),
    established: toNumber(values.established ?? values.establishedActual ?? record?.established, 0),
    noShows: toNumber(values.noShows ?? record?.noShows, 0),
    cancelled: toNumber(values.cancelled ?? record?.cancelled, 0),
    totalCalls: toNumber(values.totalCalls ?? values.callVolume ?? record?.totalCalls ?? record?.callVolume, 0),
    abandonedCalls: toNumber(values.abandonedCalls ?? record?.abandonedCalls, 0),
    visitVolume: toNumber(values.visitVolume ?? values.totalVisits ?? record?.visitVolume, 0),
    callVolume: toNumber(values.callVolume ?? values.totalCalls ?? record?.callVolume ?? record?.totalCalls, 0),
    noShowRate: toNumber(values.noShowRate ?? record?.noShowRate, 0),
    cancellationRate: toNumber(values.cancellationRate ?? record?.cancellationRate, 0),
    abandonedCallRate: toNumber(
      values.abandonedCallRate ?? values.abandonmentRate ?? record?.abandonedCallRate,
      0
    ),

    ptScheduledVisits: toNumber(values.ptScheduledVisits ?? record?.ptScheduledVisits, 0),
    ptCancellations: toNumber(values.ptCancellations ?? record?.ptCancellations, 0),
    ptNoShows: toNumber(values.ptNoShows ?? record?.ptNoShows, 0),
    ptReschedules: toNumber(values.ptReschedules ?? record?.ptReschedules, 0),
    ptTotalUnitsBilled: toNumber(values.ptTotalUnitsBilled ?? record?.ptTotalUnitsBilled, 0),
    ptVisitsSeen: toNumber(values.ptVisitsSeen ?? record?.ptVisitsSeen, 0),
    ptWorkingDays: toNumber(values.ptWorkingDays ?? record?.ptWorkingDays, 5),
    ptUnitsPerVisit: toNumber(values.ptUnitsPerVisit ?? record?.ptUnitsPerVisit, 0),
    ptVisitsPerDay: toNumber(values.ptVisitsPerDay ?? record?.ptVisitsPerDay, 0),

    ptoDays: toNumber(values.ptoDays ?? record?.ptoDays, 0),
    cashCollected: toNumber(values.cashCollected ?? values.cashActual ?? record?.cashCollected ?? record?.cashActual, 0),
    piNp: toNumber(values.piNp ?? record?.piNp, 0),
    piCashCollection: toNumber(values.piCashCollection ?? record?.piCashCollection, 0),
    operationsNarrative: toText(values.operationsNarrative ?? record?.operationsNarrative, "")
  };
}

module.exports = async function (context, req) {
  try {
    const user = getUserFromRequest(req);
    const access = resolveAccess(user);

    const authError = requireAccess(access);
    if (authError) return authError;

    const entity = String(req.query.entity || "").trim();
    const weekEnding = String(req.query.weekEnding || "").trim();

    if (!entity || !weekEnding) {
      return {
        status: 400,
        body: {
          ok: false,
          error: "Missing entity or weekEnding"
        }
      };
    }

    const entityError = requireEntityAccess(access, entity);
    if (entityError) return entityError;

    const table = getTableClient(TABLE_NAME);

    let record = null;
    try {
      record = await table.getEntity(entity, weekEnding);
    } catch (err) {
      if (err.statusCode !== 404) {
        throw err;
      }
    }

    const rawValues = record ? parseJsonSafely(record.valuesJson, {}) : {};
    const values = normalizeWeeklyValues(rawValues, record);

    return {
      status: 200,
      body: {
        ok: true,
        found: !!record,
        entity,
        weekEnding,
        values,
        raw: record || null,
        source: record?.source || null,
        status: record?.status || null,
        importedAt: record?.importedAt || null,
        createdAt: record?.createdAt || null,
        createdBy: record?.createdBy || null,
        updatedAt: record?.updatedAt || null,
        updatedBy: record?.updatedBy || null
      }
    };
  } catch (error) {
    return safeErrorResponse(context, error, "Failed to load weekly region data");
  }
};