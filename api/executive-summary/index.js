const { getUserFromRequest } = require("../shared/auth");
const { resolveAccess } = require("../shared/permissions");
const { getTableClient } = require("../shared/table");
const { monthLabelToMonthKey, getWorkingDaysForMonth } = require("../shared/budget");

const REGION_TABLE = "WeeklyRegionData";
const SHARED_TABLE = "SharedPageData";
const BUDGET_TABLE = "BudgetData";
const ENTITIES = ["LAOSS", "NES", "SpineOne", "MRO"];

// Default working days per week per entity (used when daysInPeriod not stored on record)
const ENTITY_WORKING_DAYS = {
  LAOSS:     5,
  NES:       4.5,
  SpineOne:  5,
  MRO:       5
};

function toNumber(value, fallback = 0) {
  if (value == null || value === "") return fallback;
  const n = Number(value);
  return Number.isFinite(n) ? n : fallback;
}

function safeParseJson(value, fallback = {}) {
  if (!value) return fallback;
  try {
    return typeof value === "string" ? JSON.parse(value) : value;
  } catch {
    return fallback;
  }
}

async function getEntityRecord(table, entity, weekEnding) {
  try {
    return await table.getEntity(entity, weekEnding);
  } catch (error) {
    if (error?.statusCode === 404) return null;
    throw error;
  }
}

async function getSharedRecord(table, page, weekEnding) {
  try {
    return await table.getEntity(page, weekEnding);
  } catch (error) {
    if (error?.statusCode === 404) return null;
    throw error;
  }
}

async function getBudgetRecord(table, entity, monthKey) {
  try {
    return await table.getEntity(entity, monthKey);
  } catch (error) {
    if (error?.statusCode === 404) return null;
    throw error;
  }
}

function buildMonthKey(values, weekEnding) {
  const monthTag = String(values?.monthTag || "").trim();
  if (monthTag) {
    const mk = monthLabelToMonthKey(monthTag);
    if (mk) return mk;
  }

  if (!weekEnding) return "";
  const d = new Date(`${weekEnding}T12:00:00Z`);
  const yyyy = d.getUTCFullYear();
  const mm = String(d.getUTCMonth() + 1).padStart(2, "0");
  return `${yyyy}-${mm}`;
}

function prorateMonthlyToWeek(monthly, workingDaysInMonth, daysInPeriod) {
  const monthBudget = toNumber(monthly, 0);
  const monthDays = Math.max(1, toNumber(workingDaysInMonth, 0));
  const weekDays = Math.max(0, toNumber(daysInPeriod, 0));
  if (!monthBudget || !weekDays) return 0;
  return Number(((monthBudget / monthDays) * weekDays).toFixed(2));
}

module.exports = async function (context, req) {
  try {
    const user = getUserFromRequest(req);
    const access = resolveAccess(user);

    if (!access?.authenticated) {
      return {
        status: 401,
        body: {
          ok: false,
          error: "Unauthorized"
        }
      };
    }

    const weekEnding = String(req.query.weekEnding || "").trim();
    if (!weekEnding) {
      return {
        status: 400,
        body: {
          ok: false,
          error: "Missing weekEnding"
        }
      };
    }

    const regionTable = getTableClient(REGION_TABLE);
    const sharedTable = getTableClient(SHARED_TABLE);
    const budgetTable = getTableClient(BUDGET_TABLE);

    const cxnsRecord = await getSharedRecord(sharedTable, "CXNS", weekEnding);
    const cxnsValues = safeParseJson(cxnsRecord?.valuesJson, {});

    const regions = [];

    // Scope to the caller's entity — admins see all four, regionals only their own.
    const scopeEntities = access.isAdmin
      ? ENTITIES
      : ENTITIES.filter((e) => e === access.entity);

    for (const entity of scopeEntities) {
      const record = await getEntityRecord(regionTable, entity, weekEnding);
      const values = safeParseJson(record?.valuesJson, {});

      // PT data lives on the base entity partition — never on a separate "{entity}-PT" partition.
      // getSelectedEntity() on the frontend always returns the base entity, so PT saves go to the
      // same row as ortho saves. Overlaying from a "{entity}-PT" partition here clobbered the
      // correct PT values with stale legacy data.

      // daysInPeriod is set by Excel import; manual entries won't have it — use entity-specific default
      const defaultDays = ENTITY_WORKING_DAYS[entity] ?? 5;
      const daysInPeriod = toNumber(values.daysInPeriod ?? record?.daysInPeriod, 0) || defaultDays;
      const monthKey = buildMonthKey(values, weekEnding);

      const budgetRecord = monthKey
        ? await getBudgetRecord(budgetTable, entity, monthKey)
        : null;

      const workingDaysInMonth =
        toNumber(budgetRecord?.workingDaysInMonth, 0) || getWorkingDaysForMonth(monthKey);

      const visitBudgetMonthly = toNumber(budgetRecord?.visitBudgetMonthly, 0);
      const newPatientsBudgetMonthly = toNumber(budgetRecord?.newPatientsBudgetMonthly, 0);

      // PT data is now stored directly on the WeeklyRegionData record
      const ptScheduledVisits = toNumber(record?.ptScheduledVisits ?? values.ptScheduledVisits, 0);
      const ptCancellations = toNumber(record?.ptCancellations ?? values.ptCancellations, 0);
      const ptNoShows = toNumber(record?.ptNoShows ?? values.ptNoShows, 0);
      const ptReschedules = toNumber(record?.ptReschedules ?? values.ptReschedules, 0);
      const ptTotalUnitsBilled = toNumber(record?.ptTotalUnitsBilled ?? values.ptTotalUnitsBilled, 0);
      const ptVisitsSeen = toNumber(record?.ptVisitsSeen ?? values.ptVisitsSeen, 0);
      const ptWorkingDays = toNumber(record?.ptWorkingDays ?? values.ptWorkingDays, 5);

      const ptUnitsPerVisit =
        ptVisitsSeen > 0
          ? Number((ptTotalUnitsBilled / ptVisitsSeen).toFixed(2))
          : 0;

      const ptVisitsPerDay =
        ptWorkingDays > 0 && ptVisitsSeen > 0
          ? Number((ptVisitsSeen / ptWorkingDays).toFixed(2))
          : 0;

      const region = {
        entity,
        weekEnding,
        status: record?.status || "saved",

        visitVolume: toNumber(values.totalVisits ?? values.visitVolume ?? record?.visitVolume, 0),
        callVolume: toNumber(values.totalCalls ?? values.callVolume ?? record?.callVolume, 0),
        newPatients: toNumber(values.npActual ?? values.newPatients ?? record?.newPatients, 0),
        surgeries: toNumber(values.surgeryActual ?? values.surgeries ?? record?.surgeries, 0),
        established: toNumber(values.establishedActual ?? values.established ?? record?.established, 0),
        noShows: toNumber(values.noShows ?? record?.noShows, 0),
        cancelled: toNumber(values.cancelled ?? record?.cancelled, 0),
        totalCalls: toNumber(values.totalCalls ?? values.callVolume ?? record?.totalCalls ?? record?.callVolume, 0),
        abandonedCalls: toNumber(values.abandonedCalls ?? record?.abandonedCalls, 0),

        noShowRate: (() => {
          const stored = toNumber(record?.noShowRate ?? values.noShowRate, 0);
          if (stored > 0) return stored;
          const noShowsRaw = toNumber(values.noShows ?? record?.noShows, 0);
          const cancelledRaw = toNumber(values.cancelled ?? record?.cancelled, 0);
          const visitVol = toNumber(values.totalVisits ?? values.visitVolume ?? record?.visitVolume, 0);
          const scheduled = visitVol + noShowsRaw + cancelledRaw;
          return scheduled > 0 ? Number(((noShowsRaw / scheduled) * 100).toFixed(2)) : 0;
        })(),
        cancellationRate: (() => {
          const stored = toNumber(record?.cancellationRate ?? values.cancellationRate, 0);
          if (stored > 0) return stored;
          const noShowsRaw = toNumber(values.noShows ?? record?.noShows, 0);
          const cancelledRaw = toNumber(values.cancelled ?? record?.cancelled, 0);
          const visitVol = toNumber(values.totalVisits ?? values.visitVolume ?? record?.visitVolume, 0);
          const scheduled = visitVol + noShowsRaw + cancelledRaw;
          return scheduled > 0 ? Number(((cancelledRaw / scheduled) * 100).toFixed(2)) : 0;
        })(),
        abandonedCallRate: (() => {
          const stored = toNumber(record?.abandonedCallRate ?? values.abandonedCallRate ?? values.abandonmentRate, 0);
          if (stored > 0) return stored;
          const abandonedRaw = toNumber(values.abandonedCalls ?? record?.abandonedCalls, 0);
          const callsRaw = toNumber(values.totalCalls ?? values.callVolume ?? record?.totalCalls ?? record?.callVolume, 0);
          return callsRaw > 0 ? Number(((abandonedRaw / callsRaw) * 100).toFixed(2)) : 0;
        })(),

        cashCollected: toNumber(
          values.cashCollected ??
          values.cashActual ??
          record?.cashCollected ??
          record?.cashActual,
          0
        ),
        ptoDays: toNumber(values.ptoDays ?? record?.ptoDays, 0),
        operationsNarrative: String(values.operationsNarrative ?? record?.operationsNarrative ?? "").trim(),

        piNp: entity === "SpineOne"
          ? toNumber(values.piNp ?? record?.piNp, 0)
          : 0,
        piCashCollection: entity === "SpineOne"
          ? toNumber(values.piCashCollection ?? record?.piCashCollection, 0)
          : 0,
        imaging: entity === "SpineOne"
          ? toNumber(values.imaging ?? record?.imaging, 0)
          : 0,

        budget: {
          visitVolumeBudget: prorateMonthlyToWeek(
            visitBudgetMonthly,
            workingDaysInMonth,
            daysInPeriod
          ),
          newPatientsBudget: prorateMonthlyToWeek(
            newPatientsBudgetMonthly,
            workingDaysInMonth,
            daysInPeriod
          ),
          workingDaysInMonth,
          daysInPeriod,
          monthKey,
          monthLabel: budgetRecord?.monthLabel || "",
          hasBudgetRecord: !!budgetRecord,
          visitBudgetMonthly,
          newPatientsBudgetMonthly
        },

        pt: {
          scheduledVisits: ptScheduledVisits,
          cancellations: ptCancellations,
          noShows: ptNoShows,
          reschedules: ptReschedules,
          totalUnitsBilled: ptTotalUnitsBilled,
          visitsSeen: ptVisitsSeen,
          unitsPerVisit: ptUnitsPerVisit,
          visitsPerDay: ptVisitsPerDay
        }
      };

      regions.push(region);
    }

    const totals = {
      visitVolume: regions.reduce((sum, r) => sum + toNumber(r.visitVolume, 0), 0),
      callVolume: regions.reduce((sum, r) => sum + toNumber(r.callVolume, 0), 0),
      newPatients: regions.reduce((sum, r) => sum + toNumber(r.newPatients, 0), 0),
      surgeries: regions.reduce((sum, r) => sum + toNumber(r.surgeries, 0), 0),
      cashCollected: regions.reduce((sum, r) => sum + toNumber(r.cashCollected, 0), 0),
      ptoDays: regions.reduce((sum, r) => sum + toNumber(r.ptoDays, 0), 0),
      piNp: regions.reduce((sum, r) => sum + toNumber(r.piNp, 0), 0),
      piCashCollection: regions.reduce((sum, r) => sum + toNumber(r.piCashCollection, 0), 0),
      imaging: regions.reduce((sum, r) => sum + toNumber(r.imaging, 0), 0)
    };

    const budgetTotals = {
      visitVolumeBudget: regions.reduce((sum, r) => sum + toNumber(r.budget?.visitVolumeBudget, 0), 0),
      newPatientsBudget: regions.reduce((sum, r) => sum + toNumber(r.budget?.newPatientsBudget, 0), 0)
    };

    const ptTotals = {
      scheduledVisits: regions.reduce((sum, r) => sum + toNumber(r.pt?.scheduledVisits, 0), 0),
      cancellations: regions.reduce((sum, r) => sum + toNumber(r.pt?.cancellations, 0), 0),
      noShows: regions.reduce((sum, r) => sum + toNumber(r.pt?.noShows, 0), 0),
      reschedules: regions.reduce((sum, r) => sum + toNumber(r.pt?.reschedules, 0), 0),
      totalUnitsBilled: regions.reduce((sum, r) => sum + toNumber(r.pt?.totalUnitsBilled, 0), 0),
      visitsSeen: regions.reduce((sum, r) => sum + toNumber(r.pt?.visitsSeen, 0), 0)
    };

    const avg = (key) =>
      regions.length
        ? regions.reduce((sum, r) => sum + toNumber(r[key], 0), 0) / regions.length
        : 0;

    const ptUnitsPerVisitOverall =
      ptTotals.visitsSeen > 0
        ? Number((ptTotals.totalUnitsBilled / ptTotals.visitsSeen).toFixed(2))
        : 0;

    const ptVisitsPerDayOverall =
      regions.length
        ? Number(
            (
              regions.reduce((sum, r) => sum + toNumber(r.pt?.visitsPerDay, 0), 0) /
              regions.length
            ).toFixed(2)
          )
        : 0;

    return {
      status: 200,
      body: {
        ok: true,
        weekEnding,
        entityCount: regions.length,
        totals,
        budgetTotals,
        averages: {
          noShowRate: Number(avg("noShowRate").toFixed(2)),
          cancellationRate: Number(avg("cancellationRate").toFixed(2)),
          abandonedCallRate: Number(avg("abandonedCallRate").toFixed(2)),
          ptUnitsPerVisit: ptUnitsPerVisitOverall,
          ptVisitsPerDay: ptVisitsPerDayOverall
        },
        ptTotals,
        cxns: {
          scheduledAppts: toNumber(cxnsValues.scheduledAppts, 0),
          cancellations: toNumber(cxnsValues.cancellations, 0),
          noShows: toNumber(cxnsValues.noShows, 0),
          reschedules: toNumber(cxnsValues.reschedules, 0)
        },
        regions
      }
    };
  } catch (error) {
    context.log.error("executive-summary failed", error);
    return {
      status: 500,
      body: {
        ok: false,
        error: "Failed to build executive summary",
        details: error.message
      }
    };
  }
};
