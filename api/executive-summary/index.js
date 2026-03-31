const { getUserFromRequest } = require("../shared/auth");
const { resolveAccess } = require("../shared/permissions");
const { getTableClient } = require("../shared/table");
const { monthLabelToMonthKey, getWorkingDaysForMonth } = require("../shared/budget");

const REGION_TABLE = "WeeklyRegionData";
const SHARED_TABLE = "SharedPageData";
const BUDGET_TABLE = "BudgetData";
const ENTITIES = ["LAOSS", "NES", "SpineOne", "MRO"];

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

    const ptRecord = await getSharedRecord(sharedTable, "PT", weekEnding);
    const cxnsRecord = await getSharedRecord(sharedTable, "CXNS", weekEnding);

    const ptValues = safeParseJson(ptRecord?.valuesJson, {});
    const cxnsValues = safeParseJson(cxnsRecord?.valuesJson, {});
    const ptByEntity = ptValues?.byEntity || {};

    const regions = [];

    for (const entity of ENTITIES) {
      const record = await getEntityRecord(regionTable, entity, weekEnding);
      const values = safeParseJson(record?.valuesJson, {});

      const daysInPeriod = toNumber(values.daysInPeriod ?? record?.daysInPeriod, 0);
      const monthKey = buildMonthKey(values, weekEnding);

      const budgetRecord = monthKey
        ? await getBudgetRecord(budgetTable, entity, monthKey)
        : null;

      const workingDaysInMonth =
        toNumber(budgetRecord?.workingDaysInMonth, 0) || getWorkingDaysForMonth(monthKey);

      const visitBudgetMonthly = toNumber(budgetRecord?.visitBudgetMonthly, 0);
      const newPatientsBudgetMonthly = toNumber(budgetRecord?.newPatientsBudgetMonthly, 0);

      const ptEntity = ptByEntity?.[entity] || {};
      const ptVisitsSeen =
        toNumber(ptEntity.scheduledVisits, 0) -
        toNumber(ptEntity.cancellations, 0) -
        toNumber(ptEntity.noShows, 0);

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

        noShowRate: toNumber(
          record?.noShowRate ??
          values.noShowRate,
          0
        ),
        cancellationRate: toNumber(
          record?.cancellationRate ??
          values.cancellationRate,
          0
        ),
        abandonedCallRate: toNumber(
          record?.abandonedCallRate ??
          values.abandonedCallRate ??
          values.abandonmentRate,
          0
        ),

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
          monthKey,
          monthLabel: budgetRecord?.monthLabel || ""
        },

        pt: {
          scheduledVisits: toNumber(ptEntity.scheduledVisits, 0),
          cancellations: toNumber(ptEntity.cancellations, 0),
          noShows: toNumber(ptEntity.noShows, 0),
          reschedules: toNumber(ptEntity.reschedules, 0),
          totalUnitsBilled: toNumber(ptEntity.totalUnitsBilled, 0),
          visitsSeen: ptVisitsSeen > 0 ? ptVisitsSeen : 0,
          unitsPerVisit:
            ptVisitsSeen > 0
              ? Number((toNumber(ptEntity.totalUnitsBilled, 0) / ptVisitsSeen).toFixed(2))
              : 0,
          visitsPerDay:
            daysInPeriod > 0 && ptVisitsSeen > 0
              ? Number((ptVisitsSeen / daysInPeriod).toFixed(2))
              : 0
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
      piCashCollection: regions.reduce((sum, r) => sum + toNumber(r.piCashCollection, 0), 0)
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

    const ptUnitsPerVisit =
      ptTotals.visitsSeen > 0
        ? Number((ptTotals.totalUnitsBilled / ptTotals.visitsSeen).toFixed(2))
        : 0;

    const ptVisitsPerDay =
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
          ptUnitsPerVisit,
          ptVisitsPerDay
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