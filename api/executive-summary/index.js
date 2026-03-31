const { getUserFromRequest } = require("../shared/auth");
const { resolveAccess } = require("../shared/permissions");
const { getTableClient } = require("../shared/table");

const REGION_TABLE = "WeeklyRegionData";
const BUDGET_TABLE = "BudgetData";
const ENTITIES = ["LAOSS", "NES", "SpineOne", "MRO"];

function toNumber(value) {
  if (value == null || value === "") return 0;
  const n = Number(value);
  return Number.isFinite(n) ? n : 0;
}

function parseJson(value) {
  if (!value) return {};
  try {
    return typeof value === "string" ? JSON.parse(value) : value;
  } catch {
    return {};
  }
}

function normalizeStatus(status) {
  const s = String(status || "").toLowerCase();
  if (s === "approved") return "approved";
  if (s === "submitted") return "submitted";
  return "draft";
}

function monthKeyFromWeekEnding(weekEnding) {
  const text = String(weekEnding || "").trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(text)) return "";
  return text.slice(0, 7);
}

function mapRecord(record) {
  const values = parseJson(record.valuesJson);

  return {
    entity: record.partitionKey,
    weekEnding: record.rowKey,
    status: normalizeStatus(record.status),

    daysInPeriod:
      toNumber(values.daysInPeriod) ||
      toNumber(values.workingDaysInWeek) ||
      0,

    visitVolume:
      toNumber(values.visitVolume ?? values.totalVisits) || toNumber(record.visitVolume),

    callVolume:
      toNumber(values.callVolume ?? values.totalCalls) || toNumber(record.callVolume),

    newPatients:
      toNumber(values.newPatients ?? values.npActual) || toNumber(record.newPatients),

    surgeries:
      toNumber(values.surgeries ?? values.surgeryActual) || toNumber(record.surgeries),

    established:
      toNumber(values.established ?? values.establishedActual) || toNumber(record.established),

    noShows:
      toNumber(values.noShows) || toNumber(record.noShows),

    cancelled:
      toNumber(values.cancelled) || toNumber(record.cancelled),

    totalCalls:
      toNumber(values.totalCalls ?? values.callVolume) || toNumber(record.totalCalls ?? record.callVolume),

    abandonedCalls:
      toNumber(values.abandonedCalls) || toNumber(record.abandonedCalls),

    noShowRate:
      toNumber(values.noShowRate) || toNumber(record.noShowRate),

    cancellationRate:
      toNumber(values.cancellationRate) || toNumber(record.cancellationRate),

    abandonedCallRate:
      toNumber(values.abandonmentRate) ||
      toNumber(values.abandonedCallRate) ||
      toNumber(record.abandonedCallRate)
  };
}

function buildWeeklyBudget(budgetRow, daysInPeriod) {
  if (!budgetRow) {
    return {
      visitVolumeBudget: 0,
      newPatientsBudget: 0,
      workingDaysInMonth: 0,
      monthKey: null,
      monthLabel: null
    };
  }

  const workingDaysInMonth = toNumber(budgetRow.workingDaysInMonth);
  const safeWorkingDays = workingDaysInMonth > 0 ? workingDaysInMonth : 20;
  const safeDaysInPeriod = toNumber(daysInPeriod) > 0 ? toNumber(daysInPeriod) : 5;

  const visitVolumeBudget =
    (toNumber(budgetRow.visitBudgetMonthly) / safeWorkingDays) * safeDaysInPeriod;

  const newPatientsBudget =
    (toNumber(budgetRow.newPatientsBudgetMonthly) / safeWorkingDays) * safeDaysInPeriod;

  return {
    visitVolumeBudget,
    newPatientsBudget,
    workingDaysInMonth: safeWorkingDays,
    monthKey: budgetRow.rowKey || budgetRow.monthKey || null,
    monthLabel: budgetRow.monthLabel || null
  };
}

module.exports = async function (context, req) {
  try {
    const user = getUserFromRequest(req);
    resolveAccess(user);

    const weekEnding = req.query.weekEnding;

    if (!weekEnding) {
      return {
        status: 400,
        body: { ok: false, error: "Missing weekEnding" }
      };
    }

    const regionTable = getTableClient(REGION_TABLE);
    const budgetTable = getTableClient(BUDGET_TABLE);
    const monthKey = monthKeyFromWeekEnding(weekEnding);

    const rows = [];

    for (const entity of ENTITIES) {
      try {
        const record = await regionTable.getEntity(entity, weekEnding);
        const mapped = mapRecord(record);

        if (mapped.status !== "approved") {
          continue;
        }

        let budgetRecord = null;
        try {
          budgetRecord = await budgetTable.getEntity(entity, monthKey);
        } catch (err) {
          if (err.statusCode !== 404) throw err;
        }

        const budget = buildWeeklyBudget(budgetRecord, mapped.daysInPeriod);

        rows.push({
          ...mapped,
          budget
        });
      } catch (err) {
        if (err.statusCode !== 404) throw err;
      }
    }

    const totals = {
      visitVolume: rows.reduce((s, r) => s + r.visitVolume, 0),
      callVolume: rows.reduce((s, r) => s + r.callVolume, 0),
      newPatients: rows.reduce((s, r) => s + r.newPatients, 0),
      surgeries: rows.reduce((s, r) => s + r.surgeries, 0)
    };

    const budgetTotals = {
      visitVolumeBudget: rows.reduce((s, r) => s + toNumber(r.budget?.visitVolumeBudget), 0),
      newPatientsBudget: rows.reduce((s, r) => s + toNumber(r.budget?.newPatientsBudget), 0)
    };

    const variances = {
      visitVolumeVariance: totals.visitVolume - budgetTotals.visitVolumeBudget,
      newPatientsVariance: totals.newPatients - budgetTotals.newPatientsBudget
    };

    const avg = (arr) =>
      arr.length ? arr.reduce((s, v) => s + v, 0) / arr.length : 0;

    const averages = {
      noShowRate: avg(rows.map((r) => r.noShowRate)),
      cancellationRate: avg(rows.map((r) => r.cancellationRate)),
      abandonedCallRate: avg(rows.map((r) => r.abandonedCallRate))
    };

    return {
      status: 200,
      body: {
        ok: true,
        weekEnding,
        monthKey,
        entityCount: rows.length,
        totals,
        budgetTotals,
        variances,
        averages,
        regions: rows
      }
    };
  } catch (error) {
    context.log.error("executive failed", error);

    return {
      status: 500,
      body: {
        ok: false,
        error: "Failed to load executive summary",
        details: error.message
      }
    };
  }
};