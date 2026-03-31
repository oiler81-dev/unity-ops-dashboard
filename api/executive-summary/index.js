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

function safeText(value) {
  return value == null ? "" : String(value).trim();
}

function normalizeStatus(status) {
  const s = String(status || "").toLowerCase();
  if (s === "approved") return "approved";
  if (s === "submitted") return "submitted";
  return "draft";
}

function monthKeyFromWeekEnding(weekEnding) {
  const text = safeText(weekEnding);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(text)) return "";
  return text.slice(0, 7);
}

function daysInMonthFromMonthKey(monthKey) {
  if (!/^\d{4}-\d{2}$/.test(monthKey)) return 0;
  const [year, month] = monthKey.split("-").map(Number);
  return new Date(Date.UTC(year, month, 0)).getUTCDate();
}

function dayOfMonthFromWeekEnding(weekEnding) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(weekEnding || ""))) return 0;
  return Number(String(weekEnding).slice(8, 10));
}

function calculateDaysInPeriod(weekEnding) {
  const day = dayOfMonthFromWeekEnding(weekEnding);
  if (!day) return 5;
  return day >= 5 ? 5 : day;
}

function buildBudgetMap(rows) {
  const map = {};

  for (const row of rows || []) {
    const entity = safeText(row.partitionKey || row.PartitionKey || row.entity);
    const monthKey = safeText(row.rowKey || row.RowKey || row.monthKey);

    if (!entity || !monthKey) continue;

    map[`${entity}__${monthKey}`] = {
      entity,
      monthKey,
      monthLabel: safeText(row.monthLabel),
      visitBudgetMonthly: toNumber(row.visitBudgetMonthly),
      newPatientsBudgetMonthly: toNumber(row.newPatientsBudgetMonthly),
      workingDaysInMonth: toNumber(row.workingDaysInMonth)
    };
  }

  return map;
}

function getWeeklyBudgetForEntity(budgetRow, weekEnding) {
  if (!budgetRow) {
    return {
      visitVolumeBudget: 0,
      newPatientsBudget: 0,
      workingDaysInMonth: 0,
      monthKey: monthKeyFromWeekEnding(weekEnding),
      monthLabel: ""
    };
  }

  const daysInPeriod = calculateDaysInPeriod(weekEnding);
  const workingDaysInMonth = toNumber(budgetRow.workingDaysInMonth) || 1;

  return {
    visitVolumeBudget: (toNumber(budgetRow.visitBudgetMonthly) / workingDaysInMonth) * daysInPeriod,
    newPatientsBudget: (toNumber(budgetRow.newPatientsBudgetMonthly) / workingDaysInMonth) * daysInPeriod,
    workingDaysInMonth,
    monthKey: safeText(budgetRow.monthKey),
    monthLabel: safeText(budgetRow.monthLabel)
  };
}

function mapRecord(record, budgetRow) {
  const values = parseJson(record.valuesJson);

  const weekEnding = record.rowKey;
  const budget = getWeeklyBudgetForEntity(budgetRow, weekEnding);

  return {
    entity: record.partitionKey,
    weekEnding,
    status: normalizeStatus(record.status),
    daysInPeriod: calculateDaysInPeriod(weekEnding),

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
      toNumber(values.totalCalls ?? values.callVolume) || toNumber(record.totalCalls) || toNumber(record.callVolume),

    abandonedCalls:
      toNumber(values.abandonedCalls) || toNumber(record.abandonedCalls),

    noShowRate:
      toNumber(values.noShowRate) || toNumber(record.noShowRate),

    cancellationRate:
      toNumber(values.cancellationRate) || toNumber(record.cancellationRate),

    abandonedCallRate:
      toNumber(values.abandonedCallRate ?? values.abandonmentRate) || toNumber(record.abandonedCallRate),

    ptScheduledVisits:
      toNumber(values.ptScheduledVisits) || toNumber(record.ptScheduledVisits),

    ptCancellations:
      toNumber(values.ptCancellations) || toNumber(record.ptCancellations),

    ptNoShows:
      toNumber(values.ptNoShows) || toNumber(record.ptNoShows),

    ptReschedules:
      toNumber(values.ptReschedules) || toNumber(record.ptReschedules),

    ptTotalUnitsBilled:
      toNumber(values.ptTotalUnitsBilled) || toNumber(record.ptTotalUnitsBilled),

    ptVisitsSeen:
      toNumber(values.ptVisitsSeen) || toNumber(record.ptVisitsSeen),

    ptWorkingDays:
      toNumber(values.ptWorkingDays) || toNumber(record.ptWorkingDays),

    ptUnitsPerVisit:
      toNumber(values.ptUnitsPerVisit) || toNumber(record.ptUnitsPerVisit),

    ptVisitsPerDay:
      toNumber(values.ptVisitsPerDay) || toNumber(record.ptVisitsPerDay),

    budget
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

    let budgetRows = [];
    try {
      budgetRows = await budgetTable.query(
        `RowKey eq '${String(monthKey).replace(/'/g, "''")}'`
      );
    } catch (error) {
      budgetRows = [];
    }

    const budgetMap = buildBudgetMap(budgetRows);
    const rows = [];

    for (const entity of ENTITIES) {
      try {
        const record = await regionTable.getEntity(entity, weekEnding);
        const mapped = mapRecord(record, budgetMap[`${entity}__${monthKey}`]);

        if (mapped.status === "approved") {
          rows.push(mapped);
        }
      } catch (err) {
        if (err.statusCode !== 404) throw err;
      }
    }

    const totals = {
      visitVolume: rows.reduce((s, r) => s + r.visitVolume, 0),
      callVolume: rows.reduce((s, r) => s + r.callVolume, 0),
      newPatients: rows.reduce((s, r) => s + r.newPatients, 0),
      surgeries: rows.reduce((s, r) => s + r.surgeries, 0),

      ptScheduledVisits: rows.reduce((s, r) => s + r.ptScheduledVisits, 0),
      ptCancellations: rows.reduce((s, r) => s + r.ptCancellations, 0),
      ptNoShows: rows.reduce((s, r) => s + r.ptNoShows, 0),
      ptReschedules: rows.reduce((s, r) => s + r.ptReschedules, 0),
      ptTotalUnitsBilled: rows.reduce((s, r) => s + r.ptTotalUnitsBilled, 0),
      ptVisitsSeen: rows.reduce((s, r) => s + r.ptVisitsSeen, 0)
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
      abandonedCallRate: avg(rows.map((r) => r.abandonedCallRate)),
      ptUnitsPerVisit: avg(rows.filter((r) => r.ptVisitsSeen > 0).map((r) => r.ptUnitsPerVisit)),
      ptVisitsPerDay: avg(rows.filter((r) => r.ptWorkingDays > 0).map((r) => r.ptVisitsPerDay))
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
    context.log.error("executive-summary failed", error);

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