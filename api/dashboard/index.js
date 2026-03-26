const { getUserFromRequest } = require("../shared/auth");
const { resolveAccess } = require("../shared/permissions");
const { getTableClient } = require("../shared/table");
const {
  getWeekRangeFromWeekEnding,
  getPreviousMonthDateRange,
  getMonthDateRange,
  getProratedBudgetForRange,
  isoDate
} = require("../shared/budget");

const REGION_TABLE = "WeeklyRegionData";
const TARGETS_TABLE = "ReferenceData";
const BUDGET_TABLE = "BudgetData";
const ENTITIES = ["LAOSS", "NES", "SpineOne", "MRO"];

function toNumber(value) {
  if (value == null || value === "") return null;
  const n = Number(value);
  return Number.isFinite(n) ? n : null;
}

function parseJson(value, fallback = {}) {
  if (!value) return fallback;
  try {
    return typeof value === "string" ? JSON.parse(value) : value;
  } catch {
    return fallback;
  }
}

function normalizeStatus(status) {
  const s = String(status || "").trim().toLowerCase();
  if (s === "approved") return "Approved";
  if (s === "submitted") return "submitted";
  if (s === "draft") return "draft";
  return status || "draft";
}

function average(values) {
  if (!values.length) return 0;
  return values.reduce((sum, v) => sum + v, 0) / values.length;
}

function pctVariance(actual, target) {
  const a = toNumber(actual) || 0;
  const t = toNumber(target) || 0;
  if (!t) return null;
  return ((a - t) / t) * 100;
}

function buildTrend(actual, previous) {
  const a = toNumber(actual) || 0;
  const p = toNumber(previous) || 0;
  const diff = a - p;

  return {
    current: a,
    previous: p,
    diff,
    direction: diff > 0 ? "up" : diff < 0 ? "down" : "flat"
  };
}

function getPreviousWeekEnding(weekEnding) {
  const d = new Date(`${weekEnding}T12:00:00Z`);
  d.setUTCDate(d.getUTCDate() - 7);
  return d.toISOString().split("T")[0];
}

function addDays(isoDateValue, days) {
  const d = new Date(`${isoDateValue}T12:00:00Z`);
  d.setUTCDate(d.getUTCDate() + days);
  return d.toISOString().slice(0, 10);
}

function mapRecord(record) {
  if (!record) return null;

  const values = parseJson(record.valuesJson, {});
  const visitVolume = toNumber(values.totalVisits) ?? toNumber(record.visitVolume) ?? 0;
  const callVolume = toNumber(values.totalCalls) ?? toNumber(record.callVolume) ?? 0;
  const newPatients = toNumber(values.npActual) ?? toNumber(record.newPatients) ?? 0;
  const noShowRate = toNumber(values.noShowRate) ?? toNumber(record.noShowRate) ?? 0;
  const cancellationRate =
    toNumber(values.cancellationRate) ?? toNumber(record.cancellationRate) ?? 0;
  const abandonedCallRate =
    toNumber(values.abandonmentRate) ?? toNumber(record.abandonedCallRate) ?? 0;

  return {
    entity: record.entity || record.partitionKey,
    weekEnding: record.weekEnding || record.rowKey,
    status: normalizeStatus(record.status),
    visitVolume,
    callVolume,
    newPatients,
    noShowRate,
    cancellationRate,
    abandonedCallRate,
    source: record.source || null,
    updatedAt: record.updatedAt || record.importedAt || record.approvedAt || null,
    rawValues: values
  };
}

async function getWeekRecord(table, entity, weekEnding) {
  try {
    const record = await table.getEntity(entity, weekEnding);
    return mapRecord(record);
  } catch (err) {
    if (err.statusCode === 404) return null;
    throw err;
  }
}

async function getEntityTargets(table, entity) {
  try {
    const record = await table.getEntity("Targets", entity);
    const values = parseJson(record.valuesJson, {});
    return {
      visitTarget: toNumber(values.visitTarget),
      callTarget: toNumber(values.callTarget),
      newPatientTarget: toNumber(values.newPatientTarget)
    };
  } catch (err) {
    if (err.statusCode === 404) {
      return {
        visitTarget: null,
        callTarget: null,
        newPatientTarget: null
      };
    }
    throw err;
  }
}

function buildAlertRow(entityRow, compareAgainst) {
  const alerts = [];

  if (entityRow.status !== "Approved") {
    alerts.push({
      entity: entityRow.entity,
      severity: "yellow",
      message: `${entityRow.entity} is not approved for this period`
    });
  }

  if ((entityRow.noShowRate || 0) >= 6) {
    alerts.push({
      entity: entityRow.entity,
      severity: "red",
      message: `${entityRow.entity} no-show rate is elevated`
    });
  }

  if ((entityRow.cancellationRate || 0) >= 8) {
    alerts.push({
      entity: entityRow.entity,
      severity: "red",
      message: `${entityRow.entity} cancellation rate is elevated`
    });
  }

  if ((entityRow.abandonedCallRate || 0) >= 10) {
    alerts.push({
      entity: entityRow.entity,
      severity: "red",
      message: `${entityRow.entity} abandoned call rate is elevated`
    });
  }

  if (compareAgainst === "budget") {
    const visitVar = entityRow.variance?.visitVariancePct;
    const npVar = entityRow.variance?.newPatientVariancePct;

    if (visitVar != null && visitVar < -10) {
      alerts.push({
        entity: entityRow.entity,
        severity: "yellow",
        message: `${entityRow.entity} visit volume is below budget`
      });
    }

    if (npVar != null && npVar < -10) {
      alerts.push({
        entity: entityRow.entity,
        severity: "yellow",
        message: `${entityRow.entity} new patient volume is below budget`
      });
    }
  }

  return alerts;
}

function buildWeekSets(periodType, anchorWeek, customStart, customEnd) {
  if (periodType === "currentWeek") {
    return {
      primaryWeeks: [anchorWeek],
      comparisonWeeks: [getPreviousWeekEnding(anchorWeek)],
      startDate: getWeekRangeFromWeekEnding(anchorWeek).startDate,
      endDate: getWeekRangeFromWeekEnding(anchorWeek).endDate,
      summary: `Viewing Current Week anchored to ${anchorWeek}`
    };
  }

  if (periodType === "lastWeek") {
    const primary = getPreviousWeekEnding(anchorWeek);
    const range = getWeekRangeFromWeekEnding(primary);
    return {
      primaryWeeks: [primary],
      comparisonWeeks: [getPreviousWeekEnding(primary)],
      startDate: range.startDate,
      endDate: range.endDate,
      summary: `Viewing Last Week anchored from ${anchorWeek}`
    };
  }

  if (periodType === "rolling4") {
    const primaryWeeks = [
      addDays(anchorWeek, -21),
      addDays(anchorWeek, -14),
      addDays(anchorWeek, -7),
      anchorWeek
    ];
    const comparisonWeeks = primaryWeeks.map((w) => addDays(w, -28));
    const firstRange = getWeekRangeFromWeekEnding(primaryWeeks[0]);
    const lastRange = getWeekRangeFromWeekEnding(primaryWeeks[primaryWeeks.length - 1]);

    return {
      primaryWeeks,
      comparisonWeeks,
      startDate: firstRange.startDate,
      endDate: lastRange.endDate,
      summary: `Viewing Rolling 4 Weeks ending ${anchorWeek}`
    };
  }

  if (periodType === "mtd") {
    const monthKey = anchorWeek.slice(0, 7);
    const monthRange = getMonthDateRange(monthKey);
    const primaryWeeks = [];

    const d = new Date(`${monthRange.startDate}T12:00:00Z`);
    while (d.toISOString().slice(0, 10) <= anchorWeek) {
      if (d.getUTCDay() === 5) {
        primaryWeeks.push(d.toISOString().slice(0, 10));
      }
      d.setUTCDate(d.getUTCDate() + 1);
    }

    const comparisonWeeks = primaryWeeks.map((w) => addDays(w, -28));

    return {
      primaryWeeks,
      comparisonWeeks,
      startDate: monthRange.startDate,
      endDate: anchorWeek,
      summary: `Viewing Month to Date through ${anchorWeek}`
    };
  }

  if (periodType === "lastMonth") {
    const previous = getPreviousMonthDateRange(anchorWeek);
    const primaryWeeks = [];
    const walker = new Date(`${previous.startDate}T12:00:00Z`);

    while (walker.toISOString().slice(0, 10) <= previous.endDate) {
      if (walker.getUTCDay() === 5) {
        primaryWeeks.push(walker.toISOString().slice(0, 10));
      }
      walker.setUTCDate(walker.getUTCDate() + 1);
    }

    const comparisonWeeks = primaryWeeks.map((w) => addDays(w, -28));

    return {
      primaryWeeks,
      comparisonWeeks,
      startDate: previous.startDate,
      endDate: previous.endDate,
      summary: `Viewing Last Month (${previous.startDate} to ${previous.endDate})`
    };
  }

  if (periodType === "custom" && customStart && customEnd) {
    const primaryWeeks = [];
    const walker = new Date(`${customStart}T12:00:00Z`);
    const end = new Date(`${customEnd}T12:00:00Z`);

    while (walker <= end) {
      if (walker.getUTCDay() === 5) {
        primaryWeeks.push(walker.toISOString().slice(0, 10));
      }
      walker.setUTCDate(walker.getUTCDate() + 1);
    }

    const comparisonWeeks = primaryWeeks.map((w) => addDays(w, -28));

    return {
      primaryWeeks,
      comparisonWeeks,
      startDate: customStart,
      endDate: customEnd,
      summary: `Viewing Custom Range (${customStart} to ${customEnd})`
    };
  }

  const fallbackRange = getWeekRangeFromWeekEnding(anchorWeek);
  return {
    primaryWeeks: [anchorWeek],
    comparisonWeeks: [getPreviousWeekEnding(anchorWeek)],
    startDate: fallbackRange.startDate,
    endDate: fallbackRange.endDate,
    summary: `Viewing Current Week anchored to ${anchorWeek}`
  };
}

async function aggregateEntityForWeeks(regionTable, entity, weeks) {
  const records = [];

  for (const week of weeks) {
    const record = await getWeekRecord(regionTable, entity, week);
    if (record) {
      records.push(record);
    }
  }

  if (!records.length) {
    return {
      entity,
      status: "missing",
      weekEnding: weeks[weeks.length - 1] || null,
      visitVolume: 0,
      callVolume: 0,
      newPatients: 0,
      noShowRate: 0,
      cancellationRate: 0,
      abandonedCallRate: 0,
      recordCount: 0
    };
  }

  const approvedRecords = records.filter((r) => r.status === "Approved");
  const baseRecords = approvedRecords.length ? approvedRecords : records;

  return {
    entity,
    status: approvedRecords.length === records.length ? "Approved" : records[0].status,
    weekEnding: weeks[weeks.length - 1] || records[records.length - 1].weekEnding,
    visitVolume: baseRecords.reduce((sum, r) => sum + (r.visitVolume || 0), 0),
    callVolume: baseRecords.reduce((sum, r) => sum + (r.callVolume || 0), 0),
    newPatients: baseRecords.reduce((sum, r) => sum + (r.newPatients || 0), 0),
    noShowRate: average(baseRecords.map((r) => r.noShowRate || 0)),
    cancellationRate: average(baseRecords.map((r) => r.cancellationRate || 0)),
    abandonedCallRate: average(baseRecords.map((r) => r.abandonedCallRate || 0)),
    recordCount: baseRecords.length
  };
}

function buildKpis(currentRows, compareAgainst, compareRows, benchmarkRows) {
  const approvedRows = currentRows.filter((r) => r.status === "Approved");
  const totals = {
    visitVolume: approvedRows.reduce((sum, r) => sum + (r.visitVolume || 0), 0),
    callVolume: approvedRows.reduce((sum, r) => sum + (r.callVolume || 0), 0),
    newPatients: approvedRows.reduce((sum, r) => sum + (r.newPatients || 0), 0)
  };

  const averages = {
    noShowRate: average(approvedRows.map((r) => r.noShowRate || 0)),
    cancellationRate: average(approvedRows.map((r) => r.cancellationRate || 0)),
    abandonedCallRate: average(approvedRows.map((r) => r.abandonedCallRate || 0))
  };

  let approvedMeta = `${approvedRows.length} region(s) approved`;
  let visitMeta = "Approved-region total";
  let callMeta = "Approved-region total";
  let npMeta = "Approved-region total";

  if (compareAgainst === "priorPeriod") {
    const compareApproved = compareRows.filter((r) => r.status === "Approved");
    const compareTotals = {
      visitVolume: compareApproved.reduce((sum, r) => sum + (r.visitVolume || 0), 0),
      callVolume: compareApproved.reduce((sum, r) => sum + (r.callVolume || 0), 0),
      newPatients: compareApproved.reduce((sum, r) => sum + (r.newPatients || 0), 0)
    };

    approvedMeta = `Prev ${compareApproved.length}`;
    visitMeta = `${totals.visitVolume - compareTotals.visitVolume >= 0 ? "+" : ""}${totals.visitVolume - compareTotals.visitVolume} vs prior period`;
    callMeta = `${totals.callVolume - compareTotals.callVolume >= 0 ? "+" : ""}${totals.callVolume - compareTotals.callVolume} vs prior period`;
    npMeta = `${totals.newPatients - compareTotals.newPatients >= 0 ? "+" : ""}${totals.newPatients - compareTotals.newPatients} vs prior period`;
  } else if (compareAgainst === "budget") {
    const budgetTotals = {
      visitVolume: benchmarkRows.reduce((sum, r) => sum + (r.budgetVisitVolume || 0), 0),
      newPatients: benchmarkRows.reduce((sum, r) => sum + (r.budgetNewPatients || 0), 0)
    };

    approvedMeta = "Budget benchmark";
    visitMeta = `${Math.round(totals.visitVolume - budgetTotals.visitVolume) >= 0 ? "+" : ""}${Math.round(totals.visitVolume - budgetTotals.visitVolume)} vs budget`;
    callMeta = "No call budget loaded";
    npMeta = `${Math.round(totals.newPatients - budgetTotals.newPatients) >= 0 ? "+" : ""}${Math.round(totals.newPatients - budgetTotals.newPatients)} vs budget`;
  }

  const kpis = [
    {
      key: "approvedRegions",
      label: "Approved Regions",
      value: approvedRows.length,
      meta: approvedMeta
    },
    {
      key: "visitVolume",
      label: "Visit Volume",
      value: totals.visitVolume,
      meta: visitMeta
    },
    {
      key: "callVolume",
      label: "Call Volume",
      value: totals.callVolume,
      meta: callMeta
    },
    {
      key: "newPatients",
      label: "New Patients",
      value: totals.newPatients,
      meta: npMeta
    },
    {
      key: "avgNoShowRate",
      label: "Avg No Show %",
      value: averages.noShowRate,
      format: "percent"
    },
    {
      key: "avgCancellationRate",
      label: "Avg Cancel %",
      value: averages.cancellationRate,
      format: "percent"
    },
    {
      key: "avgAbandonedCallRate",
      label: "Avg Abandoned %",
      value: averages.abandonedCallRate,
      format: "percent"
    }
  ];

  return { approvedRows, totals, averages, kpis };
}

module.exports = async function (context, req) {
  try {
    const user = getUserFromRequest(req);
    resolveAccess(user);

    const anchorWeek = String(req.query.anchorWeek || req.query.weekEnding || "").trim();
    if (!anchorWeek) {
      return {
        status: 400,
        body: {
          ok: false,
          error: "Missing weekEnding or anchorWeek"
        }
      };
    }

    const periodType = String(req.query.periodType || "currentWeek").trim();
    const compareAgainst = String(req.query.compareAgainst || "priorPeriod").trim();
    const entityScope = String(req.query.entityScope || "ALL").trim();
    const customStart = String(req.query.customStart || "").trim();
    const customEnd = String(req.query.customEnd || "").trim();

    const weekSets = buildWeekSets(periodType, anchorWeek, customStart, customEnd);

    const regionTable = getTableClient(REGION_TABLE);
    const targetsTable = getTableClient(TARGETS_TABLE);
    const budgetTable = getTableClient(BUDGET_TABLE);

    const selectedEntities =
      entityScope && entityScope !== "ALL"
        ? ENTITIES.filter((e) => e === entityScope)
        : ENTITIES.slice();

    const currentRows = [];
    const compareRows = [];
    const benchmarkRows = [];
    const alerts = [];

    for (const entity of selectedEntities) {
      const current = await aggregateEntityForWeeks(regionTable, entity, weekSets.primaryWeeks);
      const previous = await aggregateEntityForWeeks(regionTable, entity, weekSets.comparisonWeeks);
      const targets = await getEntityTargets(targetsTable, entity);

      let budget = {
        budgetVisitVolume: null,
        budgetNewPatients: null,
        budgetCallVolume: null,
        budgetMeta: null
      };

      if (compareAgainst === "budget") {
        const budgetTotals = await getProratedBudgetForRange(
          budgetTable,
          entity,
          weekSets.startDate,
          weekSets.endDate
        );

        budget = {
          budgetVisitVolume: budgetTotals.visitBudget,
          budgetNewPatients: budgetTotals.newPatientsBudget,
          budgetCallVolume: null,
          budgetMeta: {
            workingDaysUsed: budgetTotals.workingDaysUsed,
            monthBreakdown: budgetTotals.monthBreakdown
          }
        };
      }

      const row = {
        ...current,
        targets,
        budgetVisitVolume: budget.budgetVisitVolume,
        budgetNewPatients: budget.budgetNewPatients,
        budgetCallVolume: budget.budgetCallVolume,
        budgetMeta: budget.budgetMeta,
        variance: {
          visitVariancePct:
            compareAgainst === "budget"
              ? pctVariance(current.visitVolume, budget.budgetVisitVolume)
              : pctVariance(current.visitVolume, targets.visitTarget),
          callVariancePct:
            compareAgainst === "budget"
              ? null
              : pctVariance(current.callVolume, targets.callTarget),
          newPatientVariancePct:
            compareAgainst === "budget"
              ? pctVariance(current.newPatients, budget.budgetNewPatients)
              : pctVariance(current.newPatients, targets.newPatientTarget)
        },
        trends: {
          visits: buildTrend(current.visitVolume, previous.visitVolume || 0),
          calls: buildTrend(current.callVolume, previous.callVolume || 0),
          newPatients: buildTrend(current.newPatients, previous.newPatients || 0)
        }
      };

      currentRows.push(row);
      compareRows.push(previous);
      benchmarkRows.push({
        entity,
        budgetVisitVolume: budget.budgetVisitVolume || 0,
        budgetNewPatients: budget.budgetNewPatients || 0
      });
      alerts.push(...buildAlertRow(row, compareAgainst));
    }

    const summary = buildKpis(currentRows, compareAgainst, compareRows, benchmarkRows);

    return {
      status: 200,
      body: {
        ok: true,
        anchorWeek,
        periodType,
        compareAgainst,
        entityScope,
        weekSets,
        summaryText: weekSets.summary,
        startDate: weekSets.startDate,
        endDate: weekSets.endDate,
        entityCount: summary.approvedRows.length,
        totals: summary.totals,
        averages: summary.averages,
        kpis: summary.kpis,
        entities: currentRows,
        comparisonEntities: compareRows,
        alerts
      }
    };
  } catch (error) {
    context.log.error("dashboard failed", error);

    return {
      status: 500,
      body: {
        ok: false,
        error: "Failed to load dashboard",
        details: error.message
      }
    };
  }
};
