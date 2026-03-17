const { getTableClient } = require("../shared/table");

const REGION_TABLE = "WeeklyRegionData";
const SHARED_TABLE = "SharedPageData";

const REGION_ENTITIES = ["LAOSS", "NES", "SpineOne", "MRO"];
const SHARED_PAGES = ["PT", "CXNS"];

function toNumber(value, fallback = 0) {
  const n = Number(value);
  return Number.isFinite(n) ? n : fallback;
}

function pct(value, digits = 1) {
  return `${toNumber(value).toFixed(digits)}%`;
}

function whole(value) {
  return Math.round(toNumber(value)).toLocaleString();
}

function fixed(value, digits = 1) {
  return toNumber(value).toFixed(digits);
}

function statusFromChange(change, betterDirection = "up") {
  const n = toNumber(change, 0);

  if (betterDirection === "up") {
    if (n > 0) return { status: "Up", statusColor: "green" };
    if (n < 0) return { status: "Down", statusColor: "red" };
    return { status: "Flat", statusColor: "yellow" };
  }

  if (betterDirection === "down") {
    if (n < 0) return { status: "Improved", statusColor: "green" };
    if (n > 0) return { status: "Worse", statusColor: "red" };
    return { status: "Flat", statusColor: "yellow" };
  }

  return { status: "Flat", statusColor: "yellow" };
}

function parseValuesJson(record) {
  if (!record || !record.valuesJson) return {};
  try {
    return JSON.parse(record.valuesJson);
  } catch (err) {
    return {};
  }
}

function getPreviousWeekEnding(weekEnding) {
  const date = new Date(`${weekEnding}T12:00:00Z`);
  date.setUTCDate(date.getUTCDate() - 7);
  return date.toISOString().split("T")[0];
}

async function getEntityWeek(table, entity, weekEnding) {
  try {
    const record = await table.getEntity(entity, weekEnding);
    return parseValuesJson(record);
  } catch (err) {
    if (err.statusCode === 404) return {};
    throw err;
  }
}

async function getSharedWeek(table, page, weekEnding) {
  try {
    const record = await table.getEntity(page, weekEnding);
    return parseValuesJson(record);
  } catch (err) {
    if (err.statusCode === 404) return {};
    throw err;
  }
}

function aggregateRegionRows(rows) {
  const totals = {
    totalVisits: 0,
    daysInPeriod: 0,
    npActual: 0,
    establishedActual: 0,
    surgeryActual: 0,
    totalCalls: 0,
    abandonedCalls: 0,
    cashActual: 0
  };

  for (const row of rows) {
    totals.totalVisits += toNumber(row.totalVisits);
    totals.daysInPeriod += toNumber(row.daysInPeriod);
    totals.npActual += toNumber(row.npActual);
    totals.establishedActual += toNumber(row.establishedActual);
    totals.surgeryActual += toNumber(row.surgeryActual);
    totals.totalCalls += toNumber(row.totalCalls);
    totals.abandonedCalls += toNumber(row.abandonedCalls);
    totals.cashActual += toNumber(row.cashActual);
  }

  totals.visitsPerDay =
    totals.daysInPeriod > 0 ? totals.totalVisits / totals.daysInPeriod : 0;

  totals.abandonmentRate =
    totals.totalCalls > 0
      ? (totals.abandonedCalls / totals.totalCalls) * 100
      : 0;

  totals.answeredCalls = Math.max(totals.totalCalls - totals.abandonedCalls, 0);

  totals.answeredCallToNpConversion =
    totals.answeredCalls > 0
      ? (totals.npActual / totals.answeredCalls) * 100
      : 0;

  return totals;
}

function aggregateSharedRows(rowsByPage) {
  const pt = rowsByPage.PT || {};
  const cxns = rowsByPage.CXNS || {};

  const ptScheduledVisits = toNumber(pt.ptScheduledVisits);
  const ptCancellations = toNumber(pt.ptCancellations);
  const ptNoShows = toNumber(pt.ptNoShows);
  const ptReschedules = toNumber(pt.ptReschedules);
  const totalUnitsBilled = toNumber(pt.totalUnitsBilled);

  const cxnsScheduled = toNumber(cxns.scheduledAppts);
  const cxnsCancellations = toNumber(cxns.cancellations);
  const cxnsNoShows = toNumber(cxns.noShows);
  const cxnsReschedules = toNumber(cxns.reschedules);

  return {
    ptScheduledVisits,
    ptCancellations,
    ptNoShows,
    ptReschedules,
    totalUnitsBilled,
    cxnsScheduled,
    cxnsCancellations,
    cxnsNoShows,
    cxnsReschedules
  };
}

function buildKpis(currentRegions, previousRegions, currentShared, previousShared) {
  const totalVisitsChange = currentRegions.totalVisits - previousRegions.totalVisits;
  const visitsPerDayChange = currentRegions.visitsPerDay - previousRegions.visitsPerDay;
  const abandonmentRateChange =
    currentRegions.abandonmentRate - previousRegions.abandonmentRate;
  const newPatientsChange = currentRegions.npActual - previousRegions.npActual;
  const surgeriesChange = currentRegions.surgeryActual - previousRegions.surgeryActual;
  const callsChange = currentRegions.totalCalls - previousRegions.totalCalls;
  const conversionChange =
    currentRegions.answeredCallToNpConversion -
    previousRegions.answeredCallToNpConversion;
  const cashChange = currentRegions.cashActual - previousRegions.cashActual;
  const ptVisitsChange =
    currentShared.ptScheduledVisits - previousShared.ptScheduledVisits;
  const cxnsScheduledChange =
    currentShared.cxnsScheduled - previousShared.cxnsScheduled;

  return [
    {
      label: "Total Visits",
      value: whole(currentRegions.totalVisits),
      meta: `${totalVisitsChange >= 0 ? "+" : ""}${whole(totalVisitsChange)} vs prior week`,
      ...statusFromChange(totalVisitsChange, "up")
    },
    {
      label: "Visits / Day",
      value: fixed(currentRegions.visitsPerDay, 1),
      meta: `${visitsPerDayChange >= 0 ? "+" : ""}${fixed(visitsPerDayChange, 1)} vs prior week`,
      ...statusFromChange(visitsPerDayChange, "up")
    },
    {
      label: "New Patients",
      value: whole(currentRegions.npActual),
      meta: `${newPatientsChange >= 0 ? "+" : ""}${whole(newPatientsChange)} vs prior week`,
      ...statusFromChange(newPatientsChange, "up")
    },
    {
      label: "Surgical Cases",
      value: whole(currentRegions.surgeryActual),
      meta: `${surgeriesChange >= 0 ? "+" : ""}${whole(surgeriesChange)} vs prior week`,
      ...statusFromChange(surgeriesChange, "up")
    },
    {
      label: "Call Volume",
      value: whole(currentRegions.totalCalls),
      meta: `${callsChange >= 0 ? "+" : ""}${whole(callsChange)} vs prior week`,
      ...statusFromChange(callsChange, "up")
    },
    {
      label: "Abandonment Rate",
      value: pct(currentRegions.abandonmentRate, 1),
      meta: `${abandonmentRateChange >= 0 ? "+" : ""}${fixed(abandonmentRateChange, 1)} pts vs prior week`,
      ...statusFromChange(abandonmentRateChange, "down")
    },
    {
      label: "Answered Call to NP %",
      value: pct(currentRegions.answeredCallToNpConversion, 1),
      meta: `${conversionChange >= 0 ? "+" : ""}${fixed(conversionChange, 1)} pts vs prior week`,
      ...statusFromChange(conversionChange, "up")
    },
    {
      label: "Cash Collected",
      value: `$${whole(currentRegions.cashActual)}`,
      meta: `${cashChange >= 0 ? "+" : ""}$${whole(cashChange)} vs prior week`,
      ...statusFromChange(cashChange, "up")
    },
    {
      label: "PT Scheduled Visits",
      value: whole(currentShared.ptScheduledVisits),
      meta: `${ptVisitsChange >= 0 ? "+" : ""}${whole(ptVisitsChange)} vs prior week`,
      ...statusFromChange(ptVisitsChange, "up")
    },
    {
      label: "CXNS Scheduled",
      value: whole(currentShared.cxnsScheduled),
      meta: `${cxnsScheduledChange >= 0 ? "+" : ""}${whole(cxnsScheduledChange)} vs prior week`,
      ...statusFromChange(cxnsScheduledChange, "up")
    }
  ];
}

module.exports = async function (context, req) {
  try {
    const weekEnding = String(req.query.weekEnding || "").trim();

    if (!weekEnding) {
      context.res = {
        status: 400,
        body: { error: "Missing weekEnding." }
      };
      return;
    }

    const previousWeekEnding = getPreviousWeekEnding(weekEnding);

    const regionTable = getTableClient(REGION_TABLE);
    const sharedTable = getTableClient(SHARED_TABLE);

    const currentRegionRows = await Promise.all(
      REGION_ENTITIES.map((entity) => getEntityWeek(regionTable, entity, weekEnding))
    );

    const previousRegionRows = await Promise.all(
      REGION_ENTITIES.map((entity) => getEntityWeek(regionTable, entity, previousWeekEnding))
    );

    const currentSharedEntries = await Promise.all(
      SHARED_PAGES.map(async (page) => [page, await getSharedWeek(sharedTable, page, weekEnding)])
    );

    const previousSharedEntries = await Promise.all(
      SHARED_PAGES.map(async (page) => [page, await getSharedWeek(sharedTable, page, previousWeekEnding)])
    );

    const currentSharedByPage = Object.fromEntries(currentSharedEntries);
    const previousSharedByPage = Object.fromEntries(previousSharedEntries);

    const currentRegions = aggregateRegionRows(currentRegionRows);
    const previousRegions = aggregateRegionRows(previousRegionRows);

    const currentShared = aggregateSharedRows(currentSharedByPage);
    const previousShared = aggregateSharedRows(previousSharedByPage);

    const kpis = buildKpis(
      currentRegions,
      previousRegions,
      currentShared,
      previousShared
    );

    context.res = {
      status: 200,
      headers: {
        "Content-Type": "application/json"
      },
      body: {
        ok: true,
        weekEnding,
        previousWeekEnding,
        kpis,
        summary: {
          currentRegions,
          previousRegions,
          currentShared,
          previousShared
        }
      }
    };
  } catch (error) {
    context.log.error("dashboard failed", error);

    context.res = {
      status: 500,
      body: {
        error: "Failed to load dashboard.",
        details: error.message
      }
    };
  }
};
