/**
 * Azure Function: call-center
 * GET /api/call-center?startDate=YYYY-MM-DD&endDate=YYYY-MM-DD&entity=<optional>
 *
 * Returns aggregated call-center metrics across the requested date range,
 * grouped multiple ways for the Call Center Report view:
 *   - totals          — full-period totals across all entities the caller can see
 *   - byEntity        — one row per entity with totals + rates + source
 *   - byDate          — one row per day across all entities (for the trend chart)
 *   - byDayOfWeek     — Mon–Fri averages (Sat/Sun usually empty for clinic ops)
 *   - perEntityByDate — entity × date grid for stacked-bar style charts
 *
 * Caller scoping: admins see all 4 entities. Regional users see only their
 * own entity. The `entity` query param further filters within accessible set.
 */
const { TableClient } = require("@azure/data-tables");
const { getUserFromRequest } = require("../shared/auth");
const {
  resolveAccess,
  requireAccess,
  canAccessEntity,
  scopeEntitiesToAccess,
  safeErrorResponse
} = require("../shared/permissions");

const TABLE_NAME = "CallData";
const ALL_ENTITIES = ["LAOSS", "NES", "SpineOne", "MRO"];
const DAY_NAMES = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

function getTableClient() {
  return TableClient.fromConnectionString(
    process.env.AZURE_STORAGE_CONNECTION_STRING,
    TABLE_NAME
  );
}

function isValidDateString(s) {
  return /^\d{4}-\d{2}-\d{2}$/.test(s) && !Number.isNaN(Date.parse(s + "T12:00:00Z"));
}

function* eachDateInclusive(startDate, endDate) {
  const start = new Date(startDate + "T12:00:00Z");
  const end = new Date(endDate + "T12:00:00Z");
  const cur = new Date(start);
  while (cur <= end) {
    yield cur.toISOString().slice(0, 10);
    cur.setUTCDate(cur.getUTCDate() + 1);
  }
}

function dayOfWeek(yyyymmdd) {
  return new Date(yyyymmdd + "T12:00:00Z").getUTCDay();
}

function num(v) {
  const n = Number(v);
  return Number.isFinite(n) ? n : 0;
}

function formatSeconds(sec) {
  if (!sec) return "0:00";
  const m = Math.floor(sec / 60);
  const s = sec % 60;
  return `${m}:${String(s).padStart(2, "0")}`;
}

function emptyAggregate() {
  return {
    totalCalls: 0,
    answeredCalls: 0,
    abandonedCalls: 0,
    waitSecondsSum: 0,
    talkSecondsSum: 0,
    handleSecondsSum: 0,
    daysWithData: 0
  };
}

function addRecord(agg, rec) {
  agg.totalCalls     += num(rec.totalCalls);
  agg.answeredCalls  += num(rec.answeredCalls);
  agg.abandonedCalls += num(rec.abandonedCalls);
  agg.waitSecondsSum   += num(rec.avgWaitSeconds);
  agg.talkSecondsSum   += num(rec.avgTalkSeconds);
  agg.handleSecondsSum += num(rec.avgHandleSeconds);
  agg.daysWithData++;
}

function finalizeAggregate(agg) {
  const calls = agg.totalCalls;
  const days = agg.daysWithData;
  return {
    totalCalls: agg.totalCalls,
    answeredCalls: agg.answeredCalls,
    abandonedCalls: agg.abandonedCalls,
    answerRate: calls > 0 ? Number(((agg.answeredCalls / calls) * 100).toFixed(2)) : 0,
    abandonedRate: calls > 0 ? Number(((agg.abandonedCalls / calls) * 100).toFixed(2)) : 0,
    avgWaitSeconds: days > 0 ? Math.round(agg.waitSecondsSum / days) : 0,
    avgTalkSeconds: days > 0 ? Math.round(agg.talkSecondsSum / days) : 0,
    avgHandleSeconds: days > 0 ? Math.round(agg.handleSecondsSum / days) : 0,
    avgWaitFormatted: formatSeconds(days > 0 ? Math.round(agg.waitSecondsSum / days) : 0),
    avgTalkFormatted: formatSeconds(days > 0 ? Math.round(agg.talkSecondsSum / days) : 0),
    avgHandleFormatted: formatSeconds(days > 0 ? Math.round(agg.handleSecondsSum / days) : 0),
    daysWithData: days
  };
}

module.exports = async function (context, req) {
  try {
    const user = getUserFromRequest(req);
    const access = resolveAccess(user);
    const authError = requireAccess(access);
    if (authError) {
      context.res = { status: authError.status, headers: { "Content-Type": "application/json" }, body: authError.body };
      return;
    }

    const startDate = String(req.query?.startDate || "").trim();
    const endDate = String(req.query?.endDate || "").trim();
    const entityQuery = String(req.query?.entity || "").trim();

    if (!isValidDateString(startDate) || !isValidDateString(endDate)) {
      context.res = { status: 400, headers: { "Content-Type": "application/json" }, body: { ok: false, error: "Provide startDate and endDate as YYYY-MM-DD" } };
      return;
    }
    if (startDate > endDate) {
      context.res = { status: 400, headers: { "Content-Type": "application/json" }, body: { ok: false, error: "startDate must be <= endDate" } };
      return;
    }
    // Cap the range so a misclick can't pull years of data.
    const days = Math.round((Date.parse(endDate) - Date.parse(startDate)) / 86400000) + 1;
    if (days > 366) {
      context.res = { status: 400, headers: { "Content-Type": "application/json" }, body: { ok: false, error: "Range capped at 366 days" } };
      return;
    }

    let entities = scopeEntitiesToAccess(access, ALL_ENTITIES);
    if (entityQuery) {
      if (!canAccessEntity(access, entityQuery)) {
        context.res = { status: 404, headers: { "Content-Type": "application/json" }, body: { ok: false, error: "Not found" } };
        return;
      }
      entities = entities.filter((e) => e === entityQuery);
    }

    const client = getTableClient();
    const dateList = Array.from(eachDateInclusive(startDate, endDate));

    // entity → date → record
    const grid = {};
    const sources = {}; // entity → set of source labels seen

    for (const entity of entities) {
      grid[entity] = {};
      sources[entity] = new Set();
      for (const date of dateList) {
        try {
          const rec = await client.getEntity(entity, date);
          grid[entity][date] = rec;
          if (rec.source) sources[entity].add(String(rec.source));
        } catch (e) {
          if (e?.statusCode !== 404) throw e;
          // missing day — skip
        }
      }
    }

    // Top-line totals across all accessible entities
    const totalsAgg = emptyAggregate();
    for (const entity of entities) {
      for (const date of dateList) {
        const rec = grid[entity][date];
        if (rec) addRecord(totalsAgg, rec);
      }
    }

    // Per-entity totals
    const byEntity = entities.map((entity) => {
      const agg = emptyAggregate();
      for (const date of dateList) {
        const rec = grid[entity][date];
        if (rec) addRecord(agg, rec);
      }
      return {
        entity,
        sources: Array.from(sources[entity] || []),
        ...finalizeAggregate(agg)
      };
    });

    // Per-date totals (sum across entities for each date)
    const byDate = dateList.map((date) => {
      const agg = emptyAggregate();
      for (const entity of entities) {
        const rec = grid[entity][date];
        if (rec) addRecord(agg, rec);
      }
      return {
        date,
        dayOfWeek: DAY_NAMES[dayOfWeek(date)],
        ...finalizeAggregate(agg)
      };
    });

    // Day-of-week rollups (which days are busiest)
    const dowMap = {};
    for (const date of dateList) {
      const dow = dayOfWeek(date);
      if (!dowMap[dow]) dowMap[dow] = emptyAggregate();
      for (const entity of entities) {
        const rec = grid[entity][date];
        if (rec) addRecord(dowMap[dow], rec);
      }
    }
    const byDayOfWeek = DAY_NAMES.map((name, idx) => {
      const agg = dowMap[idx] || emptyAggregate();
      const finalized = finalizeAggregate(agg);
      return {
        dayOfWeek: name,
        dayIndex: idx,
        ...finalized,
        avgCallsPerDay: finalized.daysWithData > 0
          ? Math.round(finalized.totalCalls / finalized.daysWithData)
          : 0
      };
    });

    // Entity × date grid for stacked-bar / heatmap charts. Numbers only,
    // not per-cell aggregate objects, so the payload stays small.
    const perEntityByDate = entities.map((entity) => ({
      entity,
      data: dateList.map((date) => {
        const rec = grid[entity][date];
        return rec ? num(rec.totalCalls) : 0;
      })
    }));

    context.res = {
      status: 200,
      headers: { "Content-Type": "application/json" },
      body: {
        ok: true,
        startDate,
        endDate,
        days,
        entities,
        dateList,
        totals: finalizeAggregate(totalsAgg),
        byEntity,
        byDate,
        byDayOfWeek,
        perEntityByDate
      }
    };
  } catch (error) {
    return safeErrorResponse(context, error, "Failed to load call center report");
  }
};
