/**
 * Azure Function: call-data
 * GET  /api/call-data?entity=LAOSS&weekEnding=2026-04-12
 *   Returns aggregated call metrics for the week ending on that Friday
 *
 * Deploy to: /api/call-data/index.js
 */

const { TableClient } = require("@azure/data-tables");
const { getUserFromRequest } = require("../shared/auth");
const {
  resolveAccess,
  requireAccess,
  requireEntityAccess
} = require("../shared/permissions");

const TABLE_NAME = "CallData";
const ENTITIES = ["LAOSS", "NES", "SpineOne", "MRO"];

function getTableClient() {
  return TableClient.fromConnectionString(
    process.env.AZURE_STORAGE_CONNECTION_STRING, TABLE_NAME
  );
}

function getWeekDates(weekEnding) {
  // Return Mon–Fri dates for the week ending on weekEnding (Friday)
  const friday = new Date(weekEnding + "T12:00:00Z");
  const dates  = [];
  for (let i = 4; i >= 0; i--) {
    const d = new Date(friday);
    d.setUTCDate(d.getUTCDate() - i);
    dates.push(d.toISOString().slice(0, 10));
  }
  return dates; // [Mon, Tue, Wed, Thu, Fri]
}

module.exports = async function (context, req) {
  const respond = (status, body) => ({
    status,
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body)
  });

  const user = getUserFromRequest(req);
  const access = resolveAccess(user);

  const authError = requireAccess(access);
  if (authError) return respond(authError.status, authError.body);

  const entity     = String(req.query.entity     || "").trim();
  const weekEnding = String(req.query.weekEnding || "").trim();

  if (!entity || !weekEnding) {
    return respond(400, { ok: false, error: "Missing entity or weekEnding" });
  }

  if (!ENTITIES.includes(entity)) {
    return respond(400, { ok: false, error: "Invalid entity" });
  }

  if (!/^\d{4}-\d{2}-\d{2}$/.test(weekEnding)) {
    return respond(400, { ok: false, error: "Invalid weekEnding format (expected YYYY-MM-DD)" });
  }

  const entityError = requireEntityAccess(access, entity);
  if (entityError) return respond(entityError.status, entityError.body);

  try {
    const client = getTableClient();
    const dates  = getWeekDates(weekEnding);

    // Fetch all days in the week for this entity
    const records = [];
    for (const date of dates) {
      try {
        const rec = await client.getEntity(entity, date);
        records.push(rec);
      } catch (e) {
        if (e?.statusCode !== 404) throw e;
        // No data for this day — skip
      }
    }

    // Aggregate across days
    const agg = records.reduce((acc, r) => {
      acc.totalCalls     += Number(r.totalCalls     || 0);
      acc.answeredCalls  += Number(r.answeredCalls  || 0);
      acc.abandonedCalls += Number(r.abandonedCalls || 0);
      acc.totalWaitSec   += Number(r.avgWaitSeconds   || 0);
      acc.totalTalkSec   += Number(r.avgTalkSeconds   || 0);
      acc.totalHandleSec += Number(r.avgHandleSeconds || 0);
      acc.dayCount++;
      return acc;
    }, { totalCalls: 0, answeredCalls: 0, abandonedCalls: 0, totalWaitSec: 0, totalTalkSec: 0, totalHandleSec: 0, dayCount: 0 });

    const answerRate    = agg.totalCalls > 0 ? Number(((agg.answeredCalls / agg.totalCalls) * 100).toFixed(2)) : 0;
    const abandonedRate = agg.totalCalls > 0 ? Number(((agg.abandonedCalls / agg.totalCalls) * 100).toFixed(2)) : 0;
    const avgWait       = agg.dayCount > 0 ? Math.round(agg.totalWaitSec / agg.dayCount) : 0;
    const avgTalk       = agg.dayCount > 0 ? Math.round(agg.totalTalkSec / agg.dayCount) : 0;
    const avgHandle     = agg.dayCount > 0 ? Math.round(agg.totalHandleSec / agg.dayCount) : 0;

    return respond(200, {
      ok: true,
      entity,
      weekEnding,
      datesFound:      records.map(r => r.date),
      datesRequested:  dates,
      totalCalls:      agg.totalCalls,
      answeredCalls:   agg.answeredCalls,
      abandonedCalls:  agg.abandonedCalls,
      answerRate,
      abandonedRate,
      avgWaitSeconds:   avgWait,
      avgTalkSeconds:   avgTalk,
      avgHandleSeconds: avgHandle,
      // Formatted for display
      avgWaitFormatted:   formatSeconds(avgWait),
      avgTalkFormatted:   formatSeconds(avgTalk),
      avgHandleFormatted: formatSeconds(avgHandle),
      source:  records[0]?.source || "unknown",
      hasData: records.length > 0
    });
  } catch (error) {
    context.log.error("call-data error:", error.message);
    return respond(500, { ok: false, error: "Failed to load call data" });
  }
};

function formatSeconds(sec) {
  if (!sec) return "0:00";
  const m = Math.floor(sec / 60);
  const s = sec % 60;
  return `${m}:${String(s).padStart(2, "0")}`;
}
