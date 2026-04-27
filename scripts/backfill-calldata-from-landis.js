/**
 * One-shot backfill: overwrite CallData rows for LAOSS + NES with the
 * weekly totals from Landis's QueueSummaryExport (the version you sent
 * 2026-04-27, 3:22 PM). Necessary because the webhook misattributed
 * cross-coverage calls between LAOSS and NES — historical rows can't be
 * replayed (no raw payload log) so we restore truth from Landis's own
 * report.
 *
 * Strategy:
 *   - Mon–Thu of the target week → zero-fill (clears the bad attribution)
 *   - Fri (week-ending) → write the full weekly total per entity
 * The dashboard's weekly aggregate sums Mon–Fri so the total comes out
 * correct. Daily granularity for that week is sacrificed — only the
 * weekly total matters for the report tomorrow.
 *
 * Usage (from repo/unity-ops-dashboard/api/):
 *   set AZURE_STORAGE_CONNECTION_STRING=...   # admin connection string
 *   node ../scripts/backfill-calldata-from-landis.js --week 2026-04-24
 *   # review the dry-run summary, then:
 *   node ../scripts/backfill-calldata-from-landis.js --week 2026-04-24 --commit
 *
 * Only LAOSS and NES are touched. SpineOne and MRO come from the
 * RingCentral ingest path which is unaffected by the Landis bug.
 */

const { TableClient } = require("@azure/data-tables");

// ───────────────────────────────────────────────────────────────────
// LANDIS TRUTH — extracted from QueueSummaryExport_4-27-2026,3-22-07 PM
// (period: Mon 2026-04-20 through Fri 2026-04-24 per Nestor's confirmation)
// ───────────────────────────────────────────────────────────────────
//
// Each row: [queueName, total, answered, abandoned, talkSec, waitSec]
// "abandoned" here means "not answered" (= total − answered) so it
// includes Landis's Abandoned + Queue Timeout columns. The dashboard's
// abandoned-rate metric is really a not-answered rate, so this matches
// the model. Times converted from HH:MM:SS to seconds.
const LANDIS_QUEUES = [
  ["Landis-NES-Locust",                225,   172,  53,  247, 135],
  ["Landis-LAOSS-Santa Fe Springs",    928,   743, 185,  222, 226],
  ["Landis-LAOSS-WC",                 1274,   920, 354,  201, 318],
  ["Landis-LAOSS-East LA",            2296,  1865, 431,  220, 130],
  ["Landis-NES-Newberg Front Desk",     90,    76,  14,  199, 127],
  ["Landis-LAOSS-Tarzana",             901,   775, 126,  224, 106],
  ["Landis-LAOSS-Acevedo",              59,     0,  59,    0,   1],
  ["Landis-NES-CM-HT",                 460,   412,  48,  190, 154],
  ["Landis-NES-New Patients",          485,   244, 241,  335, 141],
  ["Landis-NES-Tigard Front Desk",     411,   347,  64,  231, 148],
  ["Landis-NES-West Linn Front Desk",   69,    61,   8,  223, 145],
  ["Landis-LAOSS-Glendale",            581,   515,  66,  210, 100],
  ["Landis-LAOSS-Valencia",            191,   163,  28,  207,  71],
  ["Landis-LAOSS-Referrals",           435,   426,   9,  296,  44],
  ["Landis-LAOSS-Sx",                  243,   131, 112,  300, 185],
  ["Landis-LAOSS-Wilshire",            780,   637, 143,  257, 174],
  ["Landis-NES-Milwaukie FrontDesk",   167,   143,  24,  226, 137],
  ["Landis-NES-Wilsonville Front Desk", 70,    57,  13,  152, 203],
  ["LAOSS-Tarzana-PPO",                 47,    42,   5,  264,  96],
  ["Landis-LAOSS-Disability",            1,     0,   1,    0,   1],
  ["Landis-NES- Urgent",                 1,     1,   0,   89, 145]
];

// Same word-boundary entity matcher as the webhook resolver.
const ENTITY_TOKENS = {
  LAOSS: ["LAOSS", "laorthos.com", "laorthos.org"],
  NES:   ["NES", "nespecialists.com", "NES OU"]
};
function classify(name) {
  const hayLower = String(name).toLowerCase();
  for (const [entity, tokens] of Object.entries(ENTITY_TOKENS)) {
    for (const token of tokens) {
      const re = new RegExp(`(^|[^a-z0-9])${token.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}($|[^a-z0-9])`, "i");
      if (re.test(name)) return entity;
      if (hayLower.includes(token.toLowerCase())) return entity;
    }
  }
  return null;
}

function aggregate(queues) {
  const out = {};
  for (const [name, total, ans, aban, talk, wait] of queues) {
    const entity = classify(name);
    if (!entity) {
      console.warn(`SKIP unclassified queue: ${name}`);
      continue;
    }
    if (!out[entity]) {
      out[entity] = { totalCalls: 0, answeredCalls: 0, abandonedCalls: 0, weightedTalk: 0, weightedWait: 0 };
    }
    const r = out[entity];
    r.totalCalls     += total;
    r.answeredCalls  += ans;
    r.abandonedCalls += aban;
    r.weightedTalk   += total * talk;
    r.weightedWait   += total * wait;
  }
  for (const e of Object.keys(out)) {
    const r = out[e];
    r.avgTalkSeconds = r.totalCalls > 0 ? Math.round(r.weightedTalk / r.totalCalls) : 0;
    r.avgWaitSeconds = r.totalCalls > 0 ? Math.round(r.weightedWait / r.totalCalls) : 0;
    delete r.weightedTalk;
    delete r.weightedWait;
  }
  return out;
}

function getWeekDates(weekEnding) {
  const fri = new Date(weekEnding + "T12:00:00Z");
  const out = [];
  for (let i = 4; i >= 0; i--) {
    const d = new Date(fri);
    d.setUTCDate(d.getUTCDate() - i);
    out.push(d.toISOString().slice(0, 10));
  }
  return out; // [Mon, Tue, Wed, Thu, Fri]
}

async function main() {
  const args = process.argv.slice(2);
  const weekIdx = args.indexOf("--week");
  const commit  = args.includes("--commit");
  const weekEnding = weekIdx >= 0 ? args[weekIdx + 1] : null;

  if (!weekEnding || !/^\d{4}-\d{2}-\d{2}$/.test(weekEnding)) {
    console.error("usage: node backfill-calldata-from-landis.js --week YYYY-MM-DD [--commit]");
    process.exit(1);
  }
  const conn = process.env.AZURE_STORAGE_CONNECTION_STRING;
  if (!conn) {
    console.error("AZURE_STORAGE_CONNECTION_STRING must be set");
    process.exit(1);
  }

  const totals = aggregate(LANDIS_QUEUES);
  const dates  = getWeekDates(weekEnding);
  const fri    = dates[dates.length - 1];

  console.log("\nLandis truth from QueueSummaryExport (period to overwrite):");
  console.log(JSON.stringify(totals, null, 2));
  console.log(`\nTarget week: Mon ${dates[0]} → Fri ${fri}`);
  console.log("Mon–Thu rows will be ZERO-FILLED, Fri row will receive the weekly totals.\n");

  const client = TableClient.fromConnectionString(conn, "CallData");
  await client.createTable().catch(e => { if (e?.statusCode !== 409) throw e; });

  const writes = [];
  for (const entity of Object.keys(totals)) {
    for (const date of dates) {
      const isFri = date === fri;
      const r = totals[entity];
      const row = {
        partitionKey: entity,
        rowKey: date,
        entity,
        date,
        source: "landis",
        totalCalls:       isFri ? r.totalCalls     : 0,
        answeredCalls:    isFri ? r.answeredCalls  : 0,
        abandonedCalls:   isFri ? r.abandonedCalls : 0,
        avgWaitSeconds:   isFri ? r.avgWaitSeconds : 0,
        avgTalkSeconds:   isFri ? r.avgTalkSeconds : 0,
        avgHandleSeconds: 0,
        // Wipe per-queue breakdowns so any future re-aggregation starts fresh.
        queueTotalsJson:        JSON.stringify({}),
        queueAnsweredJson:      JSON.stringify({}),
        queueAbandonedJson:     JSON.stringify({}),
        queueWaitSecondsJson:   JSON.stringify({}),
        queueHandleSecondsJson: JSON.stringify({}),
        queueCount:             0,
        answerRate:    isFri && r.totalCalls > 0 ? Number(((r.answeredCalls / r.totalCalls) * 100).toFixed(2)) : 0,
        abandonedRate: isFri && r.totalCalls > 0 ? Number(((r.abandonedCalls / r.totalCalls) * 100).toFixed(2)) : 0,
        backfilledFrom: "landis-queuesummary-export-2026-04-27",
        backfilledAt: new Date().toISOString()
      };
      writes.push(row);
    }
  }

  console.log("Planned writes:");
  for (const w of writes) {
    console.log(`  ${w.entity}/${w.date}  total=${w.totalCalls} ans=${w.answeredCalls} aban=${w.abandonedCalls} wait=${w.avgWaitSeconds}s talk=${w.avgTalkSeconds}s`);
  }

  if (!commit) {
    console.log("\nDRY RUN — pass --commit to apply.");
    return;
  }

  console.log("\nApplying writes...");
  for (const w of writes) {
    await client.upsertEntity(w, "Replace");
    console.log(`  ✓ ${w.entity}/${w.date}`);
  }
  console.log("\nDone. Refresh the dashboard to confirm.");
}

main().catch((e) => { console.error(e); process.exit(1); });
