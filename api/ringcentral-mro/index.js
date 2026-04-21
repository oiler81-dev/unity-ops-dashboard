/**
 * Azure Function: ringcentral-mro
 * Timer-triggered — runs every hour to pull RingCentral call logs for MRO (Chicago).
 * Writes normalized data to Azure Table Storage (CallData table).
 *
 * Deploy to: /api/ringcentral-mro/index.js
 *
 * Required App Settings:
 *   RC_MRO_CLIENT_ID       — RingCentral app client ID
 *   RC_MRO_CLIENT_SECRET   — RingCentral app client secret
 *   RC_MRO_USERNAME        — RingCentral account phone/email
 *   RC_MRO_PASSWORD        — RingCentral account password
 *   RC_MRO_EXTENSION       — RingCentral extension (optional, omit for main account)
 *   AZURE_STORAGE_CONNECTION_STRING
 */

const { TableClient } = require("@azure/data-tables");
const https = require("https");

const TABLE_NAME = "CallData";
const ENTITY     = "MRO";
const RC_SERVER  = "https://platform.ringcentral.com";

function getTableClient() {
  return TableClient.fromConnectionString(
    process.env.AZURE_STORAGE_CONNECTION_STRING, TABLE_NAME
  );
}

async function httpRequest(options, body) {
  return new Promise((resolve, reject) => {
    const req = https.request(options, res => {
      let data = "";
      res.on("data", chunk => data += chunk);
      res.on("end", () => {
        try { resolve({ status: res.statusCode, body: JSON.parse(data) }); }
        catch (e) { resolve({ status: res.statusCode, body: data }); }
      });
    });
    req.on("error", reject);
    if (body) req.write(body);
    req.end();
  });
}

async function getRcToken() {
  const creds = Buffer.from(
    `${process.env.RC_MRO_CLIENT_ID}:${process.env.RC_MRO_CLIENT_SECRET}`
  ).toString("base64");

  const body = new URLSearchParams({
    grant_type: "password",
    username:   process.env.RC_MRO_USERNAME,
    password:   process.env.RC_MRO_PASSWORD,
    extension:  process.env.RC_MRO_EXTENSION || ""
  }).toString();

  const url = new URL(`${RC_SERVER}/restapi/oauth/token`);
  const res = await httpRequest({
    hostname: url.hostname,
    path:     url.pathname,
    method:   "POST",
    headers: {
      "Authorization": `Basic ${creds}`,
      "Content-Type":  "application/x-www-form-urlencoded",
      "Content-Length": Buffer.byteLength(body)
    }
  }, body);

  if (res.status !== 200) throw new Error(`RC auth failed: ${JSON.stringify(res.body)}`);
  return res.body.access_token;
}

async function getCallLog(token, dateFrom, dateTo) {
  const params = new URLSearchParams({
    dateFrom, dateTo,
    view:    "Simple",
    perPage: "1000",
    type:    "Voice"
  });
  const url = new URL(`${RC_SERVER}/restapi/v1.0/account/~/call-log?${params}`);

  const res = await httpRequest({
    hostname: url.hostname,
    path:     url.pathname + url.search,
    method:   "GET",
    headers:  { "Authorization": `Bearer ${token}` }
  });

  if (res.status !== 200) throw new Error(`RC call log failed: ${JSON.stringify(res.body)}`);
  return res.body.records || [];
}

function aggregateByDate(records) {
  const byDate = {};

  for (const rec of records) {
    const date = rec.startTime?.slice(0, 10);
    if (!date) continue;

    if (!byDate[date]) {
      byDate[date] = {
        totalCalls: 0, answeredCalls: 0, abandonedCalls: 0,
        totalWaitMs: 0, totalTalkMs: 0, totalHandleMs: 0,
        waitCount: 0, talkCount: 0
      };
    }

    const d = byDate[date];
    d.totalCalls++;

    const result   = rec.result?.toLowerCase() || "";
    const duration = Number(rec.duration || 0) * 1000; // convert to ms

    if (result === "accepted" || result === "connected") {
      d.answeredCalls++;
      d.totalTalkMs   += duration;
      d.totalHandleMs += duration;
      d.talkCount++;
    } else if (result === "missed" || result === "voicemail" || result === "abandoned") {
      d.abandonedCalls++;
    }

    // Wait time from legs if available
    if (rec.legs) {
      for (const leg of rec.legs) {
        if (leg.legType === "Accept" && leg.startTime && leg.action === "HoldOn") {
          const waitMs = new Date(leg.endTime) - new Date(leg.startTime);
          d.totalWaitMs += Math.max(0, waitMs);
          d.waitCount++;
        }
      }
    }
  }

  return byDate;
}

async function saveToTable(client, dateAggs) {
  const saved = [];
  for (const [date, agg] of Object.entries(dateAggs)) {
    const avgWait   = agg.waitCount > 0 ? Math.round(agg.totalWaitMs / agg.waitCount / 1000) : 0;
    const avgTalk   = agg.talkCount > 0 ? Math.round(agg.totalTalkMs / agg.talkCount / 1000) : 0;
    const avgHandle = agg.talkCount > 0 ? Math.round(agg.totalHandleMs / agg.talkCount / 1000) : 0;

    const entity = {
      partitionKey:     ENTITY,
      rowKey:           date,
      entity:           ENTITY,
      date,
      source:           "ringcentral",
      totalCalls:       agg.totalCalls,
      answeredCalls:    agg.answeredCalls,
      abandonedCalls:   agg.abandonedCalls,
      answerRate:       agg.totalCalls > 0 ? Number(((agg.answeredCalls / agg.totalCalls) * 100).toFixed(2)) : 0,
      abandonedRate:    agg.totalCalls > 0 ? Number(((agg.abandonedCalls / agg.totalCalls) * 100).toFixed(2)) : 0,
      avgWaitSeconds:   avgWait,
      avgTalkSeconds:   avgTalk,
      avgHandleSeconds: avgHandle,
      updatedAt:        new Date().toISOString()
    };

    await client.upsertEntity(entity, "Replace");
    saved.push({ date, totalCalls: entity.totalCalls, answerRate: entity.answerRate });
  }
  return saved;
}

module.exports = async function (context) {
  context.log("RingCentral MRO sync started");

  try {
    const client = getTableClient();
    await client.createTable().catch(e => { if (e?.statusCode !== 409) throw e; });

    // Pull last 8 days to catch any gaps
    const now      = new Date();
    const dateTo   = now.toISOString();
    const dateFrom = new Date(now - 8 * 24 * 60 * 60 * 1000).toISOString();

    const token   = await getRcToken();
    const records = await getCallLog(token, dateFrom, dateTo);
    context.log(`RingCentral MRO: pulled ${records.length} call records`);

    const aggs  = aggregateByDate(records);
    const saved = await saveToTable(client, aggs);
    context.log("RingCentral MRO: saved", saved);

  } catch (error) {
    context.log.error("RingCentral MRO sync error:", error.message);
  }
};
