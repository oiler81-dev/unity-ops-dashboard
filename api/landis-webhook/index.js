/**
 * Azure Function: landis-webhook
 * Receives Landis Contact Center webhook events and writes
 * normalized call data to Azure Table Storage (CallData table).
 *
 * Deploy to: /api/landis-webhook/index.js
 *
 * In Landis:
 *   Webhook Endpoint: https://your-app.azurestaticapps.net/api/landis-webhook
 *   Events: App.QueueWebhook, App.AgentWebhook, App.IvrWebhook
 *   Header Key: x-landis-secret  |  Header Value: (set in LANDIS_WEBHOOK_SECRET env var)
 */

const { TableClient, AzureNamedKeyCredential } = require("@azure/data-tables");
const crypto = require("crypto");

function timingSafeStringCompare(a, b) {
  const ab = Buffer.from(String(a || ""), "utf8");
  const bb = Buffer.from(String(b || ""), "utf8");
  if (ab.length !== bb.length) return false;
  return crypto.timingSafeEqual(ab, bb);
}

const TABLE_NAME   = "CallData";
const ENTITY_MAP   = {
  // Map Landis queue/team names to your entity keys
  // Update these to match your actual Landis queue names
  "LAOSS":         "LAOSS",
  "LA Ortho":      "LAOSS",
  "Los Angeles":   "LAOSS",
  "NES":           "NES",
  "Portland":      "NES",
  "NE Spine":      "NES"
};

function getTableClient() {
  const connStr = process.env.AZURE_STORAGE_CONNECTION_STRING;
  if (!connStr) throw new Error("AZURE_STORAGE_CONNECTION_STRING not set");
  return TableClient.fromConnectionString(connStr, TABLE_NAME);
}

function toDateKey(dateStr) {
  // Normalize to YYYY-MM-DD
  if (!dateStr) return new Date().toISOString().slice(0, 10);
  return new Date(dateStr).toISOString().slice(0, 10);
}

function resolveEntity(payload) {
  // Try queue name, team name, or campaign name. Match exactly (case-insensitive)
  // against ENTITY_MAP keys — substring/loose matching let a short value like
  // "n" match "NES".
  const candidates = [
    payload.QueueName, payload.TeamName, payload.CampaignName,
    payload.queueName, payload.teamName, payload.Location,
    payload.location, payload.Site, payload.site
  ];
  const lowerMap = Object.fromEntries(
    Object.entries(ENTITY_MAP).map(([k, v]) => [k.toLowerCase(), v])
  );
  for (const c of candidates) {
    if (!c) continue;
    const normalized = String(c).trim().toLowerCase();
    if (normalized && lowerMap[normalized]) return lowerMap[normalized];
  }
  return null;
}

async function upsertCallData(client, entity, dateKey, updates) {
  // Read existing record first so we can merge, not overwrite
  let existing = {};
  try {
    const rec = await client.getEntity(entity, dateKey);
    existing = rec;
  } catch (e) {
    if (e?.statusCode !== 404) throw e;
  }

  const merged = {
    partitionKey:     entity,
    rowKey:           dateKey,
    entity,
    date:             dateKey,
    source:           "landis",
    totalCalls:       Math.max(Number(existing.totalCalls || 0), Number(updates.totalCalls || 0)),
    answeredCalls:    Math.max(Number(existing.answeredCalls || 0), Number(updates.answeredCalls || 0)),
    abandonedCalls:   Math.max(Number(existing.abandonedCalls || 0), Number(updates.abandonedCalls || 0)),
    avgWaitSeconds:   updates.avgWaitSeconds   ?? existing.avgWaitSeconds   ?? 0,
    avgTalkSeconds:   updates.avgTalkSeconds   ?? existing.avgTalkSeconds   ?? 0,
    avgHandleSeconds: updates.avgHandleSeconds ?? existing.avgHandleSeconds ?? 0,
    updatedAt:        new Date().toISOString()
  };

  // Compute derived rates
  merged.answerRate    = merged.totalCalls > 0
    ? Number(((merged.answeredCalls / merged.totalCalls) * 100).toFixed(2))
    : 0;
  merged.abandonedRate = merged.totalCalls > 0
    ? Number(((merged.abandonedCalls / merged.totalCalls) * 100).toFixed(2))
    : 0;

  await client.upsertEntity(merged, "Merge");
  return merged;
}

function parseQueueWebhook(payload) {
  // App.QueueWebhook — queue-level aggregates
  return {
    totalCalls:       Number(payload.TotalCalls       || payload.totalCalls       || 0),
    answeredCalls:    Number(payload.AnsweredCalls     || payload.answeredCalls    || 0),
    abandonedCalls:   Number(payload.AbandonedCalls    || payload.abandonedCalls   || 0),
    avgWaitSeconds:   Number(payload.AvgWaitTime       || payload.avgWaitTime      || payload.AvgWaitSeconds || 0),
    avgHandleSeconds: Number(payload.AvgHandleTime     || payload.avgHandleTime    || payload.AvgHandleSeconds || 0)
  };
}

function parseAgentWebhook(payload) {
  // App.AgentWebhook — agent-level, aggregate talk time
  return {
    avgTalkSeconds: Number(payload.AvgTalkTime || payload.avgTalkTime || payload.TalkTime || 0)
  };
}

module.exports = async function (context, req) {
  // Validate webhook secret. Fail closed when the env var is missing or empty so
  // a misconfiguration can't silently turn the endpoint into an open write.
  const secret = process.env.LANDIS_WEBHOOK_SECRET;
  if (!secret) {
    context.log.error("Landis webhook: LANDIS_WEBHOOK_SECRET not configured — rejecting request");
    return { status: 500, body: { ok: false, error: "Webhook not configured" } };
  }
  const incoming = req.headers["x-landis-secret"] || req.headers["x-webhook-secret"] || "";
  if (!timingSafeStringCompare(incoming, secret)) {
    context.log.warn("Landis webhook: invalid secret");
    return { status: 401, body: { ok: false, error: "Unauthorized" } };
  }

  const body = req.body || {};
  const eventType = body.EventType || body.eventType || body.Event || body.event || "";
  // Log event type only — payload bodies can include caller identifiers (PHI-adjacent).
  context.log("Landis webhook received:", eventType);

  try {
    const client = getTableClient();
    await client.createTable().catch(e => { if (e?.statusCode !== 409) throw e; });

    const entity = resolveEntity(body);
    if (!entity) {
      // Don't log raw body — may contain caller identifiers. Log event type only.
      context.log.warn("Landis webhook: could not resolve entity", eventType);
      return { status: 200, body: { ok: true, skipped: true, reason: "entity not resolved" } };
    }

    const dateKey  = toDateKey(body.Date || body.date || body.Timestamp || body.timestamp);
    let updates    = {};

    const evtLower = eventType.toLowerCase();
    if (evtLower.includes("queue")) {
      updates = parseQueueWebhook(body);
    } else if (evtLower.includes("agent")) {
      updates = parseAgentWebhook(body);
    } else if (evtLower.includes("ivr")) {
      // IVR — just count the call
      updates = { totalCalls: 1 };
    } else {
      context.log("Landis webhook: unknown event type", eventType);
      return { status: 200, body: { ok: true, skipped: true, reason: "unhandled event type" } };
    }

    const merged = await upsertCallData(client, entity, dateKey, updates);
    context.log("Landis webhook: saved", entity, dateKey);

    return {
      status: 200,
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ ok: true, entity, date: dateKey, saved: merged })
    };
  } catch (error) {
    context.log.error("Landis webhook error:", error.message);
    return {
      status: 500,
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ ok: false, error: "Webhook processing failed" })
    };
  }
};
