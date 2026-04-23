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

const TABLE_NAME = "CallData";

// Known entity tokens that appear inside Landis queue names, OU paths, interaction
// object names, and email domains. Each token is checked as a word-boundary match
// (bounded by non-alphanumerics) so short tokens like "NES" never match inside
// unrelated words.
const ENTITY_TOKENS = {
  LAOSS: ["LAOSS", "laorthos.com", "laorthos.org"],
  NES:   ["NES", "nespecialists.com", "NES OU"]
  // SpineOne and MRO are not yet served by Landis — add tokens here when onboarded.
};

function tokenMatchesEntity(haystack) {
  if (!haystack) return null;
  const hay = String(haystack);
  for (const [entity, tokens] of Object.entries(ENTITY_TOKENS)) {
    for (const token of tokens) {
      // Word-boundary: token must be surrounded by start/end or non-alphanumeric chars.
      const re = new RegExp(`(^|[^a-z0-9])${token.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}($|[^a-z0-9])`, "i");
      if (re.test(hay)) return entity;
    }
  }
  return null;
}

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
  // Landis naming: "Landis-LAOSS-Tarzana", OUPath "Default\\Landis-LAOSS-WC" or
  // "NES OU\\Landis-NES-New Patients", agent UPNs "*@laorthos.com" / "*@nespecialists.com".
  // Check each candidate with a word-boundary token match.
  const candidates = [
    payload.QueueName, payload.TeamName, payload.CampaignName,
    payload.queueName, payload.teamName, payload.Location,
    payload.location, payload.Site, payload.site,
    payload.OUPath, payload.ouPath,
    payload.InteractionObjectName, payload.interactionObjectName,
    // Agent/caller UPNs as last-resort fallback — domain identifies the entity.
    payload.AgentUpn, payload.agentUpn,
    payload.CallerUpn, payload.callerUpn
  ];
  for (const c of candidates) {
    const match = tokenMatchesEntity(c);
    if (match) return match;
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

  // Queue webhooks send the running daily aggregate — take max.
  // IVR/Agent events are per-call, so they carry { incrementCalls: 1 }
  // (or similar) and we ADD instead of max'ing. The prior code used
  // Math.max(existing, 1) for IVR events which meant the counter
  // never moved past 1 after the first event of the day.
  const base = {
    totalCalls:     Number(existing.totalCalls     || 0),
    answeredCalls:  Number(existing.answeredCalls  || 0),
    abandonedCalls: Number(existing.abandonedCalls || 0)
  };
  const merged = {
    partitionKey:     entity,
    rowKey:           dateKey,
    entity,
    date:             dateKey,
    source:           "landis",
    totalCalls:       base.totalCalls     + Number(updates.incrementCalls     || 0) +
                        Math.max(0, Number(updates.totalCalls     || 0) - base.totalCalls),
    answeredCalls:    base.answeredCalls  + Number(updates.incrementAnswered  || 0) +
                        Math.max(0, Number(updates.answeredCalls  || 0) - base.answeredCalls),
    abandonedCalls:   base.abandonedCalls + Number(updates.incrementAbandoned || 0) +
                        Math.max(0, Number(updates.abandonedCalls || 0) - base.abandonedCalls),
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
      // Log the routing-relevant fields (queue/ou/site names) so we can
      // diagnose why Landis events aren't being matched to entities.
      // These fields shouldn't carry caller PHI — they're config labels.
      context.log.warn("Landis webhook: unresolved entity", {
        eventType,
        queueName: body.QueueName || body.queueName,
        teamName: body.TeamName || body.teamName,
        location: body.Location || body.location,
        site: body.Site || body.site,
        ouPath: body.OUPath || body.ouPath,
        interactionObjectName: body.InteractionObjectName || body.interactionObjectName
      });
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
      // IVR — one call per event. incrementCalls avoids the Math.max(1,1)
      // bug where the daily total got stuck at 1 no matter how many
      // IVR events came in.
      updates = { incrementCalls: 1 };
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
