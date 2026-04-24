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

// Landis Contact Center's webhook header editor saves the value with
// literal surrounding quotation marks (visible in the admin UI as
// `"x-landis-secret"` : `"<token>"`). Their server then transmits the
// value still wrapped in quotes, so a byte-for-byte compare against our
// quote-less env-var fails. Strip a single pair of surrounding ASCII
// quotes (single or double) so the secret matches regardless of how the
// admin entered it.
function unquoteHeader(v) {
  if (v == null) return "";
  const s = String(v).trim();
  if (s.length >= 2) {
    const first = s[0];
    const last = s[s.length - 1];
    if ((first === '"' || first === "'") && first === last) {
      return s.slice(1, -1);
    }
  }
  return s;
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
  const hayLower = hay.toLowerCase();
  for (const [entity, tokens] of Object.entries(ENTITY_TOKENS)) {
    for (const token of tokens) {
      // First try a word-boundary match (preferred; safer for short tokens).
      const re = new RegExp(`(^|[^a-z0-9])${token.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}($|[^a-z0-9])`, "i");
      if (re.test(hay)) return entity;
      // Fallback: simple case-insensitive substring. Landis queue/OU names
      // sometimes contain unusual characters around the token (e.g. tabs,
      // unicode dashes, NBSPs) that defeat the word-boundary regex. The
      // tokens we use ("LAOSS", "NES", "laorthos.com", "nespecialists.com")
      // are unique enough that a substring match is safe.
      if (hayLower.includes(token.toLowerCase())) return entity;
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
  // Lookup tries both the unquoted and quoted header NAME variants because
  // Landis admin UI saves the key with surrounding quotes; some HTTP stacks
  // strip them before sending, others don't. Then strip surrounding quotes
  // from the VALUE before comparing — same Landis admin oddity.
  const incoming = unquoteHeader(
    req.headers["x-landis-secret"]
    || req.headers['"x-landis-secret"']
    || req.headers["x-webhook-secret"]
    || ""
  );
  if (!timingSafeStringCompare(incoming, secret)) {
    context.log.warn("Landis webhook: invalid secret", { headerKeys: Object.keys(req.headers || {}).filter(k => /landis|secret/i.test(k)) });
    return { status: 401, body: { ok: false, error: "Unauthorized" } };
  }

  const envelope = req.body || {};
  // Real Landis webhooks ship with a wrapper:
  //   { Id, WebhookEventUtc, Attempt, Data: { ...actual call fields... }, CreationTimeUtc }
  // The call fields (CallId, QueueName, OUPath, AgentUpn, TotalCalls, ...)
  // live INSIDE Data, not at the top level. Unwrap before routing/parsing.
  // Fallback to the envelope itself for older/test payloads that send fields flat.
  const data = (envelope && typeof envelope.Data === "object" && envelope.Data !== null)
    ? envelope.Data
    : (envelope && typeof envelope.data === "object" && envelope.data !== null)
      ? envelope.data
      : envelope;
  const eventType = envelope.EventType || envelope.eventType || envelope.Event || envelope.event
    || data.EventType || data.eventType || data.Event || data.event || "";
  // Log event type only — payload bodies can include caller identifiers (PHI-adjacent).
  context.log("Landis webhook received:", eventType, "Attempt:", envelope.Attempt);

  try {
    const client = getTableClient();
    await client.createTable().catch(e => { if (e?.statusCode !== 409) throw e; });

    const entity = resolveEntity(data);
    if (!entity) {
      // Echo back keys at BOTH layers so misshapen payloads are easy to debug
      // from Landis's "Webhook Send Attempts" view without server-side logs.
      const debug = {
        envelopeKeys: Object.keys(envelope || {}),
        dataKeys: Object.keys(data || {}),
        eventType,
        queueName: data.QueueName || data.queueName,
        teamName: data.TeamName || data.teamName,
        location: data.Location || data.location,
        site: data.Site || data.site,
        ouPath: data.OUPath || data.ouPath,
        interactionObjectName: data.InteractionObjectName || data.interactionObjectName,
        agentUpn: data.AgentUpn || data.agentUpn,
        agentName: data.AgentName || data.agentName
      };
      context.log.warn("Landis webhook: unresolved entity", debug);
      return { status: 200, body: { ok: true, skipped: true, reason: "entity not resolved", debug } };
    }

    // Landis events carry timestamps under a few possible names. Look at both
    // the envelope (CreationTimeUtc, WebhookEventUtc) and the inner data
    // payload (StartDateTime, EndDateTime, etc.).
    const dateKey = toDateKey(
      data.Date || data.date
      || data.Timestamp || data.timestamp
      || data.StartDateTime || data.startDateTime
      || data.EndDateTime || data.endDateTime
      || data.PeriodStart || data.periodStart
      || envelope.CreationTimeUtc || envelope.creationTimeUtc
      || envelope.WebhookEventUtc || envelope.webhookEventUtc
    );
    let updates    = {};

    const evtLower = eventType.toLowerCase();
    // Landis doesn't put EventType in the inner Data; the event NAME lives
    // only on the subscription. Infer the shape from fields actually present:
    //   TotalCalls/AnsweredCalls => queue aggregate
    //   AvgTalkTime alone        => agent rollup
    //   CallId on its own        => single per-call event (treat as IVR/inbound)
    const looksQueue = evtLower.includes("queue")
      || data.TotalCalls != null || data.totalCalls != null
      || data.AnsweredCalls != null || data.answeredCalls != null
      || data.AbandonedCalls != null || data.abandonedCalls != null;
    const looksAgent = evtLower.includes("agent") || (!looksQueue && (data.AvgTalkTime != null || data.avgTalkTime != null));
    const looksIvr = evtLower.includes("ivr") || (!looksQueue && !looksAgent && (data.CallId != null || data.callId != null));

    if (looksQueue) {
      // Queue events carry the running daily totals — Math.max keeps the
      // highest seen value so retries and out-of-order events are safe.
      updates = parseQueueWebhook(data);
    } else if (looksAgent) {
      // Agent events update talk-time averages but not the call counters
      // (those would double-count against queue aggregates).
      updates = parseAgentWebhook(data);
    } else if (looksIvr) {
      // Per-call IVR events. We DON'T increment totalCalls here — the
      // QueueWebhook stream already carries the running daily total for
      // every queue, so adding +1 per IVR event would double-count.
      // Treat these as a no-op heartbeat for now.
      return { status: 200, body: { ok: true, skipped: true, reason: "per-call ivr event (counted via queue aggregate)", callId: data.CallId } };
    } else {
      context.log("Landis webhook: unknown event shape", { eventType, dataKeys: Object.keys(data || {}) });
      return { status: 200, body: { ok: true, skipped: true, reason: "unhandled event shape", dataKeys: Object.keys(data || {}) } };
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
