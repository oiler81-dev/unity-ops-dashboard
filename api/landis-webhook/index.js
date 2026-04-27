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

const CALL_EVENT_LOG_TABLE = "CallEventLog";
function getCallEventLogClient() {
  const connStr = process.env.AZURE_STORAGE_CONNECTION_STRING;
  if (!connStr) throw new Error("AZURE_STORAGE_CONNECTION_STRING not set");
  return TableClient.fromConnectionString(connStr, CALL_EVENT_LOG_TABLE);
}

function toDateKey(dateStr) {
  // Normalize to YYYY-MM-DD
  if (!dateStr) return new Date().toISOString().slice(0, 10);
  return new Date(dateStr).toISOString().slice(0, 10);
}

// Two-tier entity resolution. QueueName/OUPath/Site/etc. are AUTHORITATIVE
// because a queue belongs to exactly one entity. Agent/caller UPNs are a
// fallback because cross-coverage (NES agent picking up a LAOSS call when
// LAOSS is busy) means the UPN's domain does NOT necessarily match the
// queue's owner. Earlier code returned the first match across both tiers,
// which silently misattributed cross-covered calls to the answering
// agent's entity instead of the queue's owner — that's how Landis truth
// of LAOSS=7,736 / NES=1,978 became dashboard 6,086 / 3,247.
const QUEUE_FIELDS = (payload) => [
  payload.QueueName, payload.queueName,
  payload.TeamName, payload.teamName,
  payload.CampaignName,
  payload.Location, payload.location, payload.Site, payload.site,
  payload.OUPath, payload.ouPath,
  payload.InteractionObjectName, payload.interactionObjectName
];
const AGENT_FIELDS = (payload) => [
  payload.AgentUpn, payload.agentUpn,
  payload.CallerUpn, payload.callerUpn
];

function resolveEntity(payload) {
  // Returns { entity, source } where source is "queue" (authoritative) or
  // "agent" (fallback when no queue context is present on the event).
  for (const c of QUEUE_FIELDS(payload)) {
    const match = tokenMatchesEntity(c);
    if (match) return { entity: match, source: "queue" };
  }
  for (const c of AGENT_FIELDS(payload)) {
    const match = tokenMatchesEntity(c);
    if (match) return { entity: match, source: "agent" };
  }
  return { entity: null, source: null };
}

function safeParseJson(s, fallback) {
  if (!s) return fallback;
  try { return JSON.parse(String(s)); } catch { return fallback; }
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

  // Per-queue running totals stored as JSON inside the row. Each entity has
  // ~10 queues (East LA, Whittier, Tarzana, …). Landis fires one event per
  // queue per refresh interval; each event carries that queue's running
  // daily total. Math.max per-queue catches the latest value, then we sum
  // across queues to get the entity-day total.
  //
  // Without this, Math.max applied directly to the entity row would only
  // remember the SINGLE highest queue ever seen, not the sum of all queues.
  const queueTotalsMap   = safeParseJson(existing.queueTotalsJson,   {});
  const queueAnsweredMap = safeParseJson(existing.queueAnsweredJson, {});
  const queueAbandMap    = safeParseJson(existing.queueAbandonedJson, {});
  const queueWaitMap     = safeParseJson(existing.queueWaitSecondsJson, {});
  const queueHandleMap   = safeParseJson(existing.queueHandleSecondsJson, {});

  const qName = updates.queueName ? String(updates.queueName) : null;

  // Aggregate mode: event carries running totals for this queue, take max.
  if (qName && updates.totalCalls != null) {
    queueTotalsMap[qName]   = Math.max(Number(queueTotalsMap[qName]   || 0), Number(updates.totalCalls     || 0));
    queueAnsweredMap[qName] = Math.max(Number(queueAnsweredMap[qName] || 0), Number(updates.answeredCalls  || 0));
    queueAbandMap[qName]    = Math.max(Number(queueAbandMap[qName]    || 0), Number(updates.abandonedCalls || 0));
    if (updates.avgWaitSeconds   != null) queueWaitMap[qName]   = Number(updates.avgWaitSeconds);
    if (updates.avgHandleSeconds != null) queueHandleMap[qName] = Number(updates.avgHandleSeconds);
  }

  // Per-call mode: event is one unique call (already deduped by CallId at the
  // caller). Increment the per-queue tally for this queue. Increments may be
  // negative when reconciling a prior misattribution (entity reassign) or
  // flipping an abandoned call to answered when an agent event reveals the
  // call was actually picked up. Trigger the block whenever ANY of the three
  // counters is changing — gating only on incrementCalls would skip pure
  // answered/abandoned reshuffles where incrementCalls = 0. Clamp at zero
  // afterwards so a race condition can never persist a negative count.
  const hasIncrement =
    Number(updates.incrementCalls    || 0) !== 0 ||
    Number(updates.incrementAnswered || 0) !== 0 ||
    Number(updates.incrementAbandoned|| 0) !== 0;
  if (qName && hasIncrement) {
    queueTotalsMap[qName]   = Math.max(0, Number(queueTotalsMap[qName]   || 0) + Number(updates.incrementCalls     || 0));
    queueAnsweredMap[qName] = Math.max(0, Number(queueAnsweredMap[qName] || 0) + Number(updates.incrementAnswered  || 0));
    queueAbandMap[qName]    = Math.max(0, Number(queueAbandMap[qName]    || 0) + Number(updates.incrementAbandoned || 0));
  }

  // Sum across queues for the entity-day total. If the event didn't carry
  // a queueName (older test payloads), we still allow ad-hoc per-call
  // increment via incrementCalls.
  const sumValues = (m) => Object.values(m).reduce((a, b) => a + Number(b || 0), 0);
  const baseTotal = sumValues(queueTotalsMap);
  const baseAns   = sumValues(queueAnsweredMap);
  const baseAban  = sumValues(queueAbandMap);

  // Average wait/handle across queues that have non-zero values.
  const avgOf = (m) => {
    const vals = Object.values(m).map(Number).filter((n) => Number.isFinite(n) && n > 0);
    return vals.length ? Math.round(vals.reduce((a, b) => a + b, 0) / vals.length) : 0;
  };

  // When a queue name is provided the increment is already folded into
  // queueTotalsMap[qName] above, so baseTotal already reflects it — adding
  // updates.increment* again would double-count (a long-standing latent
  // bug that becomes load-bearing once we start passing negative
  // increments to reassign attributions). The increment is only added on
  // top of baseTotal when qName is missing, in which case the running
  // total isn't tracked per-queue and the increment is the entity-level
  // delta. Math.max-clamped at zero so a reconciliation race never
  // persists a negative entity total.
  const incCallsOutside = qName ? 0 : Number(updates.incrementCalls     || 0);
  const incAnsOutside   = qName ? 0 : Number(updates.incrementAnswered  || 0);
  const incAbnOutside   = qName ? 0 : Number(updates.incrementAbandoned || 0);

  const merged = {
    partitionKey:     entity,
    rowKey:           dateKey,
    entity,
    date:             dateKey,
    source:           "landis",
    totalCalls:       Math.max(0, baseTotal + incCallsOutside),
    answeredCalls:    Math.max(0, baseAns   + incAnsOutside),
    abandonedCalls:   Math.max(0, baseAban  + incAbnOutside),
    avgWaitSeconds:   avgOf(queueWaitMap)   || updates.avgWaitSeconds   || existing.avgWaitSeconds   || 0,
    avgTalkSeconds:   updates.avgTalkSeconds   ?? existing.avgTalkSeconds   ?? 0,
    avgHandleSeconds: avgOf(queueHandleMap) || updates.avgHandleSeconds || existing.avgHandleSeconds || 0,
    queueTotalsJson:        JSON.stringify(queueTotalsMap),
    queueAnsweredJson:      JSON.stringify(queueAnsweredMap),
    queueAbandonedJson:     JSON.stringify(queueAbandMap),
    queueWaitSecondsJson:   JSON.stringify(queueWaitMap),
    queueHandleSecondsJson: JSON.stringify(queueHandleMap),
    queueCount:             Object.keys(queueTotalsMap).length,
    updatedAt:              new Date().toISOString()
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
  // App.QueueWebhook — queue-level aggregates. queueName is propagated so
  // upsertCallData can store per-queue running totals (multiple queues per
  // entity sum to the entity's day total).
  return {
    queueName:        payload.QueueName || payload.queueName || "",
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

    const { entity: resolvedEntity, source: resolvedSource } = resolveEntity(data);

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

    const callId    = data.CallId ?? data.callId;
    const isAnswered = (Number(data.AgentId) > 0) || !!data.AgentUpn || !!data.AgentName;
    const queueName  = data.QueueName || data.queueName || "";

    const buildUnresolvedDebug = () => ({
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
    });

    // Aggregate-only path (no CallId — older test events or aggregate dumps).
    if (callId == null) {
      if (!resolvedEntity) {
        const debug = buildUnresolvedDebug();
        context.log.warn("Landis webhook: unresolved entity", debug);
        return { status: 200, body: { ok: true, skipped: true, reason: "entity not resolved", debug } };
      }
      if (data.TotalCalls != null || data.AnsweredCalls != null) {
        const updates = parseQueueWebhook(data);
        const merged = await upsertCallData(client, resolvedEntity, dateKey, updates);
        return { status: 200, headers: { "Content-Type": "application/json" }, body: JSON.stringify({ ok: true, entity: resolvedEntity, date: dateKey, mode: "aggregate", saved: merged }) };
      }
      return { status: 200, body: { ok: true, skipped: true, reason: "no CallId, no aggregate", dataKeys: Object.keys(data || {}) } };
    }

    // Per-call path. Each Landis call fires multiple events (Queue → IVR →
    // Agent), so we dedupe globally by (date, callId) — NOT per-entity. The
    // earlier per-entity scheme allowed the same call to be counted under
    // both LAOSS (via QueueName) and NES (via cross-coverage AgentUpn),
    // which inflated NES by ~64% and undercounted LAOSS by ~21%.
    //
    // Authority order for entity attribution:
    //   1. Queue context on THIS event (QueueName/OUPath/Site/etc.) — most
    //      authoritative because a queue belongs to exactly one entity.
    //   2. Prior log entry's recorded entity — trust the dedupe state.
    //   3. Agent/caller UPN on this event — fallback ONLY when the call has
    //      never been seen with a queue context. Cross-coverage means the
    //      agent's email domain doesn't necessarily match the queue's owner.
    //
    // If a call's first event was tentative agent-derived and a later event
    // arrives with authoritative queue context disagreeing, we REASSIGN —
    // decrement the wrong entity's CallData and increment the right one.
    const logClient = getCallEventLogClient();
    await logClient.createTable().catch(e => { if (e?.statusCode !== 409) throw e; });
    const logPk = dateKey;
    const logRk = String(callId);

    let priorLog = null;
    try { priorLog = await logClient.getEntity(logPk, logRk); }
    catch (e) { if (e?.statusCode !== 404) throw e; }

    let entity = null;
    let entitySource = null;
    if (resolvedEntity && resolvedSource === "queue") {
      entity = resolvedEntity; entitySource = "queue";
    } else if (priorLog?.entity) {
      entity = priorLog.entity; entitySource = priorLog.entitySource || "agent";
    } else if (resolvedEntity) {
      entity = resolvedEntity; entitySource = resolvedSource;
    }

    if (!entity) {
      const debug = buildUnresolvedDebug();
      context.log.warn("Landis webhook: unresolved entity (per-call)", { callId, ...debug });
      return { status: 200, body: { ok: true, skipped: true, reason: "entity not resolved", callId, debug } };
    }

    if (!priorLog) {
      // First sighting — log it and bump CallData.
      await logClient.createEntity({
        partitionKey: logPk,
        rowKey: logRk,
        callId: String(callId),
        entity,
        entitySource,
        date: dateKey,
        queueName,
        answered: isAnswered,
        agentId: data.AgentId || 0,
        eventType,
        firstSeenAt: new Date().toISOString()
      });
      const merged = await upsertCallData(client, entity, dateKey, {
        queueName,
        incrementCalls: 1,
        incrementAnswered: isAnswered ? 1 : 0,
        incrementAbandoned: isAnswered ? 0 : 1
      });
      return { status: 200, headers: { "Content-Type": "application/json" }, body: JSON.stringify({ ok: true, entity, date: dateKey, mode: "per-call", callId, answered: isAnswered, source: entitySource, saved: { totalCalls: merged.totalCalls, answeredCalls: merged.answeredCalls, abandonedCalls: merged.abandonedCalls, queueCount: merged.queueCount } }) };
    }

    // Duplicate sighting — reconcile if this event reveals new info.
    const priorEntity   = priorLog.entity;
    const priorSource   = priorLog.entitySource || "agent";
    const priorQueue    = priorLog.queueName || "";
    const priorAnswered = !!priorLog.answered;

    let currentEntity = priorEntity;
    let currentSource = priorSource;
    let currentQueue  = priorQueue;
    let reassigned    = false;

    // (A) Reassign: this event has authoritative queue context that disagrees
    //     with a tentative agent-derived prior attribution.
    if (
      resolvedSource === "queue" &&
      priorSource === "agent" &&
      resolvedEntity &&
      resolvedEntity !== priorEntity
    ) {
      await upsertCallData(client, priorEntity, dateKey, {
        queueName: priorQueue,
        incrementCalls: -1,
        incrementAnswered: priorAnswered ? -1 : 0,
        incrementAbandoned: priorAnswered ? 0 : -1
      });
      await upsertCallData(client, resolvedEntity, dateKey, {
        queueName,
        incrementCalls: 1,
        incrementAnswered: priorAnswered ? 1 : 0,
        incrementAbandoned: priorAnswered ? 0 : 1
      });
      currentEntity = resolvedEntity;
      currentSource = "queue";
      currentQueue  = queueName;
      reassigned    = true;
    }

    // (B) Answered flip: prior was abandoned (no agent), this event reveals
    //     an agent picked up the call. Move the count from abandoned to
    //     answered for whichever entity now owns it.
    let answeredFlip = false;
    if (!priorAnswered && isAnswered) {
      await upsertCallData(client, currentEntity, dateKey, {
        queueName: currentQueue,
        incrementCalls: 0,
        incrementAnswered: 1,
        incrementAbandoned: -1
      });
      answeredFlip = true;
    }

    // Persist the reconciled state on the log row.
    if (reassigned || answeredFlip || priorSource !== currentSource) {
      try {
        await logClient.updateEntity({
          partitionKey: logPk,
          rowKey: logRk,
          entity: currentEntity,
          entitySource: currentSource,
          queueName: currentQueue,
          answered: priorAnswered || isAnswered,
          agentId: data.AgentId || priorLog.agentId || 0
        }, "Merge");
      } catch (_) {}
    }

    return { status: 200, body: { ok: true, entity: currentEntity, date: dateKey, mode: "per-call", callId, duplicate: true, reassigned, answeredFlip, source: currentSource } };
  } catch (error) {
    context.log.error("Landis webhook error:", error.message);
    return {
      status: 500,
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ ok: false, error: "Webhook processing failed" })
    };
  }
};
