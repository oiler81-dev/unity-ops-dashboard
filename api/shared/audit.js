const { ensureTable } = require("./table");
const { AUDIT_TABLE } = require("./constants");

function toIso(value) {
  try {
    return new Date(value || Date.now()).toISOString();
  } catch {
    return new Date().toISOString();
  }
}

function safeJson(value, fallback = {}) {
  try {
    return JSON.stringify(value ?? fallback);
  } catch {
    return JSON.stringify(fallback);
  }
}

function summarizeChanges(before = {}, after = {}, eventType = "update") {
  if (eventType === "create") {
    const fields = Object.keys(after || {});
    return fields.length
      ? `Created record with ${fields.length} fields`
      : "Created record";
  }

  if (eventType === "delete") {
    const fields = Object.keys(before || {});
    return fields.length
      ? `Deleted record with ${fields.length} fields`
      : "Deleted record";
  }

  const fieldSet = new Set([
    ...Object.keys(before || {}),
    ...Object.keys(after || {})
  ]);

  const changed = [];

  for (const field of fieldSet) {
    const beforeValue = before?.[field];
    const afterValue = after?.[field];

    if (JSON.stringify(beforeValue ?? null) !== JSON.stringify(afterValue ?? null)) {
      changed.push(field);
    }
  }

  if (!changed.length) {
    return "Updated record";
  }

  if (changed.length <= 4) {
    return `Updated ${changed.join(", ")}`;
  }

  return `Updated ${changed.length} fields`;
}

async function writeAuditEvent({
  eventType,
  entity,
  weekEnding,
  actorEmail,
  actorRole,
  before,
  after,
  summary,
  metadata
}) {
  const client = await ensureTable(AUDIT_TABLE);

  const timestamp = toIso();
  const partitionKey = entity || "UNKNOWN";
  const rowKey = `${timestamp}__${Math.random().toString(36).slice(2, 10)}`;

  const resolvedSummary =
    summary || summarizeChanges(before || {}, after || {}, eventType || "update");

  await client.upsertEntity({
    partitionKey,
    rowKey,
    entity: entity || "",
    weekEnding: weekEnding || "",
    eventType: eventType || "update",
    actorEmail: actorEmail || "",
    actorRole: actorRole || "",
    timestamp,
    summary: resolvedSummary,
    beforeJson: safeJson(before, {}),
    afterJson: safeJson(after, {}),
    metadataJson: safeJson(metadata, {})
  });
}

function parseJsonSafe(value, fallback = {}) {
  if (!value) return fallback;

  try {
    return typeof value === "string" ? JSON.parse(value) : value;
  } catch {
    return fallback;
  }
}

module.exports = {
  writeAuditEvent,
  parseJsonSafe,
  summarizeChanges
};