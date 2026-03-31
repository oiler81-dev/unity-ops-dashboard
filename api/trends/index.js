const { getUserFromRequest } = require("../shared/auth");
const { resolveAccess } = require("../shared/permissions");
const { getTableClient } = require("../shared/table");

const REGION_TABLE = "WeeklyRegionData";

function toNumber(value) {
  if (value == null || value === "") return null;
  const n = Number(value);
  return Number.isFinite(n) ? n : null;
}

function parseValuesJson(valuesJson) {
  if (!valuesJson) return {};
  try {
    return typeof valuesJson === "string" ? JSON.parse(valuesJson) : valuesJson;
  } catch {
    return {};
  }
}

function normalizeStatus(status) {
  const s = String(status || "").trim().toLowerCase();
  if (s === "approved") return "Approved";
  if (s === "submitted") return "submitted";
  if (s === "draft") return "draft";
  return status || "draft";
}

function mapRow(entity) {
  const values = parseValuesJson(entity.valuesJson);

  return {
    entity: entity.entity || entity.partitionKey,
    weekEnding: entity.weekEnding || entity.rowKey,
    status: normalizeStatus(entity.status),

    newPatients:
      toNumber(values.newPatients ?? values.npActual) ??
      toNumber(entity.newPatients) ??
      0,

    surgeries:
      toNumber(values.surgeries ?? values.surgeryActual) ??
      toNumber(entity.surgeries) ??
      0,

    established:
      toNumber(values.established ?? values.establishedActual) ??
      toNumber(entity.established) ??
      0,

    noShows:
      toNumber(values.noShows) ??
      toNumber(entity.noShows) ??
      0,

    cancelled:
      toNumber(values.cancelled) ??
      toNumber(entity.cancelled) ??
      0,

    totalCalls:
      toNumber(values.totalCalls ?? values.callVolume) ??
      toNumber(entity.totalCalls ?? entity.callVolume) ??
      0,

    abandonedCalls:
      toNumber(values.abandonedCalls) ??
      toNumber(entity.abandonedCalls) ??
      0,

    visitVolume:
      toNumber(values.visitVolume ?? values.totalVisits) ??
      toNumber(entity.visitVolume) ??
      0,

    callVolume:
      toNumber(values.callVolume ?? values.totalCalls) ??
      toNumber(entity.callVolume ?? entity.totalCalls) ??
      0,

    noShowRate:
      toNumber(values.noShowRate) ??
      toNumber(entity.noShowRate) ??
      0,

    cancellationRate:
      toNumber(values.cancellationRate) ??
      toNumber(entity.cancellationRate) ??
      0,

    abandonedCallRate:
      toNumber(values.abandonmentRate ?? values.abandonedCallRate) ??
      toNumber(entity.abandonedCallRate) ??
      0,

    updatedBy: entity.updatedBy || entity.submittedBy || entity.approvedBy || null,
    updatedAt: entity.updatedAt || entity.importedAt || entity.approvedAt || null,
    source: entity.source || null
  };
}

function hasRealData(row) {
  return (
    (row.visitVolume ?? 0) > 0 ||
    (row.callVolume ?? 0) > 0 ||
    (row.newPatients ?? 0) > 0 ||
    (row.surgeries ?? 0) > 0 ||
    (row.established ?? 0) > 0 ||
    (row.noShows ?? 0) > 0 ||
    (row.cancelled ?? 0) > 0 ||
    (row.totalCalls ?? 0) > 0 ||
    (row.abandonedCalls ?? 0) > 0 ||
    (row.noShowRate ?? 0) > 0 ||
    (row.cancellationRate ?? 0) > 0 ||
    (row.abandonedCallRate ?? 0) > 0
  );
}

function isFriday(isoDate) {
  const d = new Date(`${isoDate}T00:00:00Z`);
  return d.getUTCDay() === 5;
}

function previousFridayForDate(isoDate) {
  const d = new Date(`${isoDate}T00:00:00Z`);
  const day = d.getUTCDay();
  const diff = (day + 2) % 7;
  d.setUTCDate(d.getUTCDate() - diff);
  return d.toISOString().slice(0, 10);
}

function dedupeWeeks(items) {
  const fridayRows = new Map();
  const nonFridayRows = [];

  for (const item of items) {
    if (isFriday(item.weekEnding)) {
      const existing = fridayRows.get(item.weekEnding);
      if (!existing) {
        fridayRows.set(item.weekEnding, item);
        continue;
      }

      const existingScore =
        (existing.source === "workbook-import" ? 2 : 0) +
        (existing.status === "Approved" ? 1 : 0);

      const itemScore =
        (item.source === "workbook-import" ? 2 : 0) +
        (item.status === "Approved" ? 1 : 0);

      if (itemScore >= existingScore) {
        fridayRows.set(item.weekEnding, item);
      }
    } else {
      nonFridayRows.push(item);
    }
  }

  const filteredNonFriday = nonFridayRows.filter((item) => {
    const anchorFriday = previousFridayForDate(item.weekEnding);
    return !fridayRows.has(anchorFriday);
  });

  return [...fridayRows.values(), ...filteredNonFriday].sort((a, b) =>
    a.weekEnding < b.weekEnding ? 1 : -1
  );
}

async function getRowsForEntity(table, entity) {
  const filter = `PartitionKey eq '${entity.replace(/'/g, "''")}'`;
  const rows = [];

  for await (const row of table.listEntities({
    queryOptions: { filter }
  })) {
    rows.push(row);
  }

  return rows;
}

module.exports = async function (context, req) {
  try {
    const user = getUserFromRequest(req);
    const access = resolveAccess(user);

    const entity = String(req.query.entity || "").trim();
    const mode = String(req.query.mode || "recent").trim();
    const startDate = String(req.query.startDate || "").trim();
    const endDate = String(req.query.endDate || "").trim();
    const weeks = Math.max(1, Math.min(52, Number(req.query.weeks || 8) || 8));

    if (!entity) {
      return {
        status: 400,
        body: {
          ok: false,
          error: "Missing entity"
        }
      };
    }

    if (!access.isAdmin && access.entity !== entity) {
      return {
        status: 403,
        body: {
          ok: false,
          error: "Forbidden"
        }
      };
    }

    const table = getTableClient(REGION_TABLE);
    const rawRows = await getRowsForEntity(table, entity);

    let items = rawRows.map(mapRow).filter(hasRealData);
    items = dedupeWeeks(items);

    if (mode === "dateRange" && startDate && endDate) {
      items = items.filter(
        (item) => item.weekEnding >= startDate && item.weekEnding <= endDate
      );
    } else {
      items = items.slice(0, weeks);
    }

    return {
      status: 200,
      body: {
        ok: true,
        entity,
        count: items.length,
        appliedFilter:
          mode === "dateRange" && startDate && endDate
            ? { mode: "dateRange", startDate, endDate }
            : { mode: "recent", weeks },
        items
      }
    };
  } catch (error) {
    context.log.error("trends failed", error);

    return {
      status: 500,
      body: {
        ok: false,
        error: "Failed to load trends",
        details: error.message
      }
    };
  }
};