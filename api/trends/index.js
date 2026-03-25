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

function toLegacyShape(entity) {
  const values = parseValuesJson(entity.valuesJson);

  return {
    entity: entity.entity || entity.partitionKey,
    weekEnding: entity.weekEnding || entity.rowKey,
    status: normalizeStatus(entity.status),
    visitVolume: toNumber(values.totalVisits) ?? toNumber(entity.visitVolume) ?? 0,
    callVolume: toNumber(values.totalCalls) ?? toNumber(entity.callVolume) ?? 0,
    newPatients: toNumber(values.npActual) ?? toNumber(entity.newPatients) ?? 0,
    noShowRate: toNumber(values.noShowRate) ?? toNumber(entity.noShowRate) ?? 0,
    cancellationRate:
      toNumber(values.cancellationRate) ?? toNumber(entity.cancellationRate) ?? 0,
    abandonedCallRate:
      toNumber(values.abandonmentRate) ?? toNumber(entity.abandonedCallRate) ?? 0,
    updatedBy: entity.updatedBy || entity.submittedBy || entity.approvedBy || null,
    updatedAt: entity.updatedAt || entity.importedAt || entity.approvedAt || null
  };
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
      fridayRows.set(item.weekEnding, item);
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

    const entity = req.query.entity;
    const mode = req.query.mode || "recent";
    const startDate = req.query.startDate;
    const endDate = req.query.endDate;

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

    let items = rawRows.map(toLegacyShape);

    if (mode === "dateRange" && startDate && endDate) {
      items = items.filter(
        (item) => item.weekEnding >= startDate && item.weekEnding <= endDate
      );
    } else {
      items = items
        .sort((a, b) => (a.weekEnding < b.weekEnding ? 1 : -1))
        .slice(0, 8);
    }

    items = dedupeWeeks(items);

    return {
      status: 200,
      body: {
        ok: true,
        entity,
        count: items.length,
        appliedFilter:
          mode === "dateRange"
            ? {
                mode: "dateRange",
                startDate,
                endDate
              }
            : {
                mode: "recent"
              },
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
