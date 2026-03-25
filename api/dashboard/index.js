const { getUserFromRequest } = require("../shared/auth");
const { resolveAccess } = require("../shared/permissions");
const { getTableClient } = require("../shared/table");

const REGION_TABLE = "WeeklyRegionData";
const ENTITIES = ["LAOSS", "NES", "SpineOne", "MRO"];

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
    visitVolume: toNumber(values.totalVisits) ?? toNumber(entity.visitVolume) ?? 0,
    callVolume: toNumber(values.totalCalls) ?? toNumber(entity.callVolume) ?? 0,
    newPatients: toNumber(values.npActual) ?? toNumber(entity.newPatients) ?? 0,
    noShowRate: toNumber(values.noShowRate) ?? toNumber(entity.noShowRate) ?? 0,
    cancellationRate:
      toNumber(values.cancellationRate) ?? toNumber(entity.cancellationRate) ?? 0,
    abandonedCallRate:
      toNumber(values.abandonmentRate) ?? toNumber(entity.abandonedCallRate) ?? 0
  };
}

async function getWeekRow(table, entity, weekEnding) {
  try {
    const row = await table.getEntity(entity, weekEnding);
    return mapRow(row);
  } catch {
    return null;
  }
}

function average(values) {
  if (!values.length) return 0;
  return values.reduce((a, b) => a + b, 0) / values.length;
}

module.exports = async function (context, req) {
  try {
    const user = getUserFromRequest(req);
    resolveAccess(user);

    const weekEnding = String(req.query.weekEnding || "").trim();
    if (!weekEnding) {
      return {
        status: 400,
        body: {
          ok: false,
          error: "Missing weekEnding"
        }
      };
    }

    const table = getTableClient(REGION_TABLE);
    const rows = [];

    for (const entity of ENTITIES) {
      const row = await getWeekRow(table, entity, weekEnding);
      if (!row) continue;
      if (row.status !== "Approved") continue;
      rows.push(row);
    }

    const totals = {
      visitVolume: rows.reduce((sum, r) => sum + (r.visitVolume || 0), 0),
      callVolume: rows.reduce((sum, r) => sum + (r.callVolume || 0), 0),
      newPatients: rows.reduce((sum, r) => sum + (r.newPatients || 0), 0)
    };

    const kpis = [
      {
        key: "approvedRegions",
        label: "Approved Regions",
        value: rows.length,
        meta: `${rows.length} region(s) approved`
      },
      {
        key: "visitVolume",
        label: "Visit Volume",
        value: totals.visitVolume,
        meta: "Approved-region total"
      },
      {
        key: "callVolume",
        label: "Call Volume",
        value: totals.callVolume,
        meta: "Approved-region total"
      },
      {
        key: "newPatients",
        label: "New Patients",
        value: totals.newPatients,
        meta: "Approved-region total"
      },
      {
        key: "avgNoShowRate",
        label: "Avg No Show %",
        value: average(rows.map((r) => r.noShowRate || 0)),
        format: "percent"
      },
      {
        key: "avgCancellationRate",
        label: "Avg Cancel %",
        value: average(rows.map((r) => r.cancellationRate || 0)),
        format: "percent"
      },
      {
        key: "avgAbandonedCallRate",
        label: "Avg Abandoned %",
        value: average(rows.map((r) => r.abandonedCallRate || 0)),
        format: "percent"
      }
    ];

    return {
      status: 200,
      body: {
        ok: true,
        weekEnding,
        entityCount: rows.length,
        totals,
        regions: rows,
        kpis
      }
    };
  } catch (error) {
    context.log.error("dashboard failed", error);

    return {
      status: 500,
      body: {
        ok: false,
        error: "Failed to load dashboard",
        details: error.message
      }
    };
  }
};
