const { getUserFromRequest } = require("../shared/auth");
const { resolveAccess } = require("../shared/permissions");
const { getTableClient } = require("../shared/table");

const REGION_TABLE = "WeeklyRegionData";
const ENTITIES = ["LAOSS", "NES", "SpineOne", "MRO"];

function toNumber(value) {
  if (value == null || value === "") return 0;
  const n = Number(value);
  return Number.isFinite(n) ? n : 0;
}

function parseJson(value) {
  if (!value) return {};
  try {
    return typeof value === "string" ? JSON.parse(value) : value;
  } catch {
    return {};
  }
}

function normalizeStatus(status) {
  const s = String(status || "").toLowerCase();
  if (s === "approved") return "approved";
  if (s === "submitted") return "submitted";
  return "draft";
}

function mapRecord(record) {
  const values = parseJson(record.valuesJson);

  return {
    entity: record.partitionKey,
    weekEnding: record.rowKey,
    status: normalizeStatus(record.status),

    visitVolume:
      toNumber(values.totalVisits) || toNumber(record.visitVolume),

    callVolume:
      toNumber(values.totalCalls) || toNumber(record.callVolume),

    newPatients:
      toNumber(values.npActual) || toNumber(record.newPatients),

    noShowRate:
      toNumber(values.noShowRate) || toNumber(record.noShowRate),

    cancellationRate:
      toNumber(values.cancellationRate) || toNumber(record.cancellationRate),

    abandonedCallRate:
      toNumber(values.abandonmentRate) ||
      toNumber(record.abandonedCallRate)
  };
}

module.exports = async function (context, req) {
  try {
    const user = getUserFromRequest(req);
    resolveAccess(user);

    const weekEnding = req.query.weekEnding;

    if (!weekEnding) {
      return {
        status: 400,
        body: { ok: false, error: "Missing weekEnding" }
      };
    }

    const table = getTableClient(REGION_TABLE);

    const rows = [];

    for (const entity of ENTITIES) {
      try {
        const record = await table.getEntity(entity, weekEnding);
        const mapped = mapRecord(record);

        if (mapped.status === "approved") {
          rows.push(mapped);
        }
      } catch (err) {
        if (err.statusCode !== 404) throw err;
      }
    }

    const totals = {
      visitVolume: rows.reduce((s, r) => s + r.visitVolume, 0),
      callVolume: rows.reduce((s, r) => s + r.callVolume, 0),
      newPatients: rows.reduce((s, r) => s + r.newPatients, 0)
    };

    const avg = (arr) =>
      arr.length ? arr.reduce((s, v) => s + v, 0) / arr.length : 0;

    const averages = {
      noShowRate: avg(rows.map((r) => r.noShowRate)),
      cancellationRate: avg(rows.map((r) => r.cancellationRate)),
      abandonedCallRate: avg(rows.map((r) => r.abandonedCallRate))
    };

    return {
      status: 200,
      body: {
        ok: true,
        weekEnding,
        entityCount: rows.length,
        totals,
        averages,
        regions: rows
      }
    };
  } catch (error) {
    context.log.error("executive failed", error);

    return {
      status: 500,
      body: {
        ok: false,
        error: "Failed to load executive summary",
        details: error.message
      }
    };
  }
};
