const { getTableClient } = require("../shared/table");

const TABLE_NAME = "WeeklyRegionData";

function parseJson(value, fallback = {}) {
  if (!value) return fallback;
  try {
    return typeof value === "string" ? JSON.parse(value) : value;
  } catch {
    return fallback;
  }
}

function toNumber(value, fallback = 0) {
  if (value === null || value === undefined || value === "") return fallback;
  const n = Number(value);
  return Number.isFinite(n) ? n : fallback;
}

function normalizeValues(raw = {}) {
  const values = parseJson(raw, {});

  return {
    newPatients: toNumber(values.newPatients ?? values.npActual, 0),
    surgeries: toNumber(values.surgeries ?? values.surgeryActual, 0),
    established: toNumber(values.established ?? values.establishedActual, 0),
    noShows: toNumber(values.noShows, 0),
    cancelled: toNumber(values.cancelled, 0),
    totalCalls: toNumber(values.totalCalls ?? values.callVolume, 0),
    abandonedCalls: toNumber(values.abandonedCalls, 0),
    visitVolume: toNumber(values.visitVolume ?? values.totalVisits, 0),
    callVolume: toNumber(values.callVolume ?? values.totalCalls, 0),
    noShowRate: toNumber(values.noShowRate, 0),
    cancellationRate: toNumber(values.cancellationRate, 0),
    abandonedCallRate: toNumber(values.abandonedCallRate ?? values.abandonmentRate, 0)
  };
}

module.exports = async function (context, req) {
  try {
    const entity = String(req.query?.entity || "").trim();
    const weekEnding = String(req.query?.weekEnding || "").trim();

    if (!entity || !weekEnding) {
      context.res = {
        status: 400,
        headers: { "Content-Type": "application/json" },
        body: {
          ok: false,
          error: "Missing entity or weekEnding."
        }
      };
      return;
    }

    const table = getTableClient(TABLE_NAME);

    let record = null;
    try {
      record = await table.getEntity(entity, weekEnding);
    } catch (err) {
      if (err?.statusCode !== 404) {
        throw err;
      }
    }

    const values = normalizeValues(record?.valuesJson || {});

    context.res = {
      status: 200,
      headers: { "Content-Type": "application/json" },
      body: {
        ok: true,
        entity,
        weekEnding,
        status: record?.status || "draft",
        values,
        source: record?.source || "app",
        importedAt: record?.importedAt || null,
        updatedAt: record?.updatedAt || null
      }
    };
  } catch (error) {
    context.log.error("weekly GET failed", error);
    context.res = {
      status: 500,
      headers: { "Content-Type": "application/json" },
      body: {
        ok: false,
        error: "Failed to load weekly region data.",
        details: error.message
      }
    };
  }
};
