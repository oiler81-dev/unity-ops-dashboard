const { getUserFromRequest } = require("../shared/auth");
const { resolveAccess } = require("../shared/permissions");
const { getTableClient } = require("../shared/table");

const REGION_TABLE = "WeeklyRegionData";
const FORECAST_TABLE = "PTOForecastData";
const ENTITIES = ["LAOSS", "NES", "SpineOne", "MRO"];
const RATE_WEEKS = 12; // weeks of history used to compute daily rates

function toNumber(value, fallback = 0) {
  if (value == null || value === "") return fallback;
  const n = Number(value);
  return Number.isFinite(n) ? n : fallback;
}

function safeParseJson(value, fallback = {}) {
  if (!value) return fallback;
  try {
    return typeof value === "string" ? JSON.parse(value) : value;
  } catch {
    return fallback;
  }
}

// Build the rolling 3 month keys starting from a given date
function getRollingMonthKeys(fromDate = new Date()) {
  const keys = [];
  for (let i = 0; i < 3; i++) {
    const d = new Date(fromDate);
    d.setUTCDate(1);
    d.setUTCMonth(d.getUTCMonth() + i);
    const yyyy = d.getUTCFullYear();
    const mm = String(d.getUTCMonth() + 1).padStart(2, "0");
    keys.push(`${yyyy}-${mm}`);
  }
  return keys;
}

function monthKeyToLabel(monthKey) {
  const [yyyy, mm] = monthKey.split("-");
  const names = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
                 "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  return `${names[parseInt(mm, 10) - 1]} ${yyyy}`;
}

// Compute average visits/day and surgeries/day from recent weekly records
async function computeEntityRates(regionTable, entity) {
  try {
    const today = new Date().toISOString().slice(0, 10);

    // Query region records for this entity, past weeks only
    const iter = regionTable.listEntities({
      queryOptions: {
        filter: `PartitionKey eq '${entity}'`
      }
    });

    const allRows = [];
    for await (const row of iter) {
      allRows.push(row);
    }

    const recent = allRows
      .filter((r) => {
        const we = r.rowKey || "";
        return we <= today && toNumber(r.visitVolume, 0) > 0;
      })
      .sort((a, b) => {
        const wa = a.rowKey || "";
        const wb = b.rowKey || "";
        return wb.localeCompare(wa);
      })
      .slice(0, RATE_WEEKS);

    if (!recent.length) {
      return { visitsPerDay: 0, surgeriesPerDay: 0, weeksUsed: 0 };
    }

    const totalVisits = recent.reduce((sum, r) => {
      const v = safeParseJson(r.valuesJson, {});
      return sum + toNumber(v.totalVisits ?? v.visitVolume ?? r.visitVolume, 0);
    }, 0);

    const totalSurgeries = recent.reduce((sum, r) => {
      const v = safeParseJson(r.valuesJson, {});
      return sum + toNumber(v.surgeryActual ?? v.surgeries ?? r.surgeries, 0);
    }, 0);

    const totalDays = recent.reduce((sum, r) => {
      const v = safeParseJson(r.valuesJson, {});
      return sum + toNumber(v.daysInPeriod ?? r.daysInPeriod, 5);
    }, 0);

    const nonSurgicalVisits = Math.max(0, totalVisits - totalSurgeries);

    return {
      visitsPerDay: totalDays > 0
        ? Number((totalVisits / totalDays).toFixed(2))
        : 0,
      clinicalVisitsPerDay: totalDays > 0
        ? Number((nonSurgicalVisits / totalDays).toFixed(2))
        : 0,
      surgeriesPerDay: totalDays > 0
        ? Number((totalSurgeries / totalDays).toFixed(2))
        : 0,
      weeksUsed: recent.length
    };
  } catch {
    return { visitsPerDay: 0, clinicalVisitsPerDay: 0, surgeriesPerDay: 0, weeksUsed: 0 };
  }
}

// Compute forecasted impact for one entity+month given rates and PTO inputs
function computeForecast(clinicalPtoDays, surgicalPtoDays, rates) {
  const missedClinical = Number((toNumber(clinicalPtoDays) * rates.clinicalVisitsPerDay).toFixed(1));
  const missedSurgical = Number((toNumber(surgicalPtoDays) * rates.surgeriesPerDay).toFixed(1));
  const totalMissed = Number((missedClinical + missedSurgical).toFixed(1));

  return {
    missedClinicalVisits: missedClinical,
    missedSurgeries: missedSurgical,
    totalMissedVisits: totalMissed
  };
}

async function getForecastRecord(forecastTable, entity, monthKey) {
  try {
    return await forecastTable.getEntity(entity, monthKey);
  } catch (err) {
    // 404 = record not found, also swallow table-not-found gracefully
    if (err?.statusCode === 404 || err?.code === "ResourceNotFound" || err?.code === "TableNotFound") return null;
    throw err;
  }
}

async function ensureForecastTable(forecastTable) {
  try {
    await forecastTable.createTable();
  } catch {
    // Swallow all errors — table likely already exists
  }
}

module.exports = async function (context, req) {
  const respond = (status, body) => {
    context.res = {
      status,
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body)
    };
  };

  try {
    const user = getUserFromRequest(req);
    const access = resolveAccess(user);

    if (!access?.authenticated) {
      return respond(401, { ok: false, error: "Unauthorized" });
    }

    const regionTable = getTableClient(REGION_TABLE);
    const forecastTable = getTableClient(FORECAST_TABLE);

    // Ensure the forecast table exists (creates it silently if not)
    await ensureForecastTable(forecastTable);

    // ── POST: save operator PTO forecast inputs ──────────────────────────────
    if (req.method === "POST") {
      const body = req.body || {};
      const entity = String(body.entity || "").trim();
      const monthKey = String(body.monthKey || "").trim();
      const clinicalPtoDays = toNumber(body.clinicalPtoDays, 0);
      const surgicalPtoDays = toNumber(body.surgicalPtoDays, 0);

      if (!entity || !monthKey) {
        return respond(400, { ok: false, error: "Missing entity or monthKey" });
      }

      // Operators can only save for their own entity; admins can save any
      if (!access.isAdmin && access.entity !== entity) {
        return respond(403, { ok: false, error: "Forbidden" });
      }

      const rates = await computeEntityRates(regionTable, entity);
      const forecast = computeForecast(clinicalPtoDays, surgicalPtoDays, rates);

      await forecastTable.upsertEntity({
        partitionKey: entity,
        rowKey: monthKey,
        entity,
        monthKey,
        monthLabel: monthKeyToLabel(monthKey),
        clinicalPtoDays,
        surgicalPtoDays,
        missedClinicalVisits: forecast.missedClinicalVisits,
        missedSurgeries: forecast.missedSurgeries,
        totalMissedVisits: forecast.totalMissedVisits,
        clinicalVisitsPerDay: rates.clinicalVisitsPerDay,
        surgeriesPerDay: rates.surgeriesPerDay,
        weeksUsed: rates.weeksUsed,
        savedBy: access.email || user?.userDetails || null,
        savedAt: new Date().toISOString(),
        updatedAt: new Date().toISOString()
      });

      return respond(200, {
        ok: true,
        message: "Forecast saved",
        entity,
        monthKey,
        clinicalPtoDays,
        surgicalPtoDays,
        forecast,
        rates
      });
    }

    // ── GET: load forecasts + rates for all entities, rolling 3 months ───────
    const monthKeys = getRollingMonthKeys();

    // Determine which entities to return
    const scopeEntities = access.isAdmin
      ? ENTITIES
      : ENTITIES.filter((e) => e === access.entity);

    const results = await Promise.all(
      scopeEntities.map(async (entity) => {
        try {
          const rates = await computeEntityRates(regionTable, entity);

          const months = await Promise.all(
            monthKeys.map(async (monthKey) => {
              try {
                const saved = await getForecastRecord(forecastTable, entity, monthKey);

                const clinicalPtoDays = toNumber(saved?.clinicalPtoDays, 0);
                const surgicalPtoDays = toNumber(saved?.surgicalPtoDays, 0);
                const forecast = computeForecast(clinicalPtoDays, surgicalPtoDays, rates);

                return {
                  monthKey,
                  monthLabel: monthKeyToLabel(monthKey),
                  clinicalPtoDays,
                  surgicalPtoDays,
                  missedClinicalVisits: saved ? toNumber(saved.missedClinicalVisits) : forecast.missedClinicalVisits,
                  missedSurgeries: saved ? toNumber(saved.missedSurgeries) : forecast.missedSurgeries,
                  totalMissedVisits: saved ? toNumber(saved.totalMissedVisits) : forecast.totalMissedVisits,
                  savedBy: saved?.savedBy || null,
                  savedAt: saved?.savedAt || null,
                  hasSavedEntry: !!saved
                };
              } catch {
                return {
                  monthKey,
                  monthLabel: monthKeyToLabel(monthKey),
                  clinicalPtoDays: 0,
                  surgicalPtoDays: 0,
                  missedClinicalVisits: 0,
                  missedSurgeries: 0,
                  totalMissedVisits: 0,
                  savedBy: null,
                  savedAt: null,
                  hasSavedEntry: false
                };
              }
            })
          );

          return {
            entity,
            rates,
            months,
            quarterlyTotalMissed: months.reduce((sum, m) => sum + toNumber(m.totalMissedVisits), 0)
          };
        } catch {
          const emptyMonth = (mk) => ({
            monthKey: mk,
            monthLabel: monthKeyToLabel(mk),
            clinicalPtoDays: 0, surgicalPtoDays: 0,
            missedClinicalVisits: 0, missedSurgeries: 0, totalMissedVisits: 0,
            savedBy: null, savedAt: null, hasSavedEntry: false
          });
          return {
            entity,
            rates: { visitsPerDay: 0, clinicalVisitsPerDay: 0, surgeriesPerDay: 0, weeksUsed: 0 },
            months: monthKeys.map(emptyMonth),
            quarterlyTotalMissed: 0
          };
        }
      })
    );

    return respond(200, {
      ok: true,
      monthKeys,
      monthLabels: monthKeys.map(monthKeyToLabel),
      entities: results
    });
  } catch (error) {
    context.log.error("pto-forecast failed", error);
    return respond(500, { ok: false, error: "Failed to process PTO forecast", details: error.message });
  }
};
