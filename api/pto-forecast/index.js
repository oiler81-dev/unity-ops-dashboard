const { getUserFromRequest } = require("../shared/auth");
const {
  resolveAccess,
  requireAccess,
  requireEntityAccess,
  scopeEntitiesToAccess
} = require("../shared/permissions");
const { getTableClient } = require("../shared/table");

const REGION_TABLE = "WeeklyRegionData";
const FORECAST_TABLE = "PTOForecastData";
const SETTINGS_TABLE = "ProviderSettingsData";
const ENTITIES = ["LAOSS", "NES", "SpineOne", "MRO"];
const RATE_WEEKS = 12; // weeks of history used to compute daily rates

// Default working days per week per entity
const ENTITY_WORKING_DAYS = {
  LAOSS:    5,
  NES:      4.5,
  SpineOne: 5,
  MRO:      5
};

// Fallback defaults if provider settings haven't been configured yet
const PROVIDER_DEFAULTS = {
  LAOSS:    { mdCount: 21, paCount: 14, ptCount: 0 },
  NES:      { mdCount: 12, paCount: 1,  ptCount: 4 },
  SpineOne: { mdCount: 2,  paCount: 3,  ptCount: 2 },
  MRO:      { mdCount: 8,  paCount: 4,  ptCount: 3 }
};

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

    const defaultDays = ENTITY_WORKING_DAYS[entity] ?? 5;
    const totalDays = recent.reduce((sum, r) => {
      const v = safeParseJson(r.valuesJson, {});
      return sum + toNumber(v.daysInPeriod ?? r.daysInPeriod, defaultDays);
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

// Fetch provider settings for an entity, falling back to defaults
async function getProviderSettings(settingsTable, entity) {
  try {
    const saved = await settingsTable.getEntity(entity, "settings");
    const defaults = PROVIDER_DEFAULTS[entity] || { mdCount: 1, paCount: 0, ptCount: 0 };
    return {
      mdCount: toNumber(saved?.mdCount, defaults.mdCount),
      paCount: toNumber(saved?.paCount, defaults.paCount),
      ptCount: toNumber(saved?.ptCount, defaults.ptCount)
    };
  } catch {
    return PROVIDER_DEFAULTS[entity] || { mdCount: 1, paCount: 0, ptCount: 0 };
  }
}

// Compute per-provider daily rates from aggregate entity rates + provider counts
function computeProviderRates(rates, providerSettings) {
  const { mdCount, paCount, ptCount } = providerSettings;
  const totalClinical = Math.max(1, mdCount + paCount);
  const totalPt = Math.max(1, ptCount || 1);

  return {
    mdVisitsPerDay:  mdCount > 0 ? Number((rates.clinicalVisitsPerDay / totalClinical * (mdCount / totalClinical * totalClinical)).toFixed(2)) : 0,
    paVisitsPerDay:  paCount > 0 ? Number((rates.clinicalVisitsPerDay / totalClinical).toFixed(2)) : 0,
    ptVisitsPerDay:  ptCount > 0 ? Number((rates.visitsPerDay / totalPt).toFixed(2)) : 0,
    // Simplified: divide aggregate rate equally across all clinical providers
    perProviderClinical: Number((rates.clinicalVisitsPerDay / Math.max(1, totalClinical)).toFixed(2)),
    perProviderPt:       Number(((rates.visitsPerDay - rates.clinicalVisitsPerDay) / Math.max(1, totalPt)).toFixed(2))
  };
}

// Compute forecasted impact for one entity+month given provider PTO inputs
function computeForecast(mdPtoDays, paPtoDays, ptPtoDays, rates, providerSettings) {
  const perClinical = Number((rates.clinicalVisitsPerDay / Math.max(1, providerSettings.mdCount + providerSettings.paCount)).toFixed(2));
  const ptRate = providerSettings.ptCount > 0
    ? Number(((rates.visitsPerDay - rates.clinicalVisitsPerDay) / providerSettings.ptCount).toFixed(2))
    : 0;

  const missedMd = Number((toNumber(mdPtoDays) * perClinical).toFixed(1));
  const missedPa = Number((toNumber(paPtoDays) * perClinical).toFixed(1));
  const missedPt = Number((toNumber(ptPtoDays) * ptRate).toFixed(1));
  const totalMissed = Number((missedMd + missedPa + missedPt).toFixed(1));

  return {
    missedMdVisits:   missedMd,
    missedPaVisits:   missedPa,
    missedPtVisits:   missedPt,
    totalMissedVisits: totalMissed,
    perProviderClinicalRate: perClinical,
    perProviderPtRate: ptRate
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

    const authError = requireAccess(access);
    if (authError) return respond(authError.status, authError.body);

    const regionTable = getTableClient(REGION_TABLE);
    const forecastTable = getTableClient(FORECAST_TABLE);
    const settingsTable = getTableClient(SETTINGS_TABLE);

    // Ensure the forecast table exists (creates it silently if not)
    await ensureForecastTable(forecastTable);

    // ── POST: save operator PTO forecast inputs ──────────────────────────────
    if (req.method === "POST") {
      const body = req.body || {};
      const entity = String(body.entity || "").trim();
      const monthKey = String(body.monthKey || "").trim();
      const mdPtoDays = toNumber(body.mdPtoDays, 0);
      const paPtoDays = toNumber(body.paPtoDays, 0);
      const ptPtoDays = toNumber(body.ptPtoDays, 0);

      if (!entity || !monthKey) {
        return respond(400, { ok: false, error: "Missing entity or monthKey" });
      }

      if (!ENTITIES.includes(entity)) {
        return respond(400, { ok: false, error: "Invalid entity" });
      }

      if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(monthKey)) {
        return respond(400, { ok: false, error: "Invalid monthKey format (expected YYYY-MM)" });
      }

      // Operators can only save for their own entity; admins can save any
      const entityError = requireEntityAccess(access, entity);
      if (entityError) return respond(entityError.status, entityError.body);

      const rates = await computeEntityRates(regionTable, entity);
      const providerSettings = await getProviderSettings(settingsTable, entity);
      const forecast = computeForecast(mdPtoDays, paPtoDays, ptPtoDays, rates, providerSettings);

      // Normalize and persist provider breakdown so it shows up on reload
      const rawBreakdown = Array.isArray(body.providerBreakdown) ? body.providerBreakdown : [];
      const providerBreakdown = rawBreakdown
        .map((p) => ({
          name: String(p?.name || "").trim().slice(0, 120),
          type: ["MD", "PA", "PT"].includes(p?.type) ? p.type : "MD",
          days: toNumber(p?.days, 0)
        }))
        .filter((p) => p.name || p.days > 0)
        .slice(0, 50);

      await forecastTable.upsertEntity({
        partitionKey: entity,
        rowKey: monthKey,
        entity,
        monthKey,
        monthLabel: monthKeyToLabel(monthKey),
        mdPtoDays,
        paPtoDays,
        ptPtoDays,
        providerBreakdownJson: JSON.stringify(providerBreakdown),
        missedMdVisits:    forecast.missedMdVisits,
        missedPaVisits:    forecast.missedPaVisits,
        missedPtVisits:    forecast.missedPtVisits,
        totalMissedVisits: forecast.totalMissedVisits,
        clinicalVisitsPerDay: rates.clinicalVisitsPerDay,
        perProviderClinicalRate: forecast.perProviderClinicalRate,
        perProviderPtRate: forecast.perProviderPtRate,
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
        mdPtoDays,
        paPtoDays,
        ptPtoDays,
        forecast,
        rates,
        providerSettings
      });
    }

    // ── GET: load forecasts + rates for all entities, rolling 3 months ───────
    const monthKeys = getRollingMonthKeys();

    // Determine which entities to return
    const scopeEntities = scopeEntitiesToAccess(access, ENTITIES);

    const results = await Promise.all(
      scopeEntities.map(async (entity) => {
        try {
          const rates = await computeEntityRates(regionTable, entity);
          const providerSettings = await getProviderSettings(settingsTable, entity);

          const months = await Promise.all(
            monthKeys.map(async (monthKey) => {
              try {
                const saved = await getForecastRecord(forecastTable, entity, monthKey);

                const mdPtoDays = toNumber(saved?.mdPtoDays, 0);
                const paPtoDays = toNumber(saved?.paPtoDays, 0);
                const ptPtoDays = toNumber(saved?.ptPtoDays, 0);
                const forecast = computeForecast(mdPtoDays, paPtoDays, ptPtoDays, rates, providerSettings);

                return {
                  monthKey,
                  monthLabel: monthKeyToLabel(monthKey),
                  mdPtoDays,
                  paPtoDays,
                  ptPtoDays,
                  providerBreakdown: safeParseJson(saved?.providerBreakdownJson, []),
                  missedMdVisits:    saved ? toNumber(saved.missedMdVisits)    : forecast.missedMdVisits,
                  missedPaVisits:    saved ? toNumber(saved.missedPaVisits)    : forecast.missedPaVisits,
                  missedPtVisits:    saved ? toNumber(saved.missedPtVisits)    : forecast.missedPtVisits,
                  totalMissedVisits: saved ? toNumber(saved.totalMissedVisits) : forecast.totalMissedVisits,
                  savedBy: saved?.savedBy || null,
                  savedAt: saved?.savedAt || null,
                  hasSavedEntry: !!saved
                };
              } catch {
                return {
                  monthKey,
                  monthLabel: monthKeyToLabel(monthKey),
                  mdPtoDays: 0, paPtoDays: 0, ptPtoDays: 0,
                  providerBreakdown: [],
                  missedMdVisits: 0, missedPaVisits: 0, missedPtVisits: 0,
                  totalMissedVisits: 0,
                  savedBy: null, savedAt: null, hasSavedEntry: false
                };
              }
            })
          );

          return {
            entity,
            rates,
            providerSettings,
            months,
            quarterlyTotalMissed: months.reduce((sum, m) => sum + toNumber(m.totalMissedVisits), 0)
          };
        } catch {
          const emptyMonth = (mk) => ({
            monthKey: mk, monthLabel: monthKeyToLabel(mk),
            mdPtoDays: 0, paPtoDays: 0, ptPtoDays: 0,
            missedMdVisits: 0, missedPaVisits: 0, missedPtVisits: 0,
            totalMissedVisits: 0, savedBy: null, savedAt: null, hasSavedEntry: false
          });
          return {
            entity,
            rates: { visitsPerDay: 0, clinicalVisitsPerDay: 0, surgeriesPerDay: 0, weeksUsed: 0 },
            providerSettings: PROVIDER_DEFAULTS[entity] || { mdCount: 1, paCount: 0, ptCount: 0 },
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
    try { context?.log?.error?.("pto-forecast failed", error); } catch (_) {}
    return respond(500, { ok: false, error: "Failed to process PTO forecast" });
  }
};
