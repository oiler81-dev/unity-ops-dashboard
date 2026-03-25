const { getUserFromRequest } = require("../shared/auth");
const { resolveAccess } = require("../shared/permissions");
const { getTableClient } = require("../shared/table");

const REGION_TABLE = "WeeklyRegionData";
const TARGETS_TABLE = "ReferenceData";
const ENTITIES = ["LAOSS", "NES", "SpineOne", "MRO"];

function toNumber(value) {
  if (value == null || value === "") return null;
  const n = Number(value);
  return Number.isFinite(n) ? n : null;
}

function parseJson(value, fallback = {}) {
  if (!value) return fallback;
  try {
    return typeof value === "string" ? JSON.parse(value) : value;
  } catch {
    return fallback;
  }
}

function normalizeStatus(status) {
  const s = String(status || "").trim().toLowerCase();
  if (s === "approved") return "Approved";
  if (s === "submitted") return "submitted";
  if (s === "draft") return "draft";
  return status || "draft";
}

function average(values) {
  if (!values.length) return 0;
  return values.reduce((sum, v) => sum + v, 0) / values.length;
}

function pctVariance(actual, target) {
  const a = toNumber(actual) || 0;
  const t = toNumber(target) || 0;
  if (!t) return null;
  return ((a - t) / t) * 100;
}

function buildTrend(actual, previous) {
  const a = toNumber(actual) || 0;
  const p = toNumber(previous) || 0;
  const diff = a - p;

  return {
    current: a,
    previous: p,
    diff,
    direction: diff > 0 ? "up" : diff < 0 ? "down" : "flat"
  };
}

function getPreviousWeekEnding(weekEnding) {
  const d = new Date(`${weekEnding}T12:00:00Z`);
  d.setUTCDate(d.getUTCDate() - 7);
  return d.toISOString().split("T")[0];
}

function mapRecord(record) {
  if (!record) return null;

  const values = parseJson(record.valuesJson, {});
  const visitVolume = toNumber(values.totalVisits) ?? toNumber(record.visitVolume) ?? 0;
  const callVolume = toNumber(values.totalCalls) ?? toNumber(record.callVolume) ?? 0;
  const newPatients = toNumber(values.npActual) ?? toNumber(record.newPatients) ?? 0;
  const noShowRate = toNumber(values.noShowRate) ?? toNumber(record.noShowRate) ?? 0;
  const cancellationRate =
    toNumber(values.cancellationRate) ?? toNumber(record.cancellationRate) ?? 0;
  const abandonedCallRate =
    toNumber(values.abandonmentRate) ?? toNumber(record.abandonedCallRate) ?? 0;

  return {
    entity: record.entity || record.partitionKey,
    weekEnding: record.weekEnding || record.rowKey,
    status: normalizeStatus(record.status),
    visitVolume,
    callVolume,
    newPatients,
    noShowRate,
    cancellationRate,
    abandonedCallRate,
    source: record.source || null,
    updatedAt: record.updatedAt || record.importedAt || record.approvedAt || null,
    rawValues: values
  };
}

async function getWeekRecord(table, entity, weekEnding) {
  try {
    const record = await table.getEntity(entity, weekEnding);
    return mapRecord(record);
  } catch (err) {
    if (err.statusCode === 404) return null;
    throw err;
  }
}

async function getEntityTargets(table, entity) {
  try {
    const record = await table.getEntity("Targets", entity);
    const values = parseJson(record.valuesJson, {});
    return {
      visitTarget: toNumber(values.visitTarget),
      callTarget: toNumber(values.callTarget),
      newPatientTarget: toNumber(values.newPatientTarget)
    };
  } catch (err) {
    if (err.statusCode === 404) {
      return {
        visitTarget: null,
        callTarget: null,
        newPatientTarget: null
      };
    }
    throw err;
  }
}

function buildAlertRow(entityRow) {
  const alerts = [];

  if (entityRow.status !== "Approved") {
    alerts.push({
      entity: entityRow.entity,
      severity: "yellow",
      message: `${entityRow.entity} is not approved for this week`
    });
  }

  if ((entityRow.noShowRate || 0) >= 6) {
    alerts.push({
      entity: entityRow.entity,
      severity: "red",
      message: `${entityRow.entity} no-show rate is elevated`
    });
  }

  if ((entityRow.cancellationRate || 0) >= 8) {
    alerts.push({
      entity: entityRow.entity,
      severity: "red",
      message: `${entityRow.entity} cancellation rate is elevated`
    });
  }

  if ((entityRow.abandonedCallRate || 0) >= 10) {
    alerts.push({
      entity: entityRow.entity,
      severity: "red",
      message: `${entityRow.entity} abandoned call rate is elevated`
    });
  }

  return alerts;
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

    const previousWeekEnding = getPreviousWeekEnding(weekEnding);
    const regionTable = getTableClient(REGION_TABLE);
    const targetsTable = getTableClient(TARGETS_TABLE);

    const entityRows = [];
    const alerts = [];

    for (const entity of ENTITIES) {
      const current = await getWeekRecord(regionTable, entity, weekEnding);
      const previous = await getWeekRecord(regionTable, entity, previousWeekEnding);
      const targets = await getEntityTargets(targetsTable, entity);

      if (!current) {
        entityRows.push({
          entity,
          weekEnding,
          status: "missing",
          visitVolume: 0,
          callVolume: 0,
          newPatients: 0,
          noShowRate: 0,
          cancellationRate: 0,
          abandonedCallRate: 0,
          targets,
          variance: {
            visitVariancePct: null,
            callVariancePct: null,
            newPatientVariancePct: null
          },
          trends: {
            visits: buildTrend(0, previous?.visitVolume || 0),
            calls: buildTrend(0, previous?.callVolume || 0),
            newPatients: buildTrend(0, previous?.newPatients || 0)
          }
        });

        alerts.push({
          entity,
          severity: "yellow",
          message: `${entity} has no record for ${weekEnding}`
        });

        continue;
      }

      const row = {
        ...current,
        targets,
        variance: {
          visitVariancePct: pctVariance(current.visitVolume, targets.visitTarget),
          callVariancePct: pctVariance(current.callVolume, targets.callTarget),
          newPatientVariancePct: pctVariance(current.newPatients, targets.newPatientTarget)
        },
        trends: {
          visits: buildTrend(current.visitVolume, previous?.visitVolume || 0),
          calls: buildTrend(current.callVolume, previous?.callVolume || 0),
          newPatients: buildTrend(current.newPatients, previous?.newPatients || 0)
        }
      };

      entityRows.push(row);
      alerts.push(...buildAlertRow(row));
    }

    const approvedRows = entityRows.filter((r) => r.status === "Approved");

    const totals = {
      visitVolume: approvedRows.reduce((sum, r) => sum + (r.visitVolume || 0), 0),
      callVolume: approvedRows.reduce((sum, r) => sum + (r.callVolume || 0), 0),
      newPatients: approvedRows.reduce((sum, r) => sum + (r.newPatients || 0), 0)
    };

    const averages = {
      noShowRate: average(approvedRows.map((r) => r.noShowRate || 0)),
      cancellationRate: average(approvedRows.map((r) => r.cancellationRate || 0)),
      abandonedCallRate: average(approvedRows.map((r) => r.abandonedCallRate || 0))
    };

    const kpis = [
      {
        key: "approvedRegions",
        label: "Approved Regions",
        value: approvedRows.length,
        meta: `${approvedRows.length} region(s) approved`
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
        value: averages.noShowRate,
        format: "percent"
      },
      {
        key: "avgCancellationRate",
        label: "Avg Cancel %",
        value: averages.cancellationRate,
        format: "percent"
      },
      {
        key: "avgAbandonedCallRate",
        label: "Avg Abandoned %",
        value: averages.abandonedCallRate,
        format: "percent"
      }
    ];

    return {
      status: 200,
      body: {
        ok: true,
        weekEnding,
        previousWeekEnding,
        entityCount: approvedRows.length,
        totals,
        averages,
        kpis,
        entities: entityRows,
        alerts
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
