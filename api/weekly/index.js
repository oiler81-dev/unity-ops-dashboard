const { getUserFromRequest } = require("../shared/auth");
const { resolveAccess } = require("../shared/permissions");
const { getTableClient } = require("../shared/table");

const REGION_TABLE = "WeeklyRegionData";

function mapValues(valuesJson) {
  if (!valuesJson) return {};

  let v = {};
  try {
    v = typeof valuesJson === "string" ? JSON.parse(valuesJson) : valuesJson;
  } catch {
    return {};
  }

  return {
    visitVolume: v.totalVisits ?? null,
    callVolume: v.totalCalls ?? null,
    newPatients: v.npActual ?? null,

    noShowRate: v.noShowRate ?? null,
    cancellationRate: v.cancellationRate ?? null,
    abandonedCallRate: v.abandonmentRate ?? null
  };
}

module.exports = async function (context, req) {
  try {
    const user = getUserFromRequest(req);
    const access = resolveAccess(user);

    const entity = req.query.entity;
    const weekEnding = req.query.weekEnding;

    if (!entity || !weekEnding) {
      return {
        status: 400,
        body: {
          ok: false,
          error: "Missing entity or weekEnding"
        }
      };
    }

    const table = getTableClient(REGION_TABLE);

    const entityData = await table.getEntity(entity, weekEnding);

    if (!entityData) {
      return {
        status: 200,
        body: {
          ok: true,
          found: false
        }
      };
    }

    const mappedData = mapValues(entityData.valuesJson);

    return {
      status: 200,
      body: {
        ok: true,
        found: true,
        entity,
        weekEnding,
        data: mappedData,
        status: entityData.status,
        updatedAt: entityData.updatedAt
      }
    };
  } catch (error) {
    context.log.error("weekly get failed", error);

    return {
      status: 500,
      body: {
        ok: false,
        error: "Failed to load weekly data",
        details: error.message
      }
    };
  }
};
