const { getUserFromRequest } = require("../shared/auth");
const { resolveAccess } = require("../shared/permissions");
const { getTableClient } = require("../shared/table");

const REGION_TABLE = "WeeklyRegionData";

function parseValuesJson(valuesJson) {
  if (!valuesJson) return {};
  try {
    return typeof valuesJson === "string" ? JSON.parse(valuesJson) : valuesJson;
  } catch {
    return {};
  }
}

module.exports = async function (context, req) {
  try {
    const user = getUserFromRequest(req);
    const access = resolveAccess(user);

    const entity = String(req.query.entity || "").trim();
    const weekEnding = String(req.query.weekEnding || "").trim();

    if (!entity || !weekEnding) {
      return {
        status: 400,
        body: {
          ok: false,
          error: "Missing entity or weekEnding"
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

    let record = null;
    try {
      record = await table.getEntity(entity, weekEnding);
    } catch (err) {
      if (err.statusCode !== 404) {
        throw err;
      }
    }

    return {
      status: 200,
      body: {
        ok: true,
        found: !!record,
        entity,
        weekEnding,
        values: record ? parseValuesJson(record.valuesJson) : {},
        source: record?.source || null,
        status: record?.status || null,
        updatedAt: record?.updatedAt || record?.importedAt || null
      }
    };
  } catch (error) {
    context.log.error("weekly GET failed", error);

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
