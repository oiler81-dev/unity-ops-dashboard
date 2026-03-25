const { getUserFromRequest } = require("../shared/auth");
const { resolveAccess } = require("../shared/permissions");
const { ok, badRequest, forbidden, serverError } = require("../shared/response");
const { ensureTable } = require("../shared/table");
const { WEEKLY_TABLE } = require("../shared/constants");

const ENTITIES = ["LAOSS", "NES", "SpineOne", "MRO"];

module.exports = async function (context, req) {
  try {
    const user = getUserFromRequest(req);
    const access = resolveAccess(user);

    if (!access.allowed) {
      return forbidden();
    }

    const weekEnding = req.query && req.query.weekEnding ? req.query.weekEnding : null;

    if (!weekEnding) {
      return badRequest("Missing weekEnding");
    }

    const client = await ensureTable(WEEKLY_TABLE);
    const regions = [];

    for (const entity of ENTITIES) {
      try {
        const record = await client.getEntity(entity, weekEnding);

        if (record.status === "approved") {
          regions.push({
            entity,
            weekEnding,
            status: record.status || "draft",
            visitVolume: Number(record.visitVolume || 0),
            callVolume: Number(record.callVolume || 0),
            newPatients: Number(record.newPatients || 0),
            noShowRate: Number(record.noShowRate || 0),
            cancellationRate: Number(record.cancellationRate || 0),
            abandonedCallRate: Number(record.abandonedCallRate || 0),
            submittedBy: record.submittedBy || null,
            submittedAt: record.submittedAt || null,
            approvedBy: record.approvedBy || null,
            approvedAt: record.approvedAt || null,
            updatedBy: record.updatedBy || null,
            updatedAt: record.updatedAt || null
          });
        }
      } catch (error) {
        if (error?.statusCode !== 404) {
          throw error;
        }
      }
    }

    const totals = {
      visitVolume: 0,
      callVolume: 0,
      newPatients: 0
    };

    regions.forEach((region) => {
      totals.visitVolume += region.visitVolume;
      totals.callVolume += region.callVolume;
      totals.newPatients += region.newPatients;
    });

    return ok({
      ok: true,
      weekEnding,
      entityCount: regions.length,
      totals,
      regions
    });
  } catch (error) {
    context.log.error("executive-summary failed", error);
    return serverError(error, "Failed to load executive summary");
  }
};
