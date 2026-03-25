const { getUserFromRequest } = require("../shared/auth");
const { resolveAccess } = require("../shared/permissions");
const { ok, forbidden, serverError } = require("../shared/response");
const { ensureTable } = require("../shared/table");
const { WEEKLY_TABLE } = require("../shared/constants");

const ENTITIES = ["LAOSS", "NES", "SpineOne", "MRO"];

module.exports = async function (context, req) {
  try {
    const user = getUserFromRequest(req);
    const access = resolveAccess(user);

    if (!access.allowed) return forbidden();

    const weekEnding = req.query.weekEnding;

    if (!weekEnding) {
      return ok({ ok: false, error: "Missing weekEnding" });
    }

    const client = await ensureTable(WEEKLY_TABLE);

    const results = [];

    for (const entity of ENTITIES) {
      try {
        const record = await client.getEntity(entity, weekEnding);

        if (record.status === "approved") {
          results.push({
            entity,
            ...record
          });
        }
      } catch {
        // ignore missing
      }
    }

    const totals = {
      visitVolume: 0,
      callVolume: 0,
      newPatients: 0
    };

    results.forEach((r) => {
      totals.visitVolume += Number(r.visitVolume || 0);
      totals.callVolume += Number(r.callVolume || 0);
      totals.newPatients += Number(r.newPatients || 0);
    });

    return ok({
      ok: true,
      weekEnding,
      entityCount: results.length,
      totals,
      regions: results
    });

  } catch (error) {
    context.log.error(error);
    return serverError(error);
  }
};
