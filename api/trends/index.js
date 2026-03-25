const { getUserFromRequest } = require("../shared/auth");
const { resolveAccess, canAccessEntity } = require("../shared/permissions");
const { ok, badRequest, forbidden, serverError } = require("../shared/response");
const { ensureTable } = require("../shared/table");
const { WEEKLY_TABLE } = require("../shared/constants");

module.exports = async function (context, req) {
  try {
    const user = getUserFromRequest(req);
    const access = resolveAccess(user);

    if (!access.allowed) {
      return forbidden();
    }

    const entity = req.query && req.query.entity ? req.query.entity : null;
    const limitRaw = req.query && req.query.limit ? Number(req.query.limit) : 12;
    const limit = Number.isFinite(limitRaw) && limitRaw > 0 ? Math.min(limitRaw, 52) : 12;

    if (!entity) {
      return badRequest("Missing entity");
    }

    if (!canAccessEntity(access, entity)) {
      return forbidden("You cannot view this entity");
    }

    const client = await ensureTable(WEEKLY_TABLE);

    const rows = [];
    const filter = `PartitionKey eq '${entity}'`;

    for await (const record of client.listEntities({
      queryOptions: { filter }
    })) {
      rows.push({
        entity,
        weekEnding: record.rowKey,
        status: record.status || "draft",
        visitVolume: Number(record.visitVolume || 0),
        callVolume: Number(record.callVolume || 0),
        newPatients: Number(record.newPatients || 0),
        noShowRate: Number(record.noShowRate || 0),
        cancellationRate: Number(record.cancellationRate || 0),
        abandonedCallRate: Number(record.abandonedCallRate || 0),
        updatedBy: record.updatedBy || null,
        updatedAt: record.updatedAt || null
      });
    }

    rows.sort((a, b) => a.weekEnding.localeCompare(b.weekEnding));

    const trimmed = rows.slice(-limit);

    return ok({
      ok: true,
      entity,
      count: trimmed.length,
      items: trimmed
    });
  } catch (error) {
    context.log.error("trends failed", error);
    return serverError(error, "Failed to load trends");
  }
};
