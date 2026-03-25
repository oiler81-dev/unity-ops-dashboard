const { getUserFromRequest } = require("../shared/auth");
const { resolveAccess } = require("../shared/permissions");
const { ok, forbidden, serverError, badRequest } = require("../shared/response");
const { getTableClient } = require("../shared/table");

const BUDGET_TABLE = "BudgetReferenceData";

module.exports = async function (context, req) {
  try {
    const user = getUserFromRequest(req);
    const access = resolveAccess(user);

    if (!access.allowed) {
      return forbidden();
    }

    const yearMonth = String(req.query.yearMonth || "").trim();
    const entity = String(req.query.entity || "").trim();

    if (!yearMonth) {
      return badRequest("Missing yearMonth");
    }

    const client = getTableClient(BUDGET_TABLE);

    if (entity && entity !== "ALL") {
      const row = await client.getEntity(entity, yearMonth);
      return ok({
        ok: true,
        yearMonth,
        items: [row]
      });
    }

    const items = [];
    const filter = `RowKey eq '${yearMonth.replace(/'/g, "''")}'`;

    for await (const row of client.listEntities({
      queryOptions: { filter }
    })) {
      items.push(row);
    }

    return ok({
      ok: true,
      yearMonth,
      items
    });
  } catch (error) {
    if (error?.statusCode === 404) {
      return ok({
        ok: true,
        items: []
      });
    }

    context.log.error("budget failed", error);
    return serverError(error, "Failed to load budget data");
  }
};
