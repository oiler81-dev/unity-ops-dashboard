const { getUserFromRequest, getUserEmail, getDisplayName } = require("../shared/auth");
const { getPermissionByEmail } = require("../shared/permissions");
const { getTableClient } = require("../shared/table");

function toNullableNumber(value) {
  if (value === null || value === undefined || value === "") return null;
  const n = Number(value);
  return Number.isFinite(n) ? n : null;
}

module.exports = async function (context, req) {
  try {
    const user = getUserFromRequest(req);
    if (!user) {
      context.res = { status: 401, body: { error: "Not authenticated" } };
      return;
    }

    const email = getUserEmail(user);
    const permission = await getPermissionByEmail(email);
    if (!permission?.isAdmin) {
      context.res = { status: 403, body: { error: "Admin access required" } };
      return;
    }

    const body = req.body || {};
    const entity = body.entity;
    const kind = body.kind;
    const rows = Array.isArray(body.rows) ? body.rows : [];
    const year = body.year;

    let tableName = null;
    if (kind === "targets") tableName = "ReferenceTargets";
    if (kind === "thresholds") tableName = "ReferenceThresholds";
    if (kind === "holidays") tableName = "ReferenceHolidays";
    if (kind === "budget") tableName = "ReferenceBudget";

    if (!tableName) {
      context.res = { status: 400, body: { error: "Invalid kind" } };
      return;
    }

    const table = getTableClient(tableName);
    const updatedBy = permission.displayName || getDisplayName(user);
    const updatedAt = new Date().toISOString();

    if (kind === "holidays") {
      if (!year) {
        context.res = { status: 400, body: { error: "year is required for holidays" } };
        return;
      }

      for (const row of rows) {
        if (!row.date) continue;
        await table.upsertEntity({
          partitionKey: String(year),
          rowKey: row.date,
          holidayName: row.holidayName || "",
          updatedBy,
          updatedAt
        });
      }

      context.res = {
        status: 200,
        body: { ok: true, message: "holidays saved successfully" }
      };
      return;
    }

    if (!entity) {
      context.res = { status: 400, body: { error: "entity is required" } };
      return;
    }

    for (const row of rows) {
      if (kind === "targets") {
        const metricKey = row.metricKey;
        if (!metricKey) continue;

        await table.upsertEntity({
          partitionKey: entity,
          rowKey: metricKey,
          label: row.label || metricKey,
          targetValue: toNullableNumber(row.targetValue),
          updatedBy,
          updatedAt
        });
      }

      if (kind === "thresholds") {
        const metricKey = row.metricKey;
        if (!metricKey) continue;

        await table.upsertEntity({
          partitionKey: entity,
          rowKey: metricKey,
          comparisonType: row.comparisonType || "higher_better",
          greenMin: toNullableNumber(row.greenMin),
          yellowMin: toNullableNumber(row.yellowMin),
          greenMax: toNullableNumber(row.greenMax),
          yellowMax: toNullableNumber(row.yellowMax),
          updatedBy,
          updatedAt
        });
      }

      if (kind === "budget") {
        const monthKey = row.monthKey;
        if (!monthKey) continue;

        await table.upsertEntity({
          partitionKey: entity,
          rowKey: monthKey,
          budgetVisits: toNullableNumber(row.budgetVisits),
          budgetRevenue: toNullableNumber(row.budgetRevenue),
          updatedBy,
          updatedAt
        });
      }
    }

    context.res = {
      status: 200,
      body: { ok: true, message: `${kind} saved successfully` }
    };
  } catch (err) {
    context.log.error(err);
    context.res = { status: 500, body: { error: err.message } };
  }
};