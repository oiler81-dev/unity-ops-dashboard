const { getUserFromRequest, getUserEmail } = require("../shared/auth");
const { getPermissionByEmail } = require("../shared/permissions");
const { getTableClient } = require("../shared/table");

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

    const entity = req.query.entity;
    const kind = req.query.kind;
    const year = req.query.year;

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

    if (kind === "holidays") {
      if (!year) {
        context.res = { status: 400, body: { error: "year is required for holidays" } };
        return;
      }

      const rows = await table.listByPartition(String(year));
      context.res = {
        status: 200,
        body: {
          kind,
          year,
          rows: rows.map((row) => ({
            date: row.rowKey || row.RowKey,
            holidayName: row.holidayName || ""
          }))
        }
      };
      return;
    }

    if (!entity) {
      context.res = { status: 400, body: { error: "entity is required" } };
      return;
    }

    const rows = await table.listByPartition(entity);

    if (kind === "targets") {
      context.res = {
        status: 200,
        body: {
          entity,
          kind,
          rows: rows.map((row) => ({
            metricKey: row.rowKey || row.RowKey,
            label: row.label || "",
            targetValue: row.targetValue ?? null
          }))
        }
      };
      return;
    }

    if (kind === "thresholds") {
      context.res = {
        status: 200,
        body: {
          entity,
          kind,
          rows: rows.map((row) => ({
            metricKey: row.rowKey || row.RowKey,
            comparisonType: row.comparisonType || "",
            greenMin: row.greenMin ?? null,
            yellowMin: row.yellowMin ?? null,
            greenMax: row.greenMax ?? null,
            yellowMax: row.yellowMax ?? null
          }))
        }
      };
      return;
    }

    if (kind === "budget") {
      context.res = {
        status: 200,
        body: {
          entity,
          kind,
          rows: rows.map((row) => ({
            monthKey: row.rowKey || row.RowKey,
            budgetVisits: row.budgetVisits ?? null,
            budgetRevenue: row.budgetRevenue ?? null
          }))
        }
      };
    }
  } catch (err) {
    context.log.error(err);
    context.res = { status: 500, body: { error: err.message } };
  }
};