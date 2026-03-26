const { getUserFromRequest } = require("../shared/auth");
const { resolveAccess } = require("../shared/permissions");
const { getTableClient } = require("../shared/table");

const BUDGET_TABLE = "BudgetData";

function safeText(value) {
  return value == null ? "" : String(value).trim();
}

function normalizeNumber(value, fallback = 0) {
  const n = Number(value);
  return Number.isFinite(n) ? n : fallback;
}

function monthKeyFromWeekEnding(weekEnding) {
  const text = safeText(weekEnding);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(text)) return "";
  return text.slice(0, 7);
}

function toBudgetItem(entity) {
  return {
    entity: safeText(entity.entity || entity.partitionKey || entity.PartitionKey),
    monthKey: safeText(entity.monthKey || entity.rowKey || entity.RowKey),
    monthLabel: safeText(entity.monthLabel),
    visitBudgetMonthly: normalizeNumber(entity.visitBudgetMonthly, 0),
    newPatientsBudgetMonthly: normalizeNumber(entity.newPatientsBudgetMonthly, 0),
    workingDaysInMonth: normalizeNumber(entity.workingDaysInMonth, 0),
    source: safeText(entity.source),
    importedAt: safeText(entity.importedAt),
    updatedAt: safeText(entity.updatedAt)
  };
}

module.exports = async function (context, req) {
  try {
    const user = getUserFromRequest(req);
    const access = resolveAccess(user);

    if (!access?.authenticated) {
      context.res = {
        status: 401,
        headers: { "Content-Type": "application/json" },
        body: {
          ok: false,
          error: "Authentication required"
        }
      };
      return;
    }

    const entity = safeText(req.query?.entity);
    const monthKeyQuery = safeText(req.query?.monthKey);
    const weekEnding = safeText(req.query?.weekEnding);
    const includeAll = safeText(req.query?.all).toLowerCase() === "true";

    const derivedMonthKey = monthKeyQuery || monthKeyFromWeekEnding(weekEnding);
    const table = getTableClient(BUDGET_TABLE);

    let items = [];

    if (entity && derivedMonthKey) {
      try {
        const row = await table.getEntity(entity, derivedMonthKey);
        items = row ? [toBudgetItem(row)] : [];
      } catch (err) {
        if (err.statusCode === 404) {
          items = [];
        } else {
          throw err;
        }
      }
    } else if (entity) {
      const rows = await table.listByPartitionKey(entity);
      items = rows.map(toBudgetItem).sort((a, b) => a.monthKey.localeCompare(b.monthKey));
    } else if (derivedMonthKey) {
      const rows = await table.query(
        `RowKey eq '${derivedMonthKey.replace(/'/g, "''")}'`
      );
      items = rows.map(toBudgetItem).sort((a, b) => a.entity.localeCompare(b.entity));
    } else if (includeAll || access.isAdmin) {
      const rows = [];
      for await (const row of table.listEntities()) {
        rows.push(row);
      }
      items = rows
        .map(toBudgetItem)
        .sort((a, b) => {
          if (a.monthKey === b.monthKey) return a.entity.localeCompare(b.entity);
          return a.monthKey.localeCompare(b.monthKey);
        });
    } else {
      context.res = {
        status: 400,
        headers: { "Content-Type": "application/json" },
        body: {
          ok: false,
          error: "Provide entity, monthKey, weekEnding, or all=true"
        }
      };
      return;
    }

    const first = items[0] || null;

    context.res = {
      status: 200,
      headers: { "Content-Type": "application/json" },
      body: {
        ok: true,
        entity: entity || null,
        monthKey: derivedMonthKey || null,
        count: items.length,
        item: items.length === 1 ? first : null,
        items
      }
    };
  } catch (error) {
    context.log.error("budget-get failed", error);

    context.res = {
      status: 500,
      headers: { "Content-Type": "application/json" },
      body: {
        ok: false,
        error: "Failed to load budget data",
        details: error.message
      }
    };
  }
};
