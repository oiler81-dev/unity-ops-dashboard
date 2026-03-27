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

function toBudgetItem(row) {
  return {
    entity: safeText(row.entity || row.partitionKey || row.PartitionKey),
    monthKey: safeText(row.monthKey || row.rowKey || row.RowKey),
    monthLabel: safeText(row.monthLabel),
    visitBudgetMonthly: normalizeNumber(row.visitBudgetMonthly, 0),
    newPatientsBudgetMonthly: normalizeNumber(row.newPatientsBudgetMonthly, 0),
    workingDaysInMonth: normalizeNumber(row.workingDaysInMonth, 0),
    source: safeText(row.source),
    importedAt: safeText(row.importedAt),
    updatedAt: safeText(row.updatedAt)
  };
}

module.exports = async function (context, req) {
  try {
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
      items = rows
        .map(toBudgetItem)
        .sort((a, b) => a.monthKey.localeCompare(b.monthKey));
    } else if (derivedMonthKey) {
      const rows = await table.query(
        `RowKey eq '${derivedMonthKey.replace(/'/g, "''")}'`
      );
      items = rows
        .map(toBudgetItem)
        .sort((a, b) => a.entity.localeCompare(b.entity));
    } else if (includeAll) {
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

    context.res = {
      status: 200,
      headers: { "Content-Type": "application/json" },
      body: {
        ok: true,
        entity: entity || null,
        monthKey: derivedMonthKey || null,
        count: items.length,
        item: items.length === 1 ? items[0] : null,
        items
      }
    };
  } catch (error) {
    context.log.error("budget failed", error);

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
