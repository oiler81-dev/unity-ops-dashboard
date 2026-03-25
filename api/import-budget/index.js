// api/import-budget/index.js

const XLSX = require("xlsx");
const { getTableClient } = require("../shared/table");
const { resolveAccessFromRequest } = require("../shared/access");
const {
  normalizeMonthLabel,
  monthLabelToMonthKey,
  getWorkingDaysForMonth,
  normalizeEntityLabel,
  safeNumber
} = require("../shared/budget");

const BUDGET_TABLE = "BudgetData";

const VISIT_BUDGET_SHEET_NAMES = [
  "Volume_Revenue Budget",
  "Volume Revenue Budget",
  "Budget V. Monthly Results"
];

const NP_BUDGET_SHEET_NAMES = [
  "New Patient Budget",
  "New Patients Budget",
  "Budget V. Monthly Results"
];

function pickFirstSheet(workbook, names) {
  for (const name of names) {
    if (workbook.Sheets[name]) {
      return { name, sheet: workbook.Sheets[name] };
    }
  }
  return { name: null, sheet: null };
}

function sheetRows(sheet) {
  return XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null, blankrows: false });
}

function normalizeText(value) {
  return String(value || "").trim();
}

function isMonthCell(value) {
  const m = normalizeMonthLabel(value);
  return !!m;
}

function upsertPayload(base, sourceSheet, importFileName) {
  const now = new Date().toISOString();
  return {
    ...base,
    sourceSheet,
    importFileName,
    updatedAt: now,
    importedAt: now
  };
}

function parseBudgetVMonthlyResultsVisitBudgets(sheet) {
  const rows = sheetRows(sheet);
  const items = [];

  for (let i = 0; i < rows.length; i += 1) {
    const row = rows[i] || [];
    const monthLabel = normalizeMonthLabel(row[0]);
    const entity = normalizeEntityLabel(row[1]);

    if (!monthLabel || !entity) {
      continue;
    }

    const monthKey = monthLabelToMonthKey(monthLabel);
    if (!monthKey) {
      continue;
    }

    const np = safeNumber(row[2]);
    const established = safeNumber(row[3]);
    const pt = safeNumber(row[4]);
    const surgery = safeNumber(row[5]);

    const visitBudgetMonthly = [np, established, pt, surgery].reduce((sum, value) => {
      return sum + (Number.isFinite(value) ? value : 0);
    }, 0);

    items.push({
      entity,
      monthKey,
      monthLabel,
      visitBudgetMonthly,
      workingDaysInMonth: getWorkingDaysForMonth(monthKey)
    });
  }

  return items;
}

function parseBudgetVMonthlyResultsNpBudgets(sheet) {
  const rows = sheetRows(sheet);
  const items = [];

  for (let i = 0; i < rows.length; i += 1) {
    const row = rows[i] || [];
    const monthLabel = normalizeMonthLabel(row[0]);
    const entity = normalizeEntityLabel(row[1]);

    if (!monthLabel || !entity) {
      continue;
    }

    const monthKey = monthLabelToMonthKey(monthLabel);
    if (!monthKey) {
      continue;
    }

    const newPatientsBudgetMonthly = safeNumber(row[2]);
    if (!Number.isFinite(newPatientsBudgetMonthly)) {
      continue;
    }

    items.push({
      entity,
      monthKey,
      monthLabel,
      newPatientsBudgetMonthly,
      workingDaysInMonth: getWorkingDaysForMonth(monthKey)
    });
  }

  return items;
}

function detectHeaderMap(headerRow) {
  const map = {};
  for (let i = 0; i < headerRow.length; i += 1) {
    const key = normalizeText(headerRow[i]).toLowerCase().replace(/\s+/g, " ");
    if (!key) continue;
    map[key] = i;
  }
  return map;
}

function parseExpectedVisitBudgetSheet(sheet) {
  const rows = sheetRows(sheet);
  let headerIndex = -1;
  let headerMap = null;

  for (let i = 0; i < Math.min(rows.length, 20); i += 1) {
    const row = rows[i] || [];
    const map = detectHeaderMap(row);

    const hasMonth = Object.keys(map).some((k) => k === "month");
    const hasEntity = Object.keys(map).some((k) => k === "entity" || k === "region" || k === "practice");
    const hasCategory = Object.keys(map).some((k) => k.includes("category"));
    const hasBudget = Object.keys(map).some((k) => k.includes("budget"));

    if (hasMonth && hasEntity && hasBudget) {
      headerIndex = i;
      headerMap = map;
      break;
    }
  }

  if (headerIndex < 0 || !headerMap) {
    return [];
  }

  const monthIdx = headerMap.month;
  const entityIdx = headerMap.entity ?? headerMap.region ?? headerMap.practice;
  const categoryIdx = headerMap.category ?? headerMap["visit category"] ?? headerMap["volume category"];
  const budgetIdx =
    headerMap["budget"] ??
    headerMap["volume budget"] ??
    headerMap["visit budget"] ??
    headerMap["monthly budget"] ??
    headerMap["budget amount"];

  if (
    monthIdx == null ||
    entityIdx == null ||
    budgetIdx == null
  ) {
    return [];
  }

  const grouped = new Map();

  for (let i = headerIndex + 1; i < rows.length; i += 1) {
    const row = rows[i] || [];
    const monthLabel = normalizeMonthLabel(row[monthIdx]);
    const entity = normalizeEntityLabel(row[entityIdx]);
    const category = normalizeText(row[categoryIdx]).toLowerCase();
    const budget = safeNumber(row[budgetIdx]);

    if (!monthLabel || !entity || !Number.isFinite(budget)) {
      continue;
    }

    const monthKey = monthLabelToMonthKey(monthLabel);
    if (!monthKey) continue;

    const key = `${entity}|${monthKey}`;
    if (!grouped.has(key)) {
      grouped.set(key, {
        entity,
        monthKey,
        monthLabel,
        visitBudgetMonthly: 0,
        workingDaysInMonth: getWorkingDaysForMonth(monthKey)
      });
    }

    const item = grouped.get(key);

    if (!category || category === "total" || category === "total visits") {
      item.visitBudgetMonthly = budget;
    } else {
      item.visitBudgetMonthly += budget;
    }
  }

  return Array.from(grouped.values()).filter((item) => Number.isFinite(item.visitBudgetMonthly));
}

function parseExpectedNpBudgetSheet(sheet) {
  const rows = sheetRows(sheet);
  let headerIndex = -1;
  let headerMap = null;

  for (let i = 0; i < Math.min(rows.length, 20); i += 1) {
    const row = rows[i] || [];
    const map = detectHeaderMap(row);

    const hasMonth = Object.keys(map).some((k) => k === "month");
    const hasEntity = Object.keys(map).some((k) => k === "entity" || k === "region" || k === "practice");
    const hasBudget = Object.keys(map).some((k) => k.includes("budget"));

    if (hasMonth && hasEntity && hasBudget) {
      headerIndex = i;
      headerMap = map;
      break;
    }
  }

  if (headerIndex < 0 || !headerMap) {
    return [];
  }

  const monthIdx = headerMap.month;
  const entityIdx = headerMap.entity ?? headerMap.region ?? headerMap.practice;
  const budgetIdx =
    headerMap["budget"] ??
    headerMap["new patient budget"] ??
    headerMap["np budget"] ??
    headerMap["monthly budget"] ??
    headerMap["budget amount"];

  if (monthIdx == null || entityIdx == null || budgetIdx == null) {
    return [];
  }

  const items = [];

  for (let i = headerIndex + 1; i < rows.length; i += 1) {
    const row = rows[i] || [];
    const monthLabel = normalizeMonthLabel(row[monthIdx]);
    const entity = normalizeEntityLabel(row[entityIdx]);
    const newPatientsBudgetMonthly = safeNumber(row[budgetIdx]);

    if (!monthLabel || !entity || !Number.isFinite(newPatientsBudgetMonthly)) {
      continue;
    }

    const monthKey = monthLabelToMonthKey(monthLabel);
    if (!monthKey) continue;

    items.push({
      entity,
      monthKey,
      monthLabel,
      newPatientsBudgetMonthly,
      workingDaysInMonth: getWorkingDaysForMonth(monthKey)
    });
  }

  return items;
}

async function upsertBudgetRows(table, visitRows, npRows, sourceMeta) {
  const merged = new Map();

  for (const row of visitRows) {
    const key = `${row.entity}|${row.monthKey}`;
    merged.set(key, {
      partitionKey: row.entity,
      rowKey: row.monthKey,
      entity: row.entity,
      monthKey: row.monthKey,
      monthLabel: row.monthLabel,
      workingDaysInMonth: row.workingDaysInMonth,
      visitBudgetMonthly: row.visitBudgetMonthly
    });
  }

  for (const row of npRows) {
    const key = `${row.entity}|${row.monthKey}`;
    const existing = merged.get(key) || {
      partitionKey: row.entity,
      rowKey: row.monthKey,
      entity: row.entity,
      monthKey: row.monthKey,
      monthLabel: row.monthLabel,
      workingDaysInMonth: row.workingDaysInMonth
    };

    existing.newPatientsBudgetMonthly = row.newPatientsBudgetMonthly;
    merged.set(key, existing);
  }

  let written = 0;

  for (const item of merged.values()) {
    await table.upsertEntity(
      upsertPayload(
        item,
        sourceMeta.sourceSheet,
        sourceMeta.importFileName
      )
    );
    written += 1;
  }

  return written;
}

module.exports = async function (context, req) {
  try {
    const access = resolveAccessFromRequest(req);

    if (!access.isAdmin) {
      context.res = {
        status: 403,
        body: {
          ok: false,
          error: "Admin only"
        }
      };
      return;
    }

    const body = req.body || {};
    const fileBase64 = normalizeText(body.fileBase64);
    const fileName = normalizeText(body.fileName) || "budget-workbook.xlsx";

    if (!fileBase64) {
      context.res = {
        status: 400,
        body: {
          ok: false,
          error: "Missing fileBase64"
        }
      };
      return;
    }

    const buffer = Buffer.from(fileBase64, "base64");
    const workbook = XLSX.read(buffer, { type: "buffer" });

    const visitSheetInfo = pickFirstSheet(workbook, VISIT_BUDGET_SHEET_NAMES);
    const npSheetInfo = pickFirstSheet(workbook, NP_BUDGET_SHEET_NAMES);

    if (!visitSheetInfo.sheet && !npSheetInfo.sheet) {
      context.res = {
        status: 400,
        body: {
          ok: false,
          error: "No recognized budget sheets found",
          workbookSheets: workbook.SheetNames
        }
      };
      return;
    }

    let visitRows = [];
    let npRows = [];

    if (visitSheetInfo.sheet) {
      if (visitSheetInfo.name === "Budget V. Monthly Results") {
        visitRows = parseBudgetVMonthlyResultsVisitBudgets(visitSheetInfo.sheet);
      } else {
        visitRows = parseExpectedVisitBudgetSheet(visitSheetInfo.sheet);
      }
    }

    if (npSheetInfo.sheet) {
      if (npSheetInfo.name === "Budget V. Monthly Results") {
        npRows = parseBudgetVMonthlyResultsNpBudgets(npSheetInfo.sheet);
      } else {
        npRows = parseExpectedNpBudgetSheet(npSheetInfo.sheet);
      }
    }

    const table = getTableClient(BUDGET_TABLE);

    const written = await upsertBudgetRows(
      table,
      visitRows,
      npRows,
      {
        sourceSheet: [visitSheetInfo.name, npSheetInfo.name].filter(Boolean).join(" + "),
        importFileName: fileName
      }
    );

    context.res = {
      status: 200,
      headers: { "Content-Type": "application/json" },
      body: {
        ok: true,
        message: "Budget import completed",
        fileName,
        workbookSheets: workbook.SheetNames,
        visitBudgetSheet: visitSheetInfo.name,
        newPatientBudgetSheet: npSheetInfo.name,
        visitRowsParsed: visitRows.length,
        newPatientRowsParsed: npRows.length,
        rowsWritten: written
      }
    };
  } catch (error) {
    context.log.error("import-budget failed", error);
    context.res = {
      status: 500,
      body: {
        ok: false,
        error: "Failed to import budget workbook",
        details: error.message
      }
    };
  }
};
