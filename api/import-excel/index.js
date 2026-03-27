const XLSX = require("xlsx");
const { getUserFromRequest } = require("../shared/auth");
const { resolveAccess } = require("../shared/permissions");
const { getTableClient } = require("../shared/table");
const {
  monthLabelToMonthKey,
  getWorkingDaysForMonth,
  normalizeEntityLabel
} = require("../shared/budget");

const REGION_TABLE = "WeeklyRegionData";
const SHARED_TABLE = "SharedPageData";
const REFERENCE_TABLE = "ReferenceData";
const BUDGET_TABLE = "BudgetData";
const WORKBOOK_YEAR = 2026;
const MAIN_BUDGET_SHEET = "Budget V. Monthly Results";

const REGION_SHEET_TO_ENTITY = {
  LA: "LAOSS",
  Portland: "NES",
  Denver: "SpineOne",
  Chicago: "MRO"
};

function sheetRows(ws) {
  return XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
}

function safeText(value) {
  return value == null ? "" : String(value).trim();
}

function toNumber(value) {
  if (value == null || value === "") return null;

  if (typeof value === "number") {
    return Number.isFinite(value) ? value : null;
  }

  const text = String(value).replace(/,/g, "").trim();
  if (!text || text === "#N/A" || text === "#VALUE!" || text === "#DIV/0!" || text === "N/A") {
    return null;
  }

  const n = Number(text);
  return Number.isFinite(n) ? n : null;
}

function percentToDisplay(value) {
  const n = toNumber(value);
  if (n == null) return null;
  return n <= 1 ? Number((n * 100).toFixed(2)) : Number(n.toFixed(2));
}

function sumNumbers(values) {
  const nums = values.map(toNumber).filter((v) => v != null);
  if (!nums.length) return null;
  return nums.reduce((a, b) => a + b, 0);
}

function firstFridayOfYear(year) {
  const d = new Date(Date.UTC(year, 0, 1));
  while (d.getUTCDay() !== 5) {
    d.setUTCDate(d.getUTCDate() + 1);
  }
  return d;
}

function weekEndingFromWeekNumber(weekNumber) {
  const n = toNumber(weekNumber);
  if (!Number.isInteger(n) || n < 1) return null;

  const d = firstFridayOfYear(WORKBOOK_YEAR);
  d.setUTCDate(d.getUTCDate() + (n - 1) * 7);
  return d.toISOString().slice(0, 10);
}

function hasPositiveNumber(...values) {
  return values.some((value) => {
    const n = toNumber(value);
    return n != null && n > 0;
  });
}

function excelDateToIso(value) {
  if (!value && value !== 0) return null;

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value.toISOString().slice(0, 10);
  }

  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (!parsed) return null;
    const yyyy = String(parsed.y).padStart(4, "0");
    const mm = String(parsed.m).padStart(2, "0");
    const dd = String(parsed.d).padStart(2, "0");
    return `${yyyy}-${mm}-${dd}`;
  }

  if (typeof value === "string") {
    const trimmed = value.trim();
    if (!trimmed) return null;
    if (/^\d{4}-\d{2}-\d{2}$/.test(trimmed)) return trimmed;

    const d = new Date(trimmed);
    if (!Number.isNaN(d.getTime())) {
      return d.toISOString().slice(0, 10);
    }
  }

  return null;
}

function monthTagFromIsoDate(isoDate) {
  if (!isoDate) return null;

  const month = isoDate.slice(5, 7);
  const map = {
    "01": "Jan",
    "02": "Feb",
    "03": "Mar",
    "04": "Apr",
    "05": "May",
    "06": "Jun",
    "07": "Jul",
    "08": "Aug",
    "09": "Sep",
    "10": "Oct",
    "11": "Nov",
    "12": "Dec"
  };

  return map[month] || null;
}

function normalizeMonthFromBudgetSheet(value) {
  const text = safeText(value);
  if (!text) return "";

  const normalized = text.slice(0, 3).toLowerCase();
  const map = {
    jan: "Jan",
    feb: "Feb",
    mar: "Mar",
    apr: "Apr",
    may: "May",
    jun: "Jun",
    jul: "Jul",
    aug: "Aug",
    sep: "Sep",
    oct: "Oct",
    nov: "Nov",
    dec: "Dec"
  };

  return map[normalized] || "";
}

async function upsertRegionRecord(table, entity, weekEnding, values, meta = {}) {
  await table.upsertEntity({
    partitionKey: entity,
    rowKey: weekEnding,
    entity,
    weekEnding,
    valuesJson: JSON.stringify(values),
    source: "workbook-import",
    status: "Approved",
    submittedAt: new Date().toISOString(),
    submittedBy: "workbook-import",
    approvedAt: new Date().toISOString(),
    approvedBy: "workbook-import",
    importedAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    ...meta
  });
}

async function upsertSharedRecord(table, page, weekEnding, values, meta = {}) {
  await table.upsertEntity({
    partitionKey: page,
    rowKey: weekEnding,
    page,
    weekEnding,
    valuesJson: JSON.stringify(values),
    source: "workbook-import",
    status: "Approved",
    submittedAt: new Date().toISOString(),
    submittedBy: "workbook-import",
    approvedAt: new Date().toISOString(),
    approvedBy: "workbook-import",
    importedAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    ...meta
  });
}

async function upsertReferenceRecord(table, kind, rowKey, values, meta = {}) {
  await table.upsertEntity({
    partitionKey: kind,
    rowKey,
    kind,
    valuesJson: JSON.stringify(values),
    source: "workbook-import",
    importedAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    ...meta
  });
}

async function upsertBudgetRecord(table, entity, monthKey, values, meta = {}) {
  await table.upsertEntity({
    partitionKey: entity,
    rowKey: monthKey,
    entity,
    monthKey,
    monthLabel: values.monthLabel,
    visitBudgetMonthly: values.visitBudgetMonthly ?? 0,
    newPatientsBudgetMonthly: values.newPatientsBudgetMonthly ?? 0,
    workingDaysInMonth: values.workingDaysInMonth ?? getWorkingDaysForMonth(monthKey),
    source: "workbook-import",
    importedAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    ...meta
  });
}

function buildRegionValues(sheetName, rowIndex, row) {
  const weekNumber = toNumber(row[0]);
  const daysInPeriod = toNumber(row[1]);
  const monthTag = safeText(sheetName === "Denver" ? row[18] : row[17]);

  if (!weekNumber) {
    return { accept: false, reason: "missing-week-number", rowIndex };
  }

  const weekEnding = weekEndingFromWeekNumber(weekNumber);
  if (!weekEnding) {
    return { accept: false, reason: "invalid-week-ending", rowIndex, weekNumber };
  }

  const npActual = toNumber(row[5]);
  const establishedActual = toNumber(row[6]);
  const surgeryActual = toNumber(row[7]);
  const imagingActual = sheetName === "Denver" ? toNumber(row[8]) : null;

  let totalVisits = toNumber(row[2]);
  if (totalVisits == null) {
    totalVisits = sumNumbers(
      sheetName === "Denver"
        ? [npActual, establishedActual, surgeryActual, imagingActual]
        : [npActual, establishedActual, surgeryActual]
    );
  }

  const totalCalls = toNumber(sheetName === "Denver" ? row[9] : row[8]);
  const abandonedCalls = toNumber(sheetName === "Denver" ? row[10] : row[9]);
  const cashActual = toNumber(sheetName === "Denver" ? row[17] : row[16]);

  const hasRealData = hasPositiveNumber(
    totalVisits,
    npActual,
    establishedActual,
    surgeryActual,
    imagingActual,
    totalCalls,
    abandonedCalls
  );

  const hasUsableDays = daysInPeriod != null && daysInPeriod > 0;
  const hasMonthTag = !!monthTag;

  if (!hasUsableDays) {
    return {
      accept: false,
      reason: "days-not-positive",
      rowIndex,
      weekNumber,
      weekEnding,
      daysInPeriod,
      monthTag,
      totalVisits,
      totalCalls,
      cashActual
    };
  }

  if (!hasMonthTag) {
    return {
      accept: false,
      reason: "missing-month-tag",
      rowIndex,
      weekNumber,
      weekEnding,
      daysInPeriod,
      totalVisits,
      totalCalls,
      cashActual
    };
  }

  if (!hasRealData) {
    return {
      accept: false,
      reason: "no-real-data",
      rowIndex,
      weekNumber,
      weekEnding,
      daysInPeriod,
      monthTag,
      totalVisits,
      totalCalls,
      cashActual
    };
  }

  const values = {
    weekNumber,
    monthTag,
    daysInPeriod,
    totalVisits,
    visitsPerDay: toNumber(row[3]),
    npPerDay: toNumber(row[4]),
    npActual,
    establishedActual,
    surgeryActual,
    totalCalls,
    abandonedCalls,
    abandonmentRate: percentToDisplay(sheetName === "Denver" ? row[11] : row[10]),
    npToEstablishedConversion: percentToDisplay(sheetName === "Denver" ? row[12] : row[11]),
    npToSurgeryConversion: percentToDisplay(sheetName === "Denver" ? row[13] : row[12]),
    cashActual
  };

  if (sheetName === "Denver") {
    values.imagingActual = imagingActual;
    values.piNp = toNumber(row[19]);
    values.piCashCollection = toNumber(row[20]);
  }

  return {
    accept: true,
    rowIndex,
    weekNumber,
    weekEnding,
    values
  };
}

async function importRegionSheet(regionTable, ws, sheetName) {
  const entity = REGION_SHEET_TO_ENTITY[sheetName];
  const rows = sheetRows(ws);

  if (!entity) {
    return { imported: 0, entity: null, weekEndings: [], acceptedRows: [], rejectedRows: [] };
  }

  let imported = 0;
  const weekEndings = [];
  const acceptedRows = [];
  const rejectedRows = [];

  for (let r = 6; r < rows.length; r += 1) {
    const row = rows[r] || [];
    const parsed = buildRegionValues(sheetName, r + 1, row);

    if (!parsed.accept) {
      if (parsed.reason !== "missing-week-number") {
        rejectedRows.push(parsed);
      }
      continue;
    }

    await upsertRegionRecord(regionTable, entity, parsed.weekEnding, parsed.values, {
      importSourceSheet: sheetName,
      importWeekNumber: parsed.weekNumber,
      importMonthTag: parsed.values.monthTag
    });

    imported += 1;
    weekEndings.push(parsed.weekEnding);
    acceptedRows.push({
      rowIndex: parsed.rowIndex,
      weekNumber: parsed.weekNumber,
      weekEnding: parsed.weekEnding,
      monthTag: parsed.values.monthTag,
      daysInPeriod: parsed.values.daysInPeriod
    });
  }

  return { imported, entity, weekEndings, acceptedRows, rejectedRows };
}

function buildWorkingDaysMapFromRegionSheets(workbook) {
  const map = {};

  for (const sheetName of Object.keys(REGION_SHEET_TO_ENTITY)) {
    const ws = workbook.Sheets[sheetName];
    if (!ws) continue;

    const rows = sheetRows(ws);
    for (let r = 6; r < rows.length; r += 1) {
      const row = rows[r] || [];
      const weekNumber = toNumber(row[0]);
      const daysInPeriod = toNumber(row[1]);
      const weekEnding = weekEndingFromWeekNumber(weekNumber);

      if (!weekEnding || daysInPeriod == null || daysInPeriod <= 0) continue;

      if (!map[weekEnding] || daysInPeriod > map[weekEnding]) {
        map[weekEnding] = daysInPeriod;
      }
    }
  }

  return map;
}

function buildPtValues(rowIndex, row, workingDaysMap) {
  const chicagoWeek = toNumber(row[0]);
  const chicagoMonthTag = safeText(row[1]);
  const weekEnding = weekEndingFromWeekNumber(chicagoWeek);

  if (!chicagoWeek) {
    return { accept: false, reason: "missing-chicago-week-number", rowIndex };
  }

  if (!weekEnding) {
    return { accept: false, reason: "invalid-week-ending", rowIndex, chicagoWeek };
  }

  const denverWeek = toNumber(row[10]);
  const denverMonthTag = safeText(row[11]);
  const portlandWeek = toNumber(row[20]);
  const portlandMonthTag = safeText(row[21]);

  const chicagoScheduled = toNumber(row[2]);
  const chicagoCancels = toNumber(row[3]);
  const chicagoNoShows = toNumber(row[4]);
  const chicagoReschedules = toNumber(row[5]);
  const chicagoUnits = toNumber(row[6]);

  const denverScheduled = toNumber(row[12]);
  const denverCancels = toNumber(row[13]);
  const denverNoShows = toNumber(row[14]);
  const denverReschedules = toNumber(row[15]);
  const denverUnits = toNumber(row[16]);

  const portlandScheduled = toNumber(row[22]);
  const portlandCancels = toNumber(row[23]);
  const portlandNoShows = toNumber(row[24]);
  const portlandReschedules = toNumber(row[25]);
  const portlandUnits = toNumber(row[26]);

  const ptScheduledVisits = sumNumbers([
    chicagoScheduled,
    denverScheduled,
    portlandScheduled
  ]);

  const ptCancellations = sumNumbers([
    chicagoCancels,
    denverCancels,
    portlandCancels
  ]);

  const ptNoShows = sumNumbers([
    chicagoNoShows,
    denverNoShows,
    portlandNoShows
  ]);

  const ptReschedules = sumNumbers([
    chicagoReschedules,
    denverReschedules,
    portlandReschedules
  ]);

  const totalUnitsBilled = sumNumbers([
    chicagoUnits,
    denverUnits,
    portlandUnits
  ]);

  const alignedWeek = chicagoWeek === denverWeek && chicagoWeek === portlandWeek;
  const alignedMonthTag =
    !!chicagoMonthTag &&
    chicagoMonthTag === denverMonthTag &&
    chicagoMonthTag === portlandMonthTag;

  const hasWorkingDays = (workingDaysMap[weekEnding] ?? 0) > 0;

  const hasRealData = hasPositiveNumber(
    ptScheduledVisits,
    ptCancellations,
    ptNoShows,
    ptReschedules,
    totalUnitsBilled
  );

  if (!alignedWeek) {
    return {
      accept: false,
      reason: "unaligned-week",
      rowIndex,
      chicagoWeek,
      denverWeek,
      portlandWeek
    };
  }

  if (!alignedMonthTag) {
    return {
      accept: false,
      reason: "unaligned-month-tag",
      rowIndex,
      chicagoMonthTag,
      denverMonthTag,
      portlandMonthTag
    };
  }

  if (!hasWorkingDays) {
    return {
      accept: false,
      reason: "no-working-days",
      rowIndex,
      chicagoWeek,
      weekEnding
    };
  }

  if (!hasRealData) {
    return {
      accept: false,
      reason: "no-real-data",
      rowIndex,
      chicagoWeek,
      weekEnding
    };
  }

  return {
    accept: true,
    rowIndex,
    weekNumber: chicagoWeek,
    weekEnding,
    values: {
      weekNumber: chicagoWeek,
      monthTag: chicagoMonthTag,
      workingDaysInWeek: workingDaysMap[weekEnding],
      ptScheduledVisits,
      ptCancellations,
      ptNoShows,
      ptReschedules,
      totalUnitsBilled
    }
  };
}

async function importPtSheet(sharedTable, ws, workingDaysMap) {
  const rows = sheetRows(ws);
  let imported = 0;
  const weekEndings = [];
  const acceptedRows = [];
  const rejectedRows = [];

  for (let r = 12; r < rows.length; r += 1) {
    const row = rows[r] || [];
    const parsed = buildPtValues(r + 1, row, workingDaysMap);

    if (!parsed.accept) {
      if (parsed.reason !== "missing-chicago-week-number") {
        rejectedRows.push(parsed);
      }
      continue;
    }

    await upsertSharedRecord(sharedTable, "PT", parsed.weekEnding, parsed.values, {
      importSourceSheet: "PT",
      importWeekNumber: parsed.weekNumber,
      importMonthTag: parsed.values.monthTag
    });

    imported += 1;
    weekEndings.push(parsed.weekEnding);
    acceptedRows.push({
      rowIndex: parsed.rowIndex,
      weekNumber: parsed.weekNumber,
      weekEnding: parsed.weekEnding,
      monthTag: parsed.values.monthTag,
      workingDaysInWeek: parsed.values.workingDaysInWeek
    });
  }

  return { imported, weekEndings, acceptedRows, rejectedRows };
}

function buildCxnsValues(rowIndex, row, workingDaysMap) {
  const chicagoWeek = toNumber(row[0]);
  const chicagoMonthTag = safeText(row[1]);
  const weekEnding = weekEndingFromWeekNumber(chicagoWeek);

  if (!chicagoWeek) {
    return { accept: false, reason: "missing-chicago-week-number", rowIndex };
  }

  if (!weekEnding) {
    return { accept: false, reason: "invalid-week-ending", rowIndex, chicagoWeek };
  }

  const denverWeek = toNumber(row[8]);
  const denverMonthTag = safeText(row[9]);
  const portlandWeek = toNumber(row[16]);
  const portlandMonthTag = safeText(row[17]);
  const laWeek = toNumber(row[24]);
  const laMonthTag = safeText(row[25]);

  const chicagoScheduled = toNumber(row[2]);
  const chicagoCancels = toNumber(row[3]);
  const chicagoNoShows = toNumber(row[4]);
  const chicagoReschedules = toNumber(row[5]);

  const denverScheduled = toNumber(row[10]);
  const denverCancels = toNumber(row[11]);
  const denverNoShows = toNumber(row[12]);
  const denverReschedules = toNumber(row[13]);

  const portlandScheduled = toNumber(row[18]);
  const portlandCancels = toNumber(row[19]);
  const portlandNoShows = toNumber(row[20]);
  const portlandReschedules = toNumber(row[21]);

  const laScheduled = toNumber(row[26]);
  const laCancels = toNumber(row[27]);
  const laNoShows = toNumber(row[28]);
  const laReschedules = toNumber(row[29]);

  const scheduledAppts = sumNumbers([
    chicagoScheduled,
    denverScheduled,
    portlandScheduled,
    laScheduled
  ]);

  const cancellations = sumNumbers([
    chicagoCancels,
    denverCancels,
    portlandCancels,
    laCancels
  ]);

  const noShows = sumNumbers([
    chicagoNoShows,
    denverNoShows,
    portlandNoShows,
    laNoShows
  ]);

  const reschedules = sumNumbers([
    chicagoReschedules,
    denverReschedules,
    portlandReschedules,
    laReschedules
  ]);

  const alignedWeek =
    chicagoWeek === denverWeek &&
    chicagoWeek === portlandWeek &&
    chicagoWeek === laWeek;

  const alignedMonthTag =
    !!chicagoMonthTag &&
    chicagoMonthTag === denverMonthTag &&
    chicagoMonthTag === portlandMonthTag &&
    chicagoMonthTag === laMonthTag;

  const hasWorkingDays = (workingDaysMap[weekEnding] ?? 0) > 0;

  const hasRealData = hasPositiveNumber(
    scheduledAppts,
    cancellations,
    noShows,
    reschedules
  );

  if (!alignedWeek) {
    return {
      accept: false,
      reason: "unaligned-week",
      rowIndex,
      chicagoWeek,
      denverWeek,
      portlandWeek,
      laWeek
    };
  }

  if (!alignedMonthTag) {
    return {
      accept: false,
      reason: "unaligned-month-tag",
      rowIndex,
      chicagoMonthTag,
      denverMonthTag,
      portlandMonthTag,
      laMonthTag
    };
  }

  if (!hasWorkingDays) {
    return {
      accept: false,
      reason: "no-working-days",
      rowIndex,
      chicagoWeek,
      weekEnding
    };
  }

  if (!hasRealData) {
    return {
      accept: false,
      reason: "no-real-data",
      rowIndex,
      chicagoWeek,
      weekEnding
    };
  }

  return {
    accept: true,
    rowIndex,
    weekNumber: chicagoWeek,
    weekEnding,
    values: {
      weekNumber: chicagoWeek,
      monthTag: chicagoMonthTag,
      scheduledAppts,
      cancellations,
      noShows,
      reschedules
    }
  };
}

async function importCxnsSheet(sharedTable, ws, workingDaysMap) {
  const rows = sheetRows(ws);
  let imported = 0;
  const weekEndings = [];
  const acceptedRows = [];
  const rejectedRows = [];

  for (let r = 12; r < rows.length; r += 1) {
    const row = rows[r] || [];
    const parsed = buildCxnsValues(r + 1, row, workingDaysMap);

    if (!parsed.accept) {
      if (parsed.reason !== "missing-chicago-week-number") {
        rejectedRows.push(parsed);
      }
      continue;
    }

    await upsertSharedRecord(sharedTable, "CXNS", parsed.weekEnding, parsed.values, {
      importSourceSheet: "CXNS",
      importWeekNumber: parsed.weekNumber,
      importMonthTag: parsed.values.monthTag
    });

    imported += 1;
    weekEndings.push(parsed.weekEnding);
    acceptedRows.push({
      rowIndex: parsed.rowIndex,
      weekNumber: parsed.weekNumber,
      weekEnding: parsed.weekEnding,
      monthTag: parsed.values.monthTag
    });
  }

  return { imported, weekEndings, acceptedRows, rejectedRows };
}

async function importHolidaysSheet(referenceTable, ws) {
  const rows = sheetRows(ws);
  let imported = 0;
  const acceptedRows = [];
  const rejectedRows = [];

  for (let r = 1; r < rows.length; r += 1) {
    const row = rows[r] || [];

    const holidayDate = excelDateToIso(row[0]);
    const workingDays = toNumber(row[4]);
    const derivedMonthTag = monthTagFromIsoDate(holidayDate);

    if (!holidayDate || workingDays == null || !derivedMonthTag) {
      if (row.some((v) => v != null && v !== "")) {
        rejectedRows.push({
          rowIndex: r + 1,
          reason: "missing-required-fields",
          holidayDate,
          monthTag: derivedMonthTag,
          workingDays
        });
      }
      continue;
    }

    const rowKey = derivedMonthTag;

    await upsertReferenceRecord(
      referenceTable,
      "holidays",
      rowKey,
      {
        holidayDate,
        monthTag: derivedMonthTag,
        workingDays
      },
      {
        importSourceSheet: "Holidays",
        importMonthTag: derivedMonthTag
      }
    );

    imported += 1;
    acceptedRows.push({
      rowIndex: r + 1,
      holidayDate,
      monthTag: derivedMonthTag,
      workingDays
    });
  }

  return { imported, acceptedRows, rejectedRows };
}

function parseBudgetTargetsFromMainWorkbook(workbook) {
  const ws = workbook.Sheets[MAIN_BUDGET_SHEET];
  if (!ws) {
    return {
      imported: 0,
      acceptedRows: [],
      budgetSheet: null
    };
  }

  const rows = sheetRows(ws);
  const acceptedRows = [];
  const merged = new Map();

  for (let r = 2; r < rows.length; r += 1) {
    const row = rows[r] || [];

    const monthLabel = normalizeMonthFromBudgetSheet(row[0]);
    const entity = normalizeEntityLabel(row[1]);

    if (!monthLabel || !entity) {
      continue;
    }

    const monthKey = monthLabelToMonthKey(monthLabel);
    if (!monthKey) {
      continue;
    }

    const npTarget = toNumber(row[2]) ?? 0;
    const establishedTarget = toNumber(row[3]) ?? 0;
    const ptTarget = toNumber(row[4]) ?? 0;
    const surgeryTarget = toNumber(row[5]) ?? 0;

    const visitBudgetMonthly = npTarget + establishedTarget + ptTarget + surgeryTarget;
    const newPatientsBudgetMonthly = npTarget;

    const key = `${entity}|${monthKey}`;
    merged.set(key, {
      entity,
      monthKey,
      monthLabel,
      visitBudgetMonthly,
      newPatientsBudgetMonthly,
      workingDaysInMonth: getWorkingDaysForMonth(monthKey)
    });
  }

  for (const item of merged.values()) {
    acceptedRows.push(item);
  }

  return {
    imported: acceptedRows.length,
    acceptedRows,
    budgetSheet: MAIN_BUDGET_SHEET
  };
}

async function writeBudgetTargets(budgetTable, parsedBudget, fileName) {
  let imported = 0;

  for (const item of parsedBudget.acceptedRows || []) {
    await upsertBudgetRecord(
      budgetTable,
      item.entity,
      item.monthKey,
      item,
      {
        importSourceSheet: parsedBudget.budgetSheet || MAIN_BUDGET_SHEET,
        importFileName: fileName
      }
    );
    imported += 1;
  }

  return {
    imported,
    acceptedRows: parsedBudget.acceptedRows || [],
    budgetSheet: parsedBudget.budgetSheet || MAIN_BUDGET_SHEET
  };
}

async function importWeeklyWorkbook(regionTable, sharedTable, referenceTable, budgetTable, workbook, fileName) {
  const workingDaysMap = buildWorkingDaysMapFromRegionSheets(workbook);

  const results = {
    workbookFile: fileName,
    workbookSheets: workbook.SheetNames,
    workingDaysMap,
    regions: [],
    shared: [],
    reference: [],
    budget: {
      imported: 0,
      acceptedRows: [],
      budgetSheet: null
    }
  };

  for (const sheetName of ["LA", "Portland", "Denver", "Chicago"]) {
    const ws = workbook.Sheets[sheetName];
    if (!ws) continue;

    const result = await importRegionSheet(regionTable, ws, sheetName);
    results.regions.push({
      sheet: sheetName,
      entity: result.entity,
      imported: result.imported,
      weekEndings: result.weekEndings,
      acceptedRows: result.acceptedRows,
      rejectedRows: result.rejectedRows.slice(0, 15)
    });
  }

  if (workbook.Sheets.PT) {
    const ptResult = await importPtSheet(sharedTable, workbook.Sheets.PT, workingDaysMap);

    results.shared.push({
      sheet: "PT",
      page: "PT",
      imported: ptResult.imported,
      weekEndings: ptResult.weekEndings,
      acceptedRows: ptResult.acceptedRows,
      rejectedRows: ptResult.rejectedRows.slice(0, 15)
    });
  }

  if (workbook.Sheets.CXNS) {
    const cxnsResult = await importCxnsSheet(sharedTable, workbook.Sheets.CXNS, workingDaysMap);

    results.shared.push({
      sheet: "CXNS",
      page: "CXNS",
      imported: cxnsResult.imported,
      weekEndings: cxnsResult.weekEndings,
      acceptedRows: cxnsResult.acceptedRows,
      rejectedRows: cxnsResult.rejectedRows.slice(0, 15)
    });
  }

  if (workbook.Sheets.Holidays) {
    const holidaysResult = await importHolidaysSheet(referenceTable, workbook.Sheets.Holidays);

    results.reference.push({
      sheet: "Holidays",
      kind: "holidays",
      imported: holidaysResult.imported,
      acceptedRows: holidaysResult.acceptedRows,
      rejectedRows: holidaysResult.rejectedRows.slice(0, 15)
    });
  }

  if (workbook.Sheets[MAIN_BUDGET_SHEET]) {
    const parsedBudget = parseBudgetTargetsFromMainWorkbook(workbook);
    const writtenBudget = await writeBudgetTargets(budgetTable, parsedBudget, fileName);

    results.budget = {
      imported: writtenBudget.imported,
      acceptedRows: writtenBudget.acceptedRows,
      budgetSheet: writtenBudget.budgetSheet
    };
  }

  return results;
}

module.exports = async function (context, req) {
  try {
    const user = getUserFromRequest(req);
    const access = resolveAccess(user);

    if (!access.isAdmin) {
      return {
        status: 403,
        body: {
          ok: false,
          error: "Admin only"
        }
      };
    }

    const body = req.body || {};
    const fileBase64 = body.fileBase64;
    const fileName = body.fileName || "workbook.xlsx";

    if (!fileBase64) {
      return {
        status: 400,
        body: {
          ok: false,
          error: "Missing fileBase64"
        }
      };
    }

    const buffer = Buffer.from(fileBase64, "base64");
    const workbook = XLSX.read(buffer, {
      type: "buffer",
      cellDates: true,
      raw: true
    });

    const regionTable = getTableClient(REGION_TABLE);
    const sharedTable = getTableClient(SHARED_TABLE);
    const referenceTable = getTableClient(REFERENCE_TABLE);
    const budgetTable = getTableClient(BUDGET_TABLE);

    const result = await importWeeklyWorkbook(
      regionTable,
      sharedTable,
      referenceTable,
      budgetTable,
      workbook,
      fileName
    );

    return {
      status: 200,
      headers: { "Content-Type": "application/json" },
      body: {
        ok: true,
        message: "Import completed",
        importType: "weekly",
        ...result
      }
    };
  } catch (error) {
    context.log.error("import-excel failed", error);

    return {
      status: 500,
      body: {
        ok: false,
        error: "Failed to import workbook",
        details: error.message
      }
    };
  }
};
