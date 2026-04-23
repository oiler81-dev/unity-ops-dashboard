const XLSX = require("xlsx");
const { getUserFromRequest } = require("../shared/auth");
const {
  resolveAccess,
  requireAccess,
  safeErrorResponse
} = require("../shared/permissions");
const { getTableClient } = require("../shared/table");

// Uploaded workbook size cap. Base64 adds ~33% overhead so the raw body can be
// larger; we also cap the decoded buffer below. 10 MB of raw Excel is plenty
// for weekly operator data — anything bigger is almost certainly accidental
// or an attack trying to OOM the function worker.
const MAX_DECODED_BYTES = 10 * 1024 * 1024;
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

// PT sheet column offsets per region block (each block is 10 cols wide)
// Cols: Week(0), MonthTag(1), Scheduled(2), Cancels(3), NoShows(4),
//       Reschedules(5), TotalUnits(6), UnitsPerVisit(7), VisitsPerDay(8), VisitsSeen(9)
const PT_REGION_OFFSETS = {
  MRO: 0,       // Chicago block starts at col 0
  SpineOne: 10, // Denver block starts at col 10
  NES: 20       // Portland block starts at col 20
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

function normalizeMonthTag(value) {
  const text = safeText(value).toLowerCase();
  if (!text) return "";

  const map = {
    jan: "Jan",
    january: "Jan",
    feb: "Feb",
    february: "Feb",
    mar: "Mar",
    march: "Mar",
    apr: "Apr",
    april: "Apr",
    may: "May",
    jun: "Jun",
    june: "Jun",
    jul: "Jul",
    july: "Jul",
    aug: "Aug",
    august: "Aug",
    sep: "Sep",
    sept: "Sep",
    september: "Sep",
    oct: "Oct",
    october: "Oct",
    nov: "Nov",
    november: "Nov",
    dec: "Dec",
    december: "Dec"
  };

  return map[text] || map[text.slice(0, 3)] || "";
}

function detectEntityFromSheetName(sheetName) {
  const s = safeText(sheetName).toLowerCase();

  if (s.includes("la") || s.includes("laoss")) return "LAOSS";
  if (s.includes("portland") || s.includes("nes")) return "NES";
  if (s.includes("denver") || s.includes("spine")) return "SpineOne";
  if (s.includes("chicago") || s.includes("mro") || s.includes("midland") || s.includes("riverside")) return "MRO";

  return null;
}

function findSheetMonthTag(rows) {
  for (let r = 0; r < Math.min(rows.length, 5); r += 1) {
    const row = rows[r] || [];
    for (let c = 0; c < Math.min(row.length, 6); c += 1) {
      const monthTag = normalizeMonthTag(row[c]);
      if (monthTag) return monthTag;
    }
  }
  return "";
}

function scanRowForBestNumber(row) {
  const numbers = (row || [])
    .map((value) => toNumber(value))
    .filter((value) => value != null);

  if (!numbers.length) return null;

  return numbers[numbers.length - 1];
}

function buildMorPtoMap(workbook) {
  const result = {};

  for (const sheetName of workbook.SheetNames) {
    if (!/mor/i.test(sheetName)) continue;

    const entity = detectEntityFromSheetName(sheetName);
    if (!entity) continue;

    const rows = sheetRows(workbook.Sheets[sheetName]);
    const monthTag = findSheetMonthTag(rows);

    if (!monthTag) continue;
    if (!result[entity]) result[entity] = {};

    for (let r = 0; r < rows.length; r += 1) {
      const row = rows[r] || [];

      for (let c = 0; c < row.length; c += 1) {
        const text = safeText(row[c]).toLowerCase();
        if (!text) continue;

        const isPtoTotalRow =
          text === "provider pto mtd total" ||
          text.includes("provider pto mtd total");

        if (!isPtoTotalRow) continue;

        const ptoValue = scanRowForBestNumber(row);
        if (ptoValue == null) continue;

        result[entity][monthTag] = ptoValue;
      }
    }
  }

  return result;
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
    // Top-level fields for direct querying
    visitVolume: values.totalVisits ?? values.visitVolume ?? 0,
    callVolume: values.callVolume ?? values.totalCalls ?? 0,
    newPatients: values.npActual ?? values.newPatients ?? 0,
    surgeries: values.surgeryActual ?? values.surgeries ?? 0,
    established: values.establishedActual ?? values.established ?? 0,
    totalCalls: values.totalCalls ?? 0,
    abandonedCalls: values.abandonedCalls ?? 0,
    noShowRate: (function() {
      const visits = values.totalVisits ?? values.visitVolume ?? 0;
      const noShows = values.noShows ?? 0;
      const cancelled = values.cancelled ?? 0;
      const scheduled = visits + noShows + cancelled;
      return scheduled > 0 ? Number(((noShows / scheduled) * 100).toFixed(2)) : 0;
    })(),
    cancellationRate: (function() {
      const visits = values.totalVisits ?? values.visitVolume ?? 0;
      const noShows = values.noShows ?? 0;
      const cancelled = values.cancelled ?? 0;
      const scheduled = visits + noShows + cancelled;
      return scheduled > 0 ? Number(((cancelled / scheduled) * 100).toFixed(2)) : 0;
    })(),
    abandonedCallRate: (function() {
      const totalCalls = values.totalCalls ?? 0;
      const abandonedCalls = values.abandonedCalls ?? 0;
      if (values.abandonmentRate != null) return values.abandonmentRate;
      return totalCalls > 0 ? Number(((abandonedCalls / totalCalls) * 100).toFixed(2)) : 0;
    })(),
    ptoDays: values.ptoDays ?? 0,
    cashCollected: values.cashActual ?? values.cashCollected ?? 0,
    operationsNarrative: values.operationsNarrative ?? "",
    // PT fields — populated later by mergePtDataIntoRegionRecords
    ptScheduledVisits: values.ptScheduledVisits ?? 0,
    ptCancellations: values.ptCancellations ?? 0,
    ptNoShows: values.ptNoShows ?? 0,
    ptReschedules: values.ptReschedules ?? 0,
    ptTotalUnitsBilled: values.ptTotalUnitsBilled ?? 0,
    ptVisitsSeen: values.ptVisitsSeen ?? 0,
    ptWorkingDays: values.ptWorkingDays ?? 5,
    ptUnitsPerVisit: values.ptUnitsPerVisit ?? 0,
    ptVisitsPerDay: values.ptVisitsPerDay ?? 0,
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

function buildRegionValues(sheetName, rowIndex, row, ptoMap = {}) {
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

  const entity = REGION_SHEET_TO_ENTITY[sheetName];

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
    cashActual,
    cashCollected: cashActual ?? 0,
    ptoDays: ptoMap?.[entity]?.[monthTag] ?? 0,
    operationsNarrative: ""
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

async function importRegionSheet(regionTable, ws, sheetName, ptoMap, perEntityPtMap = {}) {
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
    const parsed = buildRegionValues(sheetName, r + 1, row, ptoMap);

    if (!parsed.accept) {
      if (parsed.reason !== "missing-week-number") {
        rejectedRows.push(parsed);
      }
      continue;
    }

    // Embed PT data directly if available for this entity + weekEnding
    const ptData = (perEntityPtMap[entity] || {})[parsed.weekEnding] || {};
    const valuesWithPt = Object.assign({}, parsed.values, ptData);

    await upsertRegionRecord(regionTable, entity, parsed.weekEnding, valuesWithPt, {
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
      daysInPeriod: parsed.values.daysInPeriod,
      ptoDays: parsed.values.ptoDays,
      ptVisitsSeen: ptData.ptVisitsSeen || 0
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

// Builds per-entity PT data from the PT sheet
// Returns { NES: { weekEnding: { ptVisitsSeen, ptTotalUnitsBilled, ... } }, SpineOne: {...}, MRO: {...} }
function buildPerEntityPtMap(ws) {
  const rows = sheetRows(ws);
  const result = { NES: {}, SpineOne: {}, MRO: {} };

  // Data starts at row index 12 (row 13 in sheet, 0-indexed = 12)
  for (let r = 12; r < rows.length; r += 1) {
    const row = rows[r] || [];

    for (const [entity, offset] of Object.entries(PT_REGION_OFFSETS)) {
      const weekNumber = toNumber(row[offset]);
      if (!weekNumber) continue;

      const weekEnding = weekEndingFromWeekNumber(weekNumber);
      if (!weekEnding) continue;

      const scheduled = toNumber(row[offset + 2]) ?? 0;
      const cancellations = toNumber(row[offset + 3]) ?? 0;
      const noShows = toNumber(row[offset + 4]) ?? 0;
      const reschedules = toNumber(row[offset + 5]) ?? 0;
      const totalUnitsBilled = toNumber(row[offset + 6]) ?? 0;
      const unitsPerVisit = toNumber(row[offset + 7]) ?? 0;
      const visitsPerDay = toNumber(row[offset + 8]) ?? 0;
      const visitsSeen = toNumber(row[offset + 9]) ?? 0;

      // Only store rows that have real PT data
      if (!hasPositiveNumber(scheduled, visitsSeen, totalUnitsBilled)) continue;

      result[entity][weekEnding] = {
        ptScheduledVisits: scheduled,
        ptCancellations: cancellations,
        ptNoShows: noShows,
        ptReschedules: reschedules,
        ptTotalUnitsBilled: totalUnitsBilled,
        ptVisitsSeen: visitsSeen,
        ptWorkingDays: 5,
        ptUnitsPerVisit: unitsPerVisit,
        ptVisitsPerDay: visitsPerDay
      };
    }
  }

  return result;
}

// After region records are written, patch PT fields onto matching region records
async function mergePtDataIntoRegionRecords(regionTable, ptMap) {
  const merged = { NES: 0, SpineOne: 0, MRO: 0 };

  for (const [entity, weekMap] of Object.entries(ptMap)) {
    for (const [weekEnding, ptValues] of Object.entries(weekMap)) {
      try {
        // Fetch existing record first so we preserve all region fields
        let existing = null;
        try {
          existing = await regionTable.getEntity(entity, weekEnding);
        } catch (err) {
          if (err.statusCode !== 404) throw err;
        }

        if (!existing) continue; // Only update records that already exist from region import

        await regionTable.updateEntity({
          partitionKey: entity,
          rowKey: weekEnding,
          ptScheduledVisits: ptValues.ptScheduledVisits,
          ptCancellations: ptValues.ptCancellations,
          ptNoShows: ptValues.ptNoShows,
          ptReschedules: ptValues.ptReschedules,
          ptTotalUnitsBilled: ptValues.ptTotalUnitsBilled,
          ptVisitsSeen: ptValues.ptVisitsSeen,
          ptWorkingDays: ptValues.ptWorkingDays,
          ptUnitsPerVisit: ptValues.ptUnitsPerVisit,
          ptVisitsPerDay: ptValues.ptVisitsPerDay,
          updatedAt: new Date().toISOString()
        }, "Merge");

        merged[entity] += 1;
      } catch (err) {
        // Non-fatal — log and continue
      }
    }
  }

  return merged;
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
      monthTag: parsed.values.monthTag,
      scheduledAppts: parsed.values.scheduledAppts,
      cancellations: parsed.values.cancellations,
      noShows: parsed.values.noShows,
      reschedules: parsed.values.reschedules
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
  const ptoMap = buildMorPtoMap(workbook);

  const results = {
    workbookFile: fileName,
    workbookSheets: workbook.SheetNames,
    workingDaysMap,
    ptoMap,
    regions: [],
    shared: [],
    reference: [],
    budget: {
      imported: 0,
      acceptedRows: [],
      budgetSheet: null
    },
    ptMerge: {}
  };

  // Step 1 — build per-entity PT map before region import so we can embed it in same upsert
  const perEntityPtMap = workbook.Sheets.PT
    ? buildPerEntityPtMap(workbook.Sheets.PT)
    : { NES: {}, SpineOne: {}, MRO: {} };

  // Step 2 — import region sheets with PT data embedded
  for (const sheetName of ["LA", "Portland", "Denver", "Chicago"]) {
    const ws = workbook.Sheets[sheetName];
    if (!ws) continue;

    const result = await importRegionSheet(regionTable, ws, sheetName, ptoMap, perEntityPtMap);
    results.regions.push({
      sheet: sheetName,
      entity: result.entity,
      imported: result.imported,
      weekEndings: result.weekEndings,
      acceptedRows: result.acceptedRows,
      rejectedRows: result.rejectedRows.slice(0, 15)
    });
  }

  // Step 2 — import PT sheet into SharedPageData (aggregate, unchanged)
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

    // PT data is now embedded directly during region import above
    results.ptMerge = { embedded: true };
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

    const authError = requireAccess(access);
    if (authError) return authError;

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

    // Rough size check on the base64 string first to short-circuit obvious
    // giant uploads before we allocate a decode buffer.
    if (typeof fileBase64 !== "string" || fileBase64.length > Math.ceil((MAX_DECODED_BYTES * 4) / 3)) {
      return {
        status: 413,
        body: { ok: false, error: "File too large" }
      };
    }

    const buffer = Buffer.from(fileBase64, "base64");
    if (buffer.length > MAX_DECODED_BYTES) {
      return {
        status: 413,
        body: { ok: false, error: "File too large" }
      };
    }

    // Magic-byte check: XLSX files are ZIP archives and must start with "PK\x03\x04".
    if (buffer.length < 4 || buffer[0] !== 0x50 || buffer[1] !== 0x4B || buffer[2] !== 0x03 || buffer[3] !== 0x04) {
      return {
        status: 400,
        body: { ok: false, error: "Invalid file type (expected .xlsx)" }
      };
    }

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
    return safeErrorResponse(context, error, "Failed to import workbook");
  }
};
