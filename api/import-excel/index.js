const XLSX = require("xlsx");
const { getUserFromRequest } = require("../shared/auth");
const { resolveAccess } = require("../shared/permissions");
const { getTableClient } = require("../shared/table");

const REGION_TABLE = "WeeklyRegionData";
const SHARED_TABLE = "SharedPageData";
const REFERENCE_TABLE = "ReferenceData";
const WORKBOOK_YEAR = 2026;

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
  if (!text || text === "#N/A" || text === "#VALUE!" || text === "#DIV/0!") {
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

function buildRegionValues(sheetName, row) {
  const weekNumber = toNumber(row[0]);
  const daysInPeriod = toNumber(row[1]);
  const monthTag = safeText(sheetName === "Denver" ? row[18] : row[17]);

  if (!weekNumber) return null;

  const weekEnding = weekEndingFromWeekNumber(weekNumber);
  if (!weekEnding) return null;

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

  const values = {
    weekNumber,
    monthTag: monthTag || null,
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
    npToEstablishedConversion: percentToDisplay(
      sheetName === "Denver" ? row[12] : row[11]
    ),
    npToSurgeryConversion: percentToDisplay(
      sheetName === "Denver" ? row[13] : row[12]
    ),
    cashActual
  };

  if (sheetName === "Denver") {
    values.imagingActual = imagingActual;
    values.piNp = toNumber(row[19]);
    values.piCashCollection = toNumber(row[20]);
  }

  const hasRealData = hasPositiveNumber(
    totalVisits,
    npActual,
    establishedActual,
    surgeryActual,
    imagingActual,
    totalCalls,
    abandonedCalls,
    cashActual
  );

  const hasUsableDays = daysInPeriod != null && daysInPeriod > 0;

  if (!hasRealData || !hasUsableDays) {
    return null;
  }

  return { weekNumber, weekEnding, values };
}

async function importRegionSheet(regionTable, ws, sheetName) {
  const entity = REGION_SHEET_TO_ENTITY[sheetName];
  const rows = sheetRows(ws);

  if (!entity) {
    return { imported: 0, entity: null, weeks: [] };
  }

  let imported = 0;
  const weeks = [];

  for (let r = 6; r < rows.length; r += 1) {
    const row = rows[r] || [];
    const parsed = buildRegionValues(sheetName, row);
    if (!parsed) continue;

    await upsertRegionRecord(regionTable, entity, parsed.weekEnding, parsed.values, {
      importSourceSheet: sheetName,
      importWeekNumber: parsed.weekNumber,
      importMonthTag: parsed.values.monthTag
    });

    imported += 1;
    weeks.push(parsed.weekEnding);
  }

  return { imported, entity, weeks };
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

async function importPtSheet(sharedTable, ws, workingDaysMap) {
  const rows = sheetRows(ws);
  let imported = 0;
  const weeks = [];

  for (let r = 12; r < rows.length; r += 1) {
    const row = rows[r] || [];

    const weekNumber = toNumber(row[0]);
    const monthTag = safeText(row[1]);
    const weekEnding = weekEndingFromWeekNumber(weekNumber);

    if (!weekNumber || !weekEnding) continue;

    const chicagoScheduled = toNumber(row[2]);
    const chicagoCancels = toNumber(row[3]);
    const chicagoNoShows = toNumber(row[4]);
    const chicagoReschedules = toNumber(row[5]);
    const chicagoUnits = toNumber(row[6]);

    const denverWeek = toNumber(row[10]);
    const denverMonthTag = safeText(row[11]);
    const denverScheduled = toNumber(row[12]);
    const denverCancels = toNumber(row[13]);
    const denverNoShows = toNumber(row[14]);
    const denverReschedules = toNumber(row[15]);
    const denverUnits = toNumber(row[16]);

    const portlandWeek = toNumber(row[20]);
    const portlandMonthTag = safeText(row[21]);
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

    const hasRealData = hasPositiveNumber(
      ptScheduledVisits,
      ptCancellations,
      ptNoShows,
      ptReschedules,
      totalUnitsBilled
    );

    const alignedMonthTag = monthTag && monthTag === denverMonthTag && monthTag === portlandMonthTag;
    const alignedWeek = weekNumber && weekNumber === denverWeek && weekNumber === portlandWeek;
    const hasWorkingDays = (workingDaysMap[weekEnding] ?? 0) > 0;

    if (!hasRealData || !alignedMonthTag || !alignedWeek || !hasWorkingDays) {
      continue;
    }

    const values = {
      weekNumber,
      monthTag,
      workingDaysInWeek: workingDaysMap[weekEnding],
      ptScheduledVisits,
      ptCancellations,
      ptNoShows,
      ptReschedules,
      totalUnitsBilled
    };

    await upsertSharedRecord(sharedTable, "PT", weekEnding, values, {
      importSourceSheet: "PT",
      importWeekNumber: weekNumber,
      importMonthTag: monthTag
    });

    imported += 1;
    weeks.push(weekEnding);
  }

  return { imported, weeks };
}

async function importCxnsSheet(sharedTable, ws, workingDaysMap) {
  const rows = sheetRows(ws);
  let imported = 0;
  const weeks = [];

  for (let r = 12; r < rows.length; r += 1) {
    const row = rows[r] || [];

    const weekNumber = toNumber(row[0]);
    const monthTag = safeText(row[1]);
    const weekEnding = weekEndingFromWeekNumber(weekNumber);

    if (!weekNumber || !weekEnding) continue;

    const chicagoScheduled = toNumber(row[2]);
    const chicagoCancels = toNumber(row[3]);
    const chicagoNoShows = toNumber(row[4]);
    const chicagoReschedules = toNumber(row[5]);

    const denverWeek = toNumber(row[8]);
    const denverMonthTag = safeText(row[9]);
    const denverScheduled = toNumber(row[10]);
    const denverCancels = toNumber(row[11]);
    const denverNoShows = toNumber(row[12]);
    const denverReschedules = toNumber(row[13]);

    const portlandWeek = toNumber(row[16]);
    const portlandMonthTag = safeText(row[17]);
    const portlandScheduled = toNumber(row[18]);
    const portlandCancels = toNumber(row[19]);
    const portlandNoShows = toNumber(row[20]);
    const portlandReschedules = toNumber(row[21]);

    const laWeek = toNumber(row[24]);
    const laMonthTag = safeText(row[25]);
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

    const hasRealData = hasPositiveNumber(
      scheduledAppts,
      cancellations,
      noShows,
      reschedules
    );

    const alignedMonthTag =
      monthTag &&
      monthTag === denverMonthTag &&
      monthTag === portlandMonthTag &&
      monthTag === laMonthTag;

    const alignedWeek =
      weekNumber &&
      weekNumber === denverWeek &&
      weekNumber === portlandWeek &&
      weekNumber === laWeek;

    const hasWorkingDays = (workingDaysMap[weekEnding] ?? 0) > 0;

    if (!hasRealData || !alignedMonthTag || !alignedWeek || !hasWorkingDays) {
      continue;
    }

    const values = {
      weekNumber,
      monthTag,
      scheduledAppts,
      cancellations,
      noShows,
      reschedules
    };

    await upsertSharedRecord(sharedTable, "CXNS", weekEnding, values, {
      importSourceSheet: "CXNS",
      importWeekNumber: weekNumber,
      importMonthTag: monthTag
    });

    imported += 1;
    weeks.push(weekEnding);
  }

  return { imported, weeks };
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

async function importHolidaysSheet(referenceTable, ws) {
  const rows = sheetRows(ws);
  let imported = 0;

  for (let r = 1; r < rows.length; r += 1) {
    const row = rows[r] || [];

    const holidayDate = excelDateToIso(row[0]);
    const monthTag = safeText(row[3]);
    const workingDays = toNumber(row[4]);

    if (!holidayDate || !monthTag || workingDays == null) continue;

    const rowKey = monthTag;

    await upsertReferenceRecord(
      referenceTable,
      "holidays",
      rowKey,
      {
        holidayDate,
        monthTag,
        workingDays
      },
      {
        importSourceSheet: "Holidays",
        importMonthTag: monthTag
      }
    );

    imported += 1;
  }

  return { imported };
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

    const workingDaysMap = buildWorkingDaysMapFromRegionSheets(workbook);

    const results = {
      workbookFile: fileName,
      workbookSheets: workbook.SheetNames,
      regions: [],
      shared: [],
      reference: []
    };

    for (const sheetName of ["LA", "Portland", "Denver", "Chicago"]) {
      const ws = workbook.Sheets[sheetName];
      if (!ws) continue;

      const result = await importRegionSheet(regionTable, ws, sheetName);
      results.regions.push({
        sheet: sheetName,
        entity: result.entity,
        imported: result.imported,
        weekEndings: result.weeks
      });
    }

    if (workbook.Sheets.PT) {
      const ptResult = await importPtSheet(
        sharedTable,
        workbook.Sheets.PT,
        workingDaysMap
      );

      results.shared.push({
        sheet: "PT",
        page: "PT",
        imported: ptResult.imported,
        weekEndings: ptResult.weeks
      });
    }

    if (workbook.Sheets.CXNS) {
      const cxnsResult = await importCxnsSheet(
        sharedTable,
        workbook.Sheets.CXNS,
        workingDaysMap
      );

      results.shared.push({
        sheet: "CXNS",
        page: "CXNS",
        imported: cxnsResult.imported,
        weekEndings: cxnsResult.weeks
      });
    }

    if (workbook.Sheets.Holidays) {
      const holidaysResult = await importHolidaysSheet(
        referenceTable,
        workbook.Sheets.Holidays
      );

      results.reference.push({
        sheet: "Holidays",
        kind: "holidays",
        imported: holidaysResult.imported
      });
    }

    return {
      status: 200,
      headers: { "Content-Type": "application/json" },
      body: {
        ok: true,
        message: "Import completed",
        ...results
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
