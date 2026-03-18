const XLSX = require("xlsx");
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

const MONTH_MAP = {
  jan: 1,
  january: 1,
  feb: 2,
  february: 2,
  mar: 3,
  march: 3,
  apr: 4,
  april: 4,
  may: 5,
  jun: 6,
  june: 6,
  jul: 7,
  july: 7,
  aug: 8,
  august: 8,
  sep: 9,
  sept: 9,
  september: 9,
  oct: 10,
  october: 10,
  nov: 11,
  november: 11,
  dec: 12,
  december: 12
};

function normalizeMonthLabel(value) {
  const raw = String(value || "").trim();
  if (!raw) return "";

  const lower = raw.toLowerCase();
  const monthNum = MONTH_MAP[lower];
  if (!monthNum) return raw;

  return new Date(Date.UTC(WORKBOOK_YEAR, monthNum - 1, 1)).toLocaleString("en-US", {
    month: "short",
    timeZone: "UTC"
  });
}

function monthNumberFromLabel(value) {
  const raw = String(value || "").trim().toLowerCase();
  return MONTH_MAP[raw] || null;
}

function weekEndingFromMonthAndWeek(monthLabel, weekNumber) {
  const monthNum = monthNumberFromLabel(monthLabel);
  const wk = Number(weekNumber);

  if (!monthNum || !Number.isFinite(wk) || wk < 1) return "";

  const firstOfMonth = new Date(Date.UTC(WORKBOOK_YEAR, monthNum - 1, 1));
  const firstDayDow = firstOfMonth.getUTCDay();
  const offsetToFirstSunday = (7 - firstDayDow) % 7;

  const firstSunday = new Date(firstOfMonth);
  firstSunday.setUTCDate(firstOfMonth.getUTCDate() + offsetToFirstSunday);

  const weekEnding = new Date(firstSunday);
  weekEnding.setUTCDate(firstSunday.getUTCDate() + (wk - 1) * 7);

  return weekEnding.toISOString().split("T")[0];
}

function toNumber(value) {
  if (value === null || value === undefined || value === "") return null;
  if (typeof value === "number" && Number.isFinite(value)) return value;

  const parsed = Number(String(value).replace(/,/g, "").trim());
  return Number.isFinite(parsed) ? parsed : null;
}

function safeText(value) {
  return value == null ? "" : String(value).trim();
}

function sheetRows(ws) {
  return XLSX.utils.sheet_to_json(ws, {
    header: 1,
    defval: null,
    raw: false
  });
}

function rowHasAnyData(row) {
  return Array.isArray(row) && row.some(
    (v) => v !== null && v !== undefined && String(v).trim() !== ""
  );
}

async function upsertRegionRecord(table, entity, weekEnding, values, source = "workbook-import") {
  await table.upsertEntity({
    partitionKey: entity,
    rowKey: weekEnding,
    entity,
    weekEnding,
    valuesJson: JSON.stringify(values),
    source,
    importedAt: new Date().toISOString(),
    updatedAt: new Date().toISOString()
  });
}

async function upsertSharedRecord(table, page, weekEnding, values, source = "workbook-import") {
  await table.upsertEntity({
    partitionKey: page,
    rowKey: weekEnding,
    page,
    weekEnding,
    valuesJson: JSON.stringify(values),
    source,
    importedAt: new Date().toISOString(),
    updatedAt: new Date().toISOString()
  });
}

async function upsertReferenceRecord(table, kind, rowKey, values, source = "workbook-import") {
  await table.upsertEntity({
    partitionKey: kind,
    rowKey,
    kind,
    valuesJson: JSON.stringify(values),
    source,
    importedAt: new Date().toISOString(),
    updatedAt: new Date().toISOString()
  });
}

function sumNullable(...values) {
  return values.reduce((acc, v) => acc + (toNumber(v) || 0), 0);
}

async function importRegionSheet(regionTable, ws, sheetName) {
  const entity = REGION_SHEET_TO_ENTITY[sheetName];
  if (!entity) return { sheet: sheetName, entity: null, imported: 0 };

  const rows = sheetRows(ws);
  let imported = 0;
  const skipped = [];

  for (let r = 6; r < rows.length; r += 1) {
    const row = rows[r];
    if (!rowHasAnyData(row)) continue;

    const weekNumber = toNumber(row[0]);
    const daysInPeriod = toNumber(row[1]);
    const npActual = toNumber(row[5]);
    const establishedActual = toNumber(row[6]);
    const surgeryActual = toNumber(row[7]);
    const totalCalls = sheetName === "Denver" ? toNumber(row[9]) : toNumber(row[8]);
    const abandonedCalls = sheetName === "Denver" ? toNumber(row[10]) : toNumber(row[9]);
    const cashActual = sheetName === "Denver" ? toNumber(row[17]) : toNumber(row[16]);
    const monthTag = sheetName === "Denver" ? safeText(row[18]) : safeText(row[17]);

    if (!weekNumber || !monthTag) {
      skipped.push({ row: r + 1, reason: "Missing week number or month tag" });
      continue;
    }

    const weekEnding = weekEndingFromMonthAndWeek(monthTag, weekNumber);
    if (!weekEnding) {
      skipped.push({ row: r + 1, reason: `Could not derive week ending from month "${monthTag}" and week "${weekNumber}"` });
      continue;
    }

    const totalVisits = sheetName === "Denver"
      ? sumNullable(row[5], row[6], row[7], row[8])
      : sumNullable(row[5], row[6], row[7]);

    const visitsPerDay =
      daysInPeriod && daysInPeriod > 0 ? totalVisits / daysInPeriod : 0;

    const abandonmentRate =
      totalCalls && totalCalls > 0 ? (abandonedCalls / totalCalls) * 100 : 0;

    const answeredCalls = Math.max((totalCalls || 0) - (abandonedCalls || 0), 0);

    const answeredCallToNpConversion =
      answeredCalls > 0 ? ((npActual || 0) / answeredCalls) * 100 : 0;

    const values = {
      weekNumber,
      monthTag: normalizeMonthLabel(monthTag),
      daysInPeriod: daysInPeriod || 0,
      totalVisits: totalVisits || 0,
      visitsPerDay: visitsPerDay || 0,
      npActual: npActual || 0,
      establishedActual: establishedActual || 0,
      surgeryActual: surgeryActual || 0,
      totalCalls: totalCalls || 0,
      abandonedCalls: abandonedCalls || 0,
      abandonmentRate: abandonmentRate || 0,
      answeredCallToNpConversion: answeredCallToNpConversion || 0,
      cashActual: cashActual || 0
    };

    await upsertRegionRecord(regionTable, entity, weekEnding, values);
    imported += 1;
  }

  return {
    sheet: sheetName,
    entity,
    imported,
    skipped: skipped.slice(0, 25)
  };
}

async function importPtSheet(sharedTable, ws) {
  const rows = sheetRows(ws);
  let imported = 0;

  const blocks = [
    { weekCol: 0, monthCol: 1, scheduledCol: 2, cancelCol: 3, noShowCol: 4, rescheduleCol: 5, unitsCol: 6 },
    { weekCol: 10, monthCol: 11, scheduledCol: 12, cancelCol: 13, noShowCol: 14, rescheduleCol: 15, unitsCol: 16 },
    { weekCol: 20, monthCol: 21, scheduledCol: 22, cancelCol: 23, noShowCol: 24, rescheduleCol: 25, unitsCol: 26 }
  ];

  const weeklyTotals = new Map();

  for (let r = 12; r < rows.length; r += 1) {
    const row = rows[r];
    if (!rowHasAnyData(row)) continue;

    for (const block of blocks) {
      const weekNumber = toNumber(row[block.weekCol]);
      const monthTag = safeText(row[block.monthCol]);
      const scheduled = toNumber(row[block.scheduledCol]);
      const cancellations = toNumber(row[block.cancelCol]);
      const noShows = toNumber(row[block.noShowCol]);
      const reschedules = toNumber(row[block.rescheduleCol]);
      const units = toNumber(row[block.unitsCol]);

      if (!weekNumber || !monthTag) continue;

      const weekEnding = weekEndingFromMonthAndWeek(monthTag, weekNumber);
      if (!weekEnding) continue;

      const current = weeklyTotals.get(weekEnding) || {
        weekNumber,
        monthTag: normalizeMonthLabel(monthTag),
        ptScheduledVisits: 0,
        ptCancellations: 0,
        ptNoShows: 0,
        ptReschedules: 0,
        totalUnitsBilled: 0
      };

      current.ptScheduledVisits += scheduled || 0;
      current.ptCancellations += cancellations || 0;
      current.ptNoShows += noShows || 0;
      current.ptReschedules += reschedules || 0;
      current.totalUnitsBilled += units || 0;

      weeklyTotals.set(weekEnding, current);
    }
  }

  for (const [weekEnding, values] of weeklyTotals.entries()) {
    await upsertSharedRecord(sharedTable, "PT", weekEnding, {
      ...values,
      workingDaysInWeek: 5
    });
    imported += 1;
  }

  return { sheet: "PT", page: "PT", imported };
}

async function importCxnsSheet(sharedTable, ws) {
  const rows = sheetRows(ws);
  let imported = 0;

  const blocks = [
    { weekCol: 0, monthCol: 1, scheduledCol: 2, cancelCol: 3, noShowCol: 4, rescheduleCol: 5 },
    { weekCol: 8, monthCol: 9, scheduledCol: 10, cancelCol: 11, noShowCol: 12, rescheduleCol: 13 },
    { weekCol: 16, monthCol: 17, scheduledCol: 18, cancelCol: 19, noShowCol: 20, rescheduleCol: 21 },
    { weekCol: 24, monthCol: 25, scheduledCol: 26, cancelCol: 27, noShowCol: 28, rescheduleCol: 29 }
  ];

  const weeklyTotals = new Map();

  for (let r = 12; r < rows.length; r += 1) {
    const row = rows[r];
    if (!rowHasAnyData(row)) continue;

    for (const block of blocks) {
      const weekNumber = toNumber(row[block.weekCol]);
      const monthTag = safeText(row[block.monthCol]);
      const scheduled = toNumber(row[block.scheduledCol]);
      const cancellations = toNumber(row[block.cancelCol]);
      const noShows = toNumber(row[block.noShowCol]);
      const reschedules = toNumber(row[block.rescheduleCol]);

      if (!weekNumber || !monthTag) continue;

      const weekEnding = weekEndingFromMonthAndWeek(monthTag, weekNumber);
      if (!weekEnding) continue;

      const current = weeklyTotals.get(weekEnding) || {
        weekNumber,
        monthTag: normalizeMonthLabel(monthTag),
        scheduledAppts: 0,
        cancellations: 0,
        noShows: 0,
        reschedules: 0
      };

      current.scheduledAppts += scheduled || 0;
      current.cancellations += cancellations || 0;
      current.noShows += noShows || 0;
      current.reschedules += reschedules || 0;

      weeklyTotals.set(weekEnding, current);
    }
  }

  for (const [weekEnding, values] of weeklyTotals.entries()) {
    await upsertSharedRecord(sharedTable, "CXNS", weekEnding, values);
    imported += 1;
  }

  return { sheet: "CXNS", page: "CXNS", imported };
}

async function importHolidaysSheet(referenceTable, ws) {
  const rows = sheetRows(ws);
  let imported = 0;

  for (let r = 1; r < rows.length; r += 1) {
    const row = rows[r];
    if (!rowHasAnyData(row)) continue;

    const holidayDate = row[0];
    const monthTag = normalizeMonthLabel(row[3]);
    const workingDays = toNumber(row[4]);

    if (!monthTag || workingDays == null) continue;

    await upsertReferenceRecord(referenceTable, "holidays", monthTag, {
      holidayDate,
      monthTag,
      workingDays
    });

    imported += 1;
  }

  return { sheet: "Holidays", kind: "holidays", imported };
}

module.exports = async function (context, req) {
  try {
    const body = req.body || {};
    const fileBase64 = body.fileBase64;

    if (!fileBase64) {
      context.res = {
        status: 400,
        headers: { "Content-Type": "application/json" },
        body: { ok: false, error: "Missing fileBase64." }
      };
      return;
    }

    const buffer = Buffer.from(fileBase64, "base64");
    const workbook = XLSX.read(buffer, {
      type: "buffer",
      cellFormula: false,
      cellNF: false,
      cellHTML: false,
      raw: false
    });

    const regionTable = getTableClient(REGION_TABLE);
    const sharedTable = getTableClient(SHARED_TABLE);
    const referenceTable = getTableClient(REFERENCE_TABLE);

    const results = {
      regions: [],
      shared: [],
      reference: []
    };

    for (const sheetName of ["LA", "Portland", "Denver", "Chicago"]) {
      const ws = workbook.Sheets[sheetName];
      if (!ws) continue;
      results.regions.push(await importRegionSheet(regionTable, ws, sheetName));
    }

    if (workbook.Sheets.PT) {
      results.shared.push(await importPtSheet(sharedTable, workbook.Sheets.PT));
    }

    if (workbook.Sheets.CXNS) {
      results.shared.push(await importCxnsSheet(sharedTable, workbook.Sheets.CXNS));
    }

    if (workbook.Sheets.Holidays) {
      results.reference.push(await importHolidaysSheet(referenceTable, workbook.Sheets.Holidays));
    }

    context.res = {
      status: 200,
      headers: { "Content-Type": "application/json" },
      body: {
        ok: true,
        message: "Workbook import completed.",
        results
      }
    };
  } catch (error) {
    context.log.error("import-excel failed", error);
    context.res = {
      status: 500,
      headers: { "Content-Type": "application/json" },
      body: {
        ok: false,
        error: "Workbook import failed.",
        details: error && error.message ? error.message : String(error)
      }
    };
  }
};
