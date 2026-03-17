const XLSX = require("xlsx");
const { getTableClient } = require("../shared/table");

const REGION_TABLE = "WeeklyRegionData";
const SHARED_TABLE = "SharedPageData";
const REFERENCE_TABLE = "ReferenceData";

const REGION_SHEET_TO_ENTITY = {
  LA: "MRO",
  Portland: "NES",
  Denver: "SpineOne",
  Chicago: "LAOSS"
};

function normalizeMonthLabel(value) {
  const raw = String(value || "").trim();
  const map = {
    Jan: "Jan",
    January: "Jan",
    Feb: "Feb",
    February: "Feb",
    Mar: "Mar",
    March: "Mar",
    Apr: "Apr",
    April: "Apr",
    May: "May",
    Jun: "Jun",
    June: "Jun",
    Jul: "Jul",
    July: "Jul",
    Aug: "Aug",
    August: "Aug",
    Sep: "Sep",
    Sept: "Sep",
    September: "Sep",
    Oct: "Oct",
    October: "Oct",
    Nov: "Nov",
    November: "Nov",
    Dec: "Dec",
    December: "Dec"
  };
  return map[raw] || raw;
}

function monthToNumber(monthLabel) {
  const map = {
    Jan: "01",
    Feb: "02",
    Mar: "03",
    Apr: "04",
    May: "05",
    Jun: "06",
    Jul: "07",
    Aug: "08",
    Sep: "09",
    Oct: "10",
    Nov: "11",
    Dec: "12"
  };
  return map[normalizeMonthLabel(monthLabel)] || "01";
}

function makeWeekEnding(monthLabel, weekIndex) {
  const monthNum = monthToNumber(monthLabel);
  const day = String(Math.min(28, weekIndex * 7)).padStart(2, "0");
  return `2026-${monthNum}-${day}`;
}

function value(ws, cellAddress) {
  return ws[cellAddress] ? ws[cellAddress].v : null;
}

function toNumber(v) {
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}

function isMeaningfulRow(arr) {
  return Array.isArray(arr) && arr.some((v) => v !== null && v !== undefined && String(v).trim() !== "");
}

function sheetRows(ws) {
  return XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
}

function safeText(v) {
  return v == null ? "" : String(v).trim();
}

async function upsertRegionRecord(table, entity, weekEnding, values, source = "import") {
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

async function upsertSharedRecord(table, page, weekEnding, values, source = "import") {
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

async function upsertReferenceRecord(table, kind, rowKey, values, source = "import") {
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

async function importRegionSheet(regionTable, ws, sheetName) {
  const entity = REGION_SHEET_TO_ENTITY[sheetName];
  if (!entity) return { imported: 0, entity: null };

  const rows = sheetRows(ws);
  let imported = 0;

  for (let r = 6; r < rows.length; r += 1) {
    const row = rows[r];
    if (!isMeaningfulRow(row)) continue;

    const weekLabel = safeText(row[0]);
    const daysInPeriod = toNumber(row[1]);
    const totalVisits = toNumber(row[2]);
    const visitsPerDay = toNumber(row[3]);
    const npPerDay = toNumber(row[4]);
    const npActual = toNumber(row[5]);
    const establishedActual = toNumber(row[6]);
    const surgeryActual = toNumber(row[7]);
    const totalCalls = toNumber(row[8]);
    const abandonedCalls = toNumber(row[9]);
    const abandonedRate = toNumber(row[10]);
    const answeredToNpConversion = toNumber(row[11]);
    const cashActual = toNumber(row[16]);
    const monthTag = normalizeMonthLabel(row[17] || row[18]);

    if (!weekLabel || !monthTag || daysInPeriod == null) continue;

    const weekEnding = makeWeekEnding(monthTag, imported + 1);

    const values = {
      weekLabel,
      monthTag,
      daysInPeriod,
      totalVisits,
      visitsPerDay,
      npPerDay,
      npActual,
      establishedActual,
      surgeryActual,
      totalCalls,
      abandonedCalls,
      abandonmentRate: abandonedRate != null ? abandonedRate * 100 : null,
      answeredCallToNpConversion: answeredToNpConversion != null ? answeredToNpConversion * 100 : null,
      cashActual
    };

    await upsertRegionRecord(regionTable, entity, weekEnding, values, "workbook-import");
    imported += 1;
  }

  return { imported, entity };
}

async function importPtSheet(sharedTable, ws) {
  const rows = sheetRows(ws);
  let imported = 0;

  for (let r = 12; r < rows.length; r += 1) {
    const row = rows[r];
    if (!isMeaningfulRow(row)) continue;

    for (const block of [
      { monthCol: 1, scheduledCol: 2, cancelCol: 3, noShowCol: 4, rescheduleCol: 5, unitsCol: 6 },
      { monthCol: 11, scheduledCol: 12, cancelCol: 13, noShowCol: 14, rescheduleCol: 15, unitsCol: 16 },
      { monthCol: 21, scheduledCol: 22, cancelCol: 23, noShowCol: 24, rescheduleCol: 25, unitsCol: 26 }
    ]) {
      const monthTag = normalizeMonthLabel(row[block.monthCol]);
      const ptScheduledVisits = toNumber(row[block.scheduledCol]);
      const ptCancellations = toNumber(row[block.cancelCol]);
      const ptNoShows = toNumber(row[block.noShowCol]);
      const ptReschedules = toNumber(row[block.rescheduleCol]);
      const totalUnitsBilled = toNumber(row[block.unitsCol]);

      if (!monthTag || ptScheduledVisits == null) continue;

      const weekEnding = makeWeekEnding(monthTag, imported + 1);

      const values = {
        monthTag,
        ptScheduledVisits,
        ptCancellations,
        ptNoShows,
        ptReschedules,
        totalUnitsBilled,
        workingDaysInWeek: 5
      };

      await upsertSharedRecord(sharedTable, "PT", weekEnding, values, "workbook-import");
      imported += 1;
    }
  }

  return { imported };
}

async function importCxnsSheet(sharedTable, ws) {
  const rows = sheetRows(ws);
  let imported = 0;

  for (let r = 12; r < rows.length; r += 1) {
    const row = rows[r];
    if (!isMeaningfulRow(row)) continue;

    for (const block of [
      { monthCol: 1, scheduledCol: 2, cancelCol: 3, noShowCol: 4, rescheduleCol: 5 },
      { monthCol: 9, scheduledCol: 10, cancelCol: 11, noShowCol: 12, rescheduleCol: 13 },
      { monthCol: 17, scheduledCol: 18, cancelCol: 19, noShowCol: 20, rescheduleCol: 21 },
      { monthCol: 25, scheduledCol: 26, cancelCol: 27, noShowCol: 28, rescheduleCol: 29 }
    ]) {
      const monthTag = normalizeMonthLabel(row[block.monthCol]);
      const scheduledAppts = toNumber(row[block.scheduledCol]);
      const cancellations = toNumber(row[block.cancelCol]);
      const noShows = toNumber(row[block.noShowCol]);
      const reschedules = toNumber(row[block.rescheduleCol]);

      if (!monthTag || scheduledAppts == null) continue;

      const weekEnding = makeWeekEnding(monthTag, imported + 1);

      const values = {
        monthTag,
        scheduledAppts,
        cancellations,
        noShows,
        reschedules
      };

      await upsertSharedRecord(sharedTable, "CXNS", weekEnding, values, "workbook-import");
      imported += 1;
    }
  }

  return { imported };
}

async function importHolidaysSheet(referenceTable, ws) {
  const rows = sheetRows(ws);
  let imported = 0;

  for (let r = 1; r < rows.length; r += 1) {
    const row = rows[r];
    const holidayDate = row[0];
    const monthTag = normalizeMonthLabel(row[3]);
    const workingDays = toNumber(row[4]);

    if (!monthTag || workingDays == null) continue;

    await upsertReferenceRecord(
      referenceTable,
      "holidays",
      monthTag,
      {
        holidayDate,
        monthTag,
        workingDays
      },
      "workbook-import"
    );

    imported += 1;
  }

  return { imported };
}

module.exports = async function (context, req) {
  try {
    const body = req.body || {};
    const fileBase64 = body.fileBase64;

    if (!fileBase64) {
      context.res = {
        status: 400,
        body: { error: "Missing fileBase64." }
      };
      return;
    }

    const buffer = Buffer.from(fileBase64, "base64");
    const workbook = XLSX.read(buffer, { type: "buffer" });

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
      const result = await importRegionSheet(regionTable, ws, sheetName);
      results.regions.push({ sheet: sheetName, ...result });
    }

    if (workbook.Sheets.PT) {
      results.shared.push({
        sheet: "PT",
        page: "PT",
        ...(await importPtSheet(sharedTable, workbook.Sheets.PT))
      });
    }

    if (workbook.Sheets.CXNS) {
      results.shared.push({
        sheet: "CXNS",
        page: "CXNS",
        ...(await importCxnsSheet(sharedTable, workbook.Sheets.CXNS))
      });
    }

    if (workbook.Sheets.Holidays) {
      results.reference.push({
        sheet: "Holidays",
        kind: "holidays",
        ...(await importHolidaysSheet(referenceTable, workbook.Sheets.Holidays))
      });
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
      body: {
        error: "Workbook import failed.",
        details: error.message
      }
    };
  }
};
