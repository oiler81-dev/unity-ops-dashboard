const XLSX = require("xlsx");
const { getUserFromRequest, getUserEmail, getDisplayName } = require("../shared/auth");
const { getPermissionByEmail } = require("../shared/permissions");
const { getTableClient } = require("../shared/table");

const YEAR = 2026;

const REGION_SHEETS = [
  {
    sheet: "LA",
    entity: "LAOSS",
    monthCol: 18,
    cols: { week: 1, days: 2, visitVolume: 3, newPatients: 6, establishedActual: 7, surgeryActual: 8, callVolume: 9, abandonedCalls: 10, abandonedCallRate: 11, cash: 17 }
  },
  {
    sheet: "Portland",
    entity: "NES",
    monthCol: 18,
    cols: { week: 1, days: 2, visitVolume: 3, newPatients: 6, establishedActual: 7, surgeryActual: 8, callVolume: 9, abandonedCalls: 10, abandonedCallRate: 11, cash: 17 }
  },
  {
    sheet: "Denver",
    entity: "SpineOne",
    monthCol: 19,
    cols: { week: 1, days: 2, visitVolume: 3, newPatients: 6, establishedActual: 7, surgeryActual: 8, imagingActual: 9, callVolume: 10, abandonedCalls: 11, abandonedCallRate: 12, cash: 18 }
  },
  {
    sheet: "Chicago",
    entity: "MRO",
    monthCol: 18,
    cols: { week: 1, days: 2, visitVolume: 3, newPatients: 6, establishedActual: 7, surgeryActual: 8, callVolume: 9, abandonedCalls: 10, abandonedCallRate: 11, cash: 17 }
  }
];

const CXNS_BLOCKS = [
  { entity: "MRO", startCol: 1 },
  { entity: "SpineOne", startCol: 9 },
  { entity: "NES", startCol: 17 },
  { entity: "LAOSS", startCol: 25 }
];

const PT_BLOCKS = [
  { entity: "MRO", startCol: 1 },
  { entity: "SpineOne", startCol: 11 },
  { entity: "NES", startCol: 21 }
];

function toNumber(value) {
  if (value === null || value === undefined || value === "") return null;
  const n = Number(value);
  return Number.isFinite(n) ? n : null;
}

function toPercent(value) {
  const n = toNumber(value);
  if (n === null) return null;
  if (n >= 0 && n <= 1) return Number((n * 100).toFixed(4));
  return n;
}

function getCell(row, col1Based) {
  return row[col1Based - 1];
}

function getWeekEndingFromIsoWeek(year, weekNumber) {
  const jan4 = new Date(Date.UTC(year, 0, 4));
  const jan4Day = jan4.getUTCDay() || 7;
  const mondayWeek1 = new Date(jan4);
  mondayWeek1.setUTCDate(jan4.getUTCDate() - jan4Day + 1);

  const monday = new Date(mondayWeek1);
  monday.setUTCDate(mondayWeek1.getUTCDate() + ((weekNumber - 1) * 7));

  const sunday = new Date(monday);
  sunday.setUTCDate(monday.getUTCDate() + 6);

  return sunday.toISOString().slice(0, 10);
}

function getSheetRows(workbook, sheetName) {
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) return [];
  return XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    raw: true,
    defval: null
  });
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
    const fileBase64 = body.fileBase64;
    const fileName = body.fileName || "workbook.xlsx";

    if (!fileBase64) {
      context.res = { status: 400, body: { error: "fileBase64 is required" } };
      return;
    }

    const workbookBuffer = Buffer.from(fileBase64, "base64");
    const workbook = XLSX.read(workbookBuffer, {
      type: "buffer",
      cellDates: true,
      raw: true
    });

    const weeklyInputsTable = getTableClient("WeeklyInputs");
    const submissionStatusTable = getTableClient("SubmissionStatus");

    const updatedBy = permission.displayName || getDisplayName(user);
    const updatedAt = new Date().toISOString();

    let weeklyInputCount = 0;
    let submissionCount = 0;
    const touchedWeeks = new Set();

    async function upsertMetric(partitionKey, rowKey, entity, weekEnding, metricKey, value, sourceSheet, sourceType) {
      if (value === null || value === undefined || value === "") return;

      await weeklyInputsTable.upsertEntity({
        partitionKey,
        rowKey,
        entity,
        weekEnding,
        section: sourceType,
        metricKey,
        label: metricKey,
        value,
        valueType: typeof value,
        sourceSheet,
        updatedBy,
        updatedAt
      }, "Merge");

      weeklyInputCount += 1;
      touchedWeeks.add(`${entity}|${weekEnding}`);
    }

    async function upsertStatus(entity, weekEnding) {
      const partitionKey = `${entity}|${weekEnding}`;

      await submissionStatusTable.upsertEntity({
        partitionKey,
        rowKey: "STATUS",
        entity,
        weekEnding,
        status: "Submitted",
        submittedBy: updatedBy,
        submittedAt: updatedAt,
        updatedBy,
        updatedAt
      }, "Merge");

      submissionCount += 1;
      touchedWeeks.add(`${entity}|${weekEnding}`);
    }

    for (const config of REGION_SHEETS) {
      const rows = getSheetRows(workbook, config.sheet);

      for (let i = 6; i < rows.length; i++) {
        const row = rows[i];
        const week = toNumber(getCell(row, config.cols.week));
        const monthMarker = getCell(row, config.monthCol);

        if (!week || !monthMarker) continue;

        const weekEnding = getWeekEndingFromIsoWeek(YEAR, week);
        const partitionKey = `${config.entity}|${weekEnding}`;

        const values = {
          visitVolume: toNumber(getCell(row, config.cols.visitVolume)),
          newPatients: toNumber(getCell(row, config.cols.newPatients)),
          establishedActual: toNumber(getCell(row, config.cols.establishedActual)),
          surgeryActual: toNumber(getCell(row, config.cols.surgeryActual)),
          callVolume: toNumber(getCell(row, config.cols.callVolume)),
          abandonedCalls: toNumber(getCell(row, config.cols.abandonedCalls)),
          abandonedCallRate: toPercent(getCell(row, config.cols.abandonedCallRate)),
          daysInPeriod: toNumber(getCell(row, config.cols.days)),
          cash: toNumber(getCell(row, config.cols.cash))
        };

        if (config.cols.imagingActual) {
          values.imagingActual = toNumber(getCell(row, config.cols.imagingActual));
        }

        for (const [metricKey, value] of Object.entries(values)) {
          await upsertMetric(
            partitionKey,
            `INPUT|${metricKey}`,
            config.entity,
            weekEnding,
            metricKey,
            value,
            config.sheet,
            "DynamicRegionInput"
          );
        }

        await upsertStatus(config.entity, weekEnding);
      }
    }

    const cxnsRows = getSheetRows(workbook, "CXNS");
    const cxnsSharedAgg = new Map();

    for (const block of CXNS_BLOCKS) {
      for (let i = 12; i < cxnsRows.length; i++) {
        const row = cxnsRows[i];
        const week = toNumber(getCell(row, block.startCol));
        const monthMarker = getCell(row, block.startCol + 1);

        if (!week || !monthMarker) continue;

        const scheduledVisits = toNumber(getCell(row, block.startCol + 2));
        const cancellations = toNumber(getCell(row, block.startCol + 3));
        const noShows = toNumber(getCell(row, block.startCol + 4));
        const rescheduledVisits = toNumber(getCell(row, block.startCol + 5));

        const cancellationRate =
          scheduledVisits && cancellations !== null
            ? Number(((cancellations / Math.max(scheduledVisits, 1)) * 100).toFixed(4))
            : null;

        const noShowRate =
          scheduledVisits && noShows !== null
            ? Number(((noShows / Math.max(scheduledVisits, 1)) * 100).toFixed(4))
            : null;

        const weekEnding = getWeekEndingFromIsoWeek(YEAR, week);
        const partitionKey = `${block.entity}|${weekEnding}`;

        const regionValues = {
          scheduledVisits,
          cancellations,
          noShows,
          rescheduledVisits,
          cancellationRate,
          noShowRate
        };

        for (const [metricKey, value] of Object.entries(regionValues)) {
          await upsertMetric(
            partitionKey,
            `INPUT|${metricKey}`,
            block.entity,
            weekEnding,
            metricKey,
            value,
            "CXNS",
            "DynamicRegionInput"
          );
        }

        await upsertStatus(block.entity, weekEnding);

        if (!cxnsSharedAgg.has(weekEnding)) {
          cxnsSharedAgg.set(weekEnding, {
            scheduledVisits: 0,
            cancellations: 0,
            noShows: 0,
            rescheduledVisits: 0
          });
        }

        const agg = cxnsSharedAgg.get(weekEnding);
        agg.scheduledVisits += scheduledVisits || 0;
        agg.cancellations += cancellations || 0;
        agg.noShows += noShows || 0;
        agg.rescheduledVisits += rescheduledVisits || 0;
      }
    }

    for (const [weekEnding, agg] of cxnsSharedAgg.entries()) {
      const partitionKey = `CXNS|${weekEnding}`;

      const cancellationRate =
        agg.scheduledVisits > 0
          ? Number(((agg.cancellations / agg.scheduledVisits) * 100).toFixed(4))
          : null;

      const noShowRate =
        agg.scheduledVisits > 0
          ? Number(((agg.noShows / agg.scheduledVisits) * 100).toFixed(4))
          : null;

      const sharedValues = {
        scheduledVisits: agg.scheduledVisits,
        rescheduledVisits: agg.rescheduledVisits,
        cancellationRate,
        noShowRate
      };

      for (const [metricKey, value] of Object.entries(sharedValues)) {
        await upsertMetric(
          partitionKey,
          `SHARED|${metricKey}`,
          "CXNS",
          weekEnding,
          metricKey,
          value,
          "CXNS",
          "SharedPageInput"
        );
      }

      await upsertStatus("CXNS", weekEnding);
    }

    const ptRows = getSheetRows(workbook, "PT");
    const ptSharedAgg = new Map();

    for (const block of PT_BLOCKS) {
      for (let i = 12; i < ptRows.length; i++) {
        const row = ptRows[i];
        const week = toNumber(getCell(row, block.startCol));
        const monthMarker = getCell(row, block.startCol + 1);

        if (!week || !monthMarker) continue;

        const scheduledVisits = toNumber(getCell(row, block.startCol + 2));
        const cancellations = toNumber(getCell(row, block.startCol + 3));
        const noShows = toNumber(getCell(row, block.startCol + 4));
        const rescheduledVisits = toNumber(getCell(row, block.startCol + 5));
        const ptUnits = toNumber(getCell(row, block.startCol + 6));
        const ptVisits = toNumber(getCell(row, block.startCol + 9));

        const weekEnding = getWeekEndingFromIsoWeek(YEAR, week);

        if (!ptSharedAgg.has(weekEnding)) {
          ptSharedAgg.set(weekEnding, {
            ptVisits: 0,
            ptUnits: 0,
            ptCancellations: 0,
            ptNoShows: 0,
            ptReschedules: 0,
            ptScheduledVisits: 0
          });
        }

        const agg = ptSharedAgg.get(weekEnding);
        agg.ptVisits += ptVisits || 0;
        agg.ptUnits += ptUnits || 0;
        agg.ptCancellations += cancellations || 0;
        agg.ptNoShows += noShows || 0;
        agg.ptReschedules += rescheduledVisits || 0;
        agg.ptScheduledVisits += scheduledVisits || 0;
      }
    }

    for (const [weekEnding, agg] of ptSharedAgg.entries()) {
      const partitionKey = `PT|${weekEnding}`;

      for (const [metricKey, value] of Object.entries(agg)) {
        await upsertMetric(
          partitionKey,
          `SHARED|${metricKey}`,
          "PT",
          weekEnding,
          metricKey,
          value,
          "PT",
          "SharedPageInput"
        );
      }

      await upsertStatus("PT", weekEnding);
    }

    context.res = {
      status: 200,
      body: {
        ok: true,
        fileName,
        sheets: workbook.SheetNames,
        weeklyInputCount,
        submissionCount,
        touchedWeekCount: touchedWeeks.size,
        message: "Workbook imported successfully"
      }
    };
  } catch (err) {
    context.log.error(err);
    context.res = {
      status: 500,
      body: { error: err.message }
    };
  }
};