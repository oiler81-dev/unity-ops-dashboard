const XLSX = require("xlsx");
const { getUserFromRequest } = require("../shared/auth");
const { resolveAccess } = require("../shared/permissions");
const { ok, badRequest, forbidden, serverError } = require("../shared/response");
const { ensureTable } = require("../shared/table");

const BUDGET_TABLE = "BudgetReferenceData";

function normalizeEntityName(value) {
  const raw = String(value || "").trim().toLowerCase();

  if (!raw) return null;
  if (raw === "laoss") return "LAOSS";
  if (raw === "nes") return "NES";
  if (raw === "chicago") return "MRO";
  if (raw === "mro") return "MRO";
  if (raw === "so") return "SpineOne";
  if (raw === "spineone") return "SpineOne";

  return null;
}

function monthKeyFromCell(value) {
  if (!value) return null;

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return `${value.getFullYear()}-${String(value.getMonth() + 1).padStart(2, "0")}`;
  }

  const parsed = new Date(value);
  if (!Number.isNaN(parsed.getTime())) {
    return `${parsed.getFullYear()}-${String(parsed.getMonth() + 1).padStart(2, "0")}`;
  }

  return null;
}

function toNumber(value) {
  if (value == null || value === "") return null;
  const n = Number(value);
  return Number.isFinite(n) ? n : null;
}

function getCell(sheet, row, col) {
  const address = XLSX.utils.encode_cell({ r: row - 1, c: col - 1 });
  return sheet[address] ? sheet[address].v : null;
}

function parseVolumeBudgetSheet(sheet) {
  const monthColumns = [];
  for (let col = 4; col <= 15; col += 1) {
    const monthKey = monthKeyFromCell(getCell(sheet, 8, col));
    if (!monthKey) continue;

    monthColumns.push({
      col,
      monthKey,
      workdays: toNumber(getCell(sheet, 9, col)) || 0
    });
  }

  const visitBudgetMap = {};
  const workdayMap = {};

  for (const item of monthColumns) {
    workdayMap[item.monthKey] = item.workdays;
  }

  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");

  for (let row = 11; row <= range.e.r + 1; row += 1) {
    const entity = normalizeEntityName(getCell(sheet, row, 2));
    const apptLookup = String(getCell(sheet, row, 3) || "").trim();

    if (!entity) continue;
    if (!apptLookup) continue;
    if (apptLookup.toLowerCase() === "days") continue;
    if (apptLookup.toLowerCase() === "total") continue;
    if (apptLookup.toLowerCase() === "per day") continue;

    for (const item of monthColumns) {
      const raw = toNumber(getCell(sheet, row, item.col));
      if (raw == null) continue;

      if (!visitBudgetMap[entity]) visitBudgetMap[entity] = {};
      if (!visitBudgetMap[entity][item.monthKey]) visitBudgetMap[entity][item.monthKey] = 0;

      visitBudgetMap[entity][item.monthKey] += raw;
    }
  }

  return {
    visitBudgetMap,
    workdayMap
  };
}

function parseNewPatientBudgetSheet(sheet) {
  const monthColumns = [];
  for (let col = 3; col <= 14; col += 1) {
    const monthKey = monthKeyFromCell(getCell(sheet, 6, col));
    if (!monthKey) continue;

    monthColumns.push({
      col,
      monthKey,
      workdays: toNumber(getCell(sheet, 7, col)) || 0
    });
  }

  const npBudgetMap = {};
  const workdayMap = {};

  for (const item of monthColumns) {
    workdayMap[item.monthKey] = item.workdays;
  }

  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");

  for (let row = 8; row <= range.e.r + 1; row += 1) {
    const entity = normalizeEntityName(getCell(sheet, row, 2));
    if (!entity) continue;

    for (const item of monthColumns) {
      const raw = toNumber(getCell(sheet, row, item.col));
      if (raw == null) continue;

      if (!npBudgetMap[entity]) npBudgetMap[entity] = {};
      npBudgetMap[entity][item.monthKey] = raw;
    }
  }

  return {
    npBudgetMap,
    workdayMap
  };
}

module.exports = async function (context, req) {
  try {
    const user = getUserFromRequest(req);
    const access = resolveAccess(user);

    if (!access.isAdmin) {
      return forbidden("Admin only");
    }

    const body = req.body || {};
    const fileBase64 = body.fileBase64;
    const fileName = body.fileName || "budget-workbook.xlsx";

    if (!fileBase64) {
      return badRequest("Missing fileBase64");
    }

    const buffer = Buffer.from(fileBase64, "base64");
    const workbook = XLSX.read(buffer, { type: "buffer", cellDates: true });

    const volumeSheet = workbook.Sheets["Volume_Revenue Budget"];
    const npSheet = workbook.Sheets["New Patient Budget"];

    if (!volumeSheet) {
      return badRequest("Missing 'Volume_Revenue Budget' sheet");
    }

    if (!npSheet) {
      return badRequest("Missing 'New Patient Budget' sheet");
    }

    const volumeParsed = parseVolumeBudgetSheet(volumeSheet);
    const npParsed = parseNewPatientBudgetSheet(npSheet);

    const entities = Array.from(
      new Set([
        ...Object.keys(volumeParsed.visitBudgetMap || {}),
        ...Object.keys(npParsed.npBudgetMap || {})
      ])
    );

    const monthKeys = Array.from(
      new Set([
        ...Object.keys(volumeParsed.workdayMap || {}),
        ...Object.keys(npParsed.workdayMap || {})
      ])
    ).sort();

    const client = await ensureTable(BUDGET_TABLE);
    const imported = [];

    for (const entity of entities) {
      for (const yearMonth of monthKeys) {
        const visitBudget = volumeParsed.visitBudgetMap?.[entity]?.[yearMonth] ?? null;
        const newPatientBudget = npParsed.npBudgetMap?.[entity]?.[yearMonth] ?? null;
        const workdays =
          volumeParsed.workdayMap?.[yearMonth] ??
          npParsed.workdayMap?.[yearMonth] ??
          null;

        if (visitBudget == null && newPatientBudget == null) {
          continue;
        }

        const record = {
          partitionKey: entity,
          rowKey: yearMonth,
          entity,
          yearMonth,
          visitBudget,
          newPatientBudget,
          callBudget: null,
          workdays,
          source: "budget-import",
          sourceWorkbook: fileName,
          importedBy: access.email,
          importedAt: new Date().toISOString()
        };

        await client.upsertEntity(record, "Replace");

        imported.push({
          entity,
          yearMonth,
          visitBudget,
          newPatientBudget,
          workdays
        });
      }
    }

    return ok({
      ok: true,
      message: "Budget import completed",
      workbookFile: fileName,
      workbookSheets: workbook.SheetNames,
      importedCount: imported.length,
      imported
    });
  } catch (error) {
    context.log.error("import-budget failed", error);
    return serverError(error, "Failed to import budget workbook");
  }
};
