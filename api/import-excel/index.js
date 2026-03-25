const XLSX = require("xlsx");
const { getUserFromRequest } = require("../shared/auth");
const { resolveAccess } = require("../shared/permissions");
const { ok, badRequest, forbidden, serverError } = require("../shared/response");
const { ensureTable } = require("../shared/table");
const { WEEKLY_TABLE } = require("../shared/constants");

const SHEET_ENTITY_MAP = {
  LA: "LAOSS",
  Portland: "NES",
  Denver: "SpineOne",
  Chicago: "MRO"
};

const KPI_KEY_MAP = {
  visitvolume: "visitVolume",
  callvolume: "callVolume",
  newpatients: "newPatients",
  noshowrate: "noShowRate",
  cancellationrate: "cancellationRate",
  abandonedcallrate: "abandonedCallRate"
};

function normalizeKey(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

function excelDateToIso(value) {
  if (!value && value !== 0) return null;

  if (typeof value === "string") {
    const trimmed = value.trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(trimmed)) return trimmed;

    const date = new Date(trimmed);
    if (!Number.isNaN(date.getTime())) {
      return date.toISOString().slice(0, 10);
    }
    return null;
  }

  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (!parsed) return null;
    const yyyy = String(parsed.y).padStart(4, "0");
    const mm = String(parsed.m).padStart(2, "0");
    const dd = String(parsed.d).padStart(2, "0");
    return `${yyyy}-${mm}-${dd}`;
  }

  return null;
}

function parseSheetRows(sheet) {
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: null });
  const items = [];

  for (const row of rows) {
    const normalized = {};
    for (const [key, value] of Object.entries(row)) {
      normalized[normalizeKey(key)] = value;
    }

    const weekEnding =
      excelDateToIso(normalized.weekending) ||
      excelDateToIso(normalized.weekend) ||
      excelDateToIso(normalized.date);

    if (!weekEnding) continue;

    const data = {};
    for (const [sourceKey, targetKey] of Object.entries(KPI_KEY_MAP)) {
      const raw = normalized[sourceKey];
      data[targetKey] =
        raw === null || raw === undefined || raw === "" ? null : Number(raw);
    }

    items.push({ weekEnding, data });
  }

  return items;
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
    const fileName = body.fileName || "workbook.xlsx";

    if (!fileBase64) {
      return badRequest("Missing fileBase64");
    }

    const buffer = Buffer.from(fileBase64, "base64");
    const workbook = XLSX.read(buffer, { type: "buffer" });
    const client = await ensureTable(WEEKLY_TABLE);

    const imported = [];
    const skippedSheets = [];

    for (const [sheetName, entity] of Object.entries(SHEET_ENTITY_MAP)) {
      const sheet = workbook.Sheets[sheetName];
      if (!sheet) {
        skippedSheets.push({ sheetName, reason: "Sheet not found" });
        continue;
      }

      const rows = parseSheetRows(sheet);
      let written = 0;

      for (const row of rows) {
        const entityRecord = {
          partitionKey: entity,
          rowKey: row.weekEnding,
          ...row.data,
          status: "approved",
          submittedBy: access.email,
          submittedAt: new Date().toISOString(),
          approvedBy: access.email,
          approvedAt: new Date().toISOString(),
          updatedBy: access.email,
          updatedAt: new Date().toISOString(),
          importFileName: fileName,
          importSourceSheet: sheetName
        };

        await client.upsertEntity(entityRecord, "Replace");
        written += 1;
      }

      imported.push({
        sheetName,
        entity,
        rowsRead: rows.length,
        rowsWritten: written
      });
    }

    return ok({
      ok: true,
      message: "Import completed",
      fileName,
      workbookSheets: workbook.SheetNames,
      imported,
      skippedSheets
    });
  } catch (error) {
    context.log.error("import-excel failed", error);
    return serverError(error, "Failed to import workbook");
  }
};
