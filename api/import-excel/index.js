const XLSX = require("xlsx");

module.exports = async function (context, req) {
  try {
    context.log("IMPORT STARTED");

    if (!req.body || !req.body.fileBase64) {
      return (context.res = {
        status: 400,
        body: { ok: false, error: "Missing fileBase64" }
      });
    }

    const buffer = Buffer.from(req.body.fileBase64, "base64");

    const workbook = XLSX.read(buffer, { type: "buffer" });

    const sheetNames = workbook.SheetNames;

    context.log("Sheets found:", sheetNames);

    let totalRows = 0;

    const parsedData = {};

    for (const sheetName of sheetNames) {
      const sheet = workbook.Sheets[sheetName];
      const json = XLSX.utils.sheet_to_json(sheet, { defval: null });

      parsedData[sheetName] = json;
      totalRows += json.length;
    }

    return (context.res = {
      status: 200,
      body: {
        ok: true,
        message: "Workbook parsed successfully",
        sheets: sheetNames,
        totalRows
      }
    });
  } catch (error) {
    context.log.error("IMPORT FAILED", error);

    return (context.res = {
      status: 500,
      body: {
        ok: false,
        error: "Import failed",
        details: error.message
      }
    });
  }
};
