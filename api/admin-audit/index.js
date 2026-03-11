const { getUserFromRequest, getUserEmail } = require("../shared/auth");
const { getPermissionByEmail } = require("../shared/permissions");
const { getTableClient } = require("../shared/table");

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

    const weekEnding = req.query.weekEnding || new Date().toISOString().slice(0, 10);
    const entity = req.query.entity || "";
    const auditTable = getTableClient("AuditLog");

    const partitions = entity
      ? [`${entity}|${weekEnding}`]
      : ["LAOSS", "NES", "SpineOne", "MRO", "PT", "CXNS", "Capacity", "Productivity Builder"]
          .map((key) => `${key}|${weekEnding}`);

    const rows = [];
    for (const partitionKey of partitions) {
      const items = await auditTable.listByPartition(partitionKey);
      rows.push(...items);
    }

    rows.sort((a, b) => String(b.changedAt || "").localeCompare(String(a.changedAt || "")));

    context.res = {
      status: 200,
      body: {
        weekEnding,
        entity: entity || "All",
        rows: rows.slice(0, 250).map((row) => ({
          entity: row.entity || "",
          weekEnding: row.weekEnding || "",
          section: row.section || "",
          metricKey: row.metricKey || "",
          oldValue: row.oldValue ?? "",
          newValue: row.newValue ?? "",
          changedBy: row.changedBy || "",
          changedAt: row.changedAt || "",
          changeType: row.changeType || ""
        }))
      }
    };
  } catch (err) {
    context.log.error(err);
    context.res = { status: 500, body: { error: err.message } };
  }
};