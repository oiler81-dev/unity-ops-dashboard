const { getUserFromRequest, getUserEmail } = require("../shared/auth");
const { getPermissionByEmail } = require("../shared/permissions");
const { getTableClient } = require("../shared/table");

const AUDIT_SCOPES = [
  "LAOSS",
  "NES",
  "SpineOne",
  "MRO",
  "PT",
  "CXNS",
  "Capacity",
  "Productivity Builder"
];

module.exports = async function (context, req) {
  try {
    const user = getUserFromRequest(req);
    if (!user) {
      context.res = {
        status: 401,
        body: { error: "Not authenticated" }
      };
      return;
    }

    const email = getUserEmail(user);
    const permission = await getPermissionByEmail(email);

    if (!permission || !permission.isAdmin) {
      context.res = {
        status: 403,
        body: { error: "Admin access required" }
      };
      return;
    }

    const weekEnding = req.query.weekEnding || new Date().toISOString().slice(0, 10);
    const entity = (req.query.entity || "").trim();

    const auditTable = getTableClient("AuditLog");

    const partitionKeys = entity
      ? [`${entity}|${weekEnding}`]
      : AUDIT_SCOPES.map((scope) => `${scope}|${weekEnding}`);

    const allRows = [];

    for (const partitionKey of partitionKeys) {
      const rows = await auditTable.listByPartition(partitionKey);
      allRows.push(...rows);
    }

    allRows.sort((a, b) => {
      const aTime = String(a.changedAt || "");
      const bTime = String(b.changedAt || "");
      return bTime.localeCompare(aTime);
    });

    context.res = {
      status: 200,
      body: {
        weekEnding,
        entity: entity || "All",
        rows: allRows.slice(0, 250).map((row) => ({
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
    context.res = {
      status: 500,
      body: { error: err.message }
    };
  }
};