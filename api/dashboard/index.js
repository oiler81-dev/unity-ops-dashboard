const { getUserInfo } = require("../shared/auth");
const { getDashboardData } = require("../shared/workbookLogic");
const { hasPermission } = require("../shared/permissions");

module.exports = async function (context, req) {
  context.log("Dashboard function processed a request.");

  try {
    const userInfo = getUserInfo(req);

    if (!hasPermission(userInfo, "canViewExecutiveSummary")) {
      context.res = {
        status: 403,
        headers: { "Content-Type": "application/json" },
        body: {
          error: "Forbidden"
        }
      };
      return;
    }

    const weekEnding = req.query.weekEnding;
    if (!weekEnding) {
      context.res = {
        status: 400,
        headers: { "Content-Type": "application/json" },
        body: {
          error: "Missing weekEnding query parameter."
        }
      };
      return;
    }

    const dashboardData = await getDashboardData(weekEnding);

    context.res = {
      status: 200,
      headers: { "Content-Type": "application/json" },
      body: dashboardData
    };
  } catch (error) {
    context.log.error("Failed to get dashboard data:", error);

    context.res = {
      status: 500,
      headers: { "Content-Type": "application/json" },
      body: {
        error: "Error fetching dashboard data.",
        details: error.message
      }
    };
  }
};
