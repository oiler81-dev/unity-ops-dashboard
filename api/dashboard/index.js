const { getUserInfo } = require('shared/auth');
const { getDashboardData } = require('shared/workbookLogic');
const { hasPermission } = require('shared/permissions');

module.exports = async function (context, req) {
  context.log('Dashboard function processed a request.');

  const userInfo = getUserInfo(req);
  if (!hasPermission(userInfo, 'canViewExecutiveSummary')) {
    context.res = { status: 403, body: 'Forbidden' };
    return;
  }

  const weekEnding = req.query.weekEnding;
  if (!weekEnding) {
    context.res = { status: 400, body: 'Missing weekEnding query parameter.' };
    return;
  }

  try {
    const dashboardData = await getDashboardData(weekEnding);
    context.res = {
      status: 200,
      body: dashboardData,
    };
  } catch (error) {
    context.log.error('Failed to get dashboard data:', error);
    context.res = { status: 500, body: 'Error fetching dashboard data.' };
  }
};
