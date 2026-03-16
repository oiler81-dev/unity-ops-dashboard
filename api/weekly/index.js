const { getUserInfo } = require('shared/auth');
const { hasPermission } = require('shared/permissions');
const { getWeeklyData } = require('shared/table');

module.exports = async function (context, req) {
  context.log('Weekly data function processed a request.');

  const userInfo = getUserInfo(req);
  const entity = req.query.entity;
  const weekEnding = req.query.weekEnding;

  if (!entity || !weekEnding) {
    context.res = { status: 400, body: 'Missing entity or weekEnding query parameters.' };
    return;
  }

  if (!hasPermission(userInfo, 'canViewRegion', entity)) {
    context.res = { status: 403, body: 'Forbidden' };
    return;
  }

  try {
    const weeklyData = await getWeeklyData(entity, weekEnding);
    context.res = {
      status: 200,
      body: weeklyData,
    };
  } catch (error) {
    context.log.error('Failed to get weekly data:', error);
    context.res = { status: 500, body: 'Error fetching weekly data.' };
  }
};
