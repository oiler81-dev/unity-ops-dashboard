const { getUserInfo } = require('shared/auth');
const { hasPermission } = require('shared/permissions');
const { saveWeeklyData } = require('shared/table');
const { logAuditEvent } = require('shared/audit');

module.exports = async function (context, req) {
  context.log('Weekly save function processed a request.');

  const userInfo = getUserInfo(req);
  const data = req.body;
  const { entity, weekEnding } = data;

  if (!entity || !weekEnding || !data.inputs) {
    context.res = { status: 400, body: 'Invalid payload for weekly save.' };
    return;
  }

  if (!hasPermission(userInfo, 'canEditRegion', entity)) {
    context.res = { status: 403, body: 'Forbidden' };
    return;
  }

  try {
    const result = await saveWeeklyData(data, userInfo);
    await logAuditEvent('weekly-save', userInfo, { entity, weekEnding, status: result.status });
    
    context.res = {
      status: 200,
      body: result,
    };
  } catch (error) {
    context.log.error('Failed to save weekly data:', error);
    context.res = { status: 500, body: 'Error saving weekly data.' };
  }
};
