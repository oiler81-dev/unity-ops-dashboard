const { getUserInfo } = require('shared/auth');
const { hasPermission } = require('shared/permissions');
const { saveSharedPageData } = require('shared/sharedPageLogic');
const { logAuditEvent } = require('shared/audit');

module.exports = async function (context, req) {
    context.log('Shared save function processed a request.');

    const userInfo = getUserInfo(req);
    const data = req.body;
    const { page, weekEnding } = data;

    if (!page || !weekEnding || !data.inputs) {
        context.res = { status: 400, body: 'Invalid payload for shared page save.' };
        return;
    }

    if (!hasPermission(userInfo, 'canEditSharedPage', page)) {
        context.res = { status: 403, body: 'Forbidden' };
        return;
    }

    try {
        const result = await saveSharedPageData(data, userInfo);
        await logAuditEvent('shared-save', userInfo, { page, weekEnding, status: result.status });

        context.res = {
            status: 200,
            body: result,
        };
    } catch (error) {
        context.log.error('Failed to save shared data:', error);
        context.res = { status: 500, body: 'Error saving shared data.' };
    }
};
