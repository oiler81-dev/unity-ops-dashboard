const { getUserInfo } = require('shared/auth');
const { hasPermission } = require('shared/permissions');
const { saveReferenceData } = require('shared/reference');
const { logAuditEvent } = require('shared/audit');

module.exports = async function (context, req) {
    context.log('Admin reference save function processed a request.');

    const userInfo = getUserInfo(req);
    if (!hasPermission(userInfo, 'isAdmin')) {
        context.res = { status: 403, body: 'Forbidden' };
        return;
    }

    const { kind, entity, year, rows } = req.body;
    if (!kind || !rows) {
        context.res = { status: 400, body: 'Invalid payload for reference save.' };
        return;
    }

    try {
        const result = await saveReferenceData({ kind, entity, year, rows });
        await logAuditEvent('admin-reference-save', userInfo, { kind, entity, year });

        context.res = {
            status: 200,
            body: { ok: true, message: `${kind} data saved successfully.` },
        };
    } catch (error) {
        context.log.error('Failed to save reference data:', error);
        context.res = { status: 500, body: 'Error saving reference data.' };
    }
};
