const { getUserInfo } = require('shared/auth');

module.exports = async function (context, req) {
  context.log('Me function processed a request.');

  const userInfo = getUserInfo(req);

  context.res = {
    status: 200,
    body: userInfo,
  };
};
