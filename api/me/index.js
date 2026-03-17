const { getUserInfo } = require("../shared/auth");

module.exports = async function (context, req) {
  context.log("Me function processed a request.");

  try {
    const userInfo = getUserInfo(req);

    context.res = {
      status: 200,
      headers: { "Content-Type": "application/json" },
      body: userInfo
    };
  } catch (error) {
    context.log.error("Failed in /api/me:", error);

    context.res = {
      status: 500,
      headers: { "Content-Type": "application/json" },
      body: {
        error: "Failed to load current user",
        details: error.message
      }
    };
  }
};
