module.exports = async function (context, req) {
  try {
    const principal = req.headers["x-ms-client-principal"];

    if (!principal) {
      context.res = {
        status: 200,
        headers: { "Content-Type": "application/json" },
        body: {
          authenticated: false,
          userDetails: null,
          userId: null,
          identityProvider: null,
          roles: ["anonymous"]
        }
      };
      return;
    }

    const decoded = JSON.parse(Buffer.from(principal, "base64").toString("utf8"));
    const roles = Array.isArray(decoded.userRoles) ? decoded.userRoles : [];

    context.res = {
      status: 200,
      headers: { "Content-Type": "application/json" },
      body: {
        authenticated: true,
        userDetails: decoded.userDetails || null,
        userId: decoded.userId || null,
        identityProvider: decoded.identityProvider || null,
        roles
      }
    };
  } catch (err) {
    context.log.error("api/me failed", err);

    context.res = {
      status: 500,
      headers: { "Content-Type": "application/json" },
      body: {
        error: "api/me failed",
        details: err.message
      }
    };
  }
};
