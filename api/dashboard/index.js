module.exports = async function (context, req) {
  try {
    const weekEnding = req.query.weekEnding || null;

    context.res = {
      status: 200,
      headers: { "Content-Type": "application/json" },
      body: {
        ok: true,
        weekEnding,
        message: "Dashboard API booted successfully"
      }
    };
  } catch (err) {
    context.log.error("api/dashboard failed", err);

    context.res = {
      status: 500,
      headers: { "Content-Type": "application/json" },
      body: {
        error: "api/dashboard failed",
        details: err.message
      }
    };
  }
};
