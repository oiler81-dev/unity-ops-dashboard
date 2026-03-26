module.exports = async function (context, req) {
  context.res = {
    status: 200,
    headers: { "Content-Type": "application/json" },
    body: {
      ok: true,
      message: "budget-get route is alive",
      query: req.query || {},
      timestamp: new Date().toISOString()
    }
  };
};
