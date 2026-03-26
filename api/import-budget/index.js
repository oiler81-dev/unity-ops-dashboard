module.exports = async function (context, req) {
  context.res = {
    status: 200,
    headers: { "Content-Type": "application/json" },
    body: {
      ok: true,
      message: "import-budget route is alive",
      method: req.method,
      hasBody: !!req.body,
      timestamp: new Date().toISOString()
    }
  };
};
