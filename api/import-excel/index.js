module.exports = async function (context, req) {
  context.res = {
    status: 200,
    headers: {
      "Content-Type": "application/json"
    },
    body: {
      ok: true,
      message: "import-excel route is alive",
      method: req.method || "",
      time: new Date().toISOString()
    }
  };
};
