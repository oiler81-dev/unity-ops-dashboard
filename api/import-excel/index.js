module.exports = async function (context, req) {
  try {
    context.log("import-excel invoked", {
      method: req && req.method ? req.method : "",
      url: req && req.url ? req.url : ""
    });

    context.res = {
      status: 200,
      headers: {
        "Content-Type": "application/json"
      },
      body: {
        ok: true,
        message: "import-excel route is alive",
        method: req && req.method ? req.method : "",
        time: new Date().toISOString()
      }
    };
  } catch (error) {
    context.log.error("import-excel test function failed", error);

    context.res = {
      status: 500,
      headers: {
        "Content-Type": "application/json"
      },
      body: {
        ok: false,
        error: "import-excel test function failed",
        details: error && error.message ? error.message : String(error)
      }
    };
  }
};
