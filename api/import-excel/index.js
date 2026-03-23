module.exports = async function (context, req) {
  try {
    context.log("import-excel reached");

    return (context.res = {
      status: 200,
      headers: {
        "Content-Type": "application/json"
      },
      body: {
        ok: true,
        message: "import-excel basic test alive",
        method: req?.method || "",
        hasBody: !!req?.body,
        time: new Date().toISOString()
      }
    });
  } catch (error) {
    context.log("import-excel hard fail", error);

    return (context.res = {
      status: 500,
      headers: {
        "Content-Type": "application/json"
      },
      body: {
        ok: false,
        error: "basic test failed",
        details: error?.message || String(error)
      }
    });
  }
};
