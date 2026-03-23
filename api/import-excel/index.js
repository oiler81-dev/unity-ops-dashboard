module.exports = async function (context, req) {
  context.log("import-excel fresh function hit");

  return {
    status: 200,
    headers: {
      "Content-Type": "application/json"
    },
    body: {
      ok: true,
      message: "fresh import-excel function is alive",
      method: req?.method || "",
      time: new Date().toISOString()
    }
  };
};
