module.exports = async function (context, req) {
  context.log("importexcel2 hit");

  return {
    status: 200,
    headers: {
      "Content-Type": "application/json"
    },
    body: {
      ok: true,
      message: "import-excel route is alive from importexcel2",
      method: req?.method || "",
      time: new Date().toISOString()
    }
  };
};
