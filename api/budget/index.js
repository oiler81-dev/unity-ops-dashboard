const { getTableClient } = require("../shared/table");

module.exports = async function (context, req) {
  try {
    const table = getTableClient("BudgetData");

    context.res = {
      status: 200,
      headers: { "Content-Type": "application/json" },
      body: {
        ok: true,
        message: "budget route is alive",
        tableClientLoaded: !!table
      }
    };
  } catch (error) {
    context.res = {
      status: 500,
      headers: { "Content-Type": "application/json" },
      body: {
        ok: false,
        error: error.message
      }
    };
  }
};
