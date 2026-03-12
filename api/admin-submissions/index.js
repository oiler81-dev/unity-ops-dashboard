const { getTableClient } = require("../shared/table");

module.exports = async function (context, req) {
  try {
    const table = getTableClient("DashboardSubmissions");
    const results = [];

    for await (const entity of table.listEntities()) {
      results.push(entity);
    }

    results.sort((a, b) => {
      const aTime = new Date(a.timestamp || 0).getTime();
      const bTime = new Date(b.timestamp || 0).getTime();
      return bTime - aTime;
    });

    context.res = {
      status: 200,
      headers: {
        "Content-Type": "application/json"
      },
      body: {
        ok: true,
        count: results.length,
        items: results
      }
    };
  } catch (err) {
    context.log.error(err);

    context.res = {
      status: 500,
      body: {
        ok: false,
        error: err.message
      }
    };
  }
};
