module.exports = async function (context, req) {
  context.res = {
    status: 200,
    body: {
      ok: true,
      message: "budget-get is LIVE",
      time: new Date().toISOString()
    }
  };
};
