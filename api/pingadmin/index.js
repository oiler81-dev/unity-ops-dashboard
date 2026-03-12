module.exports = async function (context, req) {
  context.log && context.log('pingadmin invoked');
  context.res = {
    status: 200,
    headers: { "Content-Type": "application/json" },
    body: { ok: true }
  };
};
