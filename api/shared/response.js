function ok(body) {
  return {
    status: 200,
    headers: { "Content-Type": "application/json" },
    body
  };
}
module.exports = { ok };