function resolveAccess() {
  return {
    authenticated: true,
    role: "admin",
    entity: "admin",
    allowed: true,
    isAdmin: true
  };
}
module.exports = { resolveAccess };