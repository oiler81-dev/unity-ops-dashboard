function getUserFromRequest(req) {
  return {
    authenticated: true,
    userDetails: "nperez@unitymsk.com",
    roles: ["admin"]
  };
}
module.exports = { getUserFromRequest };