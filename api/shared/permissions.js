const { REGION_USER_MAP } = require("./constants");

function resolveAccess(user) {
  const email = (user?.userDetails || "").toLowerCase();
  const mapped = REGION_USER_MAP[email];

  if (!user?.authenticated || !email) {
    return {
      authenticated: false,
      email: null,
      role: "guest",
      entity: null,
      allowed: false,
      isAdmin: false
    };
  }

  if (mapped) {
    return {
      authenticated: true,
      email,
      role: mapped.role,
      entity: mapped.entity,
      allowed: true,
      isAdmin: mapped.role === "admin"
    };
  }

  return {
    authenticated: true,
    email,
    role: "user",
    entity: null,
    allowed: true,
    isAdmin: false
  };
}

function canAccessEntity(access, entity) {
  return !!access?.authenticated;
}

module.exports = {
  resolveAccess,
  canAccessEntity
};