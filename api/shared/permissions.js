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
    role: "guest",
    entity: null,
    allowed: false,
    isAdmin: false
  };
}

function canAccessEntity(access, entity) {
  if (!access?.allowed) return false;
  if (access.isAdmin) return true;
  return access.entity === entity;
}

module.exports = {
  resolveAccess,
  canAccessEntity
};
