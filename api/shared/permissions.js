const { getTableClient } = require("./table");
const { normalizeEmail } = require("./auth");

const FALLBACK_USERS = {
  "nperez@unitymsk.com": {
    displayName: "Nestor Perez",
    email: "nperez@unitymsk.com",
    role: "Admin",
    entity: "All",
    isAdmin: true,
    canEdit: true,
    isActive: true
  },
  "tessa.kelley@spineone.com": {
    displayName: "Tessa Kelley",
    email: "tessa.kelley@spineone.com",
    role: "Admin",
    entity: "SpineOne",
    isAdmin: true,
    canEdit: true,
    isActive: true
  },
  "tguerrero@laorthos.com": {
    displayName: "Tony Guerrero",
    email: "tguerrero@laorthos.com",
    role: "Editor",
    entity: "LAOSS",
    isAdmin: false,
    canEdit: true,
    isActive: true
  },
  "agutierrez@nespecialists.com": {
    displayName: "Annette Gutierrez",
    email: "agutierrez@nespecialists.com",
    role: "Editor",
    entity: "NES",
    isAdmin: false,
    canEdit: true,
    isActive: true
  },
  "chris.zamucen@spineone.com": {
    displayName: "Chris Zamucen",
    email: "chris.zamucen@spineone.com",
    role: "Editor",
    entity: "SpineOne",
    isAdmin: false,
    canEdit: true,
    isActive: true
  },
  "glundgren@mrorthopedics.com": {
    displayName: "Greg Lundgren",
    email: "glundgren@mrorthopedics.com",
    role: "Editor",
    entity: "MRO",
    isAdmin: false,
    canEdit: true,
    isActive: true
  }
};

async function getPermissionByEmail(email) {
  const normalized = normalizeEmail(email);
  if (!normalized) return null;

  const table = getTableClient("UserPermissions");
  const stored = await table.getEntity("USER", normalized);

  if (stored) {
    return {
      displayName: stored.displayName,
      email: stored.email,
      role: stored.role,
      entity: stored.entity,
      isAdmin: Boolean(stored.isAdmin),
      canEdit: Boolean(stored.canEdit),
      isActive: stored.isActive !== false
    };
  }

  return FALLBACK_USERS[normalized] || null;
}

function canEditEntity(permission, entity) {
  if (!permission || !permission.isActive) return false;
  if (permission.isAdmin) return true;
  return permission.canEdit && permission.entity === entity;
}

module.exports = {
  getPermissionByEmail,
  canEditEntity
};
