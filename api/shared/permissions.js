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
      entity: mapped.role === "admin" ? null : mapped.entity,
      allowed: true,
      isAdmin: mapped.role === "admin"
    };
  }

  // Unknown authenticated user (passed Azure AD but not in region map).
  // Surface in App Insights traces so a UPN typo or new-hire onboarding
  // gap shows up in monitoring instead of waiting for a user complaint.
  // console.warn is captured by the Functions host into App Insights
  // automatically, so no context threading is needed at every caller.
  // The user already passed Azure AD, so logging their email is staff
  // identity, not a privacy leak.
  try {
    console.warn(`auth: authenticated user not in REGION_USER_MAP: ${email}`);
  } catch (_) {
    // Logging must never throw on the auth path.
  }
  return {
    authenticated: true,
    email,
    role: "user",
    entity: null,
    allowed: false,
    isAdmin: false
  };
}

// Returns true only if the caller is an admin, or the requested entity
// matches the caller's own entity. Every function that reads or writes
// entity-scoped data MUST call this before the data operation.
function canAccessEntity(access, entity) {
  if (!access?.allowed) return false;
  if (access.isAdmin) return true;
  if (!entity) return false;
  return String(entity).toLowerCase() === String(access.entity || "").toLowerCase();
}

// Filter an explicit list of entities down to the ones the caller may access.
// Admins keep the whole list; regional users get only their own entity.
// Used by endpoints that default to "all four practices" when the request
// omits an entity parameter.
function scopeEntitiesToAccess(access, entities) {
  const list = Array.isArray(entities) ? entities : [];
  if (!access?.allowed) return [];
  if (access.isAdmin) return list.slice();
  if (!access.entity) return [];
  const match = String(access.entity).toLowerCase();
  return list.filter((e) => String(e || "").toLowerCase() === match);
}

// Standard 401/403 response builder. Use on every function right after
// resolveAccess() to reject unauthenticated or unmapped callers.
function requireAccess(access) {
  if (!access?.authenticated) {
    return {
      status: 401,
      body: { ok: false, error: "Unauthorized" }
    };
  }
  if (!access.allowed) {
    return {
      status: 403,
      body: { ok: false, error: "Forbidden" }
    };
  }
  return null;
}

// Standard 403 response for an authenticated user trying to access an entity
// outside their scope. Returns 404 to avoid leaking whether the entity exists
// at all — standard IDOR prevention.
function requireEntityAccess(access, entity) {
  if (!canAccessEntity(access, entity)) {
    return {
      status: 404,
      body: { ok: false, error: "Not found" }
    };
  }
  return null;
}

// Safely log the full error server-side but return a body that doesn't leak
// stack traces, internal paths, or exception messages to the client.
function safeErrorResponse(context, error, publicMessage = "Internal server error") {
  try {
    context?.log?.error?.(publicMessage, error);
  } catch (_) {
    // If logging itself fails, swallow — never rethrow from the error path.
  }
  return {
    status: 500,
    body: { ok: false, error: publicMessage }
  };
}

// Number coercion that rejects negative values and non-finite inputs.
// Use for any user-supplied numeric form field (surgeries, visits, etc.).
function toSafeNumber(value, fallback = 0, { min = 0, max = Number.MAX_SAFE_INTEGER } = {}) {
  if (value == null || value === "") return fallback;
  const n = Number(value);
  if (!Number.isFinite(n)) return fallback;
  if (n < min) return fallback;
  if (n > max) return fallback;
  return n;
}

module.exports = {
  resolveAccess,
  canAccessEntity,
  scopeEntitiesToAccess,
  requireAccess,
  requireEntityAccess,
  safeErrorResponse,
  toSafeNumber
};
