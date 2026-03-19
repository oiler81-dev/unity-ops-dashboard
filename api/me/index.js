function decodeClientPrincipal(encoded) {
  if (!encoded) return null;

  try {
    const json = Buffer.from(encoded, "base64").toString("utf8");
    return JSON.parse(json);
  } catch {
    return null;
  }
}

function unique(values) {
  return Array.from(new Set((Array.isArray(values) ? values : []).filter(Boolean)));
}

function normalizeText(value) {
  return String(value || "").trim();
}

function normalizeEmail(value) {
  return normalizeText(value).toLowerCase();
}

function normalizeEntity(value) {
  const raw = normalizeText(value);
  const upper = raw.toUpperCase();

  if (upper === "LAOSS") return "LAOSS";
  if (upper === "NES") return "NES";
  if (upper === "SPINEONE") return "SpineOne";
  if (upper === "MRO") return "MRO";

  return raw;
}

function getConfiguredAdminEmails() {
  const raw =
    process.env.ADMIN_EMAILS ||
    process.env.ADMIN_USERS ||
    "nperez@unitymsk.com,tessa.kelley@spineone.com";

  return raw
    .split(",")
    .map((v) => normalizeEmail(v))
    .filter(Boolean);
}

function getUserEmails(principal) {
  const emails = [];

  if (principal?.userDetails) emails.push(principal.userDetails);
  if (principal?.userId) emails.push(principal.userId);

  const claims = Array.isArray(principal?.claims) ? principal.claims : [];
  for (const claim of claims) {
    if (!claim || !claim.typ) continue;

    const typ = String(claim.typ).toLowerCase();
    if (
      typ.includes("email") ||
      typ.endsWith("/upn") ||
      typ.endsWith("/name")
    ) {
      emails.push(claim.val);
    }
  }

  return unique(emails.map(normalizeEmail).filter(Boolean));
}

function hasAdminRole(roles) {
  return roles.some((role) => normalizeText(role).toLowerCase() === "admin");
}

function findEntityFromRoles(roles) {
  const normalizedRoles = roles.map((r) => normalizeText(r));
  const upperRoles = normalizedRoles.map((r) => r.toUpperCase());

  if (upperRoles.includes("LAOSS")) return "LAOSS";
  if (upperRoles.includes("NES")) return "NES";
  if (upperRoles.includes("SPINEONE")) return "SpineOne";
  if (upperRoles.includes("MRO")) return "MRO";

  return "";
}

function buildDebug(principal, roles, emails, isAdmin, entity) {
  return {
    userDetails: principal?.userDetails || "",
    userId: principal?.userId || "",
    identityProvider: principal?.identityProvider || "",
    roles,
    emails,
    isAdmin,
    entity
  };
}

module.exports = async function (context, req) {
  try {
    const headers = req.headers || {};
    const clientPrincipalHeader =
      headers["x-ms-client-principal"] ||
      headers["X-MS-CLIENT-PRINCIPAL"];

    const principal = decodeClientPrincipal(clientPrincipalHeader);

    if (!principal || !principal.userId) {
      context.res = {
        status: 200,
        headers: { "Content-Type": "application/json" },
        body: {
          authenticated: false,
          userDetails: "",
          roles: ["anonymous"],
          entity: "",
          isAdmin: false
        }
      };
      return;
    }

    const roles = unique(principal.userRoles || []);
    const emails = getUserEmails(principal);
    const configuredAdminEmails = getConfiguredAdminEmails();

    const emailMatchedAdmin = emails.some((email) =>
      configuredAdminEmails.includes(email)
    );

    const isAdmin = hasAdminRole(roles) || emailMatchedAdmin;

    let entity = "";
    if (isAdmin) {
      entity = "Admin";
    } else {
      entity = normalizeEntity(findEntityFromRoles(roles));
    }

    context.res = {
      status: 200,
      headers: { "Content-Type": "application/json" },
      body: {
        authenticated: true,
        userId: principal.userId || "",
        userDetails: principal.userDetails || principal.userId || "",
        identityProvider: principal.identityProvider || "",
        roles,
        entity,
        isAdmin,
        debug: buildDebug(principal, roles, emails, isAdmin, entity)
      }
    };
  } catch (error) {
    context.log.error("me failed", error);

    context.res = {
      status: 500,
      headers: { "Content-Type": "application/json" },
      body: {
        authenticated: false,
        error: "Failed to resolve current user.",
        details: error && error.message ? error.message : String(error)
      }
    };
  }
};
