function decodeClientPrincipal(encoded) {
  if (!encoded) return null;

  try {
    const json = Buffer.from(encoded, "base64").toString("utf8");
    return JSON.parse(json);
  } catch {
    return null;
  }
}

function normalizeText(value) {
  return String(value || "").trim();
}

function normalizeEmail(value) {
  return normalizeText(value).toLowerCase();
}

function getConfiguredAdminEmails() {
  const raw =
    process.env.ADMIN_EMAILS ||
    process.env.ADMIN_USERS ||
    "nperez@unitymsk.com,tessa.kelley@spineone.com";

  return raw
    .split(",")
    .map((item) => normalizeEmail(item))
    .filter(Boolean);
}

function getUserEmails(principal) {
  const emails = new Set();

  if (principal?.userDetails) {
    emails.add(normalizeEmail(principal.userDetails));
  }

  if (principal?.userId) {
    emails.add(normalizeEmail(principal.userId));
  }

  const claims = Array.isArray(principal?.claims) ? principal.claims : [];
  for (const claim of claims) {
    if (!claim || !claim.typ) continue;

    const typ = String(claim.typ).toLowerCase();
    if (
      typ.includes("email") ||
      typ.endsWith("/upn") ||
      typ.endsWith("/name")
    ) {
      emails.add(normalizeEmail(claim.val));
    }
  }

  return Array.from(emails);
}

function resolveAccessFromRequest(req) {
  const headers = req.headers || {};
  const principalHeader =
    headers["x-ms-client-principal"] ||
    headers["X-MS-CLIENT-PRINCIPAL"];

  const principal = decodeClientPrincipal(principalHeader);
  const roles = Array.isArray(principal?.userRoles) ? principal.userRoles : [];
  const emails = getUserEmails(principal);
  const adminEmails = getConfiguredAdminEmails();

  const isAdmin =
    roles.some((role) => String(role || "").toLowerCase() === "admin") ||
    emails.some((email) => adminEmails.includes(email));

  return {
    principal,
    roles,
    emails,
    email: emails[0] || "",
    isAdmin
  };
}

module.exports = {
  resolveAccessFromRequest
};
