function decodeClientPrincipal(encoded) {
  if (!encoded) return null;

  try {
    const json = Buffer.from(encoded, "base64").toString("utf8");
    return JSON.parse(json);
  } catch {
    return null;
  }
}

function getHeader(req, name) {
  if (!req?.headers) return "";
  return (
    req.headers[name] ||
    req.headers[name.toLowerCase()] ||
    req.headers[name.toUpperCase()] ||
    ""
  );
}

function normalizeRoles(principal) {
  const roles = Array.isArray(principal?.userRoles) ? principal.userRoles : [];
  return roles.filter(Boolean);
}

function getUserFromRequest(req) {
  const encodedPrincipal = getHeader(req, "x-ms-client-principal");
  const principal = decodeClientPrincipal(encodedPrincipal);

  if (!principal) {
    return {
      authenticated: false,
      userDetails: null,
      identityProvider: null,
      userId: null,
      roles: ["anonymous"]
    };
  }

  return {
    authenticated: true,
    userDetails:
      principal.userDetails ||
      principal.claims?.find((c) => c.typ === "preferred_username")?.val ||
      principal.claims?.find((c) => c.typ === "email")?.val ||
      null,
    identityProvider: principal.identityProvider || null,
    userId: principal.userId || null,
    roles: normalizeRoles(principal),
    claims: Array.isArray(principal.claims) ? principal.claims : []
  };
}

module.exports = { getUserFromRequest };
