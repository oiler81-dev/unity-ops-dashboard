function decodeClientPrincipal(headerValue) {
  if (!headerValue) return null;

  try {
    const json = Buffer.from(headerValue, "base64").toString("utf8");
    return JSON.parse(json);
  } catch (error) {
    return null;
  }
}

function inferEntityFromEmail(email) {
  const lower = String(email || "").toLowerCase();

  if (!lower) return "LAOSS";
  if (lower.includes("nes")) return "NES";
  if (lower.includes("spine")) return "SpineOne";
  if (lower.includes("mro")) return "MRO";

  return "LAOSS";
}

function getUserInfo(req) {
  const principal = decodeClientPrincipal(req.headers["x-ms-client-principal"]);

  if (!principal) {
    return {
      authenticated: false,
      userDetails: "",
      userId: "",
      identityProvider: "",
      roles: ["anonymous"],
      entity: "LAOSS"
    };
  }

  const userDetails = principal.userDetails || "";
  const roles = Array.isArray(principal.userRoles) ? principal.userRoles : ["authenticated"];

  return {
    authenticated: true,
    userDetails,
    userId: principal.userId || "",
    identityProvider: principal.identityProvider || "",
    roles,
    entity: inferEntityFromEmail(userDetails)
  };
}

module.exports = {
  getUserInfo
};
