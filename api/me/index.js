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

function normalizeEntity(value) {
  const raw = String(value || "").trim();

  if (!raw) return "";

  const normalized = raw.toUpperCase();

  if (normalized === "LAOSS") return "LAOSS";
  if (normalized === "NES") return "NES";
  if (normalized === "SPINEONE") return "SpineOne";
  if (normalized === "MRO") return "MRO";

  return raw;
}

function findEntityFromRoles(roles) {
  const normalizedRoles = roles.map((r) => String(r || "").trim());

  if (normalizedRoles.includes("LAOSS")) return "LAOSS";
  if (normalizedRoles.includes("NES")) return "NES";
  if (normalizedRoles.includes("SpineOne")) return "SpineOne";
  if (normalizedRoles.includes("MRO")) return "MRO";

  const upperRoles = normalizedRoles.map((r) => r.toUpperCase());

  if (upperRoles.includes("LAOSS")) return "LAOSS";
  if (upperRoles.includes("NES")) return "NES";
  if (upperRoles.includes("SPINEONE")) return "SpineOne";
  if (upperRoles.includes("MRO")) return "MRO";

  return "";
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
    const isAdmin = roles.some((role) => String(role || "").toLowerCase() === "admin");

    let entity = "";

    if (isAdmin) {
      entity = "Admin";
    } else {
      entity = findEntityFromRoles(roles);
      entity = normalizeEntity(entity);
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
        isAdmin
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
