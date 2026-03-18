module.exports = async function (context, req) {
  try {
    const headers = req.headers || {};

    function getHeader(name) {
      return (
        headers[name] ||
        headers[name.toLowerCase()] ||
        headers[name.toUpperCase()] ||
        ""
      );
    }

    function parseClientPrincipal() {
      const encoded = getHeader("x-ms-client-principal");

      if (!encoded) return null;

      try {
        const json = Buffer.from(encoded, "base64").toString("utf8");
        return JSON.parse(json);
      } catch (err) {
        context.log.warn("Unable to parse x-ms-client-principal", err.message);
        return null;
      }
    }

    function unique(values) {
      return Array.from(new Set((Array.isArray(values) ? values : []).filter(Boolean)));
    }

    function normalizeEmail(value) {
      return String(value || "").trim().toLowerCase();
    }

    function getAdminEmails() {
      const raw =
        process.env.ADMIN_EMAILS ||
        process.env.ADMIN_USERS ||
        "nperez@unitymsk.com";

      return raw
        .split(",")
        .map((v) => normalizeEmail(v))
        .filter(Boolean);
    }

    const principal = parseClientPrincipal();

    const userDetails =
      principal?.userDetails ||
      getHeader("x-ms-client-principal-name") ||
      "";

    const rolesFromPrincipal = unique(principal?.userRoles || []);
    const rolesNormalized = rolesFromPrincipal.map((r) => String(r).toLowerCase());

    const email = normalizeEmail(userDetails);
    const adminEmails = getAdminEmails();

    const isAdmin =
      rolesNormalized.includes("admin") ||
      adminEmails.includes(email);

    const roles = unique([
      ...rolesFromPrincipal,
      ...(isAdmin ? ["admin"] : []),
      ...(userDetails ? ["authenticated"] : [])
    ]);

    const authenticated = !!userDetails;

    context.res = {
      status: 200,
      headers: {
        "Content-Type": "application/json",
        "Cache-Control": "no-store"
      },
      body: {
        authenticated,
        userDetails: userDetails || "",
        userId: principal?.userId || "",
        identityProvider: principal?.identityProvider || "",
        roles,
        entity: isAdmin ? "Admin" : null,
        isAdmin
      }
    };
  } catch (error) {
    context.log.error("api/me failed", error);

    context.res = {
      status: 500,
      headers: {
        "Content-Type": "application/json",
        "Cache-Control": "no-store"
      },
      body: {
        authenticated: false,
        userDetails: "",
        roles: ["anonymous"],
        entity: null,
        isAdmin: false,
        error: "Failed to resolve user"
      }
    };
  }
};
