function getUserFromRequest(req) {
  const principal = req.headers["x-ms-client-principal"];
  if (!principal) return null;

  try {
    const decoded = Buffer.from(principal, "base64").toString("utf8");
    return JSON.parse(decoded);
  } catch {
    return null;
  }
}

function normalizeEmail(email) {
  return String(email || "").trim().toLowerCase();
}

function getUserEmail(user) {
  if (!user) return "";
  return normalizeEmail(
    user.userDetails ||
    user?.claims?.find((c) => c.typ === "preferred_username")?.val ||
    user?.claims?.find((c) => c.typ === "email")?.val ||
    ""
  );
}

function getDisplayName(user) {
  if (!user) return "Unknown User";
  return user.userDetails || "Unknown User";
}

module.exports = {
  getUserFromRequest,
  getUserEmail,
  getDisplayName,
  normalizeEmail
};
