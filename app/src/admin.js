const ADMIN_EMAILS = ["nperez@unitymsk.com", "tessa.kelley@spineone.com"];

const state = {
  authenticated: false,
  userDetails: "",
  role: "guest",
  entity: "None",
  isAdmin: false
};

const $ = (id) => document.getElementById(id);

function normalizeEmail(value) {
  return String(value || "").trim().toLowerCase();
}

function emailIsAdmin(value) {
  return ADMIN_EMAILS.includes(normalizeEmail(value));
}

async function safeGetJson(url, fallback = null) {
  try {
    const res = await fetch(url, {
      credentials: "include",
      headers: { Accept: "application/json" }
    });

    if (!res.ok) return fallback;
    return await res.json();
  } catch {
    return fallback;
  }
}

function unique(values) {
  return Array.from(new Set((Array.isArray(values) ? values : []).filter(Boolean)));
}

function normalizeAuthMe(result) {
  const principal = result?.clientPrincipal;
  if (!principal || !principal.userId) return null;

  const userDetails = principal.userDetails || principal.userId || "";
  const roles = unique(principal.userRoles || []);
  const roleAdmin = roles.some((r) => String(r || "").toLowerCase() === "admin");
  const forcedAdmin = emailIsAdmin(userDetails);

  return {
    authenticated: true,
    userDetails,
    roles,
    isAdmin: roleAdmin || forcedAdmin,
    entity: forcedAdmin ? "Admin" : ""
  };
}

function normalizeApiMe(result) {
  if (!result || !result.authenticated) return null;

  const userDetails = result.userDetails || "";
  const roles = unique(result.roles);
  const apiSaysAdmin =
    !!result.isAdmin ||
    roles.some((r) => String(r || "").toLowerCase() === "admin");
  const forcedAdmin = emailIsAdmin(userDetails);

  return {
    authenticated: true,
    userDetails,
    roles,
    isAdmin: apiSaysAdmin || forcedAdmin,
    entity: forcedAdmin ? "Admin" : (result.entity || "")
  };
}

function setText(id, value) {
  const el = $(id);
  if (el) el.textContent = value;
}

function setStatusBadge(text) {
  const el = $("statusBadge");
  if (el) el.textContent = text;
}

function updateUi() {
  setText("signedInAsText", state.authenticated ? state.userDetails : "Not signed in");
  setText("roleText", state.role);
  setText("accessText", state.isAdmin ? "Allowed" : "Denied");

  const importBtn = $("importButton");
  const fileInput = $("workbookFile");

  if (importBtn) importBtn.disabled = !state.isAdmin;
  if (fileInput) fileInput.disabled = !state.isAdmin;

  if (!state.authenticated) {
    setStatusBadge("Not signed in");
    setText("lastResultText", "Authentication required");
    setText("importLog", "You are not signed in.");
    return;
  }

  if (!state.isAdmin) {
    setStatusBadge("Authenticated");
    setText("lastResultText", "Signed in but not admin");
    setText("importLog", `Signed in as ${state.userDetails}, but this page is restricted to admins.`);
    return;
  }

  setStatusBadge("Ready");
  setText("lastResultText", "Admin access confirmed");
  setText("importLog", `Signed in as ${state.userDetails}. Admin access confirmed. You can import the workbook.`);
}

async function resolveAuth() {
  let me = normalizeAuthMe(await safeGetJson("/.auth/me", null));

  if (!me) {
    me = normalizeApiMe(await safeGetJson("/api/me", null));
  }

  if (!me || !me.authenticated) {
    state.authenticated = false;
    state.userDetails = "";
    state.role = "guest";
    state.entity = "None";
    state.isAdmin = false;
    updateUi();
    return;
  }

  state.authenticated = true;
  state.userDetails = me.userDetails || "Unknown User";
  state.isAdmin = !!me.isAdmin;
  state.role = state.isAdmin ? "admin" : "user";
  state.entity = state.isAdmin ? "Admin" : (me.entity || "LAOSS");

  updateUi();
}

async function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = () => {
      const result = String(reader.result || "");
      const base64 = result.includes(",") ? result.split(",")[1] : result;
      resolve(base64);
    };

    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

async function readResponseBody(res) {
  const text = await res.text();

  try {
    return {
      parsed: JSON.parse(text),
      raw: text
    };
  } catch {
    return {
      parsed: null,
      raw: text
    };
  }
}

async function importWorkbook() {
  if (!state.isAdmin) {
    setText("importLog", "Access denied. You must be an admin to import the workbook.");
    return;
  }

  const input = $("workbookFile");
  const file = input?.files?.[0];

  if (!file) {
    setText("importLog", "Choose the workbook file first.");
    return;
  }

  setStatusBadge("Importing");
  setText("lastResultText", "Uploading workbook...");
  setText("importLog", `Reading ${file.name} (${Math.round(file.size / 1024)} KB)...`);

  try {
    const fileBase64 = await fileToBase64(file);

    setText(
      "importLog",
      `Reading ${file.name} (${Math.round(file.size / 1024)} KB)...\nUploading workbook to /api/import-excel ...`
    );

    const res = await fetch("/api/import-excel", {
      method: "POST",
      credentials: "include",
      headers: {
        "Content-Type": "application/json",
        "Accept": "application/json, text/plain, */*"
      },
      body: JSON.stringify({ fileBase64, fileName: file.name })
    });

    const { parsed, raw } = await readResponseBody(res);

    if (!res.ok) {
      setStatusBadge("Import failed");
      setText("lastResultText", `HTTP ${res.status}`);

      setText(
        "importLog",
        [
          `Import failed.`,
          `Status: ${res.status} ${res.statusText}`,
          "",
          parsed ? JSON.stringify(parsed, null, 2) : (raw || "No response body returned.")
        ].join("\n")
      );
      return;
    }

    if (parsed && parsed.ok === false) {
      setStatusBadge("Import failed");
      setText("lastResultText", "Import failed");
      setText("importLog", JSON.stringify(parsed, null, 2));
      return;
    }

    setStatusBadge("Ready");
    setText("lastResultText", "Import completed");
    setText(
      "importLog",
      parsed
        ? `Workbook import completed successfully.\n\n${JSON.stringify(parsed, null, 2)}`
        : `Workbook import completed, but the response was not JSON.\n\n${raw || "(empty response)"}`
    );
  } catch (error) {
    setStatusBadge("Import failed");
    setText("lastResultText", "Import failed");
    setText("importLog", `Import failed.\n\n${error?.stack || error?.message || String(error)}`);
  }
}

function initButtons() {
  const importBtn = $("importButton");
  const recheckBtn = $("recheckAccessButton");

  if (importBtn) {
    importBtn.addEventListener("click", async () => {
      await importWorkbook();
    });
  }

  if (recheckBtn) {
    recheckBtn.addEventListener("click", async () => {
      setText("lastResultText", "Checking...");
      setText("importLog", "Rechecking access...");
      await resolveAuth();
    });
  }
}

async function init() {
  initButtons();
  await resolveAuth();
}

document.addEventListener("DOMContentLoaded", init);
