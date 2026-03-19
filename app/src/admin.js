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

function setText(id, value) {
  const el = $(id);
  if (el) el.textContent = value;
}

function setStatusBadge(text) {
  const el = $("statusBadge");
  if (el) el.textContent = text;
}

function setAccessText(text) {
  setText("accessText", text);
}

function setSignedInText(text) {
  setText("signedInAsText", text);
}

function setRoleText(text) {
  setText("roleText", text);
}

function setResultText(text) {
  setText("lastResultText", text);
}

function setLogText(text) {
  const el = $("importLog");
  if (el) el.textContent = text;
}

function updateUi() {
  setSignedInText(state.authenticated ? state.userDetails : "Not signed in");
  setRoleText(state.role);
  setAccessText(state.isAdmin ? "Allowed" : "Denied");
  setStatusBadge(state.isAdmin ? "Ready" : "Blocked");

  const importBtn = $("importButton");
  const recheckBtn = $("recheckAccessButton");
  const fileInput = $("workbookFile");

  if (importBtn) importBtn.disabled = !state.isAdmin;
  if (fileInput) fileInput.disabled = !state.isAdmin;
  if (recheckBtn) recheckBtn.disabled = false;

  if (!state.authenticated) {
    setResultText("Not signed in");
    setLogText("You are not signed in.");
    return;
  }

  if (!state.isAdmin) {
    setResultText("Signed in but not admin");
    setLogText(
      `Signed in as ${state.userDetails}, but this page is restricted to admins.`
    );
    return;
  }

  setResultText("Ready");
  setLogText(
    `Signed in as ${state.userDetails}. Admin access confirmed. You can import the workbook.`
  );
}

async function resolveAuth() {
  let me = normalizeApiMe(await safeGetJson("/api/me", null));

  if (!me) {
    me = normalizeAuthMe(await safeGetJson("/.auth/me", null));
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

async function importWorkbook() {
  if (!state.isAdmin) {
    setLogText("Access denied. You must be an admin to import the workbook.");
    return;
  }

  const input = $("workbookFile");
  const file = input?.files?.[0];

  if (!file) {
    setLogText("Choose the workbook file first.");
    return;
  }

  setStatusBadge("Importing");
  setResultText("Uploading workbook...");
  setLogText(`Reading ${file.name}...`);

  try {
    const fileBase64 = await fileToBase64(file);

    setLogText("Uploading workbook to /api/import-excel ...");

    const res = await fetch("/api/import-excel", {
      method: "POST",
      credentials: "include",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({ fileBase64 })
    });

    const result = await res.json().catch(() => ({}));

    if (!res.ok || result?.ok === false) {
      setStatusBadge("Blocked");
      setResultText("Import failed");
      setLogText(
        `Import failed.\n\n${JSON.stringify(result, null, 2)}`
      );
      return;
    }

    setStatusBadge("Ready");
    setResultText("Import completed");
    setLogText(
      `Workbook import completed successfully.\n\n${JSON.stringify(result, null, 2)}`
    );
  } catch (error) {
    setStatusBadge("Blocked");
    setResultText("Import failed");
    setLogText(`Import failed.\n\n${error?.message || String(error)}`);
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
      setResultText("Checking...");
      setLogText("Rechecking access...");
      await resolveAuth();
    });
  }
}

async function init() {
  initButtons();
  await resolveAuth();
}

document.addEventListener("DOMContentLoaded", init);
