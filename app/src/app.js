import {
  safeApiGet,
  apiPost,
  getDefaultWeekEnding,
  formatDate,
  handleFatalError,
  collectRegionFormValues,
  collectSharedFormValues,
  renderKpiCards
} from "./helpers.js";

/* =========================
   STATE
========================= */

const state = {
  authenticated: false,
  userDetails: "",
  role: "guest",
  entity: "LAOSS",
  weekEnding: getDefaultWeekEnding(),
  currentRoute: "dashboard",
  currentRegion: "LAOSS",
  currentSharedPage: "PT"
};

/* =========================
   HELPERS
========================= */

const $ = (id) => document.getElementById(id);

function setText(id, value) {
  const el = $(id);
  if (el) el.textContent = value;
}

function show(el) {
  if (el) el.style.display = "";
}

function hide(el) {
  if (el) el.style.display = "none";
}

function currentWeekEnding() {
  return $("weekEndingSelect")?.value || state.weekEnding || getDefaultWeekEnding();
}

function isAdmin() {
  return state.role === "admin";
}

function setLoadingHeader() {
  setText("signedInUserText", "Loading...");
  setText("assignedEntityText", "Loading...");
  setText("roleText", "Loading...");
}

function fillFormValues(values) {
  const fields = document.querySelectorAll("[data-key]");
  fields.forEach((el) => {
    const key = el.dataset.key;
    if (!key) return;
    el.value = values[key] ?? "";
  });
}

function getSignInEl() {
  return (
    $("signInButton") ||
    Array.from(document.querySelectorAll("a,button")).find(
      (el) => (el.textContent || "").trim().toLowerCase() === "sign in"
    )
  );
}

function getSignOutEl() {
  return (
    $("signOutButton") ||
    Array.from(document.querySelectorAll("a,button")).find(
      (el) => (el.textContent || "").trim().toLowerCase() === "sign out"
    )
  );
}

function getNavContainer() {
  return (
    $("dashboardNav") ||
    document.querySelector(".dashboard-nav") ||
    document.querySelector(".sidebar-nav") ||
    document.querySelector(".nav-list") ||
    document.querySelector(".nav-links") ||
    document.querySelector(".dashboard-sidebar")
  );
}

/* =========================
   AUTH
========================= */

function normalizeApiMe(result) {
  if (!result || !result.user) return null;

  const user = result.user;
  const roles = Array.isArray(user.roles) ? user.roles : [];

  return {
    authenticated: !!user.authenticated,
    userDetails: user.userDetails || "",
    roles,
    entity: user.entity || ""
  };
}

function normalizeSwaAuth(result) {
  if (!result) return null;

  // Azure Static Web Apps shape:
  // { clientPrincipal: { userDetails, userRoles, ... } }
  if (result.clientPrincipal) {
    const cp = result.clientPrincipal;
    return {
      authenticated: !!cp.userDetails,
      userDetails: cp.userDetails || "",
      roles: Array.isArray(cp.userRoles) ? cp.userRoles : [],
      entity: ""
    };
  }

  // Defensive fallback in case some other shape appears
  if (Array.isArray(result) && result.length > 0) {
    const first = result[0];
    if (first && first.clientPrincipal) {
      const cp = first.clientPrincipal;
      return {
        authenticated: !!cp.userDetails,
        userDetails: cp.userDetails || "",
        roles: Array.isArray(cp.userRoles) ? cp.userRoles : [],
        entity: ""
      };
    }
  }

  return null;
}

function deriveRole(roles) {
  if (!Array.isArray(roles)) return "guest";
  if (roles.includes("admin")) return "admin";
  if (roles.includes("authenticated")) return "user";
  return "guest";
}

async function resolveAuth() {
  setLoadingHeader();

  let auth = normalizeApiMe(await safeApiGet("/api/me", null));

  if (!auth || !auth.authenticated) {
    auth = normalizeSwaAuth(await safeApiGet("/.auth/me", null));
  }

  if (!auth || !auth.authenticated) {
    state.authenticated = false;
    state.userDetails = "";
    state.role = "guest";
    state.entity = "None";
    state.currentRegion = "LAOSS";
    syncAuthUi();
    return;
  }

  state.authenticated = true;
  state.userDetails = auth.userDetails || "Unknown User";
  state.role = deriveRole(auth.roles);
  state.entity = state.role === "admin" ? "All Regions" : (auth.entity || "LAOSS");
  state.currentRegion = state.role === "admin" ? "LAOSS" : (auth.entity || "LAOSS");

  syncAuthUi();
}

function syncAuthUi() {
  const signInEl = getSignInEl();
  const signOutEl = getSignOutEl();

  if (state.authenticated) {
    setText("signedInUserText", state.userDetails);
    setText("assignedEntityText", state.entity);
    setText("roleText", state.role);

    hide(signInEl);
    show(signOutEl);

    if (signOutEl && !signOutEl.getAttribute("href")) {
      signOutEl.onclick = () => {
        window.location.href = "/.auth/logout";
      };
    }
  } else {
    setText("signedInUserText", "Not signed in");
    setText("assignedEntityText", "None");
    setText("roleText", "guest");

    show(signInEl);
    hide(signOutEl);

    if (signInEl && !signInEl.getAttribute("href")) {
      signInEl.onclick = () => {
        window.location.href = "/.auth/login/aad";
      };
    }
  }

  ensureAdminImportLink();
}

/* =========================
   WEEK SELECTOR
========================= */

function initWeekSelector() {
  const select = $("weekEndingSelect");
  if (!select) return;

  const today = new Date();
  const weeks = [];

  for (let i = 0; i < 20; i += 1) {
    const d = new Date(today);
    d.setDate(today.getDate() - today.getDay() - i * 7);
    const iso = d.toISOString().split("T")[0];
    weeks.push(iso);
  }

  select.innerHTML = weeks
    .map((w) => `<option value="${w}">${formatDate(w)}</option>`)
    .join("");

  select.value = state.weekEnding;

  select.addEventListener("change", async () => {
    state.weekEnding = select.value;
    setText("sidebarWeekEndingText", formatDate(select.value));
    await loadCurrentView();
  });

  setText("sidebarWeekEndingText", formatDate(select.value));
}

/* =========================
   ADMIN IMPORT NAV
========================= */

function ensureAdminImportLink() {
  const existing = $("adminImportNavItem");

  if (!isAdmin()) {
    if (existing) existing.remove();
    return;
  }

  const nav = getNavContainer();
  if (!nav) return;

  if (existing) return;

  const wrapper = document.createElement("div");
  wrapper.id = "adminImportNavItem";
  wrapper.innerHTML = `
    <a class="nav-link" href="/admin-import.html">Admin Import</a>
  `;
  nav.appendChild(wrapper);
}

/* =========================
   ROUTING
========================= */

function activateNav(link) {
  document.querySelectorAll(".nav-link").forEach((el) => {
    el.classList.remove("active");
  });
  if (link) link.classList.add("active");
}

function inferRoute(link) {
  const explicit = link?.dataset?.route;
  if (explicit && explicit !== "admin-import") return explicit;

  const text = (link?.textContent || "").trim().toLowerCase();

  if (text.includes("executive")) return "dashboard";
  if (["laoss", "nes", "spineone", "mro"].includes(text)) return "region";
  if (["pt", "cxns", "capacity", "productivity builder"].includes(text)) return "shared";

  return "dashboard";
}

function inferRegion(link) {
  const explicit = link?.dataset?.entity;
  if (explicit) return explicit;

  const text = (link?.textContent || "").trim();
  if (["LAOSS", "NES", "SpineOne", "MRO"].includes(text)) return text;

  return null;
}

function inferSharedPage(link) {
  const explicit = link?.dataset?.page;
  if (explicit) return explicit;

  const text = (link?.textContent || "").trim();
  if (["PT", "CXNS", "Capacity", "Productivity Builder"].includes(text)) return text;

  return null;
}

function initNavigation() {
  document.querySelectorAll(".nav-link").forEach((link) => {
    link.addEventListener("click", async (e) => {
      const href = link.getAttribute("href") || "";
      if (href.includes("admin-import.html")) return;

      e.preventDefault();

      const route = inferRoute(link);
      const region = inferRegion(link);
      const page = inferSharedPage(link);

      state.currentRoute = route;

      if (route === "region" && region) {
        state.currentRegion = region;
      }

      if (route === "shared" && page) {
        state.currentSharedPage = page;
      }

      activateNav(link);
      await loadCurrentView();
    });
  });
}

/* =========================
   DATA LOADERS
========================= */

async function loadDashboard() {
  const week = currentWeekEnding();

  const result = await safeApiGet(
    `/api/dashboard?weekEnding=${encodeURIComponent(week)}`,
    { kpis: [] }
  );

  renderKpiCards(Array.isArray(result.kpis) ? result.kpis : []);
}

async function loadRegionPage(entity) {
  const week = currentWeekEnding();

  const data = await safeApiGet(
    `/api/weekly?entity=${encodeURIComponent(entity)}&weekEnding=${encodeURIComponent(week)}`,
    { values: {} }
  );

  fillFormValues(data.values || {});
}

async function loadSharedPage(page) {
  const week = currentWeekEnding();

  if (page === "Capacity" || page === "Productivity Builder") {
    fillFormValues({});
    return;
  }

  const data = await safeApiGet(
    `/api/shared-data?page=${encodeURIComponent(page)}&weekEnding=${encodeURIComponent(week)}`,
    { values: {} }
  );

  fillFormValues(data.values || {});
}

async function loadCurrentView() {
  try {
    if (state.currentRoute === "region") {
      await loadRegionPage(state.currentRegion);
      return;
    }

    if (state.currentRoute === "shared") {
      await loadSharedPage(state.currentSharedPage);
      return;
    }

    await loadDashboard();
  } catch (err) {
    handleFatalError(err);
  }
}

/* =========================
   SAVE
========================= */

async function saveRegion() {
  const payload = collectRegionFormValues();

  if (!payload.entity || payload.entity === "Loading..." || payload.entity === "None") {
    payload.entity = state.currentRegion;
  }

  if (!payload.weekEnding) {
    payload.weekEnding = currentWeekEnding();
  }

  const result = await apiPost("/api/weekly-save", payload);

  if (!result || result.error) {
    alert("Save failed.");
    return;
  }

  alert("Saved successfully.");
}

async function saveShared() {
  const payload = collectSharedFormValues();

  if (!payload.page || payload.page === "Loading...") {
    payload.page = state.currentSharedPage;
  }

  if (!payload.weekEnding) {
    payload.weekEnding = currentWeekEnding();
  }

  const result = await apiPost("/api/shared-save", payload);

  if (!result || result.error) {
    alert("Save failed.");
    return;
  }

  alert("Saved successfully.");
}

/* =========================
   BUTTONS
========================= */

function initButtons() {
  const saveBtn = $("saveButton");
  const submitBtn = $("submitWeekButton");
  const signInEl = getSignInEl();
  const signOutEl = getSignOutEl();

  if (saveBtn) {
    saveBtn.addEventListener("click", async () => {
      if (state.currentRoute === "region") {
        await saveRegion();
      } else if (state.currentRoute === "shared") {
        await saveShared();
      } else {
        alert("Nothing to save on this page.");
      }
    });
  }

  if (submitBtn) {
    submitBtn.addEventListener("click", async () => {
      if (state.currentRoute === "region") {
        await saveRegion();
      } else if (state.currentRoute === "shared") {
        await saveShared();
      } else {
        alert("Nothing to submit on this page.");
        return;
      }

      alert("Week submitted.");
    });
  }

  if (signInEl && !signInEl.getAttribute("href")) {
    signInEl.addEventListener("click", () => {
      window.location.href = "/.auth/login/aad";
    });
  }

  if (signOutEl && !signOutEl.getAttribute("href")) {
    signOutEl.addEventListener("click", () => {
      window.location.href = "/.auth/logout";
    });
  }
}

/* =========================
   HERO BUTTONS
========================= */

function initHeroButtons() {
  const goToRegionBtn =
    $("goToMyRegionButton") ||
    Array.from(document.querySelectorAll("a,button")).find((el) =>
      (el.textContent || "").trim().toLowerCase().includes("go to my region")
    );

  const executiveBtn =
    $("executiveSummaryButton") ||
    Array.from(document.querySelectorAll("a,button")).find((el) =>
      (el.textContent || "").trim().toLowerCase().includes("executive summary")
    );

  if (goToRegionBtn) {
    goToRegionBtn.addEventListener("click", async () => {
      state.currentRoute = "region";
      state.currentRegion = isAdmin() ? "LAOSS" : (state.entity === "All Regions" ? "LAOSS" : state.entity);

      const regionLink = Array.from(document.querySelectorAll(".nav-link")).find((el) => {
        const text = (el.textContent || "").trim();
        return text === state.currentRegion;
      });

      activateNav(regionLink);
      await loadCurrentView();
    });
  }

  if (executiveBtn) {
    executiveBtn.addEventListener("click", async () => {
      state.currentRoute = "dashboard";

      const executiveLink = Array.from(document.querySelectorAll(".nav-link")).find((el) =>
        ((el.textContent || "").trim().toLowerCase().includes("executive"))
      );

      activateNav(executiveLink);
      await loadCurrentView();
    });
  }
}

/* =========================
   INIT
========================= */

async function init() {
  try {
    setLoadingHeader();
    initWeekSelector();
    initButtons();
    initHeroButtons();
    await resolveAuth();
    initNavigation();
    await loadCurrentView();
  } catch (err) {
    handleFatalError(err);
  }
}

document.addEventListener("DOMContentLoaded", init);
