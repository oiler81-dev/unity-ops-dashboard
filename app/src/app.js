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
  authResolved: false,
  authenticated: false,
  userDetails: "",
  role: "user",
  entity: "LAOSS",
  weekEnding: getDefaultWeekEnding(),
  currentRoute: "dashboard",
  currentRegion: "LAOSS",
  currentSharedPage: "PT",
  me: null
};

/* =========================
   BASIC HELPERS
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

function bySelectors(selectors) {
  for (const selector of selectors) {
    const el = document.querySelector(selector);
    if (el) return el;
  }
  return null;
}

function currentWeekEnding() {
  const el = $("weekEndingSelect");
  return el?.value || state.weekEnding || getDefaultWeekEnding();
}

function isAdmin() {
  return state.role === "admin";
}

function setLoadingHeaderState() {
  setText("signedInUserText", "Loading...");
  setText("assignedEntityText", "Loading...");
  setText("roleText", "Loading...");
}

function fillFormValues(values) {
  const fields = document.querySelectorAll("[data-key]");
  fields.forEach((el) => {
    const key = el.dataset.key;
    if (!key) return;
    const val = values[key];
    el.value = val == null ? "" : val;
  });
}

/* =========================
   AUTH
========================= */

async function fetchApiMe() {
  const result = await safeApiGet("/api/me", null);
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

async function fetchStaticWebAppAuth() {
  const result = await safeApiGet("/.auth/me", null);

  if (!Array.isArray(result) || !result.length) return null;

  const principal = result[0] || {};
  const claims = Array.isArray(principal.userClaims) ? principal.userClaims : [];

  const nameClaim =
    claims.find((c) => c.typ === "preferred_username")?.val ||
    claims.find((c) => c.typ === "email")?.val ||
    claims.find((c) => c.typ === "name")?.val ||
    principal.userDetails ||
    "";

  const roleClaims = claims
    .filter((c) => c.typ === "roles")
    .map((c) => c.val)
    .filter(Boolean);

  const roles = Array.from(
    new Set([
      ...(Array.isArray(principal.userRoles) ? principal.userRoles : []),
      ...roleClaims
    ])
  );

  return {
    authenticated: true,
    userDetails: nameClaim,
    roles,
    entity: ""
  };
}

function normalizeRoleFromRoles(roles) {
  if (!Array.isArray(roles)) return "user";
  return roles.includes("admin") ? "admin" : "user";
}

async function resolveAuth() {
  setLoadingHeaderState();

  let auth = await fetchApiMe();

  if (!auth || !auth.authenticated) {
    const fallback = await fetchStaticWebAppAuth();
    if (fallback) auth = fallback;
  }

  if (!auth || !auth.authenticated) {
    state.authResolved = true;
    state.authenticated = false;
    state.userDetails = "";
    state.role = "user";
    state.entity = "LAOSS";
    state.currentRegion = "LAOSS";
    state.me = null;
    syncAuthUi();
    return;
  }

  state.authResolved = true;
  state.authenticated = true;
  state.userDetails = auth.userDetails || "Unknown User";
  state.role = normalizeRoleFromRoles(auth.roles);
  state.entity = state.role === "admin" ? "ALL" : (auth.entity || "LAOSS");
  state.currentRegion = state.role === "admin" ? "LAOSS" : state.entity;
  state.me = auth;

  syncAuthUi();
}

function syncAuthUi() {
  const signInButton =
    $("signInButton") ||
    bySelectors([
      '[data-action="sign-in"]',
      'a[href="/.auth/login/aad"]',
      'a[href*="/.auth/login"]'
    ]);

  const signOutButton =
    $("signOutButton") ||
    bySelectors([
      '[data-action="sign-out"]',
      'a[href="/.auth/logout"]',
      'a[href*="/.auth/logout"]'
    ]);

  if (state.authenticated) {
    setText("signedInUserText", state.userDetails);
    setText("assignedEntityText", isAdmin() ? "All Regions" : state.entity);
    setText("roleText", isAdmin() ? "admin" : "user");

    hide(signInButton);
    show(signOutButton);
  } else {
    setText("signedInUserText", "Not signed in");
    setText("assignedEntityText", "None");
    setText("roleText", "guest");

    show(signInButton);
    hide(signOutButton);
  }

  ensureAdminImportLink();
}

/* =========================
   WEEK SELECTOR
========================= */

function initWeekSelector() {
  const select = $("weekEndingSelect");
  if (!select) return;

  const base = new Date();
  const thisSunday = new Date(base);
  thisSunday.setDate(base.getDate() - base.getDay());

  const options = [];
  for (let i = 0; i < 20; i += 1) {
    const d = new Date(thisSunday);
    d.setDate(thisSunday.getDate() - i * 7);
    const iso = d.toISOString().split("T")[0];
    options.push({
      value: iso,
      label: formatDate(iso)
    });
  }

  select.innerHTML = options
    .map((opt) => `<option value="${opt.value}">${opt.label}</option>`)
    .join("");

  select.value = state.weekEnding;

  select.addEventListener("change", async () => {
    state.weekEnding = select.value;
    updateWeekBadges();
    await loadCurrentView();
  });

  updateWeekBadges();
}

function updateWeekBadges() {
  const week = currentWeekEnding();
  setText("currentWeekEndingText", formatDate(week));
  setText("sidebarWeekEndingText", formatDate(week));
}

/* =========================
   ADMIN IMPORT LINK
========================= */

function ensureAdminImportLink() {
  let existing = $("adminImportNavItem");

  if (!isAdmin()) {
    if (existing) existing.remove();
    return;
  }

  const navHost =
    $("dashboardNav") ||
    bySelectors([
      ".dashboard-nav",
      ".sidebar-nav",
      ".nav-list",
      ".nav-links",
      "aside nav",
      ".sidebar"
    ]);

  if (!navHost) return;

  if (!existing) {
    const wrapper = document.createElement("div");
    wrapper.id = "adminImportNavItem";
    wrapper.innerHTML = `
      <a href="/admin-import.html" class="nav-link nav-link-admin-import" data-route="admin-import">
        Admin Import
      </a>
    `;
    navHost.appendChild(wrapper);
  }
}

/* =========================
   NAVIGATION
========================= */

function setActiveNav(link) {
  document.querySelectorAll(".nav-link").forEach((el) => {
    el.classList.remove("active");
  });

  if (link) link.classList.add("active");
}

function inferRouteFromLink(link) {
  const explicit = link?.dataset?.route;
  if (explicit && explicit !== "admin-import") return explicit;

  const text = (link?.textContent || "").trim().toLowerCase();

  if (text.includes("executive")) return "dashboard";
  if (text === "pt") return "shared";
  if (text === "cxns") return "shared";
  if (text.includes("capacity")) return "shared";
  if (text.includes("productivity")) return "shared";
  if (["laoss", "nes", "spineone", "mro"].includes(text)) return "region";

  return "dashboard";
}

function inferRegionFromLink(link) {
  const explicit = link?.dataset?.entity;
  if (explicit) return explicit;

  const text = (link?.textContent || "").trim();
  if (["LAOSS", "NES", "SpineOne", "MRO"].includes(text)) return text;

  return null;
}

function inferSharedPageFromLink(link) {
  const explicit = link?.dataset?.page;
  if (explicit) return explicit;

  const text = (link?.textContent || "").trim();
  if (["PT", "CXNS", "Capacity", "Productivity Builder"].includes(text)) return text;

  return null;
}

function initNavigation() {
  document.querySelectorAll(".nav-link").forEach((link) => {
    link.addEventListener("click", async (event) => {
      const href = link.getAttribute("href") || "";

      if (href.includes("admin-import.html")) {
        return;
      }

      event.preventDefault();

      const route = inferRouteFromLink(link);
      const region = inferRegionFromLink(link);
      const sharedPage = inferSharedPageFromLink(link);

      state.currentRoute = route;

      if (route === "region" && region) {
        state.currentRegion = region;
      }

      if (route === "shared" && sharedPage) {
        state.currentSharedPage = sharedPage;
      }

      setActiveNav(link);
      await loadCurrentView();
    });
  });
}

/* =========================
   API LOADERS
========================= */

async function loadDashboard() {
  const weekEnding = currentWeekEnding();

  const result = await safeApiGet(
    `/api/dashboard?weekEnding=${encodeURIComponent(weekEnding)}`,
    { ok: true, kpis: [] }
  );

  renderKpiCards(Array.isArray(result.kpis) ? result.kpis : []);
}

async function loadRegion(entity) {
  const weekEnding = currentWeekEnding();

  const result = await safeApiGet(
    `/api/weekly?entity=${encodeURIComponent(entity)}&weekEnding=${encodeURIComponent(weekEnding)}`,
    { values: {} }
  );

  fillFormValues(result.values || {});
}

async function loadShared(page) {
  const weekEnding = currentWeekEnding();

  if (page === "Capacity" || page === "Productivity Builder") {
    fillFormValues({});
    return;
  }

  const result = await safeApiGet(
    `/api/shared-data?page=${encodeURIComponent(page)}&weekEnding=${encodeURIComponent(weekEnding)}`,
    { values: {} }
  );

  fillFormValues(result.values || {});
}

async function loadCurrentView() {
  try {
    if (state.currentRoute === "region") {
      await loadRegion(state.currentRegion);
      return;
    }

    if (state.currentRoute === "shared") {
      await loadShared(state.currentSharedPage);
      return;
    }

    await loadDashboard();
  } catch (err) {
    handleFatalError(err);
  }
}

/* =========================
   SAVE / SUBMIT
========================= */

async function saveCurrentView() {
  if (state.currentRoute === "region") {
    const payload = collectRegionFormValues();

    if (!payload.entity || payload.entity === "Loading...") {
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
    return;
  }

  if (state.currentRoute === "shared") {
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
    return;
  }

  alert("Nothing to save on this page.");
}

/* =========================
   BUTTONS
========================= */

function initButtons() {
  const saveBtn =
    $("saveButton") ||
    bySelectors([
      '[data-action="save"]',
      'button[title="Save"]'
    ]);

  const submitBtn =
    $("submitWeekButton") ||
    bySelectors([
      '[data-action="submit-week"]',
      'button[title="Submit Week"]'
    ]);

  const signInBtn =
    $("signInButton") ||
    bySelectors([
      '[data-action="sign-in"]'
    ]);

  const signOutBtn =
    $("signOutButton") ||
    bySelectors([
      '[data-action="sign-out"]'
    ]);

  if (saveBtn) {
    saveBtn.addEventListener("click", async () => {
      await saveCurrentView();
    });
  }

  if (submitBtn) {
    submitBtn.addEventListener("click", async () => {
      await saveCurrentView();
      alert("Week submitted.");
    });
  }

  if (signInBtn && !signInBtn.getAttribute("href")) {
    signInBtn.addEventListener("click", () => {
      window.location.href = "/.auth/login/aad";
    });
  }

  if (signOutBtn && !signOutBtn.getAttribute("href")) {
    signOutBtn.addEventListener("click", () => {
      window.location.href = "/.auth/logout";
    });
  }
}

/* =========================
   HERO BUTTONS
========================= */

function initHeroButtons() {
  const regionBtn =
    $("goToMyRegionButton") ||
    bySelectors([
      '[data-action="go-to-region"]'
    ]);

  const executiveBtn =
    $("executiveSummaryButton") ||
    bySelectors([
      '[data-action="go-to-dashboard"]'
    ]);

  if (regionBtn) {
    regionBtn.addEventListener("click", async () => {
      state.currentRoute = "region";
      state.currentRegion = isAdmin() ? "LAOSS" : state.entity;

      const regionLink = Array.from(document.querySelectorAll(".nav-link")).find((el) => {
        const text = (el.textContent || "").trim();
        return text === state.currentRegion;
      });

      setActiveNav(regionLink);
      await loadCurrentView();
    });
  }

  if (executiveBtn) {
    executiveBtn.addEventListener("click", async () => {
      state.currentRoute = "dashboard";

      const dashboardLink = Array.from(document.querySelectorAll(".nav-link")).find((el) => {
        const text = (el.textContent || "").trim().toLowerCase();
        return text.includes("executive");
      });

      setActiveNav(dashboardLink);
      await loadCurrentView();
    });
  }
}

/* =========================
   INIT
========================= */

async function init() {
  try {
    setLoadingHeaderState();
    initWeekSelector();
    initButtons();
    initHeroButtons();
    await resolveAuth();
    initNavigation();
    await loadCurrentView();
  } catch (error) {
    handleFatalError(error);
  }
}

document.addEventListener("DOMContentLoaded", init);
