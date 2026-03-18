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
  entity: "None",
  isAdmin: false,
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

function setTextBySelectors(selectors, value) {
  selectors.forEach((selector) => {
    const el = document.querySelector(selector);
    if (el) el.textContent = value;
  });
}

function setSignedInUserText(value) {
  setText("signedInUserText", value);
  setText("signedInAsText", value);
  setTextBySelectors(
    [
      "#signedInUserText",
      "#signedInAsText",
      "[data-auth-user]",
      ".signed-in-user",
      ".auth-user"
    ],
    value
  );
}

function setAssignedEntityText(value) {
  setText("assignedEntityText", value);
  setText("assignedEntityValue", value);
  setTextBySelectors(
    [
      "#assignedEntityText",
      "#assignedEntityValue",
      "[data-assigned-entity]",
      ".assigned-entity"
    ],
    value
  );
}

function setRoleText(value) {
  setText("roleText", value);
  setText("roleValue", value);
  setTextBySelectors(
    [
      "#roleText",
      "#roleValue",
      "[data-role-text]",
      ".role-text"
    ],
    value
  );
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
  return state.isAdmin === true || state.role === "admin";
}

function setLoadingHeader() {
  setSignedInUserText("Loading...");
  setAssignedEntityText("Loading...");
  setRoleText("Loading...");
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
    document.querySelector(".dashboard-sidebar") ||
    document.querySelector(".sidebar")
  );
}

/* =========================
   AUTH
========================= */

async function resolveAuth() {
  setLoadingHeader();

  const me = await safeApiGet("/api/me", null);

  if (!me || !me.authenticated) {
    state.authenticated = false;
    state.userDetails = "";
    state.role = "guest";
    state.entity = "None";
    state.isAdmin = false;
    state.currentRegion = "LAOSS";
    syncAuthUi();
    return;
  }

  state.authenticated = true;
  state.userDetails = me.userDetails || "Unknown User";
  state.isAdmin = !!me.isAdmin;
  state.role = state.isAdmin ? "admin" : "user";
  state.entity = state.isAdmin ? "Admin" : (me.entity || "LAOSS");
  state.currentRegion = state.isAdmin ? "LAOSS" : state.entity;

  syncAuthUi();
}

function syncAuthUi() {
  const signInEl = getSignInEl();
  const signOutEl = getSignOutEl();

  if (state.authenticated) {
    setSignedInUserText(state.userDetails);
    setAssignedEntityText(state.entity);
    setRoleText(state.role);

    hide(signInEl);
    show(signOutEl);

    if (signOutEl && !signOutEl.getAttribute("href")) {
      signOutEl.onclick = () => {
        window.location.href = "/.auth/logout";
      };
    }
  } else {
    setSignedInUserText("Not signed in");
    setAssignedEntityText("None");
    setRoleText("guest");

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
  wrapper.innerHTML = `<a class="nav-link" href="/admin-import.html">Admin Import</a>`;
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

  if (!payload.entity || payload.entity === "Loading..." || payload.entity === "None" || payload.entity === "Admin") {
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
      state.currentRegion = isAdmin() ? "LAOSS" : state.entity;

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
