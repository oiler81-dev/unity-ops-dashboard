import {
  safeApiGet,
  apiPost,
  getDefaultWeekEnding,
  formatDate,
  escapeHtml,
  handleFatalError,
  collectRegionFormValues,
  collectSharedFormValues,
  renderKpiCards
} from "./helpers.js";

/* =========================
   GLOBAL STATE
========================= */

const state = {
  user: null,
  entity: null,
  role: null,
  weekEnding: null,
  currentRoute: "dashboard"
};

/* =========================
   DOM HELPERS
========================= */

const $ = (id) => document.getElementById(id);

function setText(id, value) {
  const el = $(id);
  if (el) el.textContent = value;
}

/* =========================
   AUTH
========================= */

async function loadUser() {
  const result = await safeApiGet("/api/me", null);

  if (!result || !result.user) {
    setText("signedInUserText", "Not signed in");
    return;
  }

  const user = result.user;

  state.user = user.userDetails || "Unknown";
  state.role = Array.isArray(user.roles) && user.roles.includes("admin")
    ? "admin"
    : "user";

  setText("signedInUserText", state.user);

  if (state.role === "admin") {
    setText("assignedEntityText", "All Regions");
  } else {
    const entity = user.entity || "LAOSS";
    state.entity = entity;
    setText("assignedEntityText", entity);
  }

  setText("roleText", state.role);
}

/* =========================
   WEEK SELECTOR
========================= */

function initWeekSelector() {
  const select = $("weekEndingSelect");
  if (!select) return;

  const today = new Date();
  const weeks = [];

  for (let i = 0; i < 12; i++) {
    const d = new Date(today);
    d.setDate(today.getDate() - today.getDay() - i * 7);

    const iso = d.toISOString().split("T")[0];
    weeks.push(iso);
  }

  select.innerHTML = weeks
    .map((w) => `<option value="${w}">${formatDate(w)}</option>`)
    .join("");

  const defaultWeek = getDefaultWeekEnding();
  select.value = defaultWeek;

  state.weekEnding = defaultWeek;

  select.addEventListener("change", () => {
    state.weekEnding = select.value;
    loadCurrentView();
  });
}

/* =========================
   DASHBOARD
========================= */

async function loadDashboard() {
  const week = state.weekEnding || getDefaultWeekEnding();

  const result = await safeApiGet(`/api/dashboard?weekEnding=${week}`, {
    kpis: []
  });

  renderKpiCards(result.kpis || []);
}

/* =========================
   REGION PAGE
========================= */

async function loadRegionPage(entity) {
  const week = state.weekEnding;

  const data = await safeApiGet(
    `/api/weekly?entity=${entity}&weekEnding=${week}`,
    { values: {} }
  );

  const values = data.values || {};

  const fields = document.querySelectorAll("[data-key]");

  fields.forEach((el) => {
    const key = el.dataset.key;
    if (!key) return;

    el.value = values[key] ?? "";
  });
}

/* =========================
   SHARED PAGE
========================= */

async function loadSharedPage(page) {
  const week = state.weekEnding;

  const data = await safeApiGet(
    `/api/shared-data?page=${page}&weekEnding=${week}`,
    { values: {} }
  );

  const values = data.values || {};

  const fields = document.querySelectorAll("[data-key]");

  fields.forEach((el) => {
    const key = el.dataset.key;
    if (!key) return;

    el.value = values[key] ?? "";
  });
}

/* =========================
   SAVE
========================= */

async function saveRegion() {
  const payload = collectRegionFormValues();

  const result = await apiPost("/api/weekly-save", payload);

  if (!result.ok) {
    alert("Save failed.");
    return;
  }

  alert("Saved successfully.");
}

async function saveShared() {
  const payload = collectSharedFormValues();

  const result = await apiPost("/api/shared-save", payload);

  if (!result.ok) {
    alert("Save failed.");
    return;
  }

  alert("Saved successfully.");
}

/* =========================
   ROUTER
========================= */

function activateNav(route) {
  document.querySelectorAll(".nav-link").forEach((link) => {
    link.classList.remove("active");
  });

  const active = document.querySelector(`.nav-link[data-route="${route}"]`);

  if (active) active.classList.add("active");
}

async function loadCurrentView() {
  try {
    switch (state.currentRoute) {
      case "dashboard":
        await loadDashboard();
        break;

      case "region":
        await loadRegionPage(state.entity);
        break;

      case "shared":
        const page =
          document.querySelector(".nav-link.active")?.dataset.page || "PT";
        await loadSharedPage(page);
        break;
    }
  } catch (err) {
    handleFatalError(err);
  }
}

function initNavigation() {
  const links = document.querySelectorAll(".nav-link");

  links.forEach((link) => {
    link.addEventListener("click", async (e) => {
      e.preventDefault();

      const route = link.dataset.route;

      if (!route) return;

      state.currentRoute = route;

      if (route === "region") {
        state.entity = link.dataset.entity || state.entity;
      }

      activateNav(route);

      await loadCurrentView();
    });
  });
}

/* =========================
   BUTTONS
========================= */

function initButtons() {
  const saveBtn = $("saveButton");

  if (saveBtn) {
    saveBtn.addEventListener("click", async () => {
      if (state.currentRoute === "region") {
        await saveRegion();
      } else {
        await saveShared();
      }
    });
  }
}

/* =========================
   INIT
========================= */

async function init() {
  try {
    initWeekSelector();

    await loadUser();

    initNavigation();

    initButtons();

    await loadDashboard();
  } catch (err) {
    handleFatalError(err);
  }
}

document.addEventListener("DOMContentLoaded", init);
