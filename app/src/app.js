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

const REGION_KEYS = ["LAOSS", "NES", "SpineOne", "MRO"];
const SHARED_KEYS = ["PT", "CXNS", "Capacity", "Productivity Builder"];

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
  return state.isAdmin === true || state.role === "admin";
}

function getSignInEl() {
  return $("signInButton");
}

function getSignOutEl() {
  return $("signOutButton");
}

function getNavContainer() {
  return $("dashboardNav");
}

function setSignedInUserText(value) {
  setText("signedInUserText", value);
}

function setAssignedEntityText(value) {
  setText("assignedEntityText", value);
}

function setRoleText(value) {
  setText("roleText", value);
}

function setLoadingHeader() {
  setSignedInUserText("Loading...");
  setAssignedEntityText("Loading...");
  setRoleText("Loading...");
}

function setViewHeader(title, subtitle) {
  setText("dashboardTitle", title);
  setText("dashboardSubtitle", subtitle);
}

function setStatusPanelText(value) {
  setText("dashboardStatusPanel", value);
}

function unique(values) {
  return Array.from(new Set((Array.isArray(values) ? values : []).filter(Boolean)));
}

function normalizeApiMe(result) {
  if (!result || !result.authenticated) return null;

  return {
    authenticated: !!result.authenticated,
    userDetails: result.userDetails || "",
    roles: unique(result.roles),
    entity: result.entity || "",
    isAdmin: !!result.isAdmin
  };
}

function parseJsonSafely(value, fallback = {}) {
  if (!value) return fallback;
  if (typeof value === "object") return value;

  try {
    return JSON.parse(value);
  } catch {
    return fallback;
  }
}

function normalizeRegionValues(result) {
  if (!result || typeof result !== "object") return {};

  if (result.values && typeof result.values === "object") {
    return result.values;
  }

  if (result.valuesJson) {
    return parseJsonSafely(result.valuesJson, {});
  }

  const directKeys = [
    "weekNumber",
    "monthTag",
    "daysInPeriod",
    "totalVisits",
    "visitsPerDay",
    "npActual",
    "establishedActual",
    "surgeryActual",
    "totalCalls",
    "abandonedCalls",
    "abandonmentRate",
    "answeredCallToNpConversion",
    "cashActual"
  ];

  const direct = {};
  for (const key of directKeys) {
    if (key in result) direct[key] = result[key];
  }

  return direct;
}

function normalizeSharedValues(result) {
  if (!result || typeof result !== "object") return {};

  if (result.values && typeof result.values === "object") {
    return result.values;
  }

  if (result.valuesJson) {
    return parseJsonSafely(result.valuesJson, {});
  }

  const direct = {};
  for (const [key, value] of Object.entries(result)) {
    if (!["page", "weekEnding", "source", "updatedAt", "importedAt"].includes(key)) {
      direct[key] = value;
    }
  }

  return direct;
}

async function resolveAuth() {
  setLoadingHeader();

  const me = normalizeApiMe(await safeApiGet("/api/me", null));

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
  state.currentRegion = state.isAdmin ? "LAOSS" : (me.entity || "LAOSS");

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
  } else {
    setSignedInUserText("Not signed in");
    setAssignedEntityText("None");
    setRoleText("guest");

    show(signInEl);
    hide(signOutEl);
  }

  ensureAdminImportLink();
}

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

function ensureAdminImportLink() {
  const existing = $("adminImportNavItem");

  if (!isAdmin()) {
    if (existing) existing.remove();
    return;
  }

  const nav = getNavContainer();
  if (!nav || existing) return;

  const wrapper = document.createElement("div");
  wrapper.id = "adminImportNavItem";
  wrapper.className = "sidebar-admin-link-wrap";
  wrapper.innerHTML = `<a class="nav-link" href="./admin-import.html">Admin Import</a>`;
  nav.appendChild(wrapper);
}

function activateNav(link) {
  document.querySelectorAll(".nav-link").forEach((el) => {
    el.classList.remove("active");
  });
  if (link) link.classList.add("active");
}

function formatWhole(value) {
  return Math.round(Number(value || 0)).toLocaleString();
}

function formatDecimal(value, digits = 1) {
  return Number(value || 0).toFixed(digits);
}

function formatPercent(value, digits = 1) {
  return `${Number(value || 0).toFixed(digits)}%`;
}

function badgeFromStatus(status) {
  const s = String(status || "").toLowerCase();
  if (s === "up" || s === "improved") return "green";
  if (s === "down" || s === "worse") return "red";
  return "yellow";
}

function renderCardsFromItems(items) {
  renderKpiCards(
    items.map((item) => ({
      label: item.label,
      value: item.value,
      meta: item.meta || "",
      status: item.status || "",
      statusColor: item.statusColor || badgeFromStatus(item.status)
    }))
  );
}

function inferTrend(current, previous, betterDirection = "up") {
  const c = Number(current || 0);
  const p = Number(previous || 0);
  const diff = c - p;

  if (betterDirection === "down") {
    if (diff < 0) return { status: "Improved", statusColor: "green", diff };
    if (diff > 0) return { status: "Worse", statusColor: "red", diff };
    return { status: "Flat", statusColor: "yellow", diff };
  }

  if (diff > 0) return { status: "Up", statusColor: "green", diff };
  if (diff < 0) return { status: "Down", statusColor: "red", diff };
  return { status: "Flat", statusColor: "yellow", diff };
}

function buildRegionCards(current, previous = {}) {
  const visitTrend = inferTrend(current.totalVisits, previous.totalVisits, "up");
  const vpdTrend = inferTrend(current.visitsPerDay, previous.visitsPerDay, "up");
  const npTrend = inferTrend(current.npActual, previous.npActual, "up");
  const surgTrend = inferTrend(current.surgeryActual, previous.surgeryActual, "up");
  const callTrend = inferTrend(current.totalCalls, previous.totalCalls, "up");
  const abdTrend = inferTrend(current.abandonmentRate, previous.abandonmentRate, "down");
  const convTrend = inferTrend(current.answeredCallToNpConversion, previous.answeredCallToNpConversion, "up");
  const cashTrend = inferTrend(current.cashActual, previous.cashActual, "up");

  return [
    {
      label: "Total Visits",
      value: formatWhole(current.totalVisits),
      meta: `${visitTrend.diff >= 0 ? "+" : ""}${formatWhole(visitTrend.diff)} vs prior week`,
      status: visitTrend.status,
      statusColor: visitTrend.statusColor
    },
    {
      label: "Visits / Day",
      value: formatDecimal(current.visitsPerDay, 1),
      meta: `${vpdTrend.diff >= 0 ? "+" : ""}${formatDecimal(vpdTrend.diff, 1)} vs prior week`,
      status: vpdTrend.status,
      statusColor: vpdTrend.statusColor
    },
    {
      label: "New Patients",
      value: formatWhole(current.npActual),
      meta: `${npTrend.diff >= 0 ? "+" : ""}${formatWhole(npTrend.diff)} vs prior week`,
      status: npTrend.status,
      statusColor: npTrend.statusColor
    },
    {
      label: "Surgical Cases",
      value: formatWhole(current.surgeryActual),
      meta: `${surgTrend.diff >= 0 ? "+" : ""}${formatWhole(surgTrend.diff)} vs prior week`,
      status: surgTrend.status,
      statusColor: surgTrend.statusColor
    },
    {
      label: "Call Volume",
      value: formatWhole(current.totalCalls),
      meta: `${callTrend.diff >= 0 ? "+" : ""}${formatWhole(callTrend.diff)} vs prior week`,
      status: callTrend.status,
      statusColor: callTrend.statusColor
    },
    {
      label: "Abandonment Rate",
      value: formatPercent(current.abandonmentRate, 1),
      meta: `${abdTrend.diff >= 0 ? "+" : ""}${formatDecimal(abdTrend.diff, 1)} pts vs prior week`,
      status: abdTrend.status,
      statusColor: abdTrend.statusColor
    },
    {
      label: "Answered Call to NP %",
      value: formatPercent(current.answeredCallToNpConversion, 1),
      meta: `${convTrend.diff >= 0 ? "+" : ""}${formatDecimal(convTrend.diff, 1)} pts vs prior week`,
      status: convTrend.status,
      statusColor: convTrend.statusColor
    },
    {
      label: "Cash Collected",
      value: `$${formatWhole(current.cashActual)}`,
      meta: `${cashTrend.diff >= 0 ? "+" : ""}$${formatWhole(cashTrend.diff)} vs prior week`,
      status: cashTrend.status,
      statusColor: cashTrend.statusColor
    }
  ];
}

function buildSharedCards(page, current, previous = {}) {
  if (page === "PT") {
    const visitTrend = inferTrend(current.ptScheduledVisits, previous.ptScheduledVisits, "up");
    const cancelTrend = inferTrend(current.ptCancellations, previous.ptCancellations, "down");
    const noShowTrend = inferTrend(current.ptNoShows, previous.ptNoShows, "down");
    const rescheduleTrend = inferTrend(current.ptReschedules, previous.ptReschedules, "down");
    const unitsTrend = inferTrend(current.totalUnitsBilled, previous.totalUnitsBilled, "up");

    return [
      {
        label: "PT Scheduled Visits",
        value: formatWhole(current.ptScheduledVisits),
        meta: `${visitTrend.diff >= 0 ? "+" : ""}${formatWhole(visitTrend.diff)} vs prior week`,
        status: visitTrend.status,
        statusColor: visitTrend.statusColor
      },
      {
        label: "PT Cancellations",
        value: formatWhole(current.ptCancellations),
        meta: `${cancelTrend.diff >= 0 ? "+" : ""}${formatWhole(cancelTrend.diff)} vs prior week`,
        status: cancelTrend.status,
        statusColor: cancelTrend.statusColor
      },
      {
        label: "PT No Shows",
        value: formatWhole(current.ptNoShows),
        meta: `${noShowTrend.diff >= 0 ? "+" : ""}${formatWhole(noShowTrend.diff)} vs prior week`,
        status: noShowTrend.status,
        statusColor: noShowTrend.statusColor
      },
      {
        label: "PT Reschedules",
        value: formatWhole(current.ptReschedules),
        meta: `${rescheduleTrend.diff >= 0 ? "+" : ""}${formatWhole(rescheduleTrend.diff)} vs prior week`,
        status: rescheduleTrend.status,
        statusColor: rescheduleTrend.statusColor
      },
      {
        label: "Units Billed",
        value: formatWhole(current.totalUnitsBilled),
        meta: `${unitsTrend.diff >= 0 ? "+" : ""}${formatWhole(unitsTrend.diff)} vs prior week`,
        status: unitsTrend.status,
        statusColor: unitsTrend.statusColor
      }
    ];
  }

  if (page === "CXNS") {
    const scheduledTrend = inferTrend(current.scheduledAppts, previous.scheduledAppts, "up");
    const cancelTrend = inferTrend(current.cancellations, previous.cancellations, "down");
    const noShowTrend = inferTrend(current.noShows, previous.noShows, "down");
    const rescheduleTrend = inferTrend(current.reschedules, previous.reschedules, "down");

    return [
      {
        label: "CXNS Scheduled",
        value: formatWhole(current.scheduledAppts),
        meta: `${scheduledTrend.diff >= 0 ? "+" : ""}${formatWhole(scheduledTrend.diff)} vs prior week`,
        status: scheduledTrend.status,
        statusColor: scheduledTrend.statusColor
      },
      {
        label: "CXNS Cancellations",
        value: formatWhole(current.cancellations),
        meta: `${cancelTrend.diff >= 0 ? "+" : ""}${formatWhole(cancelTrend.diff)} vs prior week`,
        status: cancelTrend.status,
        statusColor: cancelTrend.statusColor
      },
      {
        label: "CXNS No Shows",
        value: formatWhole(current.noShows),
        meta: `${noShowTrend.diff >= 0 ? "+" : ""}${formatWhole(noShowTrend.diff)} vs prior week`,
        status: noShowTrend.status,
        statusColor: noShowTrend.statusColor
      },
      {
        label: "CXNS Reschedules",
        value: formatWhole(current.reschedules),
        meta: `${rescheduleTrend.diff >= 0 ? "+" : ""}${formatWhole(rescheduleTrend.diff)} vs prior week`,
        status: rescheduleTrend.status,
        statusColor: rescheduleTrend.statusColor
      }
    ];
  }

  return [
    {
      label: page,
      value: "Coming Soon",
      meta: "This page is not fully wired yet.",
      status: "Flat",
      statusColor: "yellow"
    }
  ];
}

function getPreviousWeekEnding(weekEnding) {
  const d = new Date(`${weekEnding}T12:00:00Z`);
  d.setUTCDate(d.getUTCDate() - 7);
  return d.toISOString().split("T")[0];
}

function findNavLinkForRegion(entity) {
  return Array.from(document.querySelectorAll(".nav-link")).find((el) => {
    const dataEntity = (el.dataset.entity || "").trim();
    const text = (el.textContent || "").trim();
    return dataEntity === entity || text === entity;
  });
}

function findNavLinkForShared(page) {
  return Array.from(document.querySelectorAll(".nav-link")).find((el) => {
    const dataPage = (el.dataset.page || "").trim();
    const text = (el.textContent || "").trim();
    return dataPage === page || text === page;
  });
}

function findExecutiveNavLink() {
  return Array.from(document.querySelectorAll(".nav-link")).find((el) =>
    ((el.textContent || "").trim().toLowerCase().includes("executive"))
  );
}

async function loadDashboard() {
  const week = currentWeekEnding();

  setViewHeader(
    "Executive Summary",
    "Weekly companywide KPI overview with trends, target comparisons, and submission visibility."
  );
  setStatusPanelText("Loading dashboard...");

  const result = await safeApiGet(
    `/api/dashboard?weekEnding=${encodeURIComponent(week)}`,
    { kpis: [] }
  );

  renderKpiCards(Array.isArray(result.kpis) ? result.kpis : []);
  setStatusPanelText("Executive dashboard loaded.");
}

async function loadRegionPage(entity) {
  const week = currentWeekEnding();
  const previousWeek = getPreviousWeekEnding(week);

  setViewHeader(
    `${entity} Weekly View`,
    `Weekly operational performance for ${entity}.`
  );
  setStatusPanelText(`Loading ${entity} weekly data...`);

  const currentResponse = await safeApiGet(
    `/api/region-weekly?entity=${encodeURIComponent(entity)}&weekEnding=${encodeURIComponent(week)}`,
    {}
  );

  const previousResponse = await safeApiGet(
    `/api/region-weekly?entity=${encodeURIComponent(entity)}&weekEnding=${encodeURIComponent(previousWeek)}`,
    {}
  );

  const currentValues = normalizeRegionValues(currentResponse);
  const previousValues = normalizeRegionValues(previousResponse);

  const cards = buildRegionCards(currentValues, previousValues);
  renderCardsFromItems(cards);

  const hasAnyCurrentValue = Object.values(currentValues).some((v) => v !== null && v !== undefined && v !== "");
  setStatusPanelText(
    hasAnyCurrentValue
      ? `${entity} weekly data loaded.`
      : `${entity} has no saved data for ${formatDate(week)}.`
  );
}

async function loadSharedPage(page) {
  const week = currentWeekEnding();
  const previousWeek = getPreviousWeekEnding(week);

  setViewHeader(
    `${page} Weekly View`,
    `Weekly metrics for ${page}.`
  );
  setStatusPanelText(`Loading ${page} data...`);

  if (page === "Capacity" || page === "Productivity Builder") {
    renderCardsFromItems([
      {
        label: page,
        value: "Coming Soon",
        meta: "This section is not fully wired yet.",
        status: "Flat",
        statusColor: "yellow"
      }
    ]);
    setStatusPanelText(`${page} is not fully wired yet.`);
    return;
  }

  const currentResponse = await safeApiGet(
    `/api/shared-data?page=${encodeURIComponent(page)}&weekEnding=${encodeURIComponent(week)}`,
    {}
  );

  const previousResponse = await safeApiGet(
    `/api/shared-data?page=${encodeURIComponent(page)}&weekEnding=${encodeURIComponent(previousWeek)}`,
    {}
  );

  const currentValues = normalizeSharedValues(currentResponse);
  const previousValues = normalizeSharedValues(previousResponse);

  const cards = buildSharedCards(page, currentValues, previousValues);
  renderCardsFromItems(cards);

  const hasAnyCurrentValue = Object.values(currentValues).some((v) => v !== null && v !== undefined && v !== "");
  setStatusPanelText(
    hasAnyCurrentValue
      ? `${page} weekly data loaded.`
      : `${page} has no saved data for ${formatDate(week)}.`
  );
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

function routeFromLink(link) {
  const explicit = link.dataset.route;
  if (explicit) return explicit;

  const dataEntity = (link.dataset.entity || "").trim();
  const dataPage = (link.dataset.page || "").trim();

  if (REGION_KEYS.includes(dataEntity)) return "region";
  if (SHARED_KEYS.includes(dataPage)) return "shared";

  const text = (link.textContent || "").trim();

  if (REGION_KEYS.includes(text)) return "region";
  if (SHARED_KEYS.includes(text)) return "shared";
  if (text.toLowerCase().includes("executive")) return "dashboard";

  return "dashboard";
}

function initNavigation() {
  const nav = getNavContainer();
  if (!nav) return;

  nav.addEventListener("click", async (e) => {
    const link = e.target.closest(".nav-link");
    if (!link) return;

    const href = link.getAttribute("href") || "";
    if (href.includes("admin-import.html")) return;

    e.preventDefault();

    const route = routeFromLink(link);
    state.currentRoute = route;

    if (route === "region") {
      const clickedEntity = (link.dataset.entity || link.textContent || "").trim();
      state.currentRegion = REGION_KEYS.includes(clickedEntity) ? clickedEntity : state.currentRegion;
    }

    if (route === "shared") {
      const clickedPage = (link.dataset.page || link.textContent || "").trim();
      state.currentSharedPage = SHARED_KEYS.includes(clickedPage) ? clickedPage : state.currentSharedPage;
    }

    activateNav(link);
    await loadCurrentView();
  });
}

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
  await loadCurrentView();
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
  await loadCurrentView();
}

function initButtons() {
  const saveBtn = $("saveButton");
  const submitBtn = $("submitWeekButton");

  if (saveBtn) {
    saveBtn.addEventListener("click", async () => {
      if (state.currentRoute === "region") {
        await saveRegion();
        return;
      }

      if (state.currentRoute === "shared") {
        await saveShared();
        return;
      }

      alert("Nothing to save on this page.");
    });
  }

  if (submitBtn) {
    submitBtn.addEventListener("click", async () => {
      if (state.currentRoute === "region") {
        await saveRegion();
        alert("Week submitted.");
        return;
      }

      if (state.currentRoute === "shared") {
        await saveShared();
        alert("Week submitted.");
        return;
      }

      alert("Nothing to submit on this page.");
    });
  }
}

function initHeroButtons() {
  const goToRegionBtn = $("goToMyRegionButton");
  const executiveBtn = $("executiveSummaryButton");

  if (goToRegionBtn) {
    goToRegionBtn.addEventListener("click", async () => {
      state.currentRoute = "region";
      state.currentRegion = isAdmin()
        ? "LAOSS"
        : (state.entity === "Admin" ? "LAOSS" : state.entity);

      activateNav(findNavLinkForRegion(state.currentRegion));
      await loadCurrentView();
    });
  }

  if (executiveBtn) {
    executiveBtn.addEventListener("click", async () => {
      state.currentRoute = "dashboard";
      activateNav(findExecutiveNavLink());
      await loadCurrentView();
    });
  }
}

function activateInitialNav() {
  if (state.currentRoute === "region") {
    activateNav(findNavLinkForRegion(state.currentRegion));
    return;
  }

  if (state.currentRoute === "shared") {
    activateNav(findNavLinkForShared(state.currentSharedPage));
    return;
  }

  activateNav(findExecutiveNavLink());
}

async function init() {
  try {
    setLoadingHeader();
    initWeekSelector();
    initButtons();
    initHeroButtons();
    await resolveAuth();
    initNavigation();
    activateInitialNav();
    await loadCurrentView();
  } catch (err) {
    handleFatalError(err);
  }
}

document.addEventListener("DOMContentLoaded", init);
