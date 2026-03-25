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
const ADMIN_EMAILS = ["nperez@unitymsk.com", "tessa.kelley@spineone.com"];

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

function q(selector) {
  return document.querySelector(selector);
}

function qa(selector) {
  return Array.from(document.querySelectorAll(selector));
}

function normalizeText(value) {
  return String(value || "").trim();
}

function normalizeEmail(value) {
  return normalizeText(value).toLowerCase();
}

function emailIsAdmin(value) {
  return ADMIN_EMAILS.includes(normalizeEmail(value));
}

function unique(values) {
  return Array.from(new Set((Array.isArray(values) ? values : []).filter(Boolean)));
}

function setText(id, value) {
  const el = $(id);
  if (el) el.textContent = value;
}

function setHtml(id, value) {
  const el = $(id);
  if (el) el.innerHTML = value;
}

function show(el) {
  if (el) el.style.display = "";
}

function hide(el) {
  if (el) el.style.display = "none";
}

function currentWeekEnding() {
  const candidates = [
    $("weekEndingSelect"),
    $("anchorWeekEnding"),
    $("anchorWeekEndingInput"),
    q("input[type='date'][id*='week']"),
    q("input[type='date']")
  ].filter(Boolean);

  const control = candidates[0];
  const value = normalizeText(control?.value);
  return value || state.weekEnding || getDefaultWeekEnding();
}

function currentPeriod() {
  const candidates = [
    $("periodSelect"),
    $("dashboardPeriodSelect"),
    q("select[id*='period']"),
    q("select[name='period']")
  ].filter(Boolean);

  return normalizeText(candidates[0]?.value) || "Current Week";
}

function currentCompareAgainst() {
  const candidates = [
    $("compareAgainstSelect"),
    $("compareAgainst"),
    q("select[id*='compare']"),
    q("select[name='compareAgainst']")
  ].filter(Boolean);

  return normalizeText(candidates[0]?.value) || "Prior Period";
}

function currentEntityScope() {
  const candidates = [
    $("entityScopeSelect"),
    $("entityScope"),
    q("select[id*='scope']"),
    q("select[name='entityScope']")
  ].filter(Boolean);

  return normalizeText(candidates[0]?.value) || "All Entities";
}

function isAdmin() {
  return state.isAdmin === true || state.role === "admin";
}

function getNavContainer() {
  return $("dashboardNav") || q(".sidebar-nav") || document.body;
}

function getSignInEl() {
  return $("signInButton") || q("a[href*='/.auth/login']");
}

function getSignOutEl() {
  return $("signOutButton") || q("a[href*='/.auth/logout']");
}

function setSignedInUserText(value) {
  setText("signedInUserText", value);
  setText("signedInAsText", value);
}

function setAssignedEntityText(value) {
  setText("assignedEntityText", value);
  setText("entityText", value);
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
  setText("statusMessage", value);
  setText("dashboardStatusMessage", value);
}

function setTopRightStatus(value) {
  setText("headerStatusText", value);
  setText("topRightStatusText", value);
}

function safeJsonStringify(value) {
  try {
    return JSON.stringify(value, null, 2);
  } catch {
    return String(value);
  }
}

function setDebugOutput(value) {
  const text = typeof value === "string" ? value : safeJsonStringify(value);

  [
    "dashboardDebugOutput",
    "debugOutput",
    "executiveDebugOutput",
    "trendsDebugOutput"
  ].forEach((id) => {
    const el = $(id);
    if (el) el.textContent = text;
  });
}

function normalizeApiMe(result) {
  if (!result || !result.authenticated) return null;

  const userDetails = result.userDetails || "";
  const roles = unique(result.roles);
  const apiSaysAdmin =
    !!result.isAdmin ||
    roles.some((r) => normalizeText(r).toLowerCase() === "admin");
  const forcedAdmin = emailIsAdmin(userDetails);

  return {
    authenticated: true,
    userDetails,
    roles,
    entity: forcedAdmin ? "Admin" : (result.entity || ""),
    isAdmin: apiSaysAdmin || forcedAdmin
  };
}

function normalizeAuthMe(result) {
  const principal = result?.clientPrincipal;
  if (!principal || !principal.userId) return null;

  const userDetails = principal.userDetails || principal.userId || "";
  const roles = unique(principal.userRoles || []);
  const roleAdmin = roles.some((r) => normalizeText(r).toLowerCase() === "admin");
  const forcedAdmin = emailIsAdmin(userDetails);

  return {
    authenticated: true,
    userDetails,
    roles,
    entity: forcedAdmin ? "Admin" : "",
    isAdmin: roleAdmin || forcedAdmin
  };
}

async function resolveAuth() {
  setLoadingHeader();

  let me = null;

  try {
    me = normalizeApiMe(await safeApiGet("/api/me", null));
  } catch {
    me = null;
  }

  if (!me) {
    try {
      me = normalizeAuthMe(await safeApiGet("/.auth/me", null));
    } catch {
      me = null;
    }
  }

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

  if (signInEl) signInEl.setAttribute("href", "/.auth/login/aad");
  if (signOutEl) signOutEl.setAttribute("href", "/.auth/logout");

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

function initWeekSelector() {
  const select =
    $("weekEndingSelect") ||
    $("anchorWeekEnding") ||
    $("anchorWeekEndingInput");

  if (!select || select.tagName !== "SELECT") {
    const input =
      $("anchorWeekEnding") ||
      $("anchorWeekEndingInput") ||
      q("input[type='date'][id*='week']") ||
      q("input[type='date']");

    if (input) {
      if (!normalizeText(input.value)) {
        input.value = state.weekEnding;
      }

      input.addEventListener("change", () => {
        state.weekEnding = normalizeText(input.value) || getDefaultWeekEnding();
      });
    }

    return;
  }

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

  if (weeks.includes(state.weekEnding)) {
    select.value = state.weekEnding;
  } else if (weeks.length) {
    select.value = weeks[0];
    state.weekEnding = weeks[0];
  }

  select.addEventListener("change", () => {
    state.weekEnding = normalizeText(select.value) || getDefaultWeekEnding();
  });
}

function activateNav(link) {
  qa(".nav-link").forEach((el) => el.classList.remove("active"));
  if (link) link.classList.add("active");
}

function findNavLinkByText(needle) {
  const lower = normalizeText(needle).toLowerCase();
  return qa(".nav-link").find((el) =>
    normalizeText(el.textContent).toLowerCase().includes(lower)
  );
}

function routeFromLink(link) {
  const explicit = normalizeText(link?.dataset?.route);
  if (explicit) return explicit;

  const dataEntity = normalizeText(link?.dataset?.entity);
  const dataPage = normalizeText(link?.dataset?.page);
  const text = normalizeText(link?.textContent).toLowerCase();

  if (REGION_KEYS.includes(dataEntity)) return "region";
  if (SHARED_KEYS.includes(dataPage)) return "shared";

  if (text === "dashboard") return "dashboard";
  if (text === "weekly entry") return "entry";
  if (text === "executive summary") return "dashboard";
  if (text === "trends") return "trends";
  if (text === "admin import") return "admin-import";
  if (REGION_KEYS.map((x) => x.toLowerCase()).includes(text)) return "region";
  if (SHARED_KEYS.map((x) => x.toLowerCase()).includes(text)) return "shared";

  return "dashboard";
}

function parseNumber(value) {
  const n = Number(value || 0);
  return Number.isFinite(n) ? n : 0;
}

function formatWhole(value) {
  return Math.round(parseNumber(value)).toLocaleString();
}

function formatDecimal(value, digits = 1) {
  return parseNumber(value).toFixed(digits);
}

function formatPercent(value, digits = 1) {
  return `${parseNumber(value).toFixed(digits)}%`;
}

function formatCurrency(value) {
  return `$${Math.round(parseNumber(value)).toLocaleString()}`;
}

function inferTrend(current, previous, betterDirection = "up") {
  const c = parseNumber(current);
  const p = parseNumber(previous);
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

function renderCardsFromItems(items) {
  renderKpiCards(
    (Array.isArray(items) ? items : []).map((item) => ({
      label: item.label,
      value: item.value,
      meta: item.meta || "",
      status: item.status || "",
      statusColor: item.statusColor || "yellow"
    }))
  );
}

function normalizeRegionValues(result) {
  if (!result || typeof result !== "object") return {};
  if (result.values && typeof result.values === "object") return result.values;
  if (result.valuesJson) {
    try {
      return JSON.parse(result.valuesJson);
    } catch {
      return {};
    }
  }
  return {};
}

function normalizeSharedValues(result) {
  if (!result || typeof result !== "object") return {};
  if (result.values && typeof result.values === "object") return result.values;
  if (result.valuesJson) {
    try {
      return JSON.parse(result.valuesJson);
    } catch {
      return {};
    }
  }
  return {};
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
      value: formatCurrency(current.cashActual),
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
      meta: "This section is not fully wired yet.",
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

async function loadDashboard() {
  const week = currentWeekEnding();

  setViewHeader(
    "Dashboard",
    "Multi-entity performance view with period and comparison controls."
  );
  setStatusPanelText("Loading dashboard...");
  setTopRightStatus("Loading...");

  const query = new URLSearchParams({
    weekEnding: week,
    period: currentPeriod(),
    compareAgainst: currentCompareAgainst(),
    entityScope: currentEntityScope()
  });

  const result = await safeApiGet(`/api/dashboard?${query.toString()}`, { kpis: [] });

  renderKpiCards(Array.isArray(result?.kpis) ? result.kpis : []);
  setDebugOutput(result);

  setStatusPanelText("Dashboard loaded.");
  setTopRightStatus("Ready");
}

async function loadRegionPage(entity) {
  const week = currentWeekEnding();
  const previousWeek = getPreviousWeekEnding(week);

  setViewHeader(
    `${entity} Weekly View`,
    `Weekly operational performance for ${entity}.`
  );
  setStatusPanelText(`Loading ${entity} weekly data...`);
  setTopRightStatus("Loading...");

  const currentResponse = await safeApiGet(
    `/api/weekly?entity=${encodeURIComponent(entity)}&weekEnding=${encodeURIComponent(week)}`,
    {}
  );

  const previousResponse = await safeApiGet(
    `/api/weekly?entity=${encodeURIComponent(entity)}&weekEnding=${encodeURIComponent(previousWeek)}`,
    {}
  );

  const currentValues = normalizeRegionValues(currentResponse);
  const previousValues = normalizeRegionValues(previousResponse);

  renderCardsFromItems(buildRegionCards(currentValues, previousValues));
  setDebugOutput({
    currentRoute: state.currentRoute,
    entity,
    weekEnding: week,
    current: currentResponse,
    previous: previousResponse
  });

  setStatusPanelText(`${entity} weekly data loaded.`);
  setTopRightStatus("Ready");
}

async function loadSharedPage(page) {
  const week = currentWeekEnding();
  const previousWeek = getPreviousWeekEnding(week);

  setViewHeader(
    `${page} Weekly View`,
    `Weekly metrics for ${page}.`
  );
  setStatusPanelText(`Loading ${page} data...`);
  setTopRightStatus("Loading...");

  if (page === "Capacity" || page === "Productivity Builder") {
    renderCardsFromItems(buildSharedCards(page, {}, {}));
    setDebugOutput({ page, message: "Not fully wired yet." });
    setStatusPanelText(`${page} is not fully wired yet.`);
    setTopRightStatus("Ready");
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

  renderCardsFromItems(buildSharedCards(page, currentValues, previousValues));
  setDebugOutput({
    currentRoute: state.currentRoute,
    page,
    weekEnding: week,
    current: currentResponse,
    previous: previousResponse
  });

  setStatusPanelText(`${page} weekly data loaded.`);
  setTopRightStatus("Ready");
}

async function loadExecutiveSummary() {
  const week = currentWeekEnding();

  setViewHeader(
    "Executive Summary",
    "Weekly companywide KPI overview with trends, target comparisons, and submission visibility."
  );
  setStatusPanelText("Loading executive summary...");
  setTopRightStatus("Loading...");

  const result = await safeApiGet(
    `/api/executive-summary?weekEnding=${encodeURIComponent(week)}`,
    { kpis: [] }
  );

  renderKpiCards(Array.isArray(result?.kpis) ? result.kpis : []);
  setDebugOutput(result);

  setStatusPanelText("Executive summary loaded.");
  setTopRightStatus("Ready");
}

async function loadTrends() {
  const week = currentWeekEnding();

  setViewHeader(
    "Trends",
    "Historical trend view for the selected entity and date range."
  );
  setStatusPanelText("Loading trends...");
  setTopRightStatus("Loading...");

  const query = new URLSearchParams({
    weekEnding: week,
    entity: state.currentRegion
  });

  const result = await safeApiGet(`/api/trends?${query.toString()}`, { items: [] });

  setDebugOutput(result);
  setStatusPanelText("Trends loaded.");
  setTopRightStatus("Ready");
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

    if (state.currentRoute === "trends") {
      await loadTrends();
      return;
    }

    if (state.currentRoute === "entry") {
      await loadRegionPage(state.currentRegion);
      return;
    }

    if (state.currentRoute === "executive") {
      await loadExecutiveSummary();
      return;
    }

    await loadDashboard();
  } catch (error) {
    setTopRightStatus("Error");
    setDebugOutput({
      error: error?.message || String(error),
      stack: error?.stack || null
    });
    handleFatalError(error);
  }
}

function initNavigation() {
  const nav = getNavContainer();
  if (!nav) return;

  nav.addEventListener("click", async (e) => {
    const link = e.target.closest(".nav-link");
    if (!link) return;

    const href = normalizeText(link.getAttribute("href"));
    if (href.includes("admin-import.html")) {
      return;
    }

    e.preventDefault();

    const route = routeFromLink(link);
    const text = normalizeText(link.textContent);

    if (route === "dashboard") {
      state.currentRoute = text.toLowerCase().includes("executive") ? "executive" : "dashboard";
    } else {
      state.currentRoute = route;
    }

    if (route === "region") {
      const clickedEntity = normalizeText(link.dataset.entity || text);
      if (REGION_KEYS.includes(clickedEntity)) {
        state.currentRegion = clickedEntity;
      }
    }

    if (route === "shared") {
      const clickedPage = normalizeText(link.dataset.page || text);
      if (SHARED_KEYS.includes(clickedPage)) {
        state.currentSharedPage = clickedPage;
      }
    }

    if (route === "trends") {
      state.currentRoute = "trends";
    }

    if (route === "entry") {
      state.currentRoute = "entry";
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
    setDebugOutput(result || { error: "Save failed." });
    return;
  }

  setDebugOutput(result);
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
    setDebugOutput(result || { error: "Save failed." });
    return;
  }

  setDebugOutput(result);
  alert("Saved successfully.");
  await loadCurrentView();
}

function initButtons() {
  const saveBtn = $("saveButton") || $("saveBtn");
  const submitBtn = $("submitWeekButton") || $("submitBtn");
  const loadDashboardBtn = $("loadDashboardButton") || $("loadDashboardBtn");

  if (saveBtn) {
    saveBtn.addEventListener("click", async () => {
      if (state.currentRoute === "shared") {
        await saveShared();
        return;
      }

      await saveRegion();
    });
  }

  if (submitBtn) {
    submitBtn.addEventListener("click", async () => {
      if (state.currentRoute === "shared") {
        await saveShared();
        alert("Week submitted.");
        return;
      }

      await saveRegion();
      alert("Week submitted.");
    });
  }

  if (loadDashboardBtn) {
    loadDashboardBtn.addEventListener("click", async () => {
      state.currentRoute = "dashboard";
      activateNav(findNavLinkByText("dashboard"));
      await loadCurrentView();
    });
  }
}

function initControlReloads() {
  [
    $("periodSelect"),
    $("dashboardPeriodSelect"),
    $("compareAgainstSelect"),
    $("compareAgainst"),
    $("entityScopeSelect"),
    $("entityScope"),
    $("weekEndingSelect"),
    $("anchorWeekEnding"),
    $("anchorWeekEndingInput")
  ]
    .filter(Boolean)
    .forEach((el) => {
      el.addEventListener("change", async () => {
        if (el.type === "date" || el.tagName === "SELECT") {
          state.weekEnding = currentWeekEnding();
        }
      });
    });
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

      activateNav(findNavLinkByText(state.currentRegion));
      await loadCurrentView();
    });
  }

  if (executiveBtn) {
    executiveBtn.addEventListener("click", async () => {
      state.currentRoute = "executive";
      activateNav(findNavLinkByText("executive"));
      await loadCurrentView();
    });
  }
}

function activateInitialNav() {
  const dashboardLink = findNavLinkByText("dashboard");
  if (dashboardLink) {
    activateNav(dashboardLink);
  }
}

async function init() {
  try {
    setLoadingHeader();
    setStatusPanelText("Initializing...");
    setTopRightStatus("Loading...");
    setDebugOutput("App booting...");

    initWeekSelector();
    initButtons();
    initControlReloads();
    initHeroButtons();

    await resolveAuth();

    initNavigation();
    activateInitialNav();

    await loadCurrentView();
  } catch (error) {
    setTopRightStatus("Error");
    setDebugOutput({
      error: error?.message || String(error),
      stack: error?.stack || null
    });
    handleFatalError(error);
  }
}

document.addEventListener("DOMContentLoaded", init);
