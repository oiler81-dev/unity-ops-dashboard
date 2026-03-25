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
  currentSharedPage: "PT",
  submissionStatus: "Draft",
  trendsEntity: "LAOSS",
  trendsMode: "recent",
  trendsWeeks: 8,
  trendsStartDate: "",
  trendsEndDate: ""
};

const $ = (id) => document.getElementById(id);

function setText(id, value) {
  const el = $(id);
  if (el) el.textContent = value;
}

function showElement(el) {
  if (el) el.style.display = "";
}

function hideElement(el) {
  if (el) el.style.display = "none";
}

function toNumber(value, fallback = 0) {
  const n = Number(value);
  return Number.isFinite(n) ? n : fallback;
}

function addDays(iso, days) {
  const d = new Date(`${iso}T12:00:00Z`);
  d.setUTCDate(d.getUTCDate() + days);
  return d.toISOString().slice(0, 10);
}

function currentWeekEnding() {
  return $("weekEndingSelect")?.value || state.weekEnding || getDefaultWeekEnding();
}

function isAdmin() {
  return !!state.isAdmin || state.role === "admin";
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

function activateNav(link) {
  document.querySelectorAll(".nav-link").forEach((el) => el.classList.remove("active"));
  if (link) link.classList.add("active");
}

function findNavLink(route, value) {
  return Array.from(document.querySelectorAll(".nav-link")).find((el) => {
    if (route === "dashboard") return (el.dataset.route || "") === "dashboard";
    if (route === "trends") return (el.dataset.route || "") === "trends";
    if (route === "region") return (el.dataset.entity || "") === value;
    if (route === "shared") return (el.dataset.page || "") === value;
    return false;
  });
}

function setSubmissionStatus(status) {
  state.submissionStatus = status || "Draft";
  setText("submissionStatusText", state.submissionStatus);
}

function setDebug(data) {
  const el = $("debugJsonPanel");
  if (!el) return;
  el.textContent = typeof data === "string" ? data : JSON.stringify(data, null, 2);
}

function showStandardView() {
  $("standardView")?.classList.remove("view-hidden");
  $("trendsView")?.classList.add("view-hidden");
}

function showTrendsView() {
  $("standardView")?.classList.add("view-hidden");
  $("trendsView")?.classList.remove("view-hidden");
}

function normalizeMeResult(me) {
  if (!me || !me.authenticated) {
    return {
      authenticated: false,
      userDetails: "",
      isAdmin: false,
      role: "guest",
      entity: "None"
    };
  }

  const roles = Array.isArray(me.roles) ? me.roles.map((x) => String(x).toLowerCase()) : [];
  const admin = !!me.isAdmin || roles.includes("admin");
  const resolvedRole = admin ? "admin" : (me.role || "user");
  const resolvedEntity = admin ? "Admin" : (me.entity || "None");

  return {
    authenticated: true,
    userDetails: me.userDetails || "Unknown User",
    isAdmin: admin,
    role: resolvedRole,
    entity: resolvedEntity
  };
}

function syncAuthUi() {
  setText("signedInUserText", state.authenticated ? state.userDetails : "Not signed in");
  setText("assignedEntityText", state.entity || "None");
  setText("roleText", state.role || "guest");

  if (state.authenticated) {
    hideElement($("signInButton"));
    showElement($("signOutButton"));
  } else {
    showElement($("signInButton"));
    hideElement($("signOutButton"));
  }

  ensureAdminImportLink();
  syncTrendsEntityOptions();
}

async function resolveAuth() {
  const me = await safeApiGet("/api/me", null);
  const normalized = normalizeMeResult(me);

  state.authenticated = normalized.authenticated;
  state.userDetails = normalized.userDetails;
  state.isAdmin = normalized.isAdmin;
  state.role = normalized.role;
  state.entity = normalized.entity;
  state.currentRegion = normalized.isAdmin ? "LAOSS" : (normalized.entity && normalized.entity !== "None" ? normalized.entity : "LAOSS");
  state.trendsEntity = state.currentRegion;

  syncAuthUi();
}

function ensureAdminImportLink() {
  const nav = $("dashboardNav");
  if (!nav) return;

  const existing = $("adminImportNavItem");

  if (!isAdmin()) {
    if (existing) existing.remove();
    return;
  }

  if (existing) return;

  const wrapper = document.createElement("div");
  wrapper.id = "adminImportNavItem";
  wrapper.style.marginTop = "14px";
  wrapper.innerHTML = `<a class="nav-link" href="./admin-import.html">Admin Import</a>`;
  nav.appendChild(wrapper);
}

function buildRecentWeeks() {
  const today = new Date();
  const weeks = [];

  for (let i = 0; i < 52; i += 1) {
    const d = new Date(today);
    d.setDate(today.getDate() - d.getDay() + 5 - (i * 7));
    weeks.push(d.toISOString().slice(0, 10));
  }

  return weeks;
}

function initWeekSelector() {
  const select = $("weekEndingSelect");
  if (!select) return;

  const weeks = buildRecentWeeks();
  select.innerHTML = weeks.map((week) => `<option value="${week}">${formatDate(week)}</option>`).join("");

  if (weeks.includes(state.weekEnding)) {
    select.value = state.weekEnding;
  } else {
    select.value = weeks[0];
    state.weekEnding = weeks[0];
  }

  setText("sidebarWeekEndingText", formatDate(select.value));

  select.addEventListener("change", async () => {
    state.weekEnding = select.value;
    setText("sidebarWeekEndingText", formatDate(select.value));

    if (state.currentRoute !== "trends") {
      await loadCurrentView();
    }
  });
}

function statusColor(status) {
  const s = String(status || "").toLowerCase();
  if (s.includes("up") || s.includes("improved") || s.includes("approved")) return "green";
  if (s.includes("down") || s.includes("worse")) return "red";
  return "yellow";
}

function buildKpiMarkup(items) {
  return items.map((item) => `
    <div class="kpi-card">
      <div class="kpi-label">${item.label}</div>
      <div class="kpi-value">${item.value}</div>
      ${item.meta ? `<div class="kpi-meta">${item.meta}</div>` : ""}
      ${item.status ? `<div class="kpi-status ${item.statusColor || statusColor(item.status)}">${item.status}</div>` : ""}
    </div>
  `).join("");
}

function renderCards(containerId, items) {
  const container = $(containerId);
  if (!container) return;

  if (typeof renderKpiCards === "function" && containerId === "dashboardCards") {
    try {
      renderKpiCards(items);
      return;
    } catch {
      // fall back below
    }
  }

  container.innerHTML = buildKpiMarkup(items);
}

function diffMeta(current, previous, suffix = "") {
  const c = toNumber(current, 0);
  const p = toNumber(previous, 0);
  const diff = c - p;
  const sign = diff >= 0 ? "+" : "";
  return `${sign}${diff.toLocaleString()}${suffix} vs prior week`;
}

function percentDiffMeta(current, previous) {
  const c = toNumber(current, 0);
  const p = toNumber(previous, 0);
  const diff = c - p;
  const sign = diff >= 0 ? "+" : "";
  return `${sign}${diff.toFixed(1)} pts vs prior week`;
}

function simpleTrend(current, previous, betterDirection = "up") {
  const c = toNumber(current, 0);
  const p = toNumber(previous, 0);

  if (betterDirection === "down") {
    if (c < p) return { status: "Improved", statusColor: "green" };
    if (c > p) return { status: "Worse", statusColor: "red" };
    return { status: "Flat", statusColor: "yellow" };
  }

  if (c > p) return { status: "Up", statusColor: "green" };
  if (c < p) return { status: "Down", statusColor: "red" };
  return { status: "Flat", statusColor: "yellow" };
}

function normalizeRegionValues(result) {
  if (!result) return {};
  if (result.values && typeof result.values === "object") return result.values;
  if (result.valuesJson) return parseJsonSafely(result.valuesJson, {});
  return result;
}

function normalizeSharedValues(result) {
  if (!result) return {};
  if (result.values && typeof result.values === "object") return result.values;
  if (result.valuesJson) return parseJsonSafely(result.valuesJson, {});
  return result;
}

async function loadDashboard() {
  showStandardView();

  setText("dashboardTitle", "Executive Summary");
  setText("dashboardSubtitle", "Weekly companywide KPI overview with trends, target comparisons, and submission visibility.");
  setText("dashboardStatusPanel", "Loading dashboard...");
  setSubmissionStatus("Draft");

  const weekEnding = currentWeekEnding();
  const result = await safeApiGet(`/api/dashboard?weekEnding=${encodeURIComponent(weekEnding)}`, { kpis: [] });

  const cards = Array.isArray(result?.kpis) ? result.kpis : [];
  renderCards("dashboardCards", cards);
  setText("dashboardStatusPanel", "Executive dashboard loaded.");
  setDebug(result);
}

async function loadRegionPage(entity) {
  showStandardView();

  setText("dashboardTitle", `${entity} Weekly View`);
  setText("dashboardSubtitle", `Weekly operational performance for ${entity}.`);
  setText("dashboardStatusPanel", `Loading ${entity} weekly data...`);

  const weekEnding = currentWeekEnding();
  const previousWeekEnding = addDays(weekEnding, -7);

  const currentResponse = await safeApiGet(
    `/api/weekly?entity=${encodeURIComponent(entity)}&weekEnding=${encodeURIComponent(weekEnding)}`,
    {}
  );

  const previousResponse = await safeApiGet(
    `/api/weekly?entity=${encodeURIComponent(entity)}&weekEnding=${encodeURIComponent(previousWeekEnding)}`,
    {}
  );

  const current = normalizeRegionValues(currentResponse);
  const previous = normalizeRegionValues(previousResponse);

  const visitsTrend = simpleTrend(current.totalVisits, previous.totalVisits, "up");
  const vpdTrend = simpleTrend(current.visitsPerDay, previous.visitsPerDay, "up");
  const npTrend = simpleTrend(current.npActual, previous.npActual, "up");
  const surgTrend = simpleTrend(current.surgeryActual, previous.surgeryActual, "up");

  renderCards("dashboardCards", [
    {
      label: "Total Visits",
      value: toNumber(current.totalVisits, 0).toLocaleString(),
      meta: diffMeta(current.totalVisits, previous.totalVisits),
      status: visitsTrend.status,
      statusColor: visitsTrend.statusColor
    },
    {
      label: "Visits / Day",
      value: toNumber(current.visitsPerDay, 0).toFixed(1),
      meta: diffMeta(toNumber(current.visitsPerDay, 0), toNumber(previous.visitsPerDay, 0)),
      status: vpdTrend.status,
      statusColor: vpdTrend.statusColor
    },
    {
      label: "New Patients",
      value: toNumber(current.npActual, 0).toLocaleString(),
      meta: diffMeta(current.npActual, previous.npActual),
      status: npTrend.status,
      statusColor: npTrend.statusColor
    },
    {
      label: "Surgical Cases",
      value: toNumber(current.surgeryActual, 0).toLocaleString(),
      meta: diffMeta(current.surgeryActual, previous.surgeryActual),
      status: surgTrend.status,
      statusColor: surgTrend.statusColor
    }
  ]);

  setSubmissionStatus(currentResponse?.status || "Draft");
  setText(
    "dashboardStatusPanel",
    Object.keys(current || {}).length
      ? `${entity} weekly data loaded.`
      : `${entity} has no saved data for ${formatDate(weekEnding)}.`
  );

  setDebug({
    current: currentResponse,
    previous: previousResponse
  });
}

async function loadSharedPage(page) {
  showStandardView();

  setText("dashboardTitle", `${page} Weekly View`);
  setText("dashboardSubtitle", `Weekly metrics for ${page}.`);
  setText("dashboardStatusPanel", `Loading ${page} data...`);

  const weekEnding = currentWeekEnding();

  if (page === "Capacity" || page === "Productivity Builder") {
    renderCards("dashboardCards", [
      {
        label: page,
        value: "Coming Soon",
        meta: "This section is not wired yet.",
        status: "Flat",
        statusColor: "yellow"
      }
    ]);

    setSubmissionStatus("Draft");
    setText("dashboardStatusPanel", `${page} is not fully wired yet.`);
    setDebug({ page, message: "Not wired yet" });
    return;
  }

  const result = await safeApiGet(
    `/api/shared-data?page=${encodeURIComponent(page)}&weekEnding=${encodeURIComponent(weekEnding)}`,
    {}
  );

  const values = normalizeSharedValues(result);

  if (page === "PT") {
    renderCards("dashboardCards", [
      { label: "Scheduled Visits", value: toNumber(values.ptScheduledVisits, 0).toLocaleString() },
      { label: "Cancellations", value: toNumber(values.ptCancellations, 0).toLocaleString() },
      { label: "No Shows", value: toNumber(values.ptNoShows, 0).toLocaleString() },
      { label: "Units Billed", value: toNumber(values.totalUnitsBilled, 0).toLocaleString() }
    ]);
  } else if (page === "CXNS") {
    renderCards("dashboardCards", [
      { label: "Scheduled", value: toNumber(values.scheduledAppts, 0).toLocaleString() },
      { label: "Cancellations", value: toNumber(values.cancellations, 0).toLocaleString() },
      { label: "No Shows", value: toNumber(values.noShows, 0).toLocaleString() },
      { label: "Reschedules", value: toNumber
