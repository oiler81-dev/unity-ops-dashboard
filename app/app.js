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
  currentSharedPage: "PT",
  currentSubmissionStatus: "Draft",
  trendsEntity: "LAOSS",
  trendsMode: "recent",
  trendsLimit: 12,
  trendsStartDate: "",
  trendsEndDate: ""
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

function addDays(isoDate, days) {
  const d = new Date(`${isoDate}T12:00:00Z`);
  d.setUTCDate(d.getUTCDate() + days);
  return d.toISOString().slice(0, 10);
}

function currentWeekEnding() {
  return $("weekEndingSelect")?.value || state.weekEnding || getDefaultWeekEnding();
}

function isAdmin() {
  return state.isAdmin === true || state.role === "admin";
}

function normalizeEmail(value) {
  return String(value || "").trim().toLowerCase();
}

function emailIsAdmin(value) {
  return ADMIN_EMAILS.includes(normalizeEmail(value));
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

function setSubmissionStatusText(value) {
  state.currentSubmissionStatus = value || "Draft";
  setText("submissionStatusText", state.currentSubmissionStatus);
}

function setDebugPanel(value) {
  const el = $("debugJsonPanel");
  if (!el) return;
  el.textContent = typeof value === "string" ? value : JSON.stringify(value, null, 2);
}

function unique(values) {
  return Array.from(new Set((Array.isArray(values) ? values : []).filter(Boolean)));
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
    entity: forcedAdmin ? "Admin" : (result.entity || ""),
    isAdmin: apiSaysAdmin || forcedAdmin
  };
}

function normalizeRegionValues(result) {
  if (!result || typeof result !== "object") return {};

  if (result.values && typeof result.values === "object") {
    return result.values;
  }

  if (result.valuesJson) {
    return parseJsonSafely(result.valuesJson, {});
  }

  return {};
}

function normalizeSharedValues(result) {
  if (!result || typeof result !== "object") return {};

  if (result.values && typeof result.values === "object") {
    return result.values;
  }

  if (result.valuesJson) {
    return parseJsonSafely(result.valuesJson, {});
  }

  return {};
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
    state.trendsEntity = "LAOSS";
    syncAuthUi();
    return;
  }

  state.authenticated = true;
  state.userDetails = me.userDetails || "Unknown User";
  state.isAdmin = !!me.isAdmin;
  state.role = state.isAdmin ? "admin" : "user";
  state.entity = state.isAdmin ? "Admin" : (me.entity || "LAOSS");
  state.currentRegion = state.isAdmin ? "LAOSS" : (me.entity || "LAOSS");
  state.trendsEntity = state.isAdmin ? "LAOSS" : state.currentRegion;

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
  syncTrendsEntityOptions();
}

function buildRecentWeeksList() {
  const today = new Date();
  const weeks = [];

  for (let i = 0; i < 52; i += 1) {
    const d = new Date(today);
    d.setDate(today.getDate() - today.getDay() - i * 7);
    weeks.push(d.toISOString().slice(0, 10));
  }

  return weeks;
}

function initWeekSelector() {
  const select = $("weekEndingSelect");
  if (!select) return;

  const weeks = buildRecentWeeksList();

  select.innerHTML = weeks
    .map((w) => `<option value="${w}">${formatDate(w)}</option>`)
    .join("");

  if (weeks.includes(state.weekEnding)) {
    select.value = state.weekEnding;
  } else {
    select.value = weeks[0];
    state.weekEnding = weeks[0];
  }

  select.addEventListener("change", async () => {
    state.weekEnding = select.value;
    setText("sidebarWeekEndingText", formatDate(select.value));
    if (state.currentRoute !== "trends") {
      await loadCurrentView();
    }
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
  const n = Number(value || 0);
  return Number.isFinite(n) ? Math.round(n).toLocaleString() : "0";
}

function formatDecimal(value, digits = 1) {
  const n = Number(value || 0);
  return Number.isFinite(n) ? n.toFixed(digits) : Number(0).toFixed(digits);
}

function formatPercent(value, digits = 1) {
  const n = Number(value || 0);
  return `${Number.isFinite(n) ? n.toFixed(digits) : Number(0).toFixed(digits)}%`;
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
      meta: "This section is not fully wired yet.",
      status: "Flat",
      statusColor: "yellow"
    }
  ];
}

function getPreviousWeekEnding(weekEnding) {
  return addDays(weekEnding, -7);
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

function findTrendsNavLink() {
  return Array.from(document.querySelectorAll(".nav-link")).find((el) =>
    ((el.dataset.route || "").trim() === "trends") ||
    ((el.textContent || "").trim().toLowerCase() === "trends")
  );
}

function showStandardView() {
  $("standardView")?.classList.remove("view-hidden");
  $("trendsView")?.classList.add("view-hidden");
}

function showTrendsView() {
  $("standardView")?.classList.add("view-hidden");
  $("trendsView")?.classList.remove("view-hidden");
}

async function loadDashboard() {
  showStandardView();

  const week = currentWeekEnding();

  setViewHeader(
    "Executive Summary",
    "Weekly companywide KPI overview with trends, target comparisons, and submission visibility."
  );
  setStatusPanelText("Loading dashboard...");
  setSubmissionStatusText("Draft");

  const result = await safeApiGet(
    `/api/dashboard?weekEnding=${encodeURIComponent(week)}`,
    { kpis: [] }
  );

  const kpis = Array.isArray(result?.kpis) ? result.kpis : [];
  renderKpiCards(kpis);
  setStatusPanelText("Executive dashboard loaded.");
  setDebugPanel(result);
}

async function loadRegionPage(entity) {
  showStandardView();

  const week = currentWeekEnding();
  const previousWeek = getPreviousWeekEnding(week);

  setViewHeader(
    `${entity} Weekly View`,
    `Weekly operational performance for ${entity}.`
  );
  setStatusPanelText(`Loading ${entity} weekly data...`);

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

  const cards = buildRegionCards(currentValues, previousValues);
  renderCardsFromItems(cards);

  const sourceStatus = currentResponse?.status || "Draft";
  setSubmissionStatusText(sourceStatus);

  const hasAnyCurrentValue = Object.values(currentValues).some((v) => v !== null && v !== undefined && v !== "");
  setStatusPanelText(
    hasAnyCurrentValue
      ? `${entity} weekly data loaded.`
      : `${entity} has no saved data for ${formatDate(week)}.`
  );

  setDebugPanel({
    current: currentResponse,
    previous: previousResponse
  });
}

async function loadSharedPage(page) {
  showStandardView();

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
    setSubmissionStatusText("Draft");
    setStatusPanelText(`${page} is not fully wired yet.`);
    setDebugPanel({ page, message: "Not wired yet." });
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

  setSubmissionStatusText(currentResponse?.status || "Draft");

  const hasAnyCurrentValue = Object.values(currentValues).some((v) => v !== null && v !== undefined && v !== "");
  setStatusPanelText(
    hasAnyCurrentValue
      ? `${page} weekly data loaded.`
      : `${page} has no saved data for ${formatDate(week)}.`
  );

  setDebugPanel({
    current: currentResponse,
    previous: previousResponse
  });
}

function syncTrendsEntityOptions() {
  const select = $("trendsEntitySelect");
  if (!select) return;

  const entities = isAdmin() ? REGION_KEYS : [state.currentRegion];
  const existingValue = state.trendsEntity || entities[0];

  select.innerHTML = entities
    .map((entity) => `<option value="${entity}">${entity}</option>`)
    .join("");

  if (entities.includes(existingValue)) {
    select.value = existingValue;
  } else {
    select.value = entities[0];
    state.trendsEntity = entities[0];
  }

  select.disabled = !isAdmin();
}

function syncTrendsControlsFromState() {
  const modeEl = $("trendsRangeMode");
  const limitEl = $("trendsLimit");
  const startEl = $("trendsStartDate");
  const endEl = $("trendsEndDate");

  if (modeEl) modeEl.value = state.trendsMode;
  if (limitEl) limitEl.value = String(state.trendsLimit);

  if (!state.trendsEndDate) {
    state.trendsEndDate = currentWeekEnding();
  }

  if (!state.trendsStartDate) {
    state.trendsStartDate = addDays(state.trendsEndDate, -(state.trendsLimit * 7));
  }

  if (startEl) startEl.value = state.trendsStartDate;
  if (endEl) endEl.value = state.trendsEndDate;

  syncTrendsRangeUi();
}

function syncTrendsRangeUi() {
  const recentMode = state.trendsMode === "recent";

  $("trendsLimitWrap")?.classList.toggle("toolbar-hidden", !recentMode);
  $("trendsStartWrap")?.classList.toggle("toolbar-hidden", recentMode);
  $("trendsEndWrap")?.classList.toggle("toolbar-hidden", recentMode);
}

function forceOpenNativeDatePicker(input) {
  if (!input) return;

  const tryOpen = () => {
    if (typeof input.showPicker === "function") {
      try {
        input.showPicker();
      } catch {
        input.focus();
      }
    } else {
      input.focus();
    }
  };

  input.addEventListener("click", tryOpen);
  input.addEventListener("focus", () => {
    if (document.activeElement === input) {
      tryOpen();
    }
  });
}

function wireDateInputs() {
  forceOpenNativeDatePicker($("trendsStartDate"));
  forceOpenNativeDatePicker($("trendsEndDate"));
}

function buildTrendsCards(items) {
  const latest = items.length ? items[items.length - 1] : null;
  const previous = items.length > 1 ? items[items.length - 2] : null;

  const getDiffMeta = (current, prior, format = "whole") => {
    const c = Number(current || 0);
    const p = Number(prior || 0);
    const diff = c - p;

    if (format === "percent1") {
      return `${diff >= 0 ? "+" : ""}${diff.toFixed(1)} pts vs prior`;
    }

    return `${diff >= 0 ? "+" : ""}${Math.round(diff).toLocaleString()} vs prior`;
  };

  if (!latest) {
    return [
      {
        label: "Weeks Loaded",
        value: "0",
        meta: "No data in selected range",
        status: "Flat",
        statusColor: "yellow"
      }
    ];
  }

  return [
    {
      label: "Weeks Loaded",
      value: String(items.length),
      meta: state.trendsMode === "recent" ? `Recent ${state.trendsLimit} weeks` : `${state.trendsStartDate} to ${state.trendsEndDate}`,
      status: "Flat",
      statusColor: "yellow"
    },
    {
      label: "Latest Visits",
      value: formatWhole(latest.totalVisits),
      meta: getDiffMeta(latest.totalVisits, previous?.totalVisits, "whole"),
      status: inferTrend(latest.totalVisits, previous?.totalVisits, "up").status,
      statusColor: inferTrend(latest.totalVisits, previous?.totalVisits, "up").statusColor
    },
    {
      label: "Latest New Patients",
      value: formatWhole(latest.npActual),
      meta: getDiffMeta(latest.npActual, previous?.npActual, "whole"),
      status: inferTrend(latest.npActual, previous?.npActual, "up").status,
      statusColor: inferTrend(latest.npActual, previous?.npActual, "up").statusColor
    },
    {
      label: "Latest Calls",
      value: formatWhole(latest.totalCalls),
      meta: getDiffMeta(latest.totalCalls, previous?.totalCalls, "whole"),
      status: inferTrend(latest.totalCalls, previous?.totalCalls, "up").status,
      statusColor: inferTrend(latest.totalCalls, previous?.totalCalls, "up").statusColor
    },
    {
      label: "Latest Abandonment",
      value: formatPercent(latest.abandonmentRate, 1),
      meta: getDiffMeta(latest.abandonmentRate, previous?.abandonmentRate, "percent1"),
      status: inferTrend(latest.abandonmentRate, previous?.abandonmentRate, "down").status,
      statusColor: inferTrend(latest.abandonmentRate, previous?.abandonmentRate, "down").statusColor
    },
    {
      label: "Latest Cash",
      value: `$${formatWhole(latest.cashActual)}`,
      meta: getDiffMeta(latest.cashActual, previous?.cashActual, "whole"),
      status: inferTrend(latest.cashActual, previous?.cashActual, "up").status,
      statusColor: inferTrend(latest.cashActual, previous?.cashActual, "up").statusColor
    }
  ];
}

function renderTrendsTable(items) {
  const wrap = $("trendsTableWrap");
  if (!wrap) return;

  if (!Array.isArray(items) || !items.length) {
    wrap.innerHTML = `<div>No trend data found for this selection.</div>`;
    return;
  }

  const rows = items
    .slice()
    .reverse()
    .map((item) => `
      <tr>
        <td>${item.weekEnding || ""}</td>
        <td>${item.entity || ""}</td>
        <td>${formatWhole(item.totalVisits)}</td>
        <td>${formatDecimal(item.visitsPerDay, 1)}</td>
        <td>${formatWhole(item.npActual)}</td>
        <td>${formatWhole(item.surgeryActual)}</td>
        <td>${formatWhole(item.totalCalls)}</td>
        <td>${formatPercent(item.abandonmentRate, 1)}</td>
        <td>${formatPercent(item.answeredCallToNpConversion, 1)}</td>
        <td>$${formatWhole(item.cashActual)}</td>
        <td>${item.status || "Draft"}</td>
        <td>
          ${
            isAdmin()
              ? `
              <div class="trends-actions">
                <button class="btn btn-secondary btn-small" type="button" data-action="override" data-entity="${item.entity}" data-week="${item.weekEnding}">Override</button>
                <button class="btn btn-secondary btn-small" type="button" data-action="delete" data-entity="${item.entity}" data-week="${item.weekEnding}">Delete</button>
              </div>
            `
              : ""
          }
        </td>
      </tr>
    `)
    .join("");

  wrap.innerHTML = `
    <table class="trends-table">
      <thead>
        <tr>
          <th>Week Ending</th>
          <th>Entity</th>
          <th>Visits</th>
          <th>Visits / Day</th>
          <th>New Patients</th>
          <th>Surgical Cases</th>
          <th>Calls</th>
          <th>Abandonment</th>
          <th>Call to NP %</th>
          <th>Cash</th>
          <th>Status</th>
          <th>Actions</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  `;

  if (isAdmin()) {
    wrap.querySelectorAll("[data-action='override']").forEach((btn) => {
      btn.addEventListener("click", async () => {
        await openOverrideEntry(btn.dataset.entity, btn.dataset.week);
      });
    });

    wrap.querySelectorAll("[data-action='delete']").forEach((btn) => {
      btn.addEventListener("click", async () => {
        await deleteTrendEntry(btn.dataset.entity, btn.dataset.week);
      });
    });
  }
}

async function loadTrendsView() {
  showTrendsView();

  setViewHeader(
    "Trends",
    "Review historical region performance across recent weeks or a custom date range."
  );
  setStatusPanelText("Loading trends...");
  setSubmissionStatusText("Historical");

  syncTrendsControlsFromState();

  const entity = state.trendsEntity || state.currentRegion;
  let url = `/api/trends?entity=${encodeURIComponent(entity)}`;

  if (state.trendsMode === "custom") {
    if (state.trendsStartDate) url += `&startDate=${encodeURIComponent(state.trendsStartDate)}`;
    if (state.trendsEndDate) url += `&endDate=${encodeURIComponent(state.trendsEndDate)}`;
  } else {
    url += `&limit=${encodeURIComponent(state.trendsLimit)}`;
  }

  const result = await safeApiGet(url, { items: [] });
  const items = Array.isArray(result?.items) ? result.items : [];

  renderCardsFromItems(buildTrendsCards(items));
  renderTrendsTable(items);
  setStatusPanelText(`Loaded ${items.length} trend row${items.length === 1 ? "" : "s"}.`);
  setDebugPanel(result);
}

async function openOverrideEntry(entity, weekEnding) {
  state.currentRoute = "region";
  state.currentRegion = entity;
  state.weekEnding = weekEnding;

  const weekSelect = $("weekEndingSelect");
  if (weekSelect) {
    const exists = Array.from(weekSelect.options).some((opt) => opt.value === weekEnding);
    if (!exists) {
      const option = document.createElement("option");
      option.value = weekEnding;
      option.textContent = formatDate(weekEnding);
      weekSelect.appendChild(option);
    }
    weekSelect.value = weekEnding;
  }

  setText("sidebarWeekEndingText", formatDate(weekEnding));
  activateNav(findNavLinkForRegion(entity));
  await loadRegionPage(entity);
  setStatusPanelText(`Admin override mode for ${entity} on ${formatDate(weekEnding)}.`);
}

async function deleteTrendEntry(entity, weekEnding) {
  if (!isAdmin()) {
    alert("Admin only.");
    return;
  }

  const confirmed = window.confirm(`Delete ${entity} for ${weekEnding}? This cannot be undone.`);
  if (!confirmed) return;

  const result = await apiPost("/api/delete-week", {
    entity,
    weekEnding
  });

  if (!result || result.error) {
    alert("Delete failed.");
    setDebugPanel(result || { error: "Delete failed" });
    return;
  }

  await loadTrendsView();
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
      await loadTrendsView();
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
  if (text.toLowerCase() === "trends") return "trends";
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
      state.trendsEntity = state.currentRegion;
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
    setDebugPanel(result || { error: "Save failed" });
    return;
  }

  setDebugPanel(result);
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
    setDebugPanel(result || { error: "Save failed" });
    return;
  }

  setDebugPanel(result);
  alert("Saved successfully.");
  await loadCurrentView();
}

function initButtons() {
  const saveBtn = $("saveButton");
  const submitBtn = $("submitWeekButton");
  const openTrendsBtn = $("openTrendsButton");
  const loadTrendsBtn = $("loadTrendsButton");

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

  if (openTrendsBtn) {
    openTrendsBtn.addEventListener("click", async () => {
      state.currentRoute = "trends";
      activateNav(findTrendsNavLink());
      await loadCurrentView();
    });
  }

  if (loadTrendsBtn) {
    loadTrendsBtn.addEventListener("click", async () => {
      state.currentRoute = "trends";
      await loadCurrentView();
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
      state.trendsEntity = state.currentRegion;

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

function initTrendsControls() {
  const entityEl = $("trendsEntitySelect");
  const modeEl = $("trendsRangeMode");
  const limitEl = $("trendsLimit");
  const startEl = $("trendsStartDate");
  const endEl = $("trendsEndDate");

  syncTrendsEntityOptions();

  if (!state.trendsEndDate) {
    state.trendsEndDate = currentWeekEnding();
  }
  if (!state.trendsStartDate) {
    state.trendsStartDate = addDays(state.trendsEndDate, -(state.trendsLimit * 7));
  }

  syncTrendsControlsFromState();
  wireDateInputs();

  if (entityEl) {
    entityEl.addEventListener("change", async () => {
      state.trendsEntity = entityEl.value;
      if (state.currentRoute === "trends") {
        await loadTrendsView();
      }
    });
  }

  if (modeEl) {
    modeEl.addEventListener("change", async () => {
      state.trendsMode = modeEl.value;

      if (state.trendsMode === "custom") {
        if (!state.trendsEndDate) state.trendsEndDate = currentWeekEnding();
        if (!state.trendsStartDate) state.trendsStartDate = addDays(state.trendsEndDate, -(state.trendsLimit * 7));
      }

      syncTrendsControlsFromState();

      if (state.currentRoute === "trends") {
        await loadTrendsView();
      }
    });
  }

  if (limitEl) {
    limitEl.addEventListener("change", async () => {
      state.trendsLimit = Number(limitEl.value || 12);
      if (state.trendsMode === "recent" && state.currentRoute === "trends") {
        await loadTrendsView();
      }
    });
  }

  if (startEl) {
    startEl.addEventListener("change", async () => {
      state.trendsStartDate = startEl.value;
      if (state.trendsMode === "custom" && state.currentRoute === "trends") {
        await loadTrendsView();
      }
    });
  }

  if (endEl) {
    endEl.addEventListener("change", async () => {
      state.trendsEndDate = endEl.value;
      if (state.trendsMode === "custom" && state.currentRoute === "trends") {
        await loadTrendsView();
      }
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

  if (state.currentRoute === "trends") {
    activateNav(findTrendsNavLink());
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
    initTrendsControls();
    activateInitialNav();
    await loadCurrentView();
  } catch (err) {
    handleFatalError(err);
  }
}

document.addEventListener("DOMContentLoaded", init);
