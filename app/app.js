import {
  safeApiGet,
  apiPost,
  getDefaultWeekEnding,
  formatDate,
  handleFatalError
} from "./helpers.js";

const ADMIN_EMAILS = ["nperez@unitymsk.com", "tessa.kelley@spineone.com"];
const ENTITIES = ["LAOSS", "NES", "SpineOne", "MRO"];

const state = {
  authenticated: false,
  userDetails: "",
  role: "guest",
  entity: "None",
  isAdmin: false,
  currentView: "dashboard",
  selectedEntity: "LAOSS",
  weekEnding: getDefaultWeekEnding(),
  lastDashboardResult: null,
  lastExecutiveResult: null,
  lastTrendsResult: null,
  lastWeeklyResult: null
};

function $(id) {
  return document.getElementById(id);
}

function q(selector) {
  return document.querySelector(selector);
}

function qa(selector) {
  return Array.from(document.querySelectorAll(selector));
}

function normalizeText(value) {
  return String(value || "").trim();
}

function normalizeLower(value) {
  return normalizeText(value).toLowerCase();
}

function normalizeEmail(value) {
  return normalizeLower(value);
}

function safeJson(value) {
  try {
    return JSON.stringify(value, null, 2);
  } catch {
    return String(value);
  }
}

function emailIsAdmin(email) {
  return ADMIN_EMAILS.includes(normalizeEmail(email));
}

function setText(id, value) {
  const el = $(id);
  if (el) {
    el.textContent = value;
  }
}

function setHtml(id, value) {
  const el = $(id);
  if (el) {
    el.innerHTML = value;
  }
}

function show(el) {
  if (el) {
    el.style.display = "";
  }
}

function hide(el) {
  if (el) {
    el.style.display = "none";
  }
}

function setTopStatus(text) {
  const candidates = [
    $("headerStatusText"),
    $("topRightStatusText"),
    $("statusBadge"),
    q(".status-badge"),
    q(".top-right-status"),
    q(".header-status")
  ].filter(Boolean);

  candidates.forEach((el) => {
    el.textContent = text;
  });
}

function setHeaderMeta() {
  setText("signedInUserText", state.authenticated ? state.userDetails : "Not signed in");
  setText("signedInAsText", state.authenticated ? state.userDetails : "Not signed in");
  setText("assignedEntityText", state.entity || "None");
  setText("entityText", state.entity || "None");
  setText("roleText", state.role || "guest");

  const signIn = $("signInButton") || q("a[href*='/.auth/login']");
  const signOut = $("signOutButton") || q("a[href*='/.auth/logout']");

  if (signIn) {
    signIn.setAttribute("href", "/.auth/login/aad");
  }

  if (signOut) {
    signOut.setAttribute("href", "/.auth/logout");
  }

  if (state.authenticated) {
    hide(signIn);
    show(signOut);
  } else {
    show(signIn);
    hide(signOut);
  }
}

function setDebugOutput(value) {
  const text = typeof value === "string" ? value : safeJson(value);

  [
    "dashboardDebugOutput",
    "debugOutput",
    "executiveDebugOutput",
    "trendsDebugOutput"
  ].forEach((id) => {
    const el = $(id);
    if (el) {
      el.textContent = text;
    }
  });
}

function findPanelByHeadingText(label) {
  const headings = qa("h1, h2, h3, h4");
  const wanted = normalizeLower(label);

  for (const heading of headings) {
    if (normalizeLower(heading.textContent) === wanted) {
      const panel =
        heading.closest(".panel") ||
        heading.closest(".card") ||
        heading.closest("section") ||
        heading.parentElement;

      if (panel) {
        return panel;
      }
    }
  }

  return null;
}

function ensureRenderSlot(panel, slotClassName) {
  if (!panel) return null;

  let slot = panel.querySelector(`.${slotClassName}`);
  if (slot) return slot;

  slot = document.createElement("div");
  slot.className = slotClassName;
  slot.style.marginTop = "18px";
  panel.appendChild(slot);
  return slot;
}

function getDashboardPanel() {
  return findPanelByHeadingText("Dashboard");
}

function getEntityPerformancePanel() {
  return findPanelByHeadingText("Entity Performance");
}

function getAlertsPanel() {
  return findPanelByHeadingText("Alerts");
}

function getSnapshotPanel() {
  return findPanelByHeadingText("Entity Snapshot");
}

function formatWhole(value) {
  const n = Number(value || 0);
  return Number.isFinite(n) ? Math.round(n).toLocaleString() : "0";
}

function formatPercent(value, digits = 1) {
  const n = Number(value || 0);
  return `${Number.isFinite(n) ? n.toFixed(digits) : "0.0"}%`;
}

function formatValue(value, format) {
  if (format === "percent") {
    return formatPercent(value, 1);
  }
  return formatWhole(value);
}

function currentWeekEnding() {
  const candidates = [
    $("anchorWeekEnding"),
    $("anchorWeekEndingInput"),
    $("weekEnding"),
    $("weekEndingSelect"),
    q("input[type='date']")
  ].filter(Boolean);

  const control = candidates[0];
  const value = normalizeText(control?.value);

  return value || state.weekEnding || getDefaultWeekEnding();
}

function setWeekEndingControlValue(value) {
  const candidates = [
    $("anchorWeekEnding"),
    $("anchorWeekEndingInput"),
    $("weekEnding"),
    $("weekEndingSelect"),
    q("input[type='date']")
  ].filter(Boolean);

  const control = candidates[0];
  if (!control) return;

  if (!normalizeText(control.value)) {
    control.value = value;
  }
}

function currentCompareAgainst() {
  const candidates = [
    $("compareAgainst"),
    $("compareAgainstSelect"),
    q("select[id*='compare']"),
    q("select[name='compareAgainst']")
  ].filter(Boolean);

  return normalizeText(candidates[0]?.value) || "Prior Period";
}

function currentPeriod() {
  const candidates = [
    $("period"),
    $("periodSelect"),
    $("dashboardPeriodSelect"),
    q("select[id*='period']"),
    q("select[name='period']")
  ].filter(Boolean);

  return normalizeText(candidates[0]?.value) || "Current Week";
}

function currentEntityScope() {
  const candidates = [
    $("entityScope"),
    $("entityScopeSelect"),
    q("select[id*='scope']"),
    q("select[name='entityScope']")
  ].filter(Boolean);

  return normalizeText(candidates[0]?.value) || "All Entities";
}

function normalizeApiMe(result) {
  if (!result) return null;

  if (result.authenticated) {
    const roles = Array.isArray(result.roles) ? result.roles : [];
    return {
      authenticated: true,
      userDetails: result.userDetails || "",
      entity: result.entity || "",
      isAdmin: !!result.isAdmin || roles.some((r) => normalizeLower(r) === "admin"),
      roles
    };
  }

  if (result.user && result.access) {
    const roles = Array.isArray(result.user.roles) ? result.user.roles : [];
    return {
      authenticated: !!result.user.authenticated,
      userDetails: result.user.userDetails || result.access.email || "",
      entity: result.access.entity || "",
      isAdmin: !!result.access.isAdmin || normalizeLower(result.access.role) === "admin",
      roles
    };
  }

  return null;
}

function normalizeAuthMe(result) {
  const principal = result?.clientPrincipal;
  if (!principal || !principal.userId) return null;

  const roles = Array.isArray(principal.userRoles) ? principal.userRoles : [];
  return {
    authenticated: true,
    userDetails: principal.userDetails || principal.userId || "",
    entity: "",
    isAdmin: roles.some((r) => normalizeLower(r) === "admin"),
    roles
  };
}

async function resolveAuth() {
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
    setHeaderMeta();
    return;
  }

  state.authenticated = true;
  state.userDetails = me.userDetails || "Unknown User";
  state.isAdmin = !!me.isAdmin || emailIsAdmin(me.userDetails);
  state.role = state.isAdmin ? "admin" : "user";
  state.entity = state.isAdmin ? "Admin" : (me.entity || "LAOSS");

  setHeaderMeta();
}

function renderDashboardKpis(kpis) {
  const panel = getDashboardPanel();
  const slot = ensureRenderSlot(panel, "dashboard-kpi-grid");
  if (!slot) return;

  const items = Array.isArray(kpis) ? kpis : [];

  slot.style.display = "grid";
  slot.style.gridTemplateColumns = "repeat(auto-fit, minmax(180px, 1fr))";
  slot.style.gap = "14px";

  slot.innerHTML = items.map((item) => {
    const label = item?.label || item?.key || "Metric";
    const value = formatValue(item?.value, item?.format);
    const meta = item?.meta || "";
    return `
      <div style="background:rgba(255,255,255,0.04);border:1px solid rgba(255,255,255,0.08);border-radius:16px;padding:16px;">
        <div style="font-size:13px;opacity:.85;margin-bottom:8px;">${label}</div>
        <div style="font-size:30px;font-weight:700;line-height:1.1;margin-bottom:6px;">${value}</div>
        <div style="font-size:12px;opacity:.75;">${meta}</div>
      </div>
    `;
  }).join("");
}

function renderEntityPerformance(entities) {
  const panel = getEntityPerformancePanel();
  const slot = ensureRenderSlot(panel, "entity-performance-grid");
  if (!slot) return;

  const items = Array.isArray(entities) ? entities : [];

  slot.style.display = "grid";
  slot.style.gridTemplateColumns = "repeat(auto-fit, minmax(220px, 1fr))";
  slot.style.gap = "14px";

  slot.innerHTML = items.map((item) => {
    return `
      <div style="background:rgba(255,255,255,0.04);border:1px solid rgba(255,255,255,0.08);border-radius:16px;padding:16px;">
        <div style="display:flex;justify-content:space-between;align-items:center;gap:8px;margin-bottom:10px;">
          <div style="font-size:18px;font-weight:700;">${item.entity || "Unknown"}</div>
          <div style="font-size:12px;opacity:.75;text-transform:capitalize;">${item.status || "draft"}</div>
        </div>
        <div style="font-size:13px;opacity:.85;margin-bottom:6px;">Visits: <strong>${formatWhole(item.visitVolume)}</strong></div>
        <div style="font-size:13px;opacity:.85;margin-bottom:6px;">Calls: <strong>${formatWhole(item.callVolume)}</strong></div>
        <div style="font-size:13px;opacity:.85;margin-bottom:6px;">New Patients: <strong>${formatWhole(item.newPatients)}</strong></div>
        <div style="font-size:13px;opacity:.85;margin-bottom:6px;">No Show: <strong>${formatPercent(item.noShowRate)}</strong></div>
        <div style="font-size:13px;opacity:.85;">Cancel: <strong>${formatPercent(item.cancellationRate)}</strong></div>
      </div>
    `;
  }).join("");
}

function renderAlerts(alerts) {
  const panel = getAlertsPanel();
  const slot = ensureRenderSlot(panel, "alerts-list");
  if (!slot) return;

  const items = Array.isArray(alerts) ? alerts : [];

  slot.style.display = "grid";
  slot.style.gap = "10px";

  slot.innerHTML = items.map((item) => {
    const severity = normalizeLower(item.severity || "yellow");
    const accent =
      severity === "red" ? "#ef4444" :
      severity === "green" ? "#22c55e" :
      "#facc15";

    return `
      <div style="background:rgba(255,255,255,0.04);border:1px solid rgba(255,255,255,0.08);border-left:4px solid ${accent};border-radius:14px;padding:14px;">
        <div style="font-size:14px;font-weight:700;margin-bottom:4px;">${item.entity || "System"}</div>
        <div style="font-size:13px;opacity:.9;">${item.message || ""}</div>
      </div>
    `;
  }).join("");

  if (!items.length) {
    slot.innerHTML = `<div style="font-size:13px;opacity:.8;">No alerts.</div>`;
  }
}

function renderSnapshot(entities) {
  const panel = getSnapshotPanel();
  const slot = ensureRenderSlot(panel, "snapshot-table-wrap");
  if (!slot) return;

  const items = Array.isArray(entities) ? entities : [];

  slot.innerHTML = `
    <div style="overflow:auto;">
      <table style="width:100%;border-collapse:collapse;font-size:13px;">
        <thead>
          <tr>
            <th style="text-align:left;padding:10px;border-bottom:1px solid rgba(255,255,255,0.1);">Entity</th>
            <th style="text-align:left;padding:10px;border-bottom:1px solid rgba(255,255,255,0.1);">Status</th>
            <th style="text-align:right;padding:10px;border-bottom:1px solid rgba(255,255,255,0.1);">Visits</th>
            <th style="text-align:right;padding:10px;border-bottom:1px solid rgba(255,255,255,0.1);">Calls</th>
            <th style="text-align:right;padding:10px;border-bottom:1px solid rgba(255,255,255,0.1);">New Patients</th>
            <th style="text-align:right;padding:10px;border-bottom:1px solid rgba(255,255,255,0.1);">Abandoned %</th>
          </tr>
        </thead>
        <tbody>
          ${items.map((item) => `
            <tr>
              <td style="padding:10px;border-bottom:1px solid rgba(255,255,255,0.06);">${item.entity || ""}</td>
              <td style="padding:10px;border-bottom:1px solid rgba(255,255,255,0.06);text-transform:capitalize;">${item.status || "draft"}</td>
              <td style="padding:10px;border-bottom:1px solid rgba(255,255,255,0.06);text-align:right;">${formatWhole(item.visitVolume)}</td>
              <td style="padding:10px;border-bottom:1px solid rgba(255,255,255,0.06);text-align:right;">${formatWhole(item.callVolume)}</td>
              <td style="padding:10px;border-bottom:1px solid rgba(255,255,255,0.06);text-align:right;">${formatWhole(item.newPatients)}</td>
              <td style="padding:10px;border-bottom:1px solid rgba(255,255,255,0.06);text-align:right;">${formatPercent(item.abandonedCallRate)}</td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    </div>
  `;
}

async function loadDashboard() {
  state.currentView = "dashboard";
  state.weekEnding = currentWeekEnding();

  setTopStatus("Loading...");
  setDebugOutput("Loading dashboard...");

  const query = new URLSearchParams({
    weekEnding: state.weekEnding,
    period: currentPeriod(),
    compareAgainst: currentCompareAgainst(),
    entityScope: currentEntityScope()
  });

  const result = await safeApiGet(`/api/dashboard?${query.toString()}`, {
    ok: false,
    kpis: [],
    entities: [],
    alerts: []
  });

  state.lastDashboardResult = result;

  renderDashboardKpis(result.kpis || []);
  renderEntityPerformance(result.entities || []);
  renderAlerts(result.alerts || []);
  renderSnapshot(result.entities || []);
  setDebugOutput(result);
  setTopStatus("Ready");
}

async function loadExecutiveSummary() {
  state.currentView = "executive";
  state.weekEnding = currentWeekEnding();

  setTopStatus("Loading...");
  setDebugOutput("Loading executive summary...");

  const result = await safeApiGet(
    `/api/executive-summary?weekEnding=${encodeURIComponent(state.weekEnding)}`,
    { ok: false, kpis: [], entities: [], alerts: [] }
  );

  state.lastExecutiveResult = result;

  renderDashboardKpis(result.kpis || []);
  renderEntityPerformance(result.entities || []);
  renderAlerts(result.alerts || []);
  renderSnapshot(result.entities || []);
  setDebugOutput(result);
  setTopStatus("Ready");
}

async function loadTrends() {
  state.currentView = "trends";
  state.weekEnding = currentWeekEnding();

  setTopStatus("Loading...");
  setDebugOutput("Loading trends...");

  const entityScope = currentEntityScope();
  const entity = entityScope && entityScope !== "All Entities" ? entityScope : state.selectedEntity;

  const query = new URLSearchParams({
    entity,
    weekEnding: state.weekEnding
  });

  const result = await safeApiGet(`/api/trends?${query.toString()}`, { ok: false, items: [] });

  state.lastTrendsResult = result;
  setDebugOutput(result);
  setTopStatus("Ready");
}

async function loadWeeklyEntryView() {
  state.currentView = "weekly";
  state.weekEnding = currentWeekEnding();

  setTopStatus("Loading...");
  setDebugOutput("Loading weekly entry data...");

  const entityScope = currentEntityScope();
  const entity = entityScope && entityScope !== "All Entities" ? entityScope : state.selectedEntity;

  const current = await safeApiGet(
    `/api/weekly?entity=${encodeURIComponent(entity)}&weekEnding=${encodeURIComponent(state.weekEnding)}`,
    { ok: false, values: {} }
  );

  state.lastWeeklyResult = current;
  setDebugOutput(current);
  setTopStatus("Ready");
}

function navLinks() {
  return qa(".nav-link");
}

function activateNavByText(text) {
  const target = normalizeLower(text);

  navLinks().forEach((link) => {
    const isMatch = normalizeLower(link.textContent) === target;
    link.classList.toggle("active", isMatch);
  });
}

function routeFromLink(link) {
  const text = normalizeLower(link?.textContent);

  if (text === "dashboard") return "dashboard";
  if (text === "weekly entry") return "weekly";
  if (text === "executive summary") return "executive";
  if (text === "trends") return "trends";
  if (text === "admin import") return "admin-import";

  return "dashboard";
}

async function handleNav(route) {
  if (route === "admin-import") {
    window.location.href = "./admin-import.html";
    return;
  }

  if (route === "weekly") {
    activateNavByText("weekly entry");
    await loadWeeklyEntryView();
    return;
  }

  if (route === "executive") {
    activateNavByText("executive summary");
    await loadExecutiveSummary();
    return;
  }

  if (route === "trends") {
    activateNavByText("trends");
    await loadTrends();
    return;
  }

  activateNavByText("dashboard");
  await loadDashboard();
}

function bindNavigation() {
  navLinks().forEach((link) => {
    link.addEventListener("click", async (event) => {
      const href = normalizeText(link.getAttribute("href"));
      if (href && href.includes("admin-import.html")) {
        return;
      }

      event.preventDefault();

      try {
        await handleNav(routeFromLink(link));
      } catch (error) {
        setTopStatus("Error");
        setDebugOutput({
          where: "navigation",
          message: error?.message || String(error),
          stack: error?.stack || null
        });
        handleFatalError(error);
      }
    });
  });
}

function bindDashboardButton() {
  const button =
    $("loadDashboardButton") ||
    $("loadDashboardBtn") ||
    qa("button").find((btn) => normalizeLower(btn.textContent) === "load dashboard");

  if (!button) return;

  button.addEventListener("click", async () => {
    try {
      await loadDashboard();
    } catch (error) {
      setTopStatus("Error");
      setDebugOutput({
        where: "load dashboard button",
        message: error?.message || String(error),
        stack: error?.stack || null
      });
      handleFatalError(error);
    }
  });
}

function bindControlChanges() {
  const controls = [
    $("anchorWeekEnding"),
    $("anchorWeekEndingInput"),
    $("weekEnding"),
    $("weekEndingSelect"),
    $("period"),
    $("periodSelect"),
    $("dashboardPeriodSelect"),
    $("compareAgainst"),
    $("compareAgainstSelect"),
    $("entityScope"),
    $("entityScopeSelect")
  ].filter(Boolean);

  controls.forEach((control) => {
    control.addEventListener("change", () => {
      state.weekEnding = currentWeekEnding();

      const entityScope = currentEntityScope();
      if (entityScope && entityScope !== "All Entities" && ENTITIES.includes(entityScope)) {
        state.selectedEntity = entityScope;
      }
    });
  });
}

function bindAdminImportLinkVisibility() {
  const adminLink = navLinks().find((link) => normalizeLower(link.textContent) === "admin import");
  if (!adminLink) return;

  if (state.isAdmin) {
    show(adminLink);
  } else {
    hide(adminLink);
  }
}

async function init() {
  try {
    setTopStatus("Loading...");
    setDebugOutput("App starting...");
    setWeekEndingControlValue(state.weekEnding);

    await resolveAuth();
    bindAdminImportLinkVisibility();
    bindNavigation();
    bindDashboardButton();
    bindControlChanges();

    activateNavByText("dashboard");
    await loadDashboard();
  } catch (error) {
    setTopStatus("Error");
    setDebugOutput({
      where: "init",
      message: error?.message || String(error),
      stack: error?.stack || null
    });
    handleFatalError(error);
  }
}

document.addEventListener("DOMContentLoaded", init);
