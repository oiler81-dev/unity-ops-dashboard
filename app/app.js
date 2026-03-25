const ENTITIES = ["LAOSS", "NES", "SpineOne", "MRO"];

const ENTITY_BRANDING = {
  LAOSS: {
    label: "LAOSS",
    fullName: "Los Angeles Orthopedic Surgery Specialists",
    logo: "/assets/logos/laoss.png",
    accent: "#F28C28",
    accentSoft: "rgba(242,140,40,0.12)",
    accentBorder: "rgba(242,140,40,0.35)"
  },
  NES: {
    label: "NES",
    fullName: "Northwest Extremity Specialists",
    logo: "/assets/logos/nes.png",
    accent: "#2E5B88",
    accentSoft: "rgba(46,91,136,0.12)",
    accentBorder: "rgba(46,91,136,0.35)"
  },
  SpineOne: {
    label: "SpineOne",
    fullName: "SpineOne",
    logo: "/assets/logos/spineone.png",
    accent: "#5A6F95",
    accentSoft: "rgba(90,111,149,0.12)",
    accentBorder: "rgba(90,111,149,0.35)"
  },
  MRO: {
    label: "MRO",
    fullName: "Midland & Riverside Orthopedics",
    logo: "/assets/logos/mro.png",
    accent: "#6B7E99",
    accentSoft: "rgba(107,126,153,0.12)",
    accentBorder: "rgba(107,126,153,0.35)"
  }
};

let currentUser = null;
let currentWeekData = null;

async function parseApiResponse(res) {
  const text = await res.text();
  let data = null;

  try {
    data = text ? JSON.parse(text) : null;
  } catch {
    throw new Error(text || "Invalid response");
  }

  if (!res.ok) {
    throw new Error(data?.details || data?.error || "Request failed");
  }

  return data;
}

async function apiGet(url) {
  const res = await fetch(url, {
    method: "GET",
    headers: { Accept: "application/json" }
  });
  return parseApiResponse(res);
}

async function apiPost(url, payload) {
  const res = await fetch(url, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Accept: "application/json"
    },
    body: JSON.stringify(payload)
  });
  return parseApiResponse(res);
}

function setActiveNav(buttonId) {
  [
    "navDashboardBtn",
    "navEntryBtn",
    "navExecutiveBtn",
    "navTrendsBtn",
    "navImportBtn"
  ].forEach((id) => {
    const btn = document.getElementById(id);
    if (btn) btn.classList.toggle("activeNav", id === buttonId);
  });
}

function setStatus(message, isError = false) {
  const el = document.getElementById("statusMessage");
  if (!el) return;
  el.textContent = message;
  el.style.color = isError ? "#ff8a8a" : "#7CFC98";
}

function setDebug(data) {
  const el = document.getElementById("debugOutput");
  if (!el) return;
  el.textContent = typeof data === "string" ? data : JSON.stringify(data, null, 2);
}

function setExecutiveDebug(data) {
  const el = document.getElementById("executiveDebugOutput");
  if (!el) return;
  el.textContent = typeof data === "string" ? data : JSON.stringify(data, null, 2);
}

function setTrendsDebug(data) {
  const el = document.getElementById("trendsDebugOutput");
  if (!el) return;
  el.textContent = typeof data === "string" ? data : JSON.stringify(data, null, 2);
}

function setImportStatus(message, isError = false) {
  const el = document.getElementById("importStatusMessage");
  if (!el) return;
  el.textContent = message;
  el.style.color = isError ? "#ff8a8a" : "#7CFC98";
}

function setImportDebug(data) {
  const el = document.getElementById("importDebugOutput");
  if (!el) return;
  el.textContent = typeof data === "string" ? data : JSON.stringify(data, null, 2);
}

function setDashboardDebug(data) {
  const el = document.getElementById("dashboardDebugOutput");
  if (!el) return;
  el.textContent = typeof data === "string" ? data : JSON.stringify(data, null, 2);
}

function setDashboardRangeSummary(text) {
  const el = document.getElementById("dashboardRangeSummary");
  if (el) el.textContent = text;
}

function setDashboardBenchmarkBanner(text = "", show = false) {
  const el = document.getElementById("dashboardBenchmarkBanner");
  if (!el) return;
  el.style.display = show ? "" : "none";
  el.textContent = text;
}

function getDefaultWeekEnding() {
  const d = new Date();
  const diff = (5 - d.getDay() + 7) % 7;
  d.setDate(d.getDate() + diff);
  return d.toISOString().slice(0, 10);
}

function getDateWeeksAgo(weeksAgo, fromDate = null) {
  const d = fromDate ? new Date(`${fromDate}T12:00:00Z`) : new Date();
  d.setUTCDate(d.getUTCDate() - weeksAgo * 7);
  return d.toISOString().slice(0, 10);
}

function getPreviousWeekEnding(weekEnding) {
  const d = new Date(`${weekEnding}T12:00:00Z`);
  d.setUTCDate(d.getUTCDate() - 7);
  return d.toISOString().slice(0, 10);
}

function normalizeNumber(value) {
  if (value === null || value === undefined || value === "") return 0;
  const n = Number(value);
  return Number.isFinite(n) ? n : 0;
}

function formatPercent(value, digits = 1) {
  return `${normalizeNumber(value).toFixed(digits)}%`;
}

function renderUser(userData) {
  document.getElementById("userInfo").innerText =
    `${userData.user.userDetails} (${userData.access.role})`;
}

function setupEntityDropdown(userData) {
  const select = document.getElementById("entitySelect");
  select.innerHTML = "";

  if (userData.access.isAdmin) {
    ENTITIES.forEach((entity) => {
      const option = document.createElement("option");
      option.value = entity;
      option.textContent = entity;
      select.appendChild(option);
    });
  } else {
    const option = document.createElement("option");
    option.value = userData.access.entity;
    option.textContent = userData.access.entity;
    select.appendChild(option);
    select.disabled = true;
  }
}

function setupTrendsEntityDropdown(userData) {
  const select = document.getElementById("trendsEntitySelect");
  select.innerHTML = "";

  if (userData.access.isAdmin) {
    ENTITIES.forEach((entity) => {
      const option = document.createElement("option");
      option.value = entity;
      option.textContent = entity;
      select.appendChild(option);
    });
  } else {
    const option = document.createElement("option");
    option.value = userData.access.entity;
    option.textContent = userData.access.entity;
    select.appendChild(option);
    select.disabled = true;
  }
}

function getSelectedEntity() {
  return document.getElementById("entitySelect").value;
}

function getSelectedTrendsEntity() {
  return document.getElementById("trendsEntitySelect").value;
}

function getBranding(entity) {
  return ENTITY_BRANDING[entity] || {
    label: entity,
    fullName: entity,
    logo: "",
    accent: "#3b82f6",
    accentSoft: "rgba(59,130,246,0.12)",
    accentBorder: "rgba(59,130,246,0.35)"
  };
}

function renderContextBrand(containerId, entity) {
  const container = document.getElementById(containerId);
  if (!container) return;

  if (!entity) {
    container.innerHTML = "";
    return;
  }

  const brand = getBranding(entity);

  container.innerHTML = `
    <div class="contextBrandLogo" style="box-shadow: inset 0 0 0 1px ${brand.accentBorder};">
      <img src="${brand.logo}" alt="${brand.label}" />
    </div>
    <div class="contextBrandText">
      <strong style="display:block;color:${brand.accent};">${brand.label}</strong>
      <span>${brand.fullName}</span>
    </div>
  `;
}

function renderForm() {
  const fields = [
    { key: "visitVolume", label: "Visit Volume" },
    { key: "callVolume", label: "Call Volume" },
    { key: "newPatients", label: "New Patients" },
    { key: "noShowRate", label: "No Show Rate" },
    { key: "cancellationRate", label: "Cancellation Rate" },
    { key: "abandonedCallRate", label: "Abandoned Call Rate" }
  ];

  const container = document.getElementById("kpiForm");
  container.innerHTML = "";

  fields.forEach((field) => {
    const div = document.createElement("div");
    div.innerHTML = `
      <label for="${field.key}">${field.label}</label>
      <input type="number" id="${field.key}" step="any" />
    `;
    container.appendChild(div);
  });
}

function mapWeeklyValuesToFormData(values) {
  return {
    visitVolume: values?.visitVolume ?? values?.totalVisits ?? "",
    callVolume: values?.callVolume ?? values?.totalCalls ?? "",
    newPatients: values?.newPatients ?? values?.npActual ?? "",
    noShowRate: values?.noShowRate ?? "",
    cancellationRate: values?.cancellationRate ?? "",
    abandonedCallRate: values?.abandonedCallRate ?? values?.abandonmentRate ?? ""
  };
}

function setFormValues(data) {
  const mapped = mapWeeklyValuesToFormData(data || {});
  const keys = [
    "visitVolume",
    "callVolume",
    "newPatients",
    "noShowRate",
    "cancellationRate",
    "abandonedCallRate"
  ];

  keys.forEach((key) => {
    const input = document.getElementById(key);
    if (!input) return;
    input.value = mapped[key] !== null && mapped[key] !== undefined ? mapped[key] : "";
  });
}

function getFormValues() {
  return {
    visitVolume: document.getElementById("visitVolume").value,
    callVolume: document.getElementById("callVolume").value,
    newPatients: document.getElementById("newPatients").value,
    noShowRate: document.getElementById("noShowRate").value,
    cancellationRate: document.getElementById("cancellationRate").value,
    abandonedCallRate: document.getElementById("abandonedCallRate").value
  };
}

function updateButtonState() {
  const saveBtn = document.getElementById("saveBtn");
  const submitBtn = document.getElementById("submitBtn");
  const approveBtn = document.getElementById("approveBtn");

  const status = String(currentWeekData?.status || "draft").toLowerCase();
  const isAdmin = !!currentUser?.access?.isAdmin;

  saveBtn.disabled = status === "approved" && !isAdmin;
  submitBtn.disabled = status === "submitted" || status === "approved";
  approveBtn.disabled = !isAdmin || status !== "submitted";
}

async function loadWeek() {
  const weekEnding = document.getElementById("weekEnding").value;
  const entity = getSelectedEntity();

  renderContextBrand("entryContextBrand", entity);

  setStatus("Loading...");
  const result = await apiGet(
    `/api/weekly?weekEnding=${encodeURIComponent(weekEnding)}&entity=${encodeURIComponent(entity)}`
  );

  currentWeekData = result;
  setFormValues(result.values || result.data || {});
  updateButtonState();

  setStatus(`Loaded ${entity} for ${weekEnding} (${result.status || "draft"})`);
  setDebug(result);
}

async function saveWeek() {
  const payload = {
    weekEnding: document.getElementById("weekEnding").value,
    entity: getSelectedEntity(),
    data: getFormValues()
  };

  setStatus("Saving...");
  setDebug(payload);

  const result = await apiPost("/api/weekly-save", payload);

  setStatus(result.message || "Saved successfully");
  setDebug(result);

  await loadWeek();
}

async function submitWeek() {
  const payload = {
    weekEnding: document.getElementById("weekEnding").value,
    entity: getSelectedEntity()
  };

  setStatus("Submitting...");
  setDebug(payload);

  const result = await apiPost("/api/submit-week", payload);

  setStatus(result.message || "Submitted successfully");
  setDebug(result);

  await loadWeek();
}

async function approveWeek() {
  const payload = {
    weekEnding: document.getElementById("weekEnding").value,
    entity: getSelectedEntity()
  };

  setStatus("Approving...");
  setDebug(payload);

  const result = await apiPost("/api/approve-week", payload);

  setStatus(result.message || "Approved successfully");
  setDebug(result);

  await loadWeek();
}

async function deleteWeek(entity, weekEnding) {
  if (!confirm(`Delete ${entity} for ${weekEnding}? This cannot be undone.`)) {
    return;
  }

  const result = await apiPost("/api/delete-week", { entity, weekEnding });
  setTrendsDebug(result);
  await loadTrends();
}

async function openOverride(entity, weekEnding) {
  showEntryView();
  document.getElementById("entitySelect").value = entity;
  document.getElementById("weekEnding").value = weekEnding;
  renderContextBrand("entryContextBrand", entity);
  await loadWeek();
  setStatus(`Admin override mode: ${entity} ${weekEnding}`);
}

function hideAllViews() {
  document.getElementById("dashboardView").style.display = "none";
  document.getElementById("entryView").style.display = "none";
  document.getElementById("executiveView").style.display = "none";
  document.getElementById("trendsView").style.display = "none";
  document.getElementById("importView").style.display = "none";
}

function showDashboardView() {
  hideAllViews();
  document.getElementById("dashboardView").style.display = "";
  setActiveNav("navDashboardBtn");
}

function showEntryView() {
  hideAllViews();
  document.getElementById("entryView").style.display = "";
  setActiveNav("navEntryBtn");
  renderContextBrand("entryContextBrand", getSelectedEntity());
}

function showExecutiveView() {
  hideAllViews();
  document.getElementById("executiveView").style.display = "";
  setActiveNav("navExecutiveBtn");
}

function showTrendsView() {
  hideAllViews();
  document.getElementById("trendsView").style.display = "";
  setActiveNav("navTrendsBtn");
  syncTrendsRangeUi();
  renderContextBrand("trendsContextBrand", getSelectedTrendsEntity());
}

function showImportView() {
  hideAllViews();
  document.getElementById("importView").style.display = "";
  setActiveNav("navImportBtn");
}

function renderExecutiveCards(summary) {
  const cards = document.getElementById("executiveCards");
  cards.innerHTML = "";

  const regions = summary.regions || [];
  const avg = (key) => {
    if (!regions.length) return 0;
    const total = regions.reduce((sum, r) => sum + Number(r[key] || 0), 0);
    return (total / regions.length).toFixed(1);
  };

  const cardData = [
    { label: "Approved Regions", value: summary.entityCount || 0 },
    { label: "Visit Volume", value: summary.totals?.visitVolume || 0 },
    { label: "Call Volume", value: summary.totals?.callVolume || 0 },
    { label: "New Patients", value: summary.totals?.newPatients || 0 },
    { label: "Avg No Show %", value: `${avg("noShowRate")}%` },
    { label: "Avg Cancel %", value: `${avg("cancellationRate")}%` },
    { label: "Avg Abandoned %", value: `${avg("abandonedCallRate")}%` }
  ];

  cardData.forEach((item) => {
    const div = document.createElement("div");
    div.className = "summaryCard";
    div.innerHTML = `<h3>${item.label}</h3><div class="value">${item.value}</div>`;
    cards.appendChild(div);
  });
}

function renderExecutiveRegions(summary) {
  const container = document.getElementById("executiveRegions");

  if (!summary.regions || !summary.regions.length) {
    container.innerHTML = "<p>No approved regions found for this week.</p>";
    return;
  }

  const rows = summary.regions
    .map((r) => `
      <tr>
        <td>${r.entity}</td>
        <td>${r.visitVolume}</td>
        <td>${r.callVolume}</td>
        <td>${r.newPatients}</td>
        <td>${r.noShowRate}%</td>
        <td>${r.cancellationRate}%</td>
        <td>${r.abandonedCallRate}%</td>
        <td>${r.status}</td>
      </tr>
    `)
    .join("");

  container.innerHTML = `
    <table class="regionTable">
      <thead>
        <tr>
          <th>Entity</th>
          <th>Visit</th>
          <th>Calls</th>
          <th>New</th>
          <th>No Show</th>
          <th>Cancel</th>
          <th>Abandoned</th>
          <th>Status</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  `;
}

async function loadExecutiveSummary() {
  const weekEnding = document.getElementById("executiveWeekEnding").value;
  const result = await apiGet(`/api/executive-summary?weekEnding=${encodeURIComponent(weekEnding)}`);
  renderExecutiveCards(result);
  renderExecutiveRegions(result);
  setExecutiveDebug(result);
}

function renderTrendsCards(result) {
  const cards = document.getElementById("trendsCards");
  cards.innerHTML = "";

  const items = result.items || [];
  const latest = items.length ? items[0] : null;
  const previous = items.length > 1 ? items[1] : null;

  const formatDelta = (current, prior) => {
    if (current == null) return "-";
    if (prior == null) return `${current}`;
    const diff = Number(current) - Number(prior);
    return `${current} (${diff >= 0 ? "+" : ""}${diff})`;
  };

  const cardData = [
    { label: "Weeks Loaded", value: items.length },
    {
      label: "Latest Visit Volume",
      value: latest ? formatDelta(latest.visitVolume, previous?.visitVolume) : "-"
    },
    {
      label: "Latest Call Volume",
      value: latest ? formatDelta(latest.callVolume, previous?.callVolume) : "-"
    },
    {
      label: "Latest New Patients",
      value: latest ? formatDelta(latest.newPatients, previous?.newPatients) : "-"
    }
  ];

  cardData.forEach((item) => {
    const div = document.createElement("div");
    div.className = "summaryCard";
    div.innerHTML = `<h3>${item.label}</h3><div class="value">${item.value}</div>`;
    cards.appendChild(div);
  });
}

function renderTrendsTable(result) {
  const wrap = document.getElementById("trendsTableWrap");
  const items = result.items || [];

  if (!items.length) {
    wrap.innerHTML = "<p>No trend data found for this entity.</p>";
    return;
  }

  const isAdmin = !!currentUser?.access?.isAdmin;

  const rows = items
    .map((item) => `
      <tr>
        <td>${item.weekEnding}</td>
        <td>${item.visitVolume}</td>
        <td>${item.callVolume}</td>
        <td>${item.newPatients}</td>
        <td>${item.noShowRate}%</td>
        <td>${item.cancellationRate}%</td>
        <td>${item.abandonedCallRate}%</td>
        <td>${item.status}</td>
        ${
          isAdmin
            ? `
          <td>
            <button class="actionBtn" data-action="override" data-entity="${item.entity}" data-week="${item.weekEnding}">Override</button>
            <button class="actionBtn" data-action="delete" data-entity="${item.entity}" data-week="${item.weekEnding}">Delete</button>
          </td>
        `
            : ""
        }
      </tr>
    `)
    .join("");

  wrap.innerHTML = `
    <table class="regionTable">
      <thead>
        <tr>
          <th>Week Ending</th>
          <th>Visit Volume</th>
          <th>Call Volume</th>
          <th>New Patients</th>
          <th>No Show Rate</th>
          <th>Cancellation Rate</th>
          <th>Abandoned Call Rate</th>
          <th>Status</th>
          ${isAdmin ? "<th>Actions</th>" : ""}
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  `;

  if (isAdmin) {
    wrap.querySelectorAll("button[data-action='override']").forEach((btn) => {
      btn.addEventListener("click", async () => {
        await openOverride(btn.dataset.entity, btn.dataset.week);
      });
    });

    wrap.querySelectorAll("button[data-action='delete']").forEach((btn) => {
      btn.addEventListener("click", async () => {
        await deleteWeek(btn.dataset.entity, btn.dataset.week);
      });
    });
  }
}

function syncTrendsRangeUi() {
  const mode = document.getElementById("trendsRangeMode").value;
  document.getElementById("trendsWeeksWrap").style.display = mode === "recent" ? "" : "none";
  document.getElementById("trendsStartWrap").style.display = mode === "dates" ? "" : "none";
  document.getElementById("trendsEndWrap").style.display = mode === "dates" ? "" : "none";
}

async function loadTrends() {
  const entity = getSelectedTrendsEntity();
  const mode = document.getElementById("trendsRangeMode").value;
  const weeks = document.getElementById("trendsLimit").value;
  const startDate = document.getElementById("trendsStartDate").value;
  const endDate = document.getElementById("trendsEndDate").value;

  renderContextBrand("trendsContextBrand", entity);

  let url = `/api/trends?entity=${encodeURIComponent(entity)}`;

  if (mode === "dates") {
    url += `&mode=dateRange`;
    if (startDate) url += `&startDate=${encodeURIComponent(startDate)}`;
    if (endDate) url += `&endDate=${encodeURIComponent(endDate)}`;
  } else {
    url += `&mode=recent&weeks=${encodeURIComponent(weeks)}`;
  }

  const result = await apiGet(url);
  renderTrendsCards(result);
  renderTrendsTable(result);
  setTrendsDebug(result);
}

function formatDeltaOnly(current, prior) {
  const diff = normalizeNumber(current) - normalizeNumber(prior);
  return `${diff >= 0 ? "+" : ""}${diff} vs comparison`;
}

function getEntityMap(summary) {
  const map = {};
  (summary?.regions || []).forEach((r) => {
    map[r.entity] = r;
  });
  return map;
}

function statusClass(status) {
  const s = String(status || "missing").toLowerCase();
  if (s === "approved") return "dashboardStatus-approved";
  if (s === "submitted") return "dashboardStatus-submitted";
  if (s === "draft") return "dashboardStatus-draft";
  return "dashboardStatus-missing";
}

function varianceClass(pct) {
  if (pct === null || pct === undefined || pct === "") return "varianceNeutral";
  if (Number(pct) > 0) return "variancePos";
  if (Number(pct) < 0) return "varianceNeg";
  return "varianceNeutral";
}

function pctChange(current, comparison) {
  const c = normalizeNumber(current);
  const p = normalizeNumber(comparison);
  if (!p) return null;
  return ((c - p) / p) * 100;
}

function summarizeDateRange(weeks, periodType, anchorWeek, compareAgainst, entityScope, customStart, customEnd) {
  const primary = weeks.primaryWeeks.length
    ? `${weeks.primaryWeeks[0]} to ${weeks.primaryWeeks[weeks.primaryWeeks.length - 1]}`
    : "No primary weeks";

  const comparison = weeks.comparisonWeeks.length
    ? `${weeks.comparisonWeeks[0]} to ${weeks.comparisonWeeks[weeks.comparisonWeeks.length - 1]}`
    : "No comparison weeks";

  const periodLabelMap = {
    currentWeek: "Current Week",
    lastWeek: "Last Week",
    mtd: "Month to Date",
    lastMonth: "Last Month",
    rolling4: "Rolling 4 Weeks",
    custom: "Custom Range"
  };

  const compareLabelMap = {
    priorPeriod: "Prior Period",
    target: "Target",
    budget: "Budget",
    forecast: "Forecast"
  };

  if (periodType === "custom") {
    return `Viewing ${periodLabelMap[periodType]} (${customStart} to ${customEnd}) for ${entityScope === "ALL" ? "all entities" : entityScope}. Compare: ${compareLabelMap[compareAgainst]}. Primary weeks: ${primary}. Comparison weeks: ${comparison}.`;
  }

  return `Viewing ${periodLabelMap[periodType]} anchored to ${anchorWeek} for ${entityScope === "ALL" ? "all entities" : entityScope}. Compare: ${compareLabelMap[compareAgainst]}. Primary weeks: ${primary}. Comparison weeks: ${comparison}.`;
}

function getFridayOnOrBefore(dateInput) {
  const d = new Date(`${dateInput}T12:00:00Z`);
  while (d.getUTCDay() !== 5) {
    d.setUTCDate(d.getUTCDate() - 1);
  }
  return d.toISOString().slice(0, 10);
}

function getFridayOnOrAfter(dateInput) {
  const d = new Date(`${dateInput}T12:00:00Z`);
  while (d.getUTCDay() !== 5) {
    d.setUTCDate(d.getUTCDate() + 1);
  }
  return d.toISOString().slice(0, 10);
}

function buildFridayRange(startIso, endIso) {
  const start = getFridayOnOrAfter(startIso);
  const end = getFridayOnOrBefore(endIso);

  const startDate = new Date(`${start}T12:00:00Z`);
  const endDate = new Date(`${end}T12:00:00Z`);

  if (startDate > endDate) return [];

  const out = [];
  const d = new Date(startDate);
  while (d <= endDate) {
    out.push(d.toISOString().slice(0, 10));
    d.setUTCDate(d.getUTCDate() + 7);
  }
  return out;
}

function getMonthStart(isoDate) {
  const d = new Date(`${isoDate}T12:00:00Z`);
  d.setUTCDate(1);
  return d.toISOString().slice(0, 10);
}

function getMonthEnd(isoDate) {
  const d = new Date(`${isoDate}T12:00:00Z`);
  d.setUTCMonth(d.getUTCMonth() + 1, 0);
  return d.toISOString().slice(0, 10);
}

function shiftDateDays(isoDate, days) {
  const d = new Date(`${isoDate}T12:00:00Z`);
  d.setUTCDate(d.getUTCDate() + days);
  return d.toISOString().slice(0, 10);
}

function getDashboardSelections() {
  return {
    periodType: document.getElementById("dashboardPeriodType").value,
    compareAgainst: document.getElementById("dashboardCompareAgainst").value,
    entityScope: document.getElementById("dashboardEntityScope").value,
    anchorWeek: document.getElementById("dashboardWeekEnding").value,
    customStart: document.getElementById("dashboardCustomStart").value,
    customEnd: document.getElementById("dashboardCustomEnd").value
  };
}

function syncDashboardPeriodUi() {
  const periodType = document.getElementById("dashboardPeriodType").value;
  const isCustom = periodType === "custom";

  document.getElementById("dashboardCustomStartWrap").style.display = isCustom ? "" : "none";
  document.getElementById("dashboardCustomEndWrap").style.display = isCustom ? "" : "none";
  document.getElementById("dashboardWeekWrap").style.display = isCustom ? "none" : "";
}

function buildDashboardWeekSets({ periodType, anchorWeek, customStart, customEnd }) {
  const anchor = anchorWeek || getDefaultWeekEnding();

  if (periodType === "currentWeek") {
    return {
      primaryWeeks: [anchor],
      comparisonWeeks: [getPreviousWeekEnding(anchor)]
    };
  }

  if (periodType === "lastWeek") {
    const primary = getPreviousWeekEnding(anchor);
    return {
      primaryWeeks: [primary],
      comparisonWeeks: [getPreviousWeekEnding(primary)]
    };
  }

  if (periodType === "rolling4") {
    const primaryWeeks = [
      shiftDateDays(anchor, -21),
      shiftDateDays(anchor, -14),
      shiftDateDays(anchor, -7),
      anchor
    ];
    const comparisonWeeks = primaryWeeks.map((w) => shiftDateDays(w, -28));
    return { primaryWeeks, comparisonWeeks };
  }

  if (periodType === "mtd") {
    const monthStart = getMonthStart(anchor);
    const primaryWeeks = buildFridayRange(monthStart, anchor);
    const prevMonthDate = shiftDateDays(monthStart, -1);
    const prevMonthStart = getMonthStart(prevMonthDate);
    const prevRangeWeeks = buildFridayRange(prevMonthStart, prevMonthDate);
    const comparisonWeeks = prevRangeWeeks.slice(0, primaryWeeks.length);
    return { primaryWeeks, comparisonWeeks };
  }

  if (periodType === "lastMonth") {
    const priorMonthRef = shiftDateDays(getMonthStart(anchor), -1);
    const primaryStart = getMonthStart(priorMonthRef);
    const primaryEnd = getMonthEnd(priorMonthRef);
    const primaryWeeks = buildFridayRange(primaryStart, primaryEnd);

    const comparisonMonthRef = shiftDateDays(primaryStart, -1);
    const comparisonStart = getMonthStart(comparisonMonthRef);
    const comparisonEnd = getMonthEnd(comparisonMonthRef);
    const comparisonWeeks = buildFridayRange(comparisonStart, comparisonEnd);

    return { primaryWeeks, comparisonWeeks };
  }

  if (periodType === "custom") {
    if (!customStart || !customEnd) {
      return { primaryWeeks: [], comparisonWeeks: [] };
    }

    const primaryWeeks = buildFridayRange(customStart, customEnd);
    const start = new Date(`${customStart}T12:00:00Z`);
    const end = new Date(`${customEnd}T12:00:00Z`);
    const spanDays = Math.round((end - start) / 86400000) + 1;

    const comparisonStart = shiftDateDays(customStart, -spanDays);
    const comparisonEnd = shiftDateDays(customEnd, -spanDays);
    const comparisonWeeks = buildFridayRange(comparisonStart, comparisonEnd);

    return { primaryWeeks, comparisonWeeks };
  }

  return {
    primaryWeeks: [anchor],
    comparisonWeeks: [getPreviousWeekEnding(anchor)]
  };
}

async function fetchExecutiveSummaryByWeek(weekEnding) {
  return apiGet(`/api/executive-summary?weekEnding=${encodeURIComponent(weekEnding)}`);
}

function aggregateExecutiveSummaries(summaries, entityScope) {
  const entityMap = new Map();

  summaries.forEach((summary) => {
    (summary.regions || []).forEach((region) => {
      if (entityScope !== "ALL" && region.entity !== entityScope) return;

      if (!entityMap.has(region.entity)) {
        entityMap.set(region.entity, {
          entity: region.entity,
          weekEnding: summary.weekEnding,
          status: "approved",
          visitVolume: 0,
          callVolume: 0,
          newPatients: 0,
          noShowRateTotal: 0,
          cancellationRateTotal: 0,
          abandonedCallRateTotal: 0,
          weekCount: 0
        });
      }

      const row = entityMap.get(region.entity);
      row.visitVolume += normalizeNumber(region.visitVolume);
      row.callVolume += normalizeNumber(region.callVolume);
      row.newPatients += normalizeNumber(region.newPatients);
      row.noShowRateTotal += normalizeNumber(region.noShowRate);
      row.cancellationRateTotal += normalizeNumber(region.cancellationRate);
      row.abandonedCallRateTotal += normalizeNumber(region.abandonedCallRate);
      row.weekCount += 1;
    });
  });

  const regions = Array.from(entityMap.values()).map((row) => ({
    entity: row.entity,
    status: "approved",
    visitVolume: row.visitVolume,
    callVolume: row.callVolume,
    newPatients: row.newPatients,
    noShowRate: row.weekCount ? row.noShowRateTotal / row.weekCount : 0,
    cancellationRate: row.weekCount ? row.cancellationRateTotal / row.weekCount : 0,
    abandonedCallRate: row.weekCount ? row.abandonedCallRateTotal / row.weekCount : 0
  }));

  const totals = {
    visitVolume: regions.reduce((sum, r) => sum + normalizeNumber(r.visitVolume), 0),
    callVolume: regions.reduce((sum, r) => sum + normalizeNumber(r.callVolume), 0),
    newPatients: regions.reduce((sum, r) => sum + normalizeNumber(r.newPatients), 0)
  };

  return {
    entityCount: regions.length,
    totals,
    regions
  };
}

async function loadDashboardDataForWeeks(weeks, entityScope) {
  const validWeeks = (weeks || []).filter(Boolean);
  if (!validWeeks.length) {
    return {
      entityCount: 0,
      totals: { visitVolume: 0, callVolume: 0, newPatients: 0 },
      regions: []
    };
  }

  const summaries = await Promise.all(validWeeks.map((week) => fetchExecutiveSummaryByWeek(week)));
  return aggregateExecutiveSummaries(summaries, entityScope);
}

function renderDashboardCards(current, comparison, compareAgainst) {
  const container = document.getElementById("dashboardCards");
  container.innerHTML = "";

  const compareLabelMap = {
    priorPeriod: "vs prior period",
    target: "vs target",
    budget: "vs budget",
    forecast: "vs forecast"
  };

  const labelSuffix = compareLabelMap[compareAgainst] || "vs comparison";

  const currentRegions = current?.regions || [];
  const avg = (key) => {
    if (!currentRegions.length) return 0;
    return (
      currentRegions.reduce((sum, r) => sum + normalizeNumber(r[key]), 0) /
      currentRegions.length
    );
  };

  const cardData = [
    {
      label: "Approved Regions",
      value: current?.entityCount || 0,
      sub: `Prev ${comparison?.entityCount || 0}`
    },
    {
      label: "Visit Volume",
      value: current?.totals?.visitVolume || 0,
      sub: formatDeltaOnly(current?.totals?.visitVolume || 0, comparison?.totals?.visitVolume || 0, labelSuffix)
    },
    {
      label: "Call Volume",
      value: current?.totals?.callVolume || 0,
      sub: formatDeltaOnly(current?.totals?.callVolume || 0, comparison?.totals?.callVolume || 0, labelSuffix)
    },
    {
      label: "New Patients",
      value: current?.totals?.newPatients || 0,
      sub: formatDeltaOnly(current?.totals?.newPatients || 0, comparison?.totals?.newPatients || 0, labelSuffix)
    },
    {
      label: "Avg No Show %",
      value: `${avg("noShowRate").toFixed(1)}%`,
      sub: "Across approved entities"
    },
    {
      label: "Avg Cancel %",
      value: `${avg("cancellationRate").toFixed(1)}%`,
      sub: "Across approved entities"
    },
    {
      label: "Avg Abandoned %",
      value: `${avg("abandonedCallRate").toFixed(1)}%`,
      sub: "Across approved entities"
    }
  ];

  cardData.forEach((item) => {
    const div = document.createElement("div");
    div.className = "summaryCard";
    div.innerHTML = `
      <h3>${item.label}</h3>
      <div class="value">${item.value}</div>
      <div class="metaText">${item.sub || ""}</div>
    `;
    container.appendChild(div);
  });
}

function formatDeltaOnly(current, comparison, label = "vs comparison") {
  const diff = normalizeNumber(current) - normalizeNumber(comparison);
  return `${diff >= 0 ? "+" : ""}${diff} ${label}`;
}

function getEntityMap(summary) {
  const map = {};
  (summary?.regions || []).forEach((r) => {
    map[r.entity] = r;
  });
  return map;
}

function statusClass(status) {
  const s = String(status || "missing").toLowerCase();
  if (s === "approved") return "dashboardStatus-approved";
  if (s === "submitted") return "dashboardStatus-submitted";
  if (s === "draft") return "dashboardStatus-draft";
  return "dashboardStatus-missing";
}

function varianceClass(pct) {
  if (pct === null || pct === undefined || pct === "") return "varianceNeutral";
  if (Number(pct) > 0) return "variancePos";
  if (Number(pct) < 0) return "varianceNeg";
  return "varianceNeutral";
}

function renderDashboardEntities(current, comparison, compareAgainst, entityScope) {
  const container = document.getElementById("dashboardEntities");
  container.innerHTML = "";

  const currentMap = getEntityMap(current);
  const comparisonMap = getEntityMap(comparison);
  const entitiesToShow = entityScope === "ALL" ? ENTITIES : [entityScope];

  entitiesToShow.forEach((entity) => {
    const row = currentMap[entity] || {
      entity,
      visitVolume: 0,
      callVolume: 0,
      newPatients: 0,
      noShowRate: 0,
      cancellationRate: 0,
      abandonedCallRate: 0,
      status: "missing"
    };

    const prior = comparisonMap[entity] || {
      visitVolume: 0,
      callVolume: 0,
      newPatients: 0
    };

    const brand = getBranding(entity);

    const visitPct = pctChange(row.visitVolume, prior.visitVolume);
    const callPct = pctChange(row.callVolume, prior.callVolume);
    const npPct = pctChange(row.newPatients, prior.newPatients);

    const card = document.createElement("div");
    card.className = "dashboardEntityCard";
    card.style.borderColor = brand.accentBorder;
    card.innerHTML = `
      <div class="dashboardEntityTopBar" style="background:${brand.accent};"></div>
      <div class="dashboardEntityInner">
        <div class="dashboardEntityHeader">
          <div class="dashboardEntityTitleWrap">
            <div class="dashboardStatusBadge ${statusClass(row.status)}">${row.status}</div>
            <h4>${entity}</h4>
            <div class="dashboardEntityMeta">${brand.fullName}</div>
          </div>
          <div class="dashboardEntityLogoWrap" style="box-shadow: inset 0 0 0 1px ${brand.accentBorder};">
            <img src="${brand.logo}" alt="${brand.label}" />
          </div>
        </div>

        <div class="dashboardMetricRow">
          <span>Visits</span>
          <span class="metricValue">
            ${normalizeNumber(row.visitVolume)}
            <span class="metricSub">${formatDeltaOnly(row.visitVolume, prior.visitVolume)}</span>
          </span>
        </div>

        <div class="dashboardMetricRow">
          <span>Calls</span>
          <span class="metricValue">
            ${normalizeNumber(row.callVolume)}
            <span class="metricSub">${formatDeltaOnly(row.callVolume, prior.callVolume)}</span>
          </span>
        </div>

        <div class="dashboardMetricRow">
          <span>New Patients</span>
          <span class="metricValue">
            ${normalizeNumber(row.newPatients)}
            <span class="metricSub">${formatDeltaOnly(row.newPatients, prior.newPatients)}</span>
          </span>
        </div>

        <div class="dashboardMetricRow">
          <span>No Show</span>
          <span class="metricValue">${normalizeNumber(row.noShowRate).toFixed(1)}%</span>
        </div>

        <div class="dashboardMetricRow">
          <span>Cancel</span>
          <span class="metricValue">${normalizeNumber(row.cancellationRate).toFixed(1)}%</span>
        </div>

        <div class="dashboardMetricRow">
          <span>Abandoned</span>
          <span class="metricValue">${normalizeNumber(row.abandonedCallRate).toFixed(1)}%</span>
        </div>

        <div class="dashboardVarianceRow">
          <div class="varianceChip" style="border-color:${brand.accentBorder}; background:${brand.accentSoft};">
            <span class="chipLabel">Visits ${compareAgainst === "priorPeriod" ? "vs Prior" : "Target Pending"}</span>
            <span class="chipValue ${varianceClass(visitPct)}">${compareAgainst === "priorPeriod" ? (visitPct === null ? "n/a" : `${visitPct >= 0 ? "+" : ""}${visitPct.toFixed(1)}%`) : normalizeNumber(row.visitVolume)}</span>
          </div>
          <div class="varianceChip" style="border-color:${brand.accentBorder}; background:${brand.accentSoft};">
            <span class="chipLabel">Calls ${compareAgainst === "priorPeriod" ? "vs Prior" : "Target Pending"}</span>
            <span class="chipValue ${varianceClass(callPct)}">${compareAgainst === "priorPeriod" ? (callPct === null ? "n/a" : `${callPct >= 0 ? "+" : ""}${callPct.toFixed(1)}%`) : normalizeNumber(row.callVolume)}</span>
          </div>
          <div class="varianceChip" style="border-color:${brand.accentBorder}; background:${brand.accentSoft};">
            <span class="chipLabel">NP ${compareAgainst === "priorPeriod" ? "vs Prior" : "Target Pending"}</span>
            <span class="chipValue ${varianceClass(npPct)}">${compareAgainst === "priorPeriod" ? (npPct === null ? "n/a" : `${npPct >= 0 ? "+" : ""}${npPct.toFixed(1)}%`) : normalizeNumber(row.newPatients)}</span>
          </div>
        </div>
      </div>
    `;
    container.appendChild(card);
  });
}

function buildDashboardAlerts(current, comparison, entityScope) {
  const alerts = [];
  const currentMap = getEntityMap(current);
  const previousMap = getEntityMap(comparison);
  const entitiesToShow = entityScope === "ALL" ? ENTITIES : [entityScope];

  entitiesToShow.forEach((entity) => {
    const row = currentMap[entity];
    const prior = previousMap[entity] || {};

    if (!row) {
      alerts.push({
        severity: "yellow",
        text: `${entity} has no approved record in the selected period.`
      });
      return;
    }

    if (String(row.status || "").toLowerCase() !== "approved") {
      alerts.push({
        severity: "yellow",
        text: `${entity} is not approved in the selected period.`
      });
    }

    if (normalizeNumber(row.noShowRate) >= 6) {
      alerts.push({
        severity: "red",
        text: `${entity} no show rate is elevated at ${normalizeNumber(row.noShowRate).toFixed(1)}%.`
      });
    }

    if (normalizeNumber(row.cancellationRate) >= 8) {
      alerts.push({
        severity: "red",
        text: `${entity} cancellation rate is elevated at ${normalizeNumber(row.cancellationRate).toFixed(1)}%.`
      });
    }

    if (normalizeNumber(row.abandonedCallRate) >= 10) {
      alerts.push({
        severity: "red",
        text: `${entity} abandoned call rate is elevated at ${normalizeNumber(row.abandonedCallRate).toFixed(1)}%.`
      });
    }

    const visitDrop = normalizeNumber(row.visitVolume) - normalizeNumber(prior.visitVolume);
    if (visitDrop < -100) {
      alerts.push({
        severity: "yellow",
        text: `${entity} visit volume is down ${Math.abs(visitDrop)} vs comparison period.`
      });
    }
  });

  if (!alerts.length) {
    alerts.push({
      severity: "green",
      text: "No major operational alerts for the selected period."
    });
  }

  return alerts;
}

function renderDashboardAlerts(current, comparison, entityScope) {
  const container = document.getElementById("dashboardAlerts");
  const alerts = buildDashboardAlerts(current, comparison, entityScope);

  container.innerHTML = alerts
    .map(
      (alert) => `
      <div class="dashboardAlert ${alert.severity}">
        ${alert.text}
      </div>
    `
    )
    .join("");
}

function renderDashboardSnapshot(current, entityScope) {
  const container = document.getElementById("dashboardSnapshot");
  const regions = (current?.regions || []).filter((r) => entityScope === "ALL" || r.entity === entityScope);

  if (!regions.length) {
    container.innerHTML = `<div class="dashboardEmpty">No approved entities found for the selected period.</div>`;
    return;
  }

  const rows = regions
    .map(
      (r) => `
      <tr>
        <td>${r.entity}</td>
        <td>${normalizeNumber(r.visitVolume)}</td>
        <td>${normalizeNumber(r.callVolume)}</td>
        <td>${normalizeNumber(r.newPatients)}</td>
        <td>${normalizeNumber(r.noShowRate).toFixed(1)}%</td>
        <td>${normalizeNumber(r.cancellationRate).toFixed(1)}%</td>
        <td>${normalizeNumber(r.abandonedCallRate).toFixed(1)}%</td>
        <td>${r.status}</td>
      </tr>
    `
    )
    .join("");

  container.innerHTML = `
    <table class="regionTable">
      <thead>
        <tr>
          <th>Entity</th>
          <th>Visit</th>
          <th>Calls</th>
          <th>New</th>
          <th>No Show</th>
          <th>Cancel</th>
          <th>Abandoned</th>
          <th>Status</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  `;
}

async function loadDashboardLanding() {
  const selections = getDashboardSelections();
  const weekSets = buildDashboardWeekSets(selections);

  const current = await loadDashboardDataForWeeks(weekSets.primaryWeeks, selections.entityScope);
  const comparison = await loadDashboardDataForWeeks(weekSets.comparisonWeeks, selections.entityScope);

  const compareBannerTextMap = {
    target: "Target comparison is staged but not wired yet. The dashboard is currently showing actuals while keeping the slicer structure in place.",
    budget: "Budget comparison is staged but not wired yet. The dashboard is currently showing actuals while keeping the slicer structure in place.",
    forecast: "Forecast comparison is staged but not wired yet. The dashboard is currently showing actuals while keeping the slicer structure in place."
  };

  if (["target", "budget", "forecast"].includes(selections.compareAgainst)) {
    setDashboardBenchmarkBanner(compareBannerTextMap[selections.compareAgainst], true);
  } else {
    setDashboardBenchmarkBanner("", false);
  }

  setDashboardRangeSummary(
    summarizeDateRange(
      weekSets,
      selections.periodType,
      selections.anchorWeek,
      selections.compareAgainst,
      selections.entityScope,
      selections.customStart,
      selections.customEnd
    )
  );

  renderDashboardCards(current, comparison, selections.compareAgainst);
  renderDashboardEntities(current, comparison, selections.compareAgainst, selections.entityScope);
  renderDashboardAlerts(current, comparison, selections.entityScope);
  renderDashboardSnapshot(current, selections.entityScope);

  setDashboardDebug({
    selections,
    weekSets,
    currentPeriod: current,
    comparisonPeriod: comparison
  });
}

async function readFileAsBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const result = reader.result || "";
      const base64 = String(result).split(",")[1] || "";
      resolve(base64);
    };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

async function runImport() {
  if (!currentUser?.access?.isAdmin) {
    throw new Error("Admin only");
  }

  const fileInput = document.getElementById("importFile");
  const file = fileInput.files && fileInput.files[0];

  if (!file) {
    throw new Error("Select a workbook file first");
  }

  setImportStatus("Reading workbook...");
  const fileBase64 = await readFileAsBase64(file);

  const payload = {
    fileName: file.name,
    fileBase64
  };

  setImportStatus("Importing workbook...");
  setImportDebug(payload.fileName);

  const result = await apiPost("/api/import-excel", payload);

  setImportStatus(result.message || "Import completed");
  setImportDebug(result);
}

(async function init() {
  try {
    currentUser = await apiGet("/api/me");

    renderUser(currentUser);
    setupEntityDropdown(currentUser);
    setupTrendsEntityDropdown(currentUser);
    renderForm();

    const defaultWeek = getDefaultWeekEnding();
    document.getElementById("dashboardWeekEnding").value = defaultWeek;
    document.getElementById("dashboardCustomEnd").value = defaultWeek;
    document.getElementById("dashboardCustomStart").value = getDateWeeksAgo(8, defaultWeek);

    document.getElementById("weekEnding").value = defaultWeek;
    document.getElementById("executiveWeekEnding").value = defaultWeek;
    document.getElementById("trendsStartDate").value = getDateWeeksAgo(12);
    document.getElementById("trendsEndDate").value = defaultWeek;

    syncTrendsRangeUi();
    syncDashboardPeriodUi();
    renderContextBrand("entryContextBrand", getSelectedEntity());
    renderContextBrand("trendsContextBrand", getSelectedTrendsEntity());

    document.getElementById("entitySelect").addEventListener("change", async () => {
      renderContextBrand("entryContextBrand", getSelectedEntity());
      try {
        await loadWeek();
      } catch (e) {
        setStatus(e.message, true);
        setDebug(String(e));
      }
    });

    document.getElementById("weekEnding").addEventListener("change", async () => {
      try {
        await loadWeek();
      } catch (e) {
        setStatus(e.message, true);
        setDebug(String(e));
      }
    });

    document.getElementById("saveBtn").addEventListener("click", async () => {
      try {
        await saveWeek();
      } catch (e) {
        setStatus(e.message, true);
        setDebug(String(e));
      }
    });

    document.getElementById("submitBtn").addEventListener("click", async () => {
      try {
        await submitWeek();
      } catch (e) {
        setStatus(e.message, true);
        setDebug(String(e));
      }
    });

    document.getElementById("approveBtn").addEventListener("click", async () => {
      try {
        await approveWeek();
      } catch (e) {
        setStatus(e.message, true);
        setDebug(String(e));
      }
    });

    document.getElementById("navDashboardBtn").addEventListener("click", async () => {
      showDashboardView();
      try {
        await loadDashboardLanding();
      } catch (e) {
        setDashboardDebug(String(e));
      }
    });

    document.getElementById("navEntryBtn").addEventListener("click", showEntryView);

    document.getElementById("navExecutiveBtn").addEventListener("click", async () => {
      showExecutiveView();
      try {
        await loadExecutiveSummary();
      } catch (e) {
        setExecutiveDebug(String(e));
      }
    });

    document.getElementById("navTrendsBtn").addEventListener("click", async () => {
      showTrendsView();
      try {
        await loadTrends();
      } catch (e) {
        setTrendsDebug(String(e));
      }
    });

    document.getElementById("navImportBtn").addEventListener("click", showImportView);

    document.getElementById("loadDashboardBtn").addEventListener("click", async () => {
      try {
        await loadDashboardLanding();
      } catch (e) {
        setDashboardDebug(String(e));
      }
    });

    document.getElementById("dashboardPeriodType").addEventListener("change", async () => {
      syncDashboardPeriodUi();
      try {
        await loadDashboardLanding();
      } catch (e) {
        setDashboardDebug(String(e));
      }
    });

    document.getElementById("dashboardCompareAgainst").addEventListener("change", async () => {
      try {
        await loadDashboardLanding();
      } catch (e) {
        setDashboardDebug(String(e));
      }
    });

    document.getElementById("dashboardEntityScope").addEventListener("change", async () => {
      try {
        await loadDashboardLanding();
      } catch (e) {
        setDashboardDebug(String(e));
      }
    });

    document.getElementById("dashboardWeekEnding").addEventListener("change", async () => {
      if (document.getElementById("dashboardPeriodType").value !== "custom") {
        try {
          await loadDashboardLanding();
        } catch (e) {
          setDashboardDebug(String(e));
        }
      }
    });

    document.getElementById("dashboardCustomStart").addEventListener("change", async () => {
      if (document.getElementById("dashboardPeriodType").value === "custom") {
        try {
          await loadDashboardLanding();
        } catch (e) {
          setDashboardDebug(String(e));
        }
      }
    });

    document.getElementById("dashboardCustomEnd").addEventListener("change", async () => {
      if (document.getElementById("dashboardPeriodType").value === "custom") {
        try {
          await loadDashboardLanding();
        } catch (e) {
          setDashboardDebug(String(e));
        }
      }
    });

    document.getElementById("loadExecutiveBtn").addEventListener("click", async () => {
      try {
        await loadExecutiveSummary();
      } catch (e) {
        setExecutiveDebug(String(e));
      }
    });

    document.getElementById("loadTrendsBtn").addEventListener("click", async () => {
      try {
        await loadTrends();
      } catch (e) {
        setTrendsDebug(String(e));
      }
    });

    document.getElementById("trendsEntitySelect").addEventListener("change", async () => {
      renderContextBrand("trendsContextBrand", getSelectedTrendsEntity());
      try {
        await loadTrends();
      } catch (e) {
        setTrendsDebug(String(e));
      }
    });

    document.getElementById("trendsLimit").addEventListener("change", async () => {
      if (document.getElementById("trendsRangeMode").value === "recent") {
        try {
          await loadTrends();
        } catch (e) {
          setTrendsDebug(String(e));
        }
      }
    });

    document.getElementById("trendsRangeMode").addEventListener("change", async () => {
      syncTrendsRangeUi();
      try {
        await loadTrends();
      } catch (e) {
        setTrendsDebug(String(e));
      }
    });

    document.getElementById("trendsStartDate").addEventListener("change", async () => {
      if (document.getElementById("trendsRangeMode").value === "dates") {
        try {
          await loadTrends();
        } catch (e) {
          setTrendsDebug(String(e));
        }
      }
    });

    document.getElementById("trendsEndDate").addEventListener("change", async () => {
      if (document.getElementById("trendsRangeMode").value === "dates") {
        try {
          await loadTrends();
        } catch (e) {
          setTrendsDebug(String(e));
        }
      }
    });

    document.getElementById("runImportBtn").addEventListener("click", async () => {
      try {
        await runImport();
      } catch (e) {
        setImportStatus(e.message, true);
        setImportDebug(String(e));
      }
    });

    showDashboardView();
    await loadDashboardLanding();
  } catch (error) {
    setStatus(error.message || "Failed to load app", true);
    setDebug(String(error));
    setDashboardDebug(String(error));
  }
})();
