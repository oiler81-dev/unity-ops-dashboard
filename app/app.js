const ENTITIES = ["LAOSS", "NES", "SpineOne", "MRO"];

const ENTITY_BRANDING = {
  LAOSS: {
    label: "LAOSS",
    fullName: "Los Angeles Orthopedic Surgery Specialists",
    logo: "./assets/logos/laoss.png",
    accent: "#F28C28"
  },
  NES: {
    label: "NES",
    fullName: "Northwest Extremity Specialists",
    logo: "./assets/logos/nes.png",
    accent: "#2E5B88"
  },
  SpineOne: {
    label: "SpineOne",
    fullName: "SpineOne",
    logo: "./assets/logos/spineone.png",
    accent: "#5A6F95"
  },
  MRO: {
    label: "MRO",
    fullName: "Midland & Riverside Orthopedics",
    logo: "./assets/logos/mro.png",
    accent: "#6B7E99"
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
    data = null;
  }

  if (!res.ok) {
    const message =
      data?.details ||
      data?.error ||
      text ||
      `Request failed with status ${res.status}`;

    throw new Error(message);
  }

  if (data !== null) {
    return data;
  }

  return {
    ok: true,
    raw: text
  };
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

function byId(id) {
  return document.getElementById(id);
}

function firstExistingId(ids) {
  for (const id of ids) {
    const el = byId(id);
    if (el) return el;
  }
  return null;
}

function setStatus(message, isError = false) {
  const el = byId("statusMessage");
  if (!el) return;
  el.textContent = message;
  el.style.color = isError ? "#ff8a8a" : "#7CFC98";
}

function setDebug(data) {
  const el = byId("debugOutput");
  if (!el) return;
  el.textContent = typeof data === "string" ? data : JSON.stringify(data, null, 2);
}

function setExecutiveDebug(data) {
  const el = byId("executiveDebugOutput");
  if (!el) return;
  el.textContent = typeof data === "string" ? data : JSON.stringify(data, null, 2);
}

function setTrendsDebug(data) {
  const el = byId("trendsDebugOutput");
  if (!el) return;
  el.textContent = typeof data === "string" ? data : JSON.stringify(data, null, 2);
}

function setImportStatus(message, isError = false) {
  const el =
    firstExistingId([
      "importStatusMessage",
      "adminImportStatusMessage",
      "budgetImportStatusMessage"
    ]);

  if (!el) return;
  el.textContent = message;
  el.style.color = isError ? "#ff8a8a" : "#7CFC98";
}

function setImportDebug(data) {
  const el =
    firstExistingId([
      "importDebugOutput",
      "adminImportDebugOutput"
    ]);

  if (!el) return;
  el.textContent = typeof data === "string" ? data : JSON.stringify(data, null, 2);
}

function setDashboardDebug(data) {
  const el = byId("dashboardDebugOutput");
  if (!el) return;
  el.textContent = typeof data === "string" ? data : JSON.stringify(data, null, 2);
}

function getDefaultWeekEnding() {
  const d = new Date();
  const diff = (5 - d.getDay() + 7) % 7;
  d.setDate(d.getDate() + diff);
  return d.toISOString().slice(0, 10);
}

function addDays(isoDate, days) {
  const d = new Date(`${isoDate}T12:00:00Z`);
  d.setUTCDate(d.getUTCDate() + days);
  return d.toISOString().slice(0, 10);
}

function getDateWeeksAgo(weeksAgo, anchor = null) {
  const d = anchor ? new Date(`${anchor}T12:00:00Z`) : new Date();
  d.setUTCDate(d.getUTCDate() - weeksAgo * 7);
  return d.toISOString().slice(0, 10);
}

function getPreviousWeekEnding(weekEnding) {
  return addDays(weekEnding, -7);
}

function normalizeNumber(value) {
  if (value === null || value === undefined || value === "") return 0;
  const n = Number(value);
  return Number.isFinite(n) ? n : 0;
}

function renderUser(userData) {
  const label = `${userData.user.userDetails} (${userData.access.role})`;
  byId("userInfo").innerText = label;
}

function setupEntityDropdown(userData) {
  const select = byId("entitySelect");
  if (!select) return;

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
  const select = byId("trendsEntitySelect");
  if (!select) return;

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
  return byId("entitySelect").value;
}

function getSelectedTrendsEntity() {
  return byId("trendsEntitySelect").value;
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

  const container = byId("kpiForm");
  if (!container) return;

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
    const input = byId(key);
    if (!input) return;
    input.value = mapped[key] !== null && mapped[key] !== undefined ? mapped[key] : "";
  });
}

function getFormValues() {
  return {
    visitVolume: byId("visitVolume").value,
    callVolume: byId("callVolume").value,
    newPatients: byId("newPatients").value,
    noShowRate: byId("noShowRate").value,
    cancellationRate: byId("cancellationRate").value,
    abandonedCallRate: byId("abandonedCallRate").value
  };
}

function updateButtonState() {
  const saveBtn = byId("saveBtn");
  const submitBtn = byId("submitBtn");
  const approveBtn = byId("approveBtn");

  if (!saveBtn || !submitBtn || !approveBtn) return;

  const status = String(currentWeekData?.status || "draft").toLowerCase();
  const isAdmin = !!currentUser?.access?.isAdmin;

  saveBtn.disabled = status === "approved" && !isAdmin;
  submitBtn.disabled = status === "submitted" || status === "approved";
  approveBtn.disabled = !isAdmin || status !== "submitted";
}

function getBranding(entity) {
  return ENTITY_BRANDING[entity] || {
    label: entity,
    fullName: entity,
    logo: "",
    accent: "#4b88c7"
  };
}

function renderEntityBrand(containerId, entity) {
  const container = byId(containerId);
  if (!container) return;

  if (!entity) {
    container.innerHTML = "";
    return;
  }

  const brand = getBranding(entity);
  container.innerHTML = `
    <div style="display:flex; align-items:center; gap:12px; padding:10px 12px; border:1px solid #1d435b; border-radius:10px; background:#0a2233;">
      <div style="width:110px; height:48px; background:#fff; border-radius:8px; padding:6px; display:flex; align-items:center; justify-content:center; flex-shrink:0;">
        <img src="${brand.logo}" alt="${brand.label}" style="max-width:100%; max-height:100%; object-fit:contain;" />
      </div>
      <div>
        <div style="font-weight:bold; color:${brand.accent};">${brand.label}</div>
        <div style="font-size:12px; opacity:0.9;">${brand.fullName}</div>
      </div>
    </div>
  `;
}

function renderMetricCards(containerId, items) {
  const container = byId(containerId);
  if (!container) return;

  container.innerHTML = items.map((item) => `
    <div class="summaryCard">
      <h3>${item.label}</h3>
      <div class="value">${item.value}</div>
      ${item.meta ? `<div style="margin-top:6px; font-size:12px; opacity:0.85;">${item.meta}</div>` : ""}
    </div>
  `).join("");
}

async function loadWeek() {
  const weekEnding = byId("weekEnding").value;
  const entity = getSelectedEntity();

  renderEntityBrand("entryBrandWrap", entity);
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
    weekEnding: byId("weekEnding").value,
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
    weekEnding: byId("weekEnding").value,
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
    weekEnding: byId("weekEnding").value,
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
  byId("entitySelect").value = entity;
  byId("weekEnding").value = weekEnding;
  await loadWeek();
  setStatus(`Admin override mode: ${entity} ${weekEnding}`);
}

function hideAllViews() {
  byId("dashboardView").style.display = "none";
  byId("entryView").style.display = "none";
  byId("executiveView").style.display = "none";
  byId("trendsView").style.display = "none";
  byId("importView").style.display = "none";
}

function setActiveNav(buttonId) {
  [
    "navDashboardBtn",
    "navEntryBtn",
    "navExecutiveBtn",
    "navTrendsBtn",
    "navImportBtn"
  ].forEach((id) => {
    const btn = byId(id);
    if (!btn) return;
    if (id === buttonId) {
      btn.style.boxShadow = "inset 0 0 0 1px #f7c62f";
    } else {
      btn.style.boxShadow = "";
    }
  });
}

function showDashboardView() {
  hideAllViews();
  byId("dashboardView").style.display = "";
  setActiveNav("navDashboardBtn");
}

function showEntryView() {
  hideAllViews();
  byId("entryView").style.display = "";
  setActiveNav("navEntryBtn");
  renderEntityBrand("entryBrandWrap", getSelectedEntity());
}

function showExecutiveView() {
  hideAllViews();
  byId("executiveView").style.display = "";
  setActiveNav("navExecutiveBtn");
}

function showTrendsView() {
  hideAllViews();
  byId("trendsView").style.display = "";
  setActiveNav("navTrendsBtn");
  syncTrendsRangeUi();
  renderEntityBrand("trendsBrandWrap", getSelectedTrendsEntity());
}

function showImportView() {
  hideAllViews();
  byId("importView").style.display = "";
  setActiveNav("navImportBtn");
}

function buildWeekSets() {
  const periodType = byId("dashboardPeriodType").value;
  const anchorWeek = byId("dashboardWeekEnding").value || getDefaultWeekEnding();
  const customStart = byId("dashboardCustomStart").value;
  const customEnd = byId("dashboardCustomEnd").value;

  if (periodType === "currentWeek") {
    return {
      primaryWeeks: [anchorWeek],
      comparisonWeeks: [getPreviousWeekEnding(anchorWeek)],
      summary: `Viewing Current Week anchored to ${anchorWeek}`
    };
  }

  if (periodType === "lastWeek") {
    const primary = getPreviousWeekEnding(anchorWeek);
    return {
      primaryWeeks: [primary],
      comparisonWeeks: [getPreviousWeekEnding(primary)],
      summary: `Viewing Last Week anchored from ${anchorWeek}`
    };
  }

  if (periodType === "rolling4") {
    const primaryWeeks = [
      addDays(anchorWeek, -21),
      addDays(anchorWeek, -14),
      addDays(anchorWeek, -7),
      anchorWeek
    ];
    const comparisonWeeks = primaryWeeks.map((w) => addDays(w, -28));
    return {
      primaryWeeks,
      comparisonWeeks,
      summary: `Viewing Rolling 4 Weeks ending ${anchorWeek}`
    };
  }

  if (periodType === "mtd") {
    const d = new Date(`${anchorWeek}T12:00:00Z`);
    d.setUTCDate(1);

    const primaryWeeks = [];
    while (d.toISOString().slice(0, 10) <= anchorWeek) {
      if (d.getUTCDay() === 5) primaryWeeks.push(d.toISOString().slice(0, 10));
      d.setUTCDate(d.getUTCDate() + 1);
    }

    const comparisonWeeks = primaryWeeks.map((w) => addDays(w, -28));
    return {
      primaryWeeks,
      comparisonWeeks,
      summary: `Viewing Month to Date through ${anchorWeek}`
    };
  }

  if (periodType === "lastMonth") {
    const anchor = new Date(`${anchorWeek}T12:00:00Z`);
    anchor.setUTCDate(1);
    anchor.setUTCDate(anchor.getUTCDate() - 1);

    const monthEnd = anchor.toISOString().slice(0, 10);
    anchor.setUTCDate(1);
    const monthStart = anchor.toISOString().slice(0, 10);

    const primaryWeeks = [];
    const walker = new Date(`${monthStart}T12:00:00Z`);
    while (walker.toISOString().slice(0, 10) <= monthEnd) {
      if (walker.getUTCDay() === 5) primaryWeeks.push(walker.toISOString().slice(0, 10));
      walker.setUTCDate(walker.getUTCDate() + 1);
    }

    const comparisonWeeks = primaryWeeks.map((w) => addDays(w, -28));
    return {
      primaryWeeks,
      comparisonWeeks,
      summary: `Viewing Last Month (${monthStart} to ${monthEnd})`
    };
  }

  if (periodType === "custom") {
    const primaryWeeks = [];
    const start = new Date(`${customStart}T12:00:00Z`);
    const end = new Date(`${customEnd}T12:00:00Z`);

    const walker = new Date(start);
    while (walker <= end) {
      if (walker.getUTCDay() === 5) primaryWeeks.push(walker.toISOString().slice(0, 10));
      walker.setUTCDate(walker.getUTCDate() + 1);
    }

    const comparisonWeeks = primaryWeeks.map((w) => addDays(w, -28));
    return {
      primaryWeeks,
      comparisonWeeks,
      summary: `Viewing Custom Range (${customStart} to ${customEnd})`
    };
  }

  return {
    primaryWeeks: [anchorWeek],
    comparisonWeeks: [getPreviousWeekEnding(anchorWeek)],
    summary: `Viewing Current Week anchored to ${anchorWeek}`
  };
}

function syncDashboardPeriodUi() {
  const periodType = byId("dashboardPeriodType").value;
  const custom = periodType === "custom";

  byId("dashboardWeekWrap").style.display = custom ? "none" : "";
  byId("dashboardCustomStartWrap").style.display = custom ? "" : "none";
  byId("dashboardCustomEndWrap").style.display = custom ? "" : "none";
}

async function fetchExecutiveSummaryByWeek(weekEnding) {
  return apiGet(`/api/executive-summary?weekEnding=${encodeURIComponent(weekEnding)}`);
}

function aggregateExecutiveSummaries(summaries, entityScope) {
  const map = new Map();

  summaries.forEach((summary) => {
    (summary.regions || []).forEach((region) => {
      if (entityScope !== "ALL" && region.entity !== entityScope) return;

      if (!map.has(region.entity)) {
        map.set(region.entity, {
          entity: region.entity,
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

      const row = map.get(region.entity);
      row.visitVolume += normalizeNumber(region.visitVolume);
      row.callVolume += normalizeNumber(region.callVolume);
      row.newPatients += normalizeNumber(region.newPatients);
      row.noShowRateTotal += normalizeNumber(region.noShowRate);
      row.cancellationRateTotal += normalizeNumber(region.cancellationRate);
      row.abandonedCallRateTotal += normalizeNumber(region.abandonedCallRate);
      row.weekCount += 1;
    });
  });

  const regions = Array.from(map.values()).map((row) => ({
    entity: row.entity,
    status: row.status,
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
  const cards = [
    {
      label: "Approved Regions",
      value: current.entityCount || 0,
      meta: compareAgainst === "priorPeriod"
        ? `Prev ${comparison.entityCount || 0}`
        : `${compareAgainst.charAt(0).toUpperCase() + compareAgainst.slice(1)} comparison pending`
    },
    {
      label: "Visit Volume",
      value: current.totals?.visitVolume || 0,
      meta: compareAgainst === "priorPeriod"
        ? `${(current.totals?.visitVolume || 0) - (comparison.totals?.visitVolume || 0) >= 0 ? "+" : ""}${(current.totals?.visitVolume || 0) - (comparison.totals?.visitVolume || 0)} vs prior period`
        : `Actuals loaded • ${compareAgainst} pending`
    },
    {
      label: "Call Volume",
      value: current.totals?.callVolume || 0,
      meta: compareAgainst === "priorPeriod"
        ? `${(current.totals?.callVolume || 0) - (comparison.totals?.callVolume || 0) >= 0 ? "+" : ""}${(current.totals?.callVolume || 0) - (comparison.totals?.callVolume || 0)} vs prior period`
        : `Actuals loaded • ${compareAgainst} pending`
    },
    {
      label: "New Patients",
      value: current.totals?.newPatients || 0,
      meta: compareAgainst === "priorPeriod"
        ? `${(current.totals?.newPatients || 0) - (comparison.totals?.newPatients || 0) >= 0 ? "+" : ""}${(current.totals?.newPatients || 0) - (comparison.totals?.newPatients || 0)} vs prior period`
        : `Actuals loaded • ${compareAgainst} pending`
    },
    {
      label: "Avg No Show %",
      value: `${averageMetric(current.regions, "noShowRate").toFixed(1)}%`,
      meta: "Across approved entities"
    },
    {
      label: "Avg Cancel %",
      value: `${averageMetric(current.regions, "cancellationRate").toFixed(1)}%`,
      meta: "Across approved entities"
    },
    {
      label: "Avg Abandoned %",
      value: `${averageMetric(current.regions, "abandonedCallRate").toFixed(1)}%`,
      meta: "Across approved entities"
    }
  ];

  renderMetricCards("dashboardCards", cards);
}

function averageMetric(rows, key) {
  if (!rows || !rows.length) return 0;
  return rows.reduce((sum, row) => sum + normalizeNumber(row[key]), 0) / rows.length;
}

function getEntityMap(summary) {
  const map = {};
  (summary.regions || []).forEach((r) => {
    map[r.entity] = r;
  });
  return map;
}

function buildVariancePct(current, comparison) {
  const c = normalizeNumber(current);
  const p = normalizeNumber(comparison);
  if (!p) return null;
  return ((c - p) / p) * 100;
}

function renderDashboardEntities(current, comparison, compareAgainst, entityScope) {
  const container = byId("dashboardEntities");
  const currentMap = getEntityMap(current);
  const comparisonMap = getEntityMap(comparison);
  const entities = entityScope === "ALL" ? ENTITIES : [entityScope];

  container.innerHTML = `
    <div style="display:grid; grid-template-columns:repeat(auto-fit,minmax(260px,1fr)); gap:16px;">
      ${entities.map((entity) => {
        const brand = getBranding(entity);
        const row = currentMap[entity] || {
          entity,
          status: "missing",
          visitVolume: 0,
          callVolume: 0,
          newPatients: 0,
          noShowRate: 0,
          cancellationRate: 0,
          abandonedCallRate: 0
        };
        const prior = comparisonMap[entity] || {
          visitVolume: 0,
          callVolume: 0,
          newPatients: 0
        };

        const visitPct = buildVariancePct(row.visitVolume, prior.visitVolume);
        const callPct = buildVariancePct(row.callVolume, prior.callVolume);
        const npPct = buildVariancePct(row.newPatients, prior.newPatients);

        const fmtPct = (value) => value === null ? "n/a" : `${value >= 0 ? "+" : ""}${value.toFixed(1)}%`;

        return `
          <div style="background:#123851; border:1px solid #285a77; border-top:4px solid ${brand.accent}; border-radius:10px; padding:16px;">
            <div style="display:flex; justify-content:space-between; gap:10px; align-items:flex-start; margin-bottom:12px;">
              <div>
                <div style="display:inline-block; padding:4px 8px; border-radius:999px; font-size:12px; font-weight:bold; background:rgba(124,252,152,0.14); color:#7CFC98; margin-bottom:8px;">
                  ${row.status || "missing"}
                </div>
                <div style="font-size:18px; font-weight:bold;">${entity}</div>
                <div style="font-size:12px; opacity:0.85;">${brand.fullName}</div>
              </div>
              <div style="width:84px; height:42px; background:#fff; border-radius:8px; padding:6px; display:flex; align-items:center; justify-content:center;">
                <img src="${brand.logo}" alt="${brand.label}" style="max-width:100%; max-height:100%; object-fit:contain;" />
              </div>
            </div>

            <div style="display:grid; gap:10px;">
              <div style="display:flex; justify-content:space-between; gap:10px; border-bottom:1px solid #1d435b; padding-bottom:8px;">
                <span>Visits</span>
                <strong>${normalizeNumber(row.visitVolume)}</strong>
              </div>
              <div style="display:flex; justify-content:space-between; gap:10px; border-bottom:1px solid #1d435b; padding-bottom:8px;">
                <span>Calls</span>
                <strong>${normalizeNumber(row.callVolume)}</strong>
              </div>
              <div style="display:flex; justify-content:space-between; gap:10px; border-bottom:1px solid #1d435b; padding-bottom:8px;">
                <span>New Patients</span>
                <strong>${normalizeNumber(row.newPatients)}</strong>
              </div>
              <div style="display:flex; justify-content:space-between; gap:10px;">
                <span>No Show</span>
                <strong>${normalizeNumber(row.noShowRate).toFixed(1)}%</strong>
              </div>
              <div style="display:flex; justify-content:space-between; gap:10px;">
                <span>Cancel</span>
                <strong>${normalizeNumber(row.cancellationRate).toFixed(1)}%</strong>
              </div>
              <div style="display:flex; justify-content:space-between; gap:10px;">
                <span>Abandoned</span>
                <strong>${normalizeNumber(row.abandonedCallRate).toFixed(1)}%</strong>
              </div>
            </div>

            <div style="display:grid; grid-template-columns:repeat(3,1fr); gap:8px; margin-top:12px;">
              <div style="background:#1a4361; border:1px solid #285a77; border-radius:8px; padding:8px;">
                <div style="font-size:11px; opacity:0.8; margin-bottom:4px;">Visits ${compareAgainst === "priorPeriod" ? "vs Prior" : "Pending"}</div>
                <div style="font-weight:bold;">${compareAgainst === "priorPeriod" ? fmtPct(visitPct) : normalizeNumber(row.visitVolume)}</div>
              </div>
              <div style="background:#1a4361; border:1px solid #285a77; border-radius:8px; padding:8px;">
                <div style="font-size:11px; opacity:0.8; margin-bottom:4px;">Calls ${compareAgainst === "priorPeriod" ? "vs Prior" : "Pending"}</div>
                <div style="font-weight:bold;">${compareAgainst === "priorPeriod" ? fmtPct(callPct) : normalizeNumber(row.callVolume)}</div>
              </div>
              <div style="background:#1a4361; border:1px solid #285a77; border-radius:8px; padding:8px;">
                <div style="font-size:11px; opacity:0.8; margin-bottom:4px;">NP ${compareAgainst === "priorPeriod" ? "vs Prior" : "Pending"}</div>
                <div style="font-weight:bold;">${compareAgainst === "priorPeriod" ? fmtPct(npPct) : normalizeNumber(row.newPatients)}</div>
              </div>
            </div>
          </div>
        `;
      }).join("")}
    </div>
  `;
}

function renderDashboardAlerts(current, comparison, entityScope) {
  const container = byId("dashboardAlerts");
  const currentMap = getEntityMap(current);
  const comparisonMap = getEntityMap(comparison);
  const entities = entityScope === "ALL" ? ENTITIES : [entityScope];
  const alerts = [];

  entities.forEach((entity) => {
    const row = currentMap[entity];
    const prior = comparisonMap[entity] || {};

    if (!row) {
      alerts.push({ severity: "warning", text: `${entity} has no approved record in the selected period.` });
      return;
    }

    if (normalizeNumber(row.noShowRate) >= 6) {
      alerts.push({ severity: "bad", text: `${entity} no show rate is elevated at ${normalizeNumber(row.noShowRate).toFixed(1)}%.` });
    }

    if (normalizeNumber(row.cancellationRate) >= 8) {
      alerts.push({ severity: "bad", text: `${entity} cancellation rate is elevated at ${normalizeNumber(row.cancellationRate).toFixed(1)}%.` });
    }

    if (normalizeNumber(row.abandonedCallRate) >= 10) {
      alerts.push({ severity: "bad", text: `${entity} abandoned call rate is elevated at ${normalizeNumber(row.abandonedCallRate).toFixed(1)}%.` });
    }

    const visitDiff = normalizeNumber(row.visitVolume) - normalizeNumber(prior.visitVolume);
    if (compareAgainst === "priorPeriod" && visitDiff < -100) {
      alerts.push({ severity: "warning", text: `${entity} visit volume is down ${Math.abs(visitDiff)} vs comparison period.` });
    }
  });

  if (!alerts.length) {
    alerts.push({ severity: "good", text: "No major operational alerts for the selected period." });
  }

  container.innerHTML = alerts.map((alert) => `
    <div class="${alert.severity}" style="margin-bottom:8px; padding:12px; border:1px solid #1d435b; border-radius:8px; background:#0a2233;">
      ${alert.text}
    </div>
  `).join("");
}

function renderDashboardSnapshot(current, entityScope) {
  const container = byId("dashboardSnapshot");
  const rows = (current.regions || []).filter((r) => entityScope === "ALL" || r.entity === entityScope);

  if (!rows.length) {
    container.innerHTML = "<p>No approved entities found for the selected period.</p>";
    return;
  }

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
      <tbody>
        ${rows.map((r) => `
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
        `).join("")}
      </tbody>
    </table>
  `;
}

async function loadDashboardLanding() {
  const compareAgainst = byId("dashboardCompareAgainst").value;
  const entityScope = byId("dashboardEntityScope").value;
  const weekSets = buildWeekSets();

  const current = await loadDashboardDataForWeeks(weekSets.primaryWeeks, entityScope);

  let comparison;
  if (compareAgainst === "priorPeriod") {
    comparison = await loadDashboardDataForWeeks(weekSets.comparisonWeeks, entityScope);
  } else {
    comparison = {
      entityCount: 0,
      totals: { visitVolume: 0, callVolume: 0, newPatients: 0 },
      regions: []
    };
  }

  const summaryEl = byId("dashboardSummaryText");
  if (summaryEl) {
    summaryEl.innerHTML = `<div style="font-size:13px; opacity:0.85;">${weekSets.summary}${entityScope !== "ALL" ? ` • Scope: ${entityScope}` : " • Scope: All Entities"}</div>`;
  }

  const noticeEl = byId("dashboardBenchmarkNotice");
  if (noticeEl) {
    if (["budget", "target", "forecast"].includes(compareAgainst)) {
      noticeEl.innerHTML = `
        <div class="warning" style="padding:10px 12px; border:1px solid #1d435b; border-radius:8px; background:#0a2233;">
          ${compareAgainst.charAt(0).toUpperCase() + compareAgainst.slice(1)} comparison structure is in place, but the benchmark data layer is not wired into this front-end yet. Current view is showing actuals cleanly so the site stays stable.
        </div>
      `;
    } else {
      noticeEl.innerHTML = "";
    }
  }

  renderDashboardCards(current, comparison, compareAgainst);
  renderDashboardEntities(current, comparison, compareAgainst, entityScope);
  renderDashboardAlerts(current, comparison, entityScope);
  renderDashboardSnapshot(current, entityScope);

  setDashboardDebug({
    compareAgainst,
    entityScope,
    weekSets,
    current,
    comparison
  });
}

function renderExecutiveCards(summary) {
  const rows = summary.regions || [];
  const avg = (key) => {
    if (!rows.length) return 0;
    return rows.reduce((sum, r) => sum + normalizeNumber(r[key]), 0) / rows.length;
  };

  renderMetricCards("executiveCards", [
    { label: "Approved Regions", value: summary.entityCount || 0 },
    { label: "Visit Volume", value: summary.totals?.visitVolume || 0 },
    { label: "Call Volume", value: summary.totals?.callVolume || 0 },
    { label: "New Patients", value: summary.totals?.newPatients || 0 },
    { label: "Avg No Show %", value: `${avg("noShowRate").toFixed(1)}%` },
    { label: "Avg Cancel %", value: `${avg("cancellationRate").toFixed(1)}%` },
    { label: "Avg Abandoned %", value: `${avg("abandonedCallRate").toFixed(1)}%` }
  ]);
}

function renderExecutiveRegions(summary) {
  const container = byId("executiveRegions");

  if (!summary.regions || !summary.regions.length) {
    container.innerHTML = "<p>No approved regions found for this week.</p>";
    return;
  }

  const rows = summary.regions.map((r) => `
    <tr>
      <td>${r.entity}</td>
      <td>${r.visitVolume}</td>
      <td>${r.callVolume}</td>
      <td>${r.newPatients}</td>
      <td>${normalizeNumber(r.noShowRate).toFixed(1)}%</td>
      <td>${normalizeNumber(r.cancellationRate).toFixed(1)}%</td>
      <td>${normalizeNumber(r.abandonedCallRate).toFixed(1)}%</td>
      <td>${r.status}</td>
    </tr>
  `).join("");

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
  const weekEnding = byId("executiveWeekEnding").value;
  const result = await apiGet(`/api/executive-summary?weekEnding=${encodeURIComponent(weekEnding)}`);
  renderExecutiveCards(result);
  renderExecutiveRegions(result);
  setExecutiveDebug(result);
}

function renderTrendsCards(result) {
  const items = result.items || [];
  const latest = items.length ? items[0] : null;
  const previous = items.length > 1 ? items[1] : null;

  const formatDelta = (current, prior) => {
    const c = normalizeNumber(current);
    const p = normalizeNumber(prior);
    const diff = c - p;
    return `${c} (${diff >= 0 ? "+" : ""}${diff})`;
  };

  renderMetricCards("trendsCards", [
    { label: "Weeks Loaded", value: items.length },
    { label: "Latest Visit Volume", value: latest ? formatDelta(latest.visitVolume, previous?.visitVolume) : "-" },
    { label: "Latest Call Volume", value: latest ? formatDelta(latest.callVolume, previous?.callVolume) : "-" },
    { label: "Latest New Patients", value: latest ? formatDelta(latest.newPatients, previous?.newPatients) : "-" }
  ]);
}

function renderTrendsTable(result) {
  const wrap = byId("trendsTableWrap");
  const items = result.items || [];

  if (!items.length) {
    wrap.innerHTML = "<p>No trend data found for this entity.</p>";
    return;
  }

  const isAdmin = !!currentUser?.access?.isAdmin;

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
      <tbody>
        ${items.map((item) => `
          <tr>
            <td>${item.weekEnding}</td>
            <td>${normalizeNumber(item.visitVolume)}</td>
            <td>${normalizeNumber(item.callVolume)}</td>
            <td>${normalizeNumber(item.newPatients)}</td>
            <td>${normalizeNumber(item.noShowRate).toFixed(1)}%</td>
            <td>${normalizeNumber(item.cancellationRate).toFixed(1)}%</td>
            <td>${normalizeNumber(item.abandonedCallRate).toFixed(1)}%</td>
            <td>${item.status}</td>
            ${isAdmin ? `
              <td>
                <button class="actionBtn" data-action="override" data-entity="${item.entity}" data-week="${item.weekEnding}">Override</button>
                <button class="actionBtn" data-action="delete" data-entity="${item.entity}" data-week="${item.weekEnding}">Delete</button>
              </td>
            ` : ""}
          </tr>
        `).join("")}
      </tbody>
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
  const mode = byId("trendsRangeMode").value;
  byId("trendsWeeksWrap").style.display = mode === "recent" ? "" : "none";
  byId("trendsStartWrap").style.display = mode === "dateRange" ? "" : "none";
  byId("trendsEndWrap").style.display = mode === "dateRange" ? "" : "none";
}

async function loadTrends() {
  const entity = getSelectedTrendsEntity();
  const mode = byId("trendsRangeMode").value;
  const weeks = byId("trendsLimit").value;
  const startDate = byId("trendsStartDate").value;
  const endDate = byId("trendsEndDate").value;

  renderEntityBrand("trendsBrandWrap", entity);

  let url = `/api/trends?entity=${encodeURIComponent(entity)}`;

  if (mode === "dateRange") {
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

function getWeeklyImportFileInput() {
  return firstExistingId([
    "importFile",
    "weeklyImportFile",
    "historicalImportFile"
  ]);
}

function getBudgetImportFileInput() {
  return firstExistingId([
    "budgetImportFile",
    "budgetFile",
    "importBudgetFile",
    "budgetWorkbookFile"
  ]);
}

function getWeeklyImportButton() {
  return firstExistingId([
    "runImportBtn",
    "runWeeklyImportBtn"
  ]);
}

function getBudgetImportButton() {
  return firstExistingId([
    "runBudgetImportBtn",
    "budgetImportBtn"
  ]);
}

async function runImport() {
  if (!currentUser?.access?.isAdmin) {
    throw new Error("Admin only");
  }

  const fileInput = getWeeklyImportFileInput();
  const file = fileInput?.files && fileInput.files[0];

  if (!file) {
    throw new Error("Select a weekly workbook file first");
  }

  setImportStatus("Reading weekly workbook...");
  const fileBase64 = await readFileAsBase64(file);

  const payload = {
    fileName: file.name,
    fileBase64
  };

  setImportStatus("Importing weekly workbook...");
  setImportDebug({
    route: "/api/import-excel",
    fileName: payload.fileName
  });

  const result = await apiPost("/api/import-excel", payload);

  setImportStatus(result.message || "Weekly import completed");
  setImportDebug(result);
}

async function runBudgetImport() {
  if (!currentUser?.access?.isAdmin) {
    throw new Error("Admin only");
  }

  const fileInput = getBudgetImportFileInput();
  const file = fileInput?.files && fileInput.files[0];

  if (!file) {
    throw new Error("Select a budget workbook file first");
  }

  setImportStatus("Reading budget workbook...");
  const fileBase64 = await readFileAsBase64(file);

  const payload = {
    fileName: file.name,
    fileBase64
  };

  setImportStatus("Importing budget workbook...");
  setImportDebug({
    route: "/api/import-budget",
    fileName: payload.fileName
  });

  const result = await apiPost("/api/import-budget", payload);

  setImportStatus(result.message || "Budget import completed");
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

    if (byId("dashboardWeekEnding")) byId("dashboardWeekEnding").value = defaultWeek;
    if (byId("dashboardCustomEnd")) byId("dashboardCustomEnd").value = defaultWeek;
    if (byId("dashboardCustomStart")) byId("dashboardCustomStart").value = getDateWeeksAgo(8, defaultWeek);
    if (byId("weekEnding")) byId("weekEnding").value = defaultWeek;
    if (byId("executiveWeekEnding")) byId("executiveWeekEnding").value = defaultWeek;
    if (byId("trendsStartDate")) byId("trendsStartDate").value = getDateWeeksAgo(12, defaultWeek);
    if (byId("trendsEndDate")) byId("trendsEndDate").value = defaultWeek;

    if (byId("trendsRangeMode")) syncTrendsRangeUi();
    if (byId("dashboardPeriodType")) syncDashboardPeriodUi();

    renderEntityBrand("entryBrandWrap", byId("entitySelect") ? getSelectedEntity() : "");
    renderEntityBrand("trendsBrandWrap", byId("trendsEntitySelect") ? getSelectedTrendsEntity() : "");

    if (byId("entitySelect")) {
      byId("entitySelect").addEventListener("change", async () => {
        renderEntityBrand("entryBrandWrap", getSelectedEntity());
        try {
          await loadWeek();
        } catch (e) {
          setStatus(e.message, true);
          setDebug(String(e));
        }
      });
    }

    if (byId("weekEnding")) {
      byId("weekEnding").addEventListener("change", async () => {
        try {
          await loadWeek();
        } catch (e) {
          setStatus(e.message, true);
          setDebug(String(e));
        }
      });
    }

    if (byId("saveBtn")) {
      byId("saveBtn").addEventListener("click", async () => {
        try {
          await saveWeek();
        } catch (e) {
          setStatus(e.message, true);
          setDebug(String(e));
        }
      });
    }

    if (byId("submitBtn")) {
      byId("submitBtn").addEventListener("click", async () => {
        try {
          await submitWeek();
        } catch (e) {
          setStatus(e.message, true);
          setDebug(String(e));
        }
      });
    }

    if (byId("approveBtn")) {
      byId("approveBtn").addEventListener("click", async () => {
        try {
          await approveWeek();
        } catch (e) {
          setStatus(e.message, true);
          setDebug(String(e));
        }
      });
    }

    if (byId("navDashboardBtn")) {
      byId("navDashboardBtn").addEventListener("click", async () => {
        showDashboardView();
        try {
          await loadDashboardLanding();
        } catch (e) {
          setDashboardDebug(String(e));
        }
      });
    }

    if (byId("navEntryBtn")) {
      byId("navEntryBtn").addEventListener("click", showEntryView);
    }

    if (byId("navExecutiveBtn")) {
      byId("navExecutiveBtn").addEventListener("click", async () => {
        showExecutiveView();
        try {
          await loadExecutiveSummary();
        } catch (e) {
          setExecutiveDebug(String(e));
        }
      });
    }

    if (byId("navTrendsBtn")) {
      byId("navTrendsBtn").addEventListener("click", async () => {
        showTrendsView();
        try {
          await loadTrends();
        } catch (e) {
          setTrendsDebug(String(e));
        }
      });
    }

    if (byId("navImportBtn")) {
      byId("navImportBtn").addEventListener("click", showImportView);
    }

    if (byId("loadDashboardBtn")) {
      byId("loadDashboardBtn").addEventListener("click", async () => {
        try {
          await loadDashboardLanding();
        } catch (e) {
          setDashboardDebug(String(e));
        }
      });
    }

    if (byId("dashboardPeriodType")) {
      byId("dashboardPeriodType").addEventListener("change", async () => {
        syncDashboardPeriodUi();
        try {
          await loadDashboardLanding();
        } catch (e) {
          setDashboardDebug(String(e));
        }
      });
    }

    if (byId("dashboardCompareAgainst")) {
      byId("dashboardCompareAgainst").addEventListener("change", async () => {
        try {
          await loadDashboardLanding();
        } catch (e) {
          setDashboardDebug(String(e));
        }
      });
    }

    if (byId("dashboardEntityScope")) {
      byId("dashboardEntityScope").addEventListener("change", async () => {
        try {
          await loadDashboardLanding();
        } catch (e) {
          setDashboardDebug(String(e));
        }
      });
    }

    if (byId("dashboardWeekEnding")) {
      byId("dashboardWeekEnding").addEventListener("change", async () => {
        if (byId("dashboardPeriodType").value !== "custom") {
          try {
            await loadDashboardLanding();
          } catch (e) {
            setDashboardDebug(String(e));
          }
        }
      });
    }

    if (byId("dashboardCustomStart")) {
      byId("dashboardCustomStart").addEventListener("change", async () => {
        if (byId("dashboardPeriodType").value === "custom") {
          try {
            await loadDashboardLanding();
          } catch (e) {
            setDashboardDebug(String(e));
          }
        }
      });
    }

    if (byId("dashboardCustomEnd")) {
      byId("dashboardCustomEnd").addEventListener("change", async () => {
        if (byId("dashboardPeriodType").value === "custom") {
          try {
            await loadDashboardLanding();
          } catch (e) {
            setDashboardDebug(String(e));
          }
        }
      });
    }

    if (byId("loadExecutiveBtn")) {
      byId("loadExecutiveBtn").addEventListener("click", async () => {
        try {
          await loadExecutiveSummary();
        } catch (e) {
          setExecutiveDebug(String(e));
        }
      });
    }

    if (byId("loadTrendsBtn")) {
      byId("loadTrendsBtn").addEventListener("click", async () => {
        try {
          await loadTrends();
        } catch (e) {
          setTrendsDebug(String(e));
        }
      });
    }

    if (byId("trendsEntitySelect")) {
      byId("trendsEntitySelect").addEventListener("change", async () => {
        renderEntityBrand("trendsBrandWrap", getSelectedTrendsEntity());
        try {
          await loadTrends();
        } catch (e) {
          setTrendsDebug(String(e));
        }
      });
    }

    if (byId("trendsLimit")) {
      byId("trendsLimit").addEventListener("change", async () => {
        if (byId("trendsRangeMode").value === "recent") {
          try {
            await loadTrends();
          } catch (e) {
            setTrendsDebug(String(e));
          }
        }
      });
    }

    if (byId("trendsRangeMode")) {
      byId("trendsRangeMode").addEventListener("change", async () => {
        syncTrendsRangeUi();
        try {
          await loadTrends();
        } catch (e) {
          setTrendsDebug(String(e));
        }
      });
    }

    if (byId("trendsStartDate")) {
      byId("trendsStartDate").addEventListener("change", async () => {
        if (byId("trendsRangeMode").value === "dateRange") {
          try {
            await loadTrends();
          } catch (e) {
            setTrendsDebug(String(e));
          }
        }
      });
    }

    if (byId("trendsEndDate")) {
      byId("trendsEndDate").addEventListener("change", async () => {
        if (byId("trendsRangeMode").value === "dateRange") {
          try {
            await loadTrends();
          } catch (e) {
            setTrendsDebug(String(e));
          }
        }
      });
    }

    const weeklyImportBtn = getWeeklyImportButton();
    if (weeklyImportBtn) {
      weeklyImportBtn.addEventListener("click", async () => {
        try {
          await runImport();
        } catch (e) {
          setImportStatus(e.message, true);
          setImportDebug(String(e));
        }
      });
    }

    const budgetImportBtn = getBudgetImportButton();
    if (budgetImportBtn) {
      budgetImportBtn.addEventListener("click", async () => {
        try {
          await runBudgetImport();
        } catch (e) {
          setImportStatus(e.message, true);
          setImportDebug(String(e));
        }
      });
    }

    showDashboardView();
    await loadDashboardLanding();
  } catch (error) {
    setStatus(error.message || "Failed to load app", true);
    setDebug(String(error));
    setDashboardDebug(String(error));
    setImportDebug(String(error));
  }
})();
