const ENTITIES = ["LAOSS", "NES", "SpineOne", "MRO"];

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

function getDefaultWeekEnding() {
  const d = new Date();
  const diff = (5 - d.getDay() + 7) % 7;
  d.setDate(d.getDate() + diff);
  return d.toISOString().slice(0, 10);
}

function getDateWeeksAgo(weeksAgo) {
  const d = new Date();
  d.setDate(d.getDate() - weeksAgo * 7);
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
}

function showEntryView() {
  hideAllViews();
  document.getElementById("entryView").style.display = "";
}

function showExecutiveView() {
  hideAllViews();
  document.getElementById("executiveView").style.display = "";
}

function showTrendsView() {
  hideAllViews();
  document.getElementById("trendsView").style.display = "";
  syncTrendsRangeUi();
}

function showImportView() {
  hideAllViews();
  document.getElementById("importView").style.display = "";
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

function renderDashboardCards(current, previous) {
  const container = document.getElementById("dashboardCards");
  container.innerHTML = "";

  const prevTotals = previous?.totals || {};
  const currentRegions = current?.regions || [];
  const avg = (key) => {
    if (!currentRegions.length) return 0;
    return (
      currentRegions.reduce((sum, r) => sum + normalizeNumber(r[key]), 0) /
      currentRegions.length
    );
  };

  const currentCards = [
    {
      label: "Approved Regions",
      value: current?.entityCount || 0,
      sub: `${(previous?.entityCount || 0) > 0 ? `Prev ${previous.entityCount}` : "Prev 0"}`
    },
    {
      label: "Visit Volume",
      value: current?.totals?.visitVolume || 0,
      sub: formatDeltaOnly(current?.totals?.visitVolume || 0, prevTotals.visitVolume || 0)
    },
    {
      label: "Call Volume",
      value: current?.totals?.callVolume || 0,
      sub: formatDeltaOnly(current?.totals?.callVolume || 0, prevTotals.callVolume || 0)
    },
    {
      label: "New Patients",
      value: current?.totals?.newPatients || 0,
      sub: formatDeltaOnly(current?.totals?.newPatients || 0, prevTotals.newPatients || 0)
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

  currentCards.forEach((item) => {
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

function formatDeltaOnly(current, prior) {
  const diff = normalizeNumber(current) - normalizeNumber(prior);
  return `${diff >= 0 ? "+" : ""}${diff} vs prior week`;
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

function renderDashboardEntities(current, previous) {
  const container = document.getElementById("dashboardEntities");
  container.innerHTML = "";

  const currentMap = getEntityMap(current);
  const previousMap = getEntityMap(previous);

  ENTITIES.forEach((entity) => {
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

    const prior = previousMap[entity] || {
      visitVolume: 0,
      callVolume: 0,
      newPatients: 0
    };

    const card = document.createElement("div");
    card.className = "dashboardEntityCard";
    card.innerHTML = `
      <div class="dashboardStatusBadge ${statusClass(row.status)}">${row.status}</div>
      <h4>${entity}</h4>
      <div class="dashboardEntityMeta">Target / forecast wiring pending</div>

      <div class="dashboardMetricRow">
        <span>Visits</span>
        <span class="metricValue">${normalizeNumber(row.visitVolume)} <span class="metricSub">${formatDeltaOnly(row.visitVolume, prior.visitVolume)}</span></span>
      </div>

      <div class="dashboardMetricRow">
        <span>Calls</span>
        <span class="metricValue">${normalizeNumber(row.callVolume)} <span class="metricSub">${formatDeltaOnly(row.callVolume, prior.callVolume)}</span></span>
      </div>

      <div class="dashboardMetricRow">
        <span>New Patients</span>
        <span class="metricValue">${normalizeNumber(row.newPatients)} <span class="metricSub">${formatDeltaOnly(row.newPatients, prior.newPatients)}</span></span>
      </div>

      <div class="dashboardMetricRow">
        <span>No Show</span>
        <span class="metricValue">${normalizeNumber(row.noShowRate)}%</span>
      </div>

      <div class="dashboardMetricRow">
        <span>Cancel</span>
        <span class="metricValue">${normalizeNumber(row.cancellationRate)}%</span>
      </div>

      <div class="dashboardMetricRow">
        <span>Abandoned</span>
        <span class="metricValue">${normalizeNumber(row.abandonedCallRate)}%</span>
      </div>
    `;
    container.appendChild(card);
  });
}

function buildDashboardAlerts(current, previous) {
  const alerts = [];
  const currentMap = getEntityMap(current);
  const previousMap = getEntityMap(previous);

  ENTITIES.forEach((entity) => {
    const row = currentMap[entity];
    const prior = previousMap[entity] || {};

    if (!row) {
      alerts.push({
        severity: "yellow",
        text: `${entity} has no approved record for this week.`
      });
      return;
    }

    if (String(row.status || "").toLowerCase() !== "approved") {
      alerts.push({
        severity: "yellow",
        text: `${entity} is not approved for this week.`
      });
    }

    if (normalizeNumber(row.noShowRate) >= 6) {
      alerts.push({
        severity: "red",
        text: `${entity} no show rate is elevated at ${row.noShowRate}%.`
      });
    }

    if (normalizeNumber(row.cancellationRate) >= 8) {
      alerts.push({
        severity: "red",
        text: `${entity} cancellation rate is elevated at ${row.cancellationRate}%.`
      });
    }

    if (normalizeNumber(row.abandonedCallRate) >= 10) {
      alerts.push({
        severity: "red",
        text: `${entity} abandoned call rate is elevated at ${row.abandonedCallRate}%.`
      });
    }

    const visitDrop = normalizeNumber(row.visitVolume) - normalizeNumber(prior.visitVolume);
    if (visitDrop < -100) {
      alerts.push({
        severity: "yellow",
        text: `${entity} visit volume is down ${Math.abs(visitDrop)} vs prior week.`
      });
    }
  });

  if (!alerts.length) {
    alerts.push({
      severity: "green",
      text: "No major operational alerts for this week."
    });
  }

  return alerts;
}

function renderDashboardAlerts(current, previous) {
  const container = document.getElementById("dashboardAlerts");
  const alerts = buildDashboardAlerts(current, previous);

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

function renderDashboardSnapshot(current) {
  const container = document.getElementById("dashboardSnapshot");
  const regions = current?.regions || [];

  if (!regions.length) {
    container.innerHTML = `<div class="dashboardEmpty">No approved entities found for this week.</div>`;
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
        <td>${normalizeNumber(r.noShowRate)}%</td>
        <td>${normalizeNumber(r.cancellationRate)}%</td>
        <td>${normalizeNumber(r.abandonedCallRate)}%</td>
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
  const weekEnding = document.getElementById("dashboardWeekEnding").value;
  const previousWeekEnding = getPreviousWeekEnding(weekEnding);

  const current = await apiGet(`/api/executive-summary?weekEnding=${encodeURIComponent(weekEnding)}`);
  const previous = await apiGet(`/api/executive-summary?weekEnding=${encodeURIComponent(previousWeekEnding)}`);

  renderDashboardCards(current, previous);
  renderDashboardEntities(current, previous);
  renderDashboardAlerts(current, previous);
  renderDashboardSnapshot(current);
  setDashboardDebug({
    currentWeek: current,
    previousWeek: previous
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
    document.getElementById("weekEnding").value = defaultWeek;
    document.getElementById("executiveWeekEnding").value = defaultWeek;
    document.getElementById("trendsStartDate").value = getDateWeeksAgo(12);
    document.getElementById("trendsEndDate").value = defaultWeek;

    syncTrendsRangeUi();

    document.getElementById("entitySelect").addEventListener("change", async () => {
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
