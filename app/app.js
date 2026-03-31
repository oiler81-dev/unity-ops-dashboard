const ENTITY_OPTIONS = [
  { key: "LAOSS", baseEntity: "LAOSS", label: "LAOSS", isPt: false },
  { key: "NES", baseEntity: "NES", label: "NES", isPt: false },
  { key: "NES-PT", baseEntity: "NES", label: "NES-PT", isPt: true },
  { key: "SpineOne", baseEntity: "SpineOne", label: "SpineOne", isPt: false },
  { key: "SpineOne-PT", baseEntity: "SpineOne", label: "SpineOne-PT", isPt: true },
  { key: "MRO", baseEntity: "MRO", label: "MRO", isPt: false },
  { key: "MRO-PT", baseEntity: "MRO", label: "MRO-PT", isPt: true }
];

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

function setActivityDebug(data) {
  const el = byId("activityDebugOutput");
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

function clamp(value, min, max) {
  return Math.min(Math.max(value, min), max);
}

function formatWhole(value) {
  return Math.round(normalizeNumber(value)).toLocaleString();
}

function formatPercent(value, digits = 1) {
  return `${normalizeNumber(value).toFixed(digits)}%`;
}

function formatVariance(actual, target) {
  const diff = normalizeNumber(actual) - normalizeNumber(target);
  return `${diff >= 0 ? "+" : ""}${Math.round(diff).toLocaleString()}`;
}

function formatVariancePct(actual, target) {
  const t = normalizeNumber(target);
  if (!t) return "n/a";
  const pct = ((normalizeNumber(actual) - t) / t) * 100;
  return `${pct >= 0 ? "+" : ""}${pct.toFixed(1)}%`;
}

function formatToGoal(actual, target) {
  const t = normalizeNumber(target);
  if (!t) return "n/a";
  return `${((normalizeNumber(actual) / t) * 100).toFixed(1)}%`;
}

function getGoalStatus(actual, target, inverse = false) {
  const a = normalizeNumber(actual);
  const t = normalizeNumber(target);

  if (!t && !inverse) return "neutral";
  if (inverse) {
    if (a <= t) return "good";
    if (a <= t * 1.15) return "warning";
    return "bad";
  }

  const pct = t ? a / t : 0;
  if (pct >= 1) return "good";
  if (pct >= 0.92) return "warning";
  return "bad";
}

function getMetricChipClass(status) {
  if (status === "good") return "metricChip metricChipGood";
  if (status === "warning") return "metricChip metricChipWarning";
  if (status === "bad") return "metricChip metricChipBad";
  return "metricChip";
}

function getPerformanceSummary(row, compareAgainst) {
  if (compareAgainst !== "budget") {
    const abandoned = normalizeNumber(row.abandonedCallRate);
    if (abandoned >= 10) return { text: "Call access needs attention", tone: "bad" };
    if (normalizeNumber(row.visitVolume) > 0 && normalizeNumber(row.newPatients) > 0) {
      return { text: "Operating normally", tone: "good" };
    }
    return { text: "Review actuals", tone: "warning" };
  }

  const visitDelta = normalizeNumber(row.visitVolume) - normalizeNumber(row.visitVolumeBudget);
  const npDelta = normalizeNumber(row.newPatients) - normalizeNumber(row.newPatientsBudget);

  if (visitDelta >= 0 && npDelta >= 0) {
    return { text: "Above budget on visits and NP", tone: "good" };
  }

  if (visitDelta < 0 && npDelta < 0) {
    return { text: "Below budget on visits and NP", tone: "bad" };
  }

  if (visitDelta >= 0 && npDelta < 0) {
    return { text: "Visits strong, NP below budget", tone: "warning" };
  }

  if (visitDelta < 0 && npDelta >= 0) {
    return { text: "NP strong, visits below budget", tone: "warning" };
  }

  return { text: "Mixed performance", tone: "warning" };
}

function progressPercent(actual, target) {
  const t = normalizeNumber(target);
  if (!t) return 0;
  return clamp((normalizeNumber(actual) / t) * 100, 0, 140);
}

function accessPercentFromAbandoned(rate, threshold = 10) {
  const r = normalizeNumber(rate);
  const pct = 100 - (r / threshold) * 100;
  return clamp(pct, 0, 100);
}

function getTrendClass(current, comparison) {
  const diff = normalizeNumber(current) - normalizeNumber(comparison);
  if (diff > 0) return "kpi-positive";
  if (diff < 0) return "kpi-negative";
  return "kpi-neutral";
}

function getEntityOptionByKey(key) {
  return ENTITY_OPTIONS.find((x) => x.key === key) || ENTITY_OPTIONS[0];
}

function getCurrentEntityOption() {
  const selectedKey = byId("entitySelect")?.value || "LAOSS";
  return getEntityOptionByKey(selectedKey);
}

function getSelectedEntity() {
  return getCurrentEntityOption().baseEntity;
}

function getSelectedEntryLabel() {
  return getCurrentEntityOption().label;
}

function entityHasPtEntry() {
  return !!getCurrentEntityOption().isPt;
}

function formatDateTime(value) {
  if (!value) return "n/a";
  const d = new Date(value);
  if (Number.isNaN(d.getTime())) return value;
  return d.toLocaleString();
}

function calculateDerivedMetrics(values = {}) {
  const newPatients = normalizeNumber(values.newPatients);
  const surgeries = normalizeNumber(values.surgeries);
  const established = normalizeNumber(values.established);
  const noShows = normalizeNumber(values.noShows);
  const cancelled = normalizeNumber(values.cancelled);
  const totalCalls = normalizeNumber(values.totalCalls);
  const abandonedCalls = normalizeNumber(values.abandonedCalls);

  const visitVolume = newPatients + surgeries + established;
  const scheduledAppointments = visitVolume + noShows + cancelled;

  const noShowRate = scheduledAppointments > 0 ? (noShows / scheduledAppointments) * 100 : 0;
  const cancellationRate = scheduledAppointments > 0 ? (cancelled / scheduledAppointments) * 100 : 0;
  const abandonedCallRate = totalCalls > 0 ? (abandonedCalls / totalCalls) * 100 : 0;

  const ptScheduledVisits = normalizeNumber(values.ptScheduledVisits);
  const ptCancellations = normalizeNumber(values.ptCancellations);
  const ptNoShows = normalizeNumber(values.ptNoShows);
  const ptReschedules = normalizeNumber(values.ptReschedules);
  const ptTotalUnitsBilled = normalizeNumber(values.ptTotalUnitsBilled);
  const ptVisitsSeen = normalizeNumber(values.ptVisitsSeen);
  const ptWorkingDays = Math.max(1, normalizeNumber(values.ptWorkingDays || 5));

  const ptUnitsPerVisit = ptVisitsSeen > 0 ? ptTotalUnitsBilled / ptVisitsSeen : 0;
  const ptVisitsPerDay = ptWorkingDays > 0 ? ptVisitsSeen / ptWorkingDays : 0;

  return {
    newPatients,
    surgeries,
    established,
    noShows,
    cancelled,
    totalCalls,
    abandonedCalls,
    visitVolume,
    callVolume: totalCalls,
    noShowRate,
    cancellationRate,
    abandonedCallRate,

    ptScheduledVisits,
    ptCancellations,
    ptNoShows,
    ptReschedules,
    ptTotalUnitsBilled,
    ptVisitsSeen,
    ptWorkingDays,
    ptUnitsPerVisit,
    ptVisitsPerDay
  };
}

function updateDerivedDisplays() {
  const currentValues = {
    newPatients: byId("newPatients")?.value || "",
    surgeries: byId("surgeries")?.value || "",
    established: byId("established")?.value || "",
    noShows: byId("noShows")?.value || "",
    cancelled: byId("cancelled")?.value || "",
    totalCalls: byId("totalCalls")?.value || "",
    abandonedCalls: byId("abandonedCalls")?.value || "",

    ptScheduledVisits: byId("ptScheduledVisits")?.value || "",
    ptCancellations: byId("ptCancellations")?.value || "",
    ptNoShows: byId("ptNoShows")?.value || "",
    ptReschedules: byId("ptReschedules")?.value || "",
    ptTotalUnitsBilled: byId("ptTotalUnitsBilled")?.value || "",
    ptVisitsSeen: byId("ptVisitsSeen")?.value || "",
    ptWorkingDays: byId("ptWorkingDays")?.value || "5"
  };

  const derived = calculateDerivedMetrics(currentValues);

  if (byId("visitVolume")) byId("visitVolume").value = derived.visitVolume ? String(derived.visitVolume) : "0";
  if (byId("noShowRate")) byId("noShowRate").value = derived.noShowRate.toFixed(1);
  if (byId("cancellationRate")) byId("cancellationRate").value = derived.cancellationRate.toFixed(1);
  if (byId("abandonedCallRate")) byId("abandonedCallRate").value = derived.abandonedCallRate.toFixed(1);

  if (byId("ptUnitsPerVisit")) byId("ptUnitsPerVisit").value = derived.ptUnitsPerVisit.toFixed(2);
  if (byId("ptVisitsPerDay")) byId("ptVisitsPerDay").value = derived.ptVisitsPerDay.toFixed(2);
}

function renderUser(userData) {
  const label = `${userData.user.userDetails} (${userData.access.role})`;
  const el = byId("userInfo");
  if (el) {
    el.innerText = label;
  }
}

function getAllowedEntryOptions() {
  return ENTITY_OPTIONS;
}

function setupEntityDropdown() {
  const select = byId("entitySelect");
  if (!select) return;

  const options = getAllowedEntryOptions();

  select.innerHTML = "";

  options.forEach((entry) => {
    const option = document.createElement("option");
    option.value = entry.key;
    option.textContent = entry.label;
    select.appendChild(option);
  });
}

function setupTrendsEntityDropdown() {
  const select = byId("trendsEntitySelect");
  if (!select) return;

  select.innerHTML = "";

  ENTITIES.forEach((entity) => {
    const option = document.createElement("option");
    option.value = entity;
    option.textContent = entity;
    select.appendChild(option);
  });
}

function getSelectedTrendsEntity() {
  const el = byId("trendsEntitySelect");
  return el ? el.value : "LAOSS";
}

function renderForm() {
  const container = byId("kpiForm");
  if (!container) return;

  container.innerHTML = `
    <div class="nonPtField">
      <div class="formSectionBreak">
        <h4>Team Inputs</h4>
      </div>
    </div>

    <div class="nonPtField">
      <label for="newPatients">New Patients</label>
      <input type="number" id="newPatients" step="1" min="0" />
    </div>

    <div class="nonPtField">
      <label for="surgeries">Surgeries</label>
      <input type="number" id="surgeries" step="1" min="0" />
    </div>

    <div class="nonPtField">
      <label for="established">Established</label>
      <input type="number" id="established" step="1" min="0" />
    </div>

    <div class="nonPtField">
      <label for="noShows">No Shows</label>
      <input type="number" id="noShows" step="1" min="0" />
    </div>

    <div class="nonPtField">
      <label for="cancelled">Cancelled</label>
      <input type="number" id="cancelled" step="1" min="0" />
    </div>

    <div class="nonPtField">
      <label for="totalCalls">Total Calls</label>
      <input type="number" id="totalCalls" step="1" min="0" />
    </div>

    <div class="nonPtField">
      <label for="abandonedCalls">Abandoned Calls</label>
      <input type="number" id="abandonedCalls" step="1" min="0" />
    </div>

    <div class="ptField" style="display:none;">
      <div class="formSectionBreak">
        <h4>PT Inputs</h4>
      </div>
    </div>

    <div class="ptField" style="display:none;">
      <label for="ptScheduledVisits">PT Scheduled Visits</label>
      <input type="number" id="ptScheduledVisits" step="1" min="0" />
    </div>

    <div class="ptField" style="display:none;">
      <label for="ptCancellations">PT Cancellations</label>
      <input type="number" id="ptCancellations" step="1" min="0" />
    </div>

    <div class="ptField" style="display:none;">
      <label for="ptNoShows">PT No Shows</label>
      <input type="number" id="ptNoShows" step="1" min="0" />
    </div>

    <div class="ptField" style="display:none;">
      <label for="ptReschedules">PT Reschedules</label>
      <input type="number" id="ptReschedules" step="1" min="0" />
    </div>

    <div class="ptField" style="display:none;">
      <label for="ptTotalUnitsBilled">PT Total Units Billed</label>
      <input type="number" id="ptTotalUnitsBilled" step="1" min="0" />
    </div>

    <div class="ptField" style="display:none;">
      <label for="ptVisitsSeen">PT Visits Seen (wk)</label>
      <input type="number" id="ptVisitsSeen" step="1" min="0" />
    </div>

    <div class="ptField" style="display:none;">
      <label for="ptWorkingDays">PT Working Days</label>
      <input type="number" id="ptWorkingDays" step="1" min="1" value="5" />
    </div>

    <div class="nonPtField">
      <div class="formSectionBreak">
        <h4>Calculated</h4>
      </div>
    </div>

    <div class="nonPtField">
      <label for="visitVolume">Visit Volume</label>
      <input type="number" id="visitVolume" step="1" readonly />
    </div>

    <div class="nonPtField">
      <label for="noShowRate">No Show %</label>
      <input type="number" id="noShowRate" step="0.1" readonly />
    </div>

    <div class="nonPtField">
      <label for="cancellationRate">Cancellation %</label>
      <input type="number" id="cancellationRate" step="0.1" readonly />
    </div>

    <div class="nonPtField">
      <label for="abandonedCallRate">Abandoned Call %</label>
      <input type="number" id="abandonedCallRate" step="0.1" readonly />
    </div>

    <div class="ptField" style="display:none;">
      <div class="formSectionBreak">
        <h4>PT Calculated</h4>
      </div>
    </div>

    <div class="ptField" style="display:none;">
      <label for="ptUnitsPerVisit">PT Units/Visit</label>
      <input type="number" id="ptUnitsPerVisit" step="0.01" readonly />
    </div>

    <div class="ptField" style="display:none;">
      <label for="ptVisitsPerDay">PT Visits/Day (wk)</label>
      <input type="number" id="ptVisitsPerDay" step="0.01" readonly />
    </div>
  `;

  [
    "newPatients",
    "surgeries",
    "established",
    "noShows",
    "cancelled",
    "totalCalls",
    "abandonedCalls",
    "ptScheduledVisits",
    "ptCancellations",
    "ptNoShows",
    "ptReschedules",
    "ptTotalUnitsBilled",
    "ptVisitsSeen",
    "ptWorkingDays"
  ].forEach((id) => {
    const input = byId(id);
    if (input) {
      input.addEventListener("input", updateDerivedDisplays);
    }
  });

  syncEntryModeVisibility();
  updateDerivedDisplays();
}

function syncEntryModeVisibility() {
  const ptMode = entityHasPtEntry();

  document.querySelectorAll(".ptField").forEach((el) => {
    el.style.display = ptMode ? "" : "none";
  });

  document.querySelectorAll(".nonPtField").forEach((el) => {
    el.style.display = ptMode ? "none" : "";
  });

  if (ptMode) {
    [
      "newPatients",
      "surgeries",
      "established",
      "noShows",
      "cancelled",
      "totalCalls",
      "abandonedCalls",
      "visitVolume",
      "noShowRate",
      "cancellationRate",
      "abandonedCallRate"
    ].forEach((id) => {
      const input = byId(id);
      if (!input) return;
      input.value = "";
    });
  } else {
    [
      "ptScheduledVisits",
      "ptCancellations",
      "ptNoShows",
      "ptReschedules",
      "ptTotalUnitsBilled",
      "ptVisitsSeen"
    ].forEach((id) => {
      const input = byId(id);
      if (!input) return;
      input.value = "";
    });

    if (byId("ptWorkingDays")) byId("ptWorkingDays").value = "5";
    if (byId("ptUnitsPerVisit")) byId("ptUnitsPerVisit").value = "0.00";
    if (byId("ptVisitsPerDay")) byId("ptVisitsPerDay").value = "0.00";
  }
}

function mapWeeklyValuesToFormData(values) {
  const mapped = {
    newPatients: values?.newPatients ?? values?.npActual ?? "",
    surgeries: values?.surgeries ?? "",
    established: values?.established ?? "",
    noShows: values?.noShows ?? "",
    cancelled: values?.cancelled ?? "",
    totalCalls: values?.totalCalls ?? values?.callVolume ?? values?.totalCallsActual ?? "",
    abandonedCalls: values?.abandonedCalls ?? "",
    visitVolume: values?.visitVolume ?? values?.totalVisits ?? "",
    noShowRate: values?.noShowRate ?? "",
    cancellationRate: values?.cancellationRate ?? "",
    abandonedCallRate: values?.abandonedCallRate ?? values?.abandonmentRate ?? "",

    ptScheduledVisits: values?.ptScheduledVisits ?? "",
    ptCancellations: values?.ptCancellations ?? "",
    ptNoShows: values?.ptNoShows ?? "",
    ptReschedules: values?.ptReschedules ?? "",
    ptTotalUnitsBilled: values?.ptTotalUnitsBilled ?? "",
    ptVisitsSeen: values?.ptVisitsSeen ?? "",
    ptWorkingDays: values?.ptWorkingDays ?? 5,
    ptUnitsPerVisit: values?.ptUnitsPerVisit ?? "",
    ptVisitsPerDay: values?.ptVisitsPerDay ?? ""
  };

  const derived = calculateDerivedMetrics(mapped);

  return {
    newPatients: mapped.newPatients,
    surgeries: mapped.surgeries,
    established: mapped.established,
    noShows: mapped.noShows,
    cancelled: mapped.cancelled,
    totalCalls: mapped.totalCalls,
    abandonedCalls: mapped.abandonedCalls,
    visitVolume: derived.visitVolume,
    noShowRate: derived.noShowRate,
    cancellationRate: derived.cancellationRate,
    abandonedCallRate: derived.abandonedCallRate,

    ptScheduledVisits: mapped.ptScheduledVisits,
    ptCancellations: mapped.ptCancellations,
    ptNoShows: mapped.ptNoShows,
    ptReschedules: mapped.ptReschedules,
    ptTotalUnitsBilled: mapped.ptTotalUnitsBilled,
    ptVisitsSeen: mapped.ptVisitsSeen,
    ptWorkingDays: mapped.ptWorkingDays,
    ptUnitsPerVisit: derived.ptUnitsPerVisit,
    ptVisitsPerDay: derived.ptVisitsPerDay
  };
}

function setFormValues(data) {
  const mapped = mapWeeklyValuesToFormData(data || {});
  const keys = [
    "newPatients",
    "surgeries",
    "established",
    "noShows",
    "cancelled",
    "totalCalls",
    "abandonedCalls",
    "ptScheduledVisits",
    "ptCancellations",
    "ptNoShows",
    "ptReschedules",
    "ptTotalUnitsBilled",
    "ptVisitsSeen",
    "ptWorkingDays"
  ];

  keys.forEach((key) => {
    const input = byId(key);
    if (!input) return;
    input.value = mapped[key] !== null && mapped[key] !== undefined ? mapped[key] : "";
  });

  syncEntryModeVisibility();
  updateDerivedDisplays();
}

function getFormValues() {
  const ptMode = entityHasPtEntry();

  const raw = {
    newPatients: ptMode ? 0 : (byId("newPatients")?.value || ""),
    surgeries: ptMode ? 0 : (byId("surgeries")?.value || ""),
    established: ptMode ? 0 : (byId("established")?.value || ""),
    noShows: ptMode ? 0 : (byId("noShows")?.value || ""),
    cancelled: ptMode ? 0 : (byId("cancelled")?.value || ""),
    totalCalls: ptMode ? 0 : (byId("totalCalls")?.value || ""),
    abandonedCalls: ptMode ? 0 : (byId("abandonedCalls")?.value || ""),

    ptScheduledVisits: ptMode ? (byId("ptScheduledVisits")?.value || "") : 0,
    ptCancellations: ptMode ? (byId("ptCancellations")?.value || "") : 0,
    ptNoShows: ptMode ? (byId("ptNoShows")?.value || "") : 0,
    ptReschedules: ptMode ? (byId("ptReschedules")?.value || "") : 0,
    ptTotalUnitsBilled: ptMode ? (byId("ptTotalUnitsBilled")?.value || "") : 0,
    ptVisitsSeen: ptMode ? (byId("ptVisitsSeen")?.value || "") : 0,
    ptWorkingDays: ptMode ? (byId("ptWorkingDays")?.value || "5") : 5
  };

  const derived = calculateDerivedMetrics(raw);

  return {
    newPatients: derived.newPatients,
    surgeries: derived.surgeries,
    established: derived.established,
    noShows: derived.noShows,
    cancelled: derived.cancelled,
    totalCalls: derived.totalCalls,
    abandonedCalls: derived.abandonedCalls,
    visitVolume: derived.visitVolume,
    callVolume: derived.callVolume,
    noShowRate: Number(derived.noShowRate.toFixed(2)),
    cancellationRate: Number(derived.cancellationRate.toFixed(2)),
    abandonedCallRate: Number(derived.abandonedCallRate.toFixed(2)),

    ptScheduledVisits: derived.ptScheduledVisits,
    ptCancellations: derived.ptCancellations,
    ptNoShows: derived.ptNoShows,
    ptReschedules: derived.ptReschedules,
    ptTotalUnitsBilled: derived.ptTotalUnitsBilled,
    ptVisitsSeen: derived.ptVisitsSeen,
    ptWorkingDays: derived.ptWorkingDays,
    ptUnitsPerVisit: Number(derived.ptUnitsPerVisit.toFixed(2)),
    ptVisitsPerDay: Number(derived.ptVisitsPerDay.toFixed(2))
  };
}

function renderEntryAuditSummary(data) {
  const el = byId("entryAuditSummary");
  if (!el) return;

  const found = !!data?.found;
  if (!found) {
    el.innerHTML = `
      <strong>New entry.</strong><br />
      This week does not exist yet. Saving will create a new record.
    `;
    return;
  }

  el.innerHTML = `
      <strong>Created by:</strong> ${data.createdBy || "n/a"}<br />
      <strong>Created at:</strong> ${formatDateTime(data.createdAt)}<br />
      <strong>Last updated by:</strong> ${data.updatedBy || "n/a"}<br />
      <strong>Last updated at:</strong> ${formatDateTime(data.updatedAt)}<br />
      <strong>Status:</strong> ${data.status || "saved"}
    `;
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
    <div class="summaryCard ${item.className || ""}">
      <h3>${item.label}</h3>
      <div class="value">${item.value}</div>
      ${item.meta ? `<div style="margin-top:6px; font-size:12px; opacity:0.85; white-space:pre-line;">${item.meta}</div>` : ""}
    </div>
  `).join("");
}

async function loadWeek() {
  const weekEnding = byId("weekEnding")?.value || getDefaultWeekEnding();
  const entity = getSelectedEntity();
  const selectedLabel = getSelectedEntryLabel();

  renderEntityBrand("entryBrandWrap", entity);
  setStatus("Loading...");

  const result = await apiGet(
    `/api/weekly?weekEnding=${encodeURIComponent(weekEnding)}&entity=${encodeURIComponent(entity)}`
  );

  currentWeekData = result;
  setFormValues(result.values || result.data || {});
  renderEntryAuditSummary(result);

  setStatus(`Loaded ${selectedLabel} for ${weekEnding} (${result.status || "saved"})`);
  setDebug({
    selectedEntry: selectedLabel,
    baseEntity: entity,
    result
  });
}

async function saveWeek() {
  const payload = {
    weekEnding: byId("weekEnding")?.value || "",
    entity: getSelectedEntity(),
    data: getFormValues()
  };

  setStatus("Saving...");
  setDebug({
    selectedEntry: getSelectedEntryLabel(),
    baseEntity: getSelectedEntity(),
    payload
  });

  const result = await apiPost("/api/weekly-save", payload);

  setStatus(result.message || "Saved successfully");
  setDebug({
    selectedEntry: getSelectedEntryLabel(),
    baseEntity: getSelectedEntity(),
    result
  });

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

  const select = byId("entitySelect");
  if (select) {
    const option = ENTITY_OPTIONS.find((x) => x.baseEntity === entity && !x.isPt) || ENTITY_OPTIONS[0];
    select.value = option.key;
  }

  if (byId("weekEnding")) byId("weekEnding").value = weekEnding;
  await loadWeek();
  setStatus(`Admin edit mode: ${entity} ${weekEnding}`);
}

function hideAllViews() {
  const ids = ["dashboardView", "entryView", "executiveView", "trendsView", "activityView", "importView"];
  ids.forEach((id) => {
    const el = byId(id);
    if (el) el.style.display = "none";
  });
}

function setActiveNav(buttonId) {
  [
    "navDashboardBtn",
    "navEntryBtn",
    "navExecutiveBtn",
    "navTrendsBtn",
    "navActivityBtn",
    "navImportBtn"
  ].forEach((id) => {
    const btn = byId(id);
    if (!btn) return;
    btn.style.boxShadow = id === buttonId ? "inset 0 0 0 1px #f7c62f" : "";
  });
}

function showDashboardView() {
  hideAllViews();
  const el = byId("dashboardView");
  if (el) el.style.display = "";
  setActiveNav("navDashboardBtn");
}

function showEntryView() {
  hideAllViews();
  const el = byId("entryView");
  if (el) el.style.display = "";
  setActiveNav("navEntryBtn");
  renderEntityBrand("entryBrandWrap", getSelectedEntity());
  syncEntryModeVisibility();
}

function showExecutiveView() {
  hideAllViews();
  const el = byId("executiveView");
  if (el) el.style.display = "";
  setActiveNav("navExecutiveBtn");
}

function showTrendsView() {
  hideAllViews();
  const el = byId("trendsView");
  if (el) el.style.display = "";
  setActiveNav("navTrendsBtn");
  syncTrendsRangeUi();
  renderEntityBrand("trendsBrandWrap", getSelectedTrendsEntity());
}

function showActivityView() {
  hideAllViews();
  const el = byId("activityView");
  if (el) el.style.display = "";
  setActiveNav("navActivityBtn");
}

function showImportView() {
  hideAllViews();
  const el = byId("importView");
  if (el) el.style.display = "";
  setActiveNav("navImportBtn");
}

function buildWeekSets() {
  const periodType = byId("dashboardPeriodType")?.value || "lastWeek";
  const anchorWeek = byId("dashboardWeekEnding")?.value || getDefaultWeekEnding();
  const customStart = byId("dashboardCustomStart")?.value || "";
  const customEnd = byId("dashboardCustomEnd")?.value || "";

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
    const now = new Date();
    const year = now.getFullYear();
    const month = now.getMonth();

    const firstDay = new Date(year, month, 1);
    const todayIso = new Date(year, month, now.getDate()).toISOString().slice(0, 10);

    const primaryWeeks = [];
    const walker = new Date(firstDay);

    while (walker <= now) {
      if (walker.getDay() === 5) {
        primaryWeeks.push(new Date(walker).toISOString().slice(0, 10));
      }
      walker.setDate(walker.getDate() + 1);
    }

    const comparisonWeeks = primaryWeeks.map((w) => addDays(w, -28));

    return {
      primaryWeeks,
      comparisonWeeks,
      summary: `Viewing Month to Date through ${todayIso}`
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

  const primary = getPreviousWeekEnding(anchorWeek);
  return {
    primaryWeeks: [primary],
    comparisonWeeks: [getPreviousWeekEnding(primary)],
    summary: `Viewing Last Week anchored from ${anchorWeek}`
  };
}

function syncDashboardPeriodUi() {
  const periodType = byId("dashboardPeriodType")?.value || "lastWeek";
  const custom = periodType === "custom";
  const hideAnchor = periodType === "mtd";

  const weekWrap = byId("dashboardWeekWrap");
  const startWrap = byId("dashboardCustomStartWrap");
  const endWrap = byId("dashboardCustomEndWrap");

  if (weekWrap) weekWrap.style.display = custom || hideAnchor ? "none" : "";
  if (startWrap) startWrap.style.display = custom ? "" : "none";
  if (endWrap) endWrap.style.display = custom ? "" : "none";
}

function setActiveQuickPreset(preset) {
  document.querySelectorAll(".quickPresetPill").forEach((btn) => {
    btn.classList.toggle("quickPresetActive", btn.dataset.preset === preset);
  });
}

async function applyDashboardPreset(preset) {
  const periodType = byId("dashboardPeriodType");
  const anchorInput = byId("dashboardWeekEnding");
  const customStart = byId("dashboardCustomStart");
  const customEnd = byId("dashboardCustomEnd");

  const anchorWeek = anchorInput?.value || getDefaultWeekEnding();

  if (!periodType) return;

  if (preset === "lastWeek") {
    periodType.value = "lastWeek";
  } else if (preset === "mtd") {
    periodType.value = "mtd";
  } else if (preset === "rolling4") {
    periodType.value = "rolling4";
  }

  if (preset === "custom" && customStart && customEnd) {
    customEnd.value = anchorWeek;
    customStart.value = getDateWeeksAgo(8, anchorWeek);
  }

  syncDashboardPeriodUi();
  setActiveQuickPreset(preset);
  await loadDashboardLanding();
}

async function fetchExecutiveSummaryByWeek(weekEnding) {
  return apiGet(`/api/executive-summary?weekEnding=${encodeURIComponent(weekEnding)}`);
}

function aggregateExecutiveSummaries(summaries, entityScope, options = {}) {
  const includeBudget = !!options.includeBudget;
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
          surgeries: 0,
          noShowRateTotal: 0,
          cancellationRateTotal: 0,
          abandonedCallRateTotal: 0,
          weekCount: 0,
          visitVolumeBudget: 0,
          newPatientsBudget: 0,
          ptScheduledVisits: 0,
          ptCancellations: 0,
          ptNoShows: 0,
          ptReschedules: 0,
          ptTotalUnitsBilled: 0,
          ptVisitsSeen: 0,
          ptUnitsPerVisitTotal: 0,
          ptVisitsPerDayTotal: 0,
          weekEntries: []
        });
      }

      const row = map.get(region.entity);
      row.visitVolume += normalizeNumber(region.visitVolume);
      row.callVolume += normalizeNumber(region.callVolume);
      row.newPatients += normalizeNumber(region.newPatients);
      row.surgeries += normalizeNumber(region.surgeries);
      row.noShowRateTotal += normalizeNumber(region.noShowRate);
      row.cancellationRateTotal += normalizeNumber(region.cancellationRate);
      row.abandonedCallRateTotal += normalizeNumber(region.abandonedCallRate);
      row.weekCount += 1;

      if (includeBudget) {
        row.visitVolumeBudget += normalizeNumber(region.budget?.visitVolumeBudget);
        row.newPatientsBudget += normalizeNumber(region.budget?.newPatientsBudget);
      }

      row.ptScheduledVisits += normalizeNumber(region.pt?.scheduledVisits);
      row.ptCancellations += normalizeNumber(region.pt?.cancellations);
      row.ptNoShows += normalizeNumber(region.pt?.noShows);
      row.ptReschedules += normalizeNumber(region.pt?.reschedules);
      row.ptTotalUnitsBilled += normalizeNumber(region.pt?.totalUnitsBilled);
      row.ptVisitsSeen += normalizeNumber(region.pt?.visitsSeen);
      row.ptUnitsPerVisitTotal += normalizeNumber(region.pt?.unitsPerVisit);
      row.ptVisitsPerDayTotal += normalizeNumber(region.pt?.visitsPerDay);

      row.weekEntries.push({
        weekEnding: summary.weekEnding || region.weekEnding || "",
        visitVolume: normalizeNumber(region.visitVolume),
        callVolume: normalizeNumber(region.callVolume),
        newPatients: normalizeNumber(region.newPatients),
        visitVolumeBudget: normalizeNumber(region.budget?.visitVolumeBudget),
        newPatientsBudget: normalizeNumber(region.budget?.newPatientsBudget),
        ptScheduledVisits: normalizeNumber(region.pt?.scheduledVisits),
        ptCancellations: normalizeNumber(region.pt?.cancellations),
        ptNoShows: normalizeNumber(region.pt?.noShows),
        ptReschedules: normalizeNumber(region.pt?.reschedules),
        ptTotalUnitsBilled: normalizeNumber(region.pt?.totalUnitsBilled),
        ptVisitsSeen: normalizeNumber(region.pt?.visitsSeen),
        ptUnitsPerVisit: normalizeNumber(region.pt?.unitsPerVisit),
        ptVisitsPerDay: normalizeNumber(region.pt?.visitsPerDay)
      });
    });
  });

  const regions = Array.from(map.values()).map((row) => ({
    entity: row.entity,
    status: row.status,
    visitVolume: row.visitVolume,
    callVolume: row.callVolume,
    newPatients: row.newPatients,
    surgeries: row.surgeries,
    noShowRate: row.weekCount ? row.noShowRateTotal / row.weekCount : 0,
    cancellationRate: row.weekCount ? row.cancellationRateTotal / row.weekCount : 0,
    abandonedCallRate: row.weekCount ? row.abandonedCallRateTotal / row.weekCount : 0,
    visitVolumeBudget: row.visitVolumeBudget,
    newPatientsBudget: row.newPatientsBudget,
    pt: {
      scheduledVisits: row.ptScheduledVisits,
      cancellations: row.ptCancellations,
      noShows: row.ptNoShows,
      reschedules: row.ptReschedules,
      totalUnitsBilled: row.ptTotalUnitsBilled,
      visitsSeen: row.ptVisitsSeen,
      unitsPerVisit: row.weekCount ? row.ptUnitsPerVisitTotal / row.weekCount : 0,
      visitsPerDay: row.weekCount ? row.ptVisitsPerDayTotal / row.weekCount : 0
    },
    weekEntries: row.weekEntries
  }));

  const totals = {
    visitVolume: regions.reduce((sum, r) => sum + normalizeNumber(r.visitVolume), 0),
    callVolume: regions.reduce((sum, r) => sum + normalizeNumber(r.callVolume), 0),
    newPatients: regions.reduce((sum, r) => sum + normalizeNumber(r.newPatients), 0),
    surgeries: regions.reduce((sum, r) => sum + normalizeNumber(r.surgeries), 0)
  };

  const budgetTotals = includeBudget
    ? {
        visitVolumeBudget: regions.reduce((sum, r) => sum + normalizeNumber(r.visitVolumeBudget), 0),
        newPatientsBudget: regions.reduce((sum, r) => sum + normalizeNumber(r.newPatientsBudget), 0)
      }
    : {
        visitVolumeBudget: 0,
        newPatientsBudget: 0
      };

  const ptTotals = {
    scheduledVisits: regions.reduce((sum, r) => sum + normalizeNumber(r.pt?.scheduledVisits), 0),
    cancellations: regions.reduce((sum, r) => sum + normalizeNumber(r.pt?.cancellations), 0),
    noShows: regions.reduce((sum, r) => sum + normalizeNumber(r.pt?.noShows), 0),
    reschedules: regions.reduce((sum, r) => sum + normalizeNumber(r.pt?.reschedules), 0),
    totalUnitsBilled: regions.reduce((sum, r) => sum + normalizeNumber(r.pt?.totalUnitsBilled), 0),
    visitsSeen: regions.reduce((sum, r) => sum + normalizeNumber(r.pt?.visitsSeen), 0)
  };

  const ptEntityRows = regions.filter((r) => {
    const pt = r.pt || {};
    return (
      normalizeNumber(pt.scheduledVisits) > 0 ||
      normalizeNumber(pt.visitsSeen) > 0 ||
      normalizeNumber(pt.totalUnitsBilled) > 0 ||
      normalizeNumber(pt.cancellations) > 0 ||
      normalizeNumber(pt.noShows) > 0 ||
      normalizeNumber(pt.reschedules) > 0
    );
  });

  const ptAverages = {
    unitsPerVisit: ptEntityRows.length
      ? ptEntityRows.reduce((sum, r) => sum + normalizeNumber(r.pt?.unitsPerVisit), 0) / ptEntityRows.length
      : 0,
    visitsPerDay: ptEntityRows.length
      ? ptEntityRows.reduce((sum, r) => sum + normalizeNumber(r.pt?.visitsPerDay), 0) / ptEntityRows.length
      : 0
  };

  return {
    entityCount: regions.length,
    totals,
    budgetTotals,
    ptTotals,
    ptAverages,
    regions
  };
}

async function loadDashboardDataForWeeks(weeks, entityScope, options = {}) {
  const validWeeks = (weeks || []).filter(Boolean);

  if (!validWeeks.length) {
    return {
      entityCount: 0,
      totals: { visitVolume: 0, callVolume: 0, newPatients: 0, surgeries: 0 },
      budgetTotals: { visitVolumeBudget: 0, newPatientsBudget: 0 },
      ptTotals: {
        scheduledVisits: 0,
        cancellations: 0,
        noShows: 0,
        reschedules: 0,
        totalUnitsBilled: 0,
        visitsSeen: 0
      },
      ptAverages: {
        unitsPerVisit: 0,
        visitsPerDay: 0
      },
      regions: []
    };
  }

  const summaries = await Promise.all(validWeeks.map((week) => fetchExecutiveSummaryByWeek(week)));
  return aggregateExecutiveSummaries(summaries, entityScope, options);
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

function renderDashboardCards(current, comparison, compareAgainst) {
  const ptVisitsCurrent = normalizeNumber(current.ptTotals?.visitsSeen);
  const ptVisitsComparison = normalizeNumber(comparison.ptTotals?.visitsSeen);

  const avgNoShow = averageMetric(current.regions, "noShowRate");
  const avgCancel = averageMetric(current.regions, "cancellationRate");
  const avgCxnsCombined = avgNoShow + avgCancel;

  const cxnsMeta = `
<details class="metricDetails">
  <summary>View split</summary>
  <div>No Show ${avgNoShow.toFixed(1)}%</div>
  <div>Cancel ${avgCancel.toFixed(1)}%</div>
</details>`;

  if (compareAgainst === "budget") {
    const visitActual = normalizeNumber(current.totals?.visitVolume);
    const visitBudget = normalizeNumber(current.budgetTotals?.visitVolumeBudget);
    const npActual = normalizeNumber(current.totals?.newPatients);
    const npBudget = normalizeNumber(current.budgetTotals?.newPatientsBudget);

    const cards = [
      {
        label: "Visit Volume",
        value: formatWhole(visitActual),
        meta: `Budget ${formatWhole(visitBudget)}
Variance ${formatVariance(visitActual, visitBudget)} (${formatVariancePct(visitActual, visitBudget)})
To Goal ${formatToGoal(visitActual, visitBudget)}`,
        className: getTrendClass(visitActual, visitBudget)
      },
      {
        label: "PT Visits",
        value: formatWhole(ptVisitsCurrent),
        meta: `Units ${formatWhole(current.ptTotals?.totalUnitsBilled || 0)}
Avg Units/Visit ${normalizeNumber(current.ptAverages?.unitsPerVisit).toFixed(2)}`,
        className: "kpi-neutral"
      },
      {
        label: "Total Surgeries",
        value: formatWhole(current.totals?.surgeries || 0),
        meta: "Actual only",
        className: "kpi-neutral"
      },
      {
        label: "Call Volume",
        value: formatWhole(current.totals?.callVolume || 0),
        meta: "Actual only",
        className: "kpi-neutral"
      },
      {
        label: "New Patients",
        value: formatWhole(npActual),
        meta: `Budget ${formatWhole(npBudget)}
Variance ${formatVariance(npActual, npBudget)} (${formatVariancePct(npActual, npBudget)})
To Goal ${formatToGoal(npActual, npBudget)}`,
        className: getTrendClass(npActual, npBudget)
      },
      {
        label: "Avg CXNS %",
        value: `${avgCxnsCombined.toFixed(1)}%`,
        meta: cxnsMeta,
        className: "kpi-neutral"
      },
      {
        label: "Avg Abandoned %",
        value: formatPercent(averageMetric(current.regions, "abandonedCallRate")),
        meta: "Across saved entities",
        className: "kpi-neutral"
      }
    ];

    renderMetricCards("dashboardCards", cards);
    return;
  }

  const visitCurrent = normalizeNumber(current.totals?.visitVolume);
  const visitComparison = normalizeNumber(comparison.totals?.visitVolume);

  const surgeriesCurrent = normalizeNumber(current.totals?.surgeries);
  const surgeriesComparison = normalizeNumber(comparison.totals?.surgeries);

  const callCurrent = normalizeNumber(current.totals?.callVolume);
  const callComparison = normalizeNumber(comparison.totals?.callVolume);

  const npCurrent = normalizeNumber(current.totals?.newPatients);
  const npComparison = normalizeNumber(comparison.totals?.newPatients);

  const cards = [
    {
      label: "Visit Volume",
      value: visitCurrent,
      meta: `${visitCurrent - visitComparison >= 0 ? "+" : ""}${visitCurrent - visitComparison} vs prior period`,
      className: getTrendClass(visitCurrent, visitComparison)
    },
    {
      label: "PT Visits",
      value: ptVisitsCurrent,
      meta: `${ptVisitsCurrent - ptVisitsComparison >= 0 ? "+" : ""}${ptVisitsCurrent - ptVisitsComparison} vs prior period`,
      className: getTrendClass(ptVisitsCurrent, ptVisitsComparison)
    },
    {
      label: "Total Surgeries",
      value: surgeriesCurrent,
      meta: `${surgeriesCurrent - surgeriesComparison >= 0 ? "+" : ""}${surgeriesCurrent - surgeriesComparison} vs prior period`,
      className: getTrendClass(surgeriesCurrent, surgeriesComparison)
    },
    {
      label: "Call Volume",
      value: callCurrent,
      meta: `${callCurrent - callComparison >= 0 ? "+" : ""}${callCurrent - callComparison} vs prior period`,
      className: getTrendClass(callCurrent, callComparison)
    },
    {
      label: "New Patients",
      value: npCurrent,
      meta: `${npCurrent - npComparison >= 0 ? "+" : ""}${npCurrent - npComparison} vs prior period`,
      className: getTrendClass(npCurrent, npComparison)
    },
    {
      label: "Avg CXNS %",
      value: `${avgCxnsCombined.toFixed(1)}%`,
      meta: cxnsMeta,
      className: "kpi-neutral"
    },
    {
      label: "Avg Abandoned %",
      value: `${averageMetric(current.regions, "abandonedCallRate").toFixed(1)}%`,
      meta: "Across saved entities",
      className: "kpi-neutral"
    }
  ];

  renderMetricCards("dashboardCards", cards);
}

function renderDashboardEntities(current, comparison, compareAgainst, entityScope) {
  const container = byId("dashboardEntities");
  if (!container) return;

  const currentMap = getEntityMap(current);
  const comparisonMap = getEntityMap(comparison);
  const entities = entityScope === "ALL" ? ENTITIES : [entityScope];

  container.innerHTML = `
    <div class="entityCardGrid">
      ${entities.map((entity) => {
        const brand = getBranding(entity);
        const row = currentMap[entity] || {
          entity,
          status: "missing",
          visitVolume: 0,
          callVolume: 0,
          newPatients: 0,
          surgeries: 0,
          noShowRate: 0,
          cancellationRate: 0,
          abandonedCallRate: 0,
          visitVolumeBudget: 0,
          newPatientsBudget: 0,
          pt: {}
        };
        const prior = comparisonMap[entity] || {
          visitVolume: 0,
          callVolume: 0,
          newPatients: 0,
          surgeries: 0,
          visitVolumeBudget: 0,
          newPatientsBudget: 0,
          pt: {}
        };

        const visitPct = buildVariancePct(row.visitVolume, prior.visitVolume);
        const callPct = buildVariancePct(row.callVolume, prior.callVolume);
        const npPct = buildVariancePct(row.newPatients, prior.newPatients);

        const visitBudgetStatus = getGoalStatus(row.visitVolume, row.visitVolumeBudget);
        const npBudgetStatus = getGoalStatus(row.newPatients, row.newPatientsBudget);
        const callStatus = getGoalStatus(row.abandonedCallRate, 10, true);
        const summary = getPerformanceSummary(row, compareAgainst);

        const visitGoalPct = progressPercent(row.visitVolume, row.visitVolumeBudget);
        const npGoalPct = progressPercent(row.newPatients, row.newPatientsBudget);
        const accessPct = accessPercentFromAbandoned(row.abandonedCallRate, 10);

        const fmtPct = (value) => value === null ? "n/a" : `${value >= 0 ? "+" : ""}${value.toFixed(1)}%`;

        const comparisonBlock = compareAgainst === "budget"
          ? `
            <div class="entityBudgetGrid">
              <div class="entityBudgetTile">
                <div class="entityBudgetLabel">Visit Budget</div>
                <div class="entityBudgetValue">${formatWhole(row.visitVolumeBudget)}</div>
                <div class="${getMetricChipClass(visitBudgetStatus)}">
                  ${visitBudgetStatus === "good" ? "Above Goal" : visitBudgetStatus === "warning" ? "Near Goal" : "Below Goal"}
                </div>
                <div class="entityBudgetMeta">Variance ${formatVariance(row.visitVolume, row.visitVolumeBudget)}</div>
                <div class="entityBudgetMeta">To Goal ${formatToGoal(row.visitVolume, row.visitVolumeBudget)}</div>
                <div class="entityProgressWrap">
                  <div class="entityProgressLabelRow">
                    <span>Goal Progress</span>
                    <strong>${formatToGoal(row.visitVolume, row.visitVolumeBudget)}</strong>
                  </div>
                  <div class="entityProgressTrack">
                    <div class="entityProgressBar ${visitBudgetStatus === "good" ? "entityProgressGood" : visitBudgetStatus === "warning" ? "entityProgressWarning" : "entityProgressBad"}" style="width:${visitGoalPct}%"></div>
                  </div>
                </div>
              </div>

              <div class="entityBudgetTile">
                <div class="entityBudgetLabel">NP Budget</div>
                <div class="entityBudgetValue">${formatWhole(row.newPatientsBudget)}</div>
                <div class="${getMetricChipClass(npBudgetStatus)}">
                  ${npBudgetStatus === "good" ? "Above Goal" : npBudgetStatus === "warning" ? "Near Goal" : "Below Goal"}
                </div>
                <div class="entityBudgetMeta">Variance ${formatVariance(row.newPatients, row.newPatientsBudget)}</div>
                <div class="entityBudgetMeta">To Goal ${formatToGoal(row.newPatients, row.newPatientsBudget)}</div>
                <div class="entityProgressWrap">
                  <div class="entityProgressLabelRow">
                    <span>Goal Progress</span>
                    <strong>${formatToGoal(row.newPatients, row.newPatientsBudget)}</strong>
                  </div>
                  <div class="entityProgressTrack">
                    <div class="entityProgressBar ${npBudgetStatus === "good" ? "entityProgressGood" : npBudgetStatus === "warning" ? "entityProgressWarning" : "entityProgressBad"}" style="width:${npGoalPct}%"></div>
                  </div>
                </div>
              </div>
            </div>
          `
          : `
            <div class="entityCompareGrid">
              <div class="entityMiniStat">
                <span class="entityMiniLabel">Visits vs Prior</span>
                <strong>${fmtPct(visitPct)}</strong>
              </div>
              <div class="entityMiniStat">
                <span class="entityMiniLabel">Calls vs Prior</span>
                <strong>${fmtPct(callPct)}</strong>
              </div>
              <div class="entityMiniStat">
                <span class="entityMiniLabel">NP vs Prior</span>
                <strong>${fmtPct(npPct)}</strong>
              </div>
            </div>
          `;

        return `
          <div class="entityCard" style="border-top:4px solid ${brand.accent};">
            <div class="entityCardHeader">
              <div class="entityHeaderLeft">
                <div class="entityStatusRow">
                  <span class="entityStatusPill">${row.status || "missing"}</span>
                  <span class="${summary.tone === "good" ? "metricChip metricChipGood" : summary.tone === "warning" ? "metricChip metricChipWarning" : "metricChip metricChipBad"}">${summary.text}</span>
                </div>
                <div class="entityTitle">${entity}</div>
                <div class="entitySubtitle">${brand.fullName}</div>
              </div>

              <div class="entityLogoWrap">
                <img src="${brand.logo}" alt="${brand.label}" class="entityLogo" />
              </div>
            </div>

            <div class="entityTopMetrics">
              <div class="entityMetricHero">
                <span class="entityMetricLabel">Visits</span>
                <strong>${formatWhole(row.visitVolume)}</strong>
              </div>
              <div class="entityMetricHero">
                <span class="entityMetricLabel">New Patients</span>
                <strong>${formatWhole(row.newPatients)}</strong>
              </div>
              <div class="entityMetricHero">
                <span class="entityMetricLabel">Calls</span>
                <strong>${formatWhole(row.callVolume)}</strong>
              </div>
            </div>

            ${comparisonBlock}

            <div class="entityHealthRow">
              <div class="entityHealthItem">
                <span>No Show</span>
                <strong>${formatPercent(row.noShowRate)}</strong>
              </div>
              <div class="entityHealthItem">
                <span>Cancel</span>
                <strong>${formatPercent(row.cancellationRate)}</strong>
              </div>
              <div class="entityHealthItem">
                <span>Abandoned</span>
                <strong>${formatPercent(row.abandonedCallRate)}</strong>
              </div>
            </div>

            <div class="entityAccessPanel">
              <div class="entityProgressLabelRow">
                <span>Access Health</span>
                <strong>${Math.round(accessPct)}%</strong>
              </div>
              <div class="entityProgressTrack">
                <div class="entityProgressBar ${callStatus === "good" ? "entityProgressGood" : callStatus === "warning" ? "entityProgressWarning" : "entityProgressBad"}" style="width:${accessPct}%"></div>
              </div>
            </div>

            <details class="entityDetailDrawer">
              <summary>More detail</summary>
              <div class="entityDetailGrid">
                <div class="entityDetailTile">
                  <div class="entityDetailLabel">Visit Variance</div>
                  <div class="entityDetailValue">${compareAgainst === "budget" ? formatVariance(row.visitVolume, row.visitVolumeBudget) : fmtPct(visitPct)}</div>
                </div>
                <div class="entityDetailTile">
                  <div class="entityDetailLabel">NP Variance</div>
                  <div class="entityDetailValue">${compareAgainst === "budget" ? formatVariance(row.newPatients, row.newPatientsBudget) : fmtPct(npPct)}</div>
                </div>
                <div class="entityDetailTile">
                  <div class="entityDetailLabel">PT Visits</div>
                  <div class="entityDetailValue">${formatWhole(row.pt?.visitsSeen || 0)}</div>
                </div>
              </div>
            </details>
          </div>
        `;
      }).join("")}
    </div>
  `;
}

function renderDashboardAlerts(current, comparison, entityScope, compareAgainst) {
  const container = byId("dashboardAlerts");
  if (!container) return;

  const currentMap = getEntityMap(current);
  const comparisonMap = getEntityMap(comparison);
  const entities = entityScope === "ALL" ? ENTITIES : [entityScope];
  const alerts = [];

  entities.forEach((entity) => {
    const row = currentMap[entity];
    const prior = comparisonMap[entity] || {};

    if (!row) {
      alerts.push({ severity: "warning", text: `${entity} has no saved record in the selected period.` });
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

    if (compareAgainst === "priorPeriod") {
      const visitDiff = normalizeNumber(row.visitVolume) - normalizeNumber(prior.visitVolume);
      if (visitDiff < -100) {
        alerts.push({ severity: "warning", text: `${entity} visit volume is down ${Math.abs(visitDiff)} vs prior period.` });
      }
      if (visitDiff > 100) {
        alerts.push({ severity: "good", text: `${entity} visit volume is up ${Math.abs(visitDiff)} vs prior period.` });
      }
    }

    if (compareAgainst === "budget") {
      const visitGap = normalizeNumber(row.visitVolume) - normalizeNumber(row.visitVolumeBudget);
      const npGap = normalizeNumber(row.newPatients) - normalizeNumber(row.newPatientsBudget);

      if (normalizeNumber(row.visitVolumeBudget) > 0 && visitGap < 0) {
        alerts.push({ severity: "warning", text: `${entity} is ${Math.abs(Math.round(visitGap))} visits below budget.` });
      }

      if (normalizeNumber(row.newPatientsBudget) > 0 && npGap < 0) {
        alerts.push({ severity: "warning", text: `${entity} is ${Math.abs(Math.round(npGap))} new patients below budget.` });
      }

      if (visitGap > 0 && npGap > 0) {
        alerts.push({ severity: "good", text: `${entity} is above budget on visits and new patients.` });
      }
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

function renderDashboardWins(current, comparison, entityScope, compareAgainst) {
  const container = byId("dashboardWins");
  if (!container) return;

  const currentMap = getEntityMap(current);
  const comparisonMap = getEntityMap(comparison);
  const entities = entityScope === "ALL" ? ENTITIES : [entityScope];
  const wins = [];

  entities.forEach((entity) => {
    const row = currentMap[entity];
    const prior = comparisonMap[entity] || {};

    if (!row) return;

    if (compareAgainst === "budget") {
      const visitGap = normalizeNumber(row.visitVolume) - normalizeNumber(row.visitVolumeBudget);
      const npGap = normalizeNumber(row.newPatients) - normalizeNumber(row.newPatientsBudget);

      if (normalizeNumber(row.visitVolumeBudget) > 0 && visitGap > 0) {
        wins.push({
          severity: "good",
          text: `${entity} is ${Math.round(visitGap)} visits above budget.`
        });
      }

      if (normalizeNumber(row.newPatientsBudget) > 0 && npGap > 0) {
        wins.push({
          severity: "good",
          text: `${entity} is ${Math.round(npGap)} new patients above budget.`
        });
      }

      if (visitGap > 0 && npGap > 0) {
        wins.push({
          severity: "good",
          text: `${entity} is beating budget on both visits and new patients.`
        });
      }
    } else {
      const visitDiff = normalizeNumber(row.visitVolume) - normalizeNumber(prior.visitVolume);
      const npDiff = normalizeNumber(row.newPatients) - normalizeNumber(prior.newPatients);
      const callDiff = normalizeNumber(row.callVolume) - normalizeNumber(prior.callVolume);
      const ptDiff = normalizeNumber(row.pt?.visitsSeen) - normalizeNumber(prior.pt?.visitsSeen);

      if (visitDiff > 100) {
        wins.push({
          severity: "good",
          text: `${entity} visits improved by ${Math.round(visitDiff)} vs prior period.`
        });
      }

      if (npDiff > 15) {
        wins.push({
          severity: "good",
          text: `${entity} new patients improved by ${Math.round(npDiff)} vs prior period.`
        });
      }

      if (callDiff > 100) {
        wins.push({
          severity: "good",
          text: `${entity} call volume increased by ${Math.round(callDiff)} vs prior period.`
        });
      }

      if (ptDiff > 0) {
        wins.push({
          severity: "good",
          text: `${entity} PT visits improved by ${Math.round(ptDiff)} vs prior period.`
        });
      }
    }

    if (normalizeNumber(row.abandonedCallRate) > 0 && normalizeNumber(row.abandonedCallRate) < 5) {
      wins.push({
        severity: "good",
        text: `${entity} has strong call handling with only ${normalizeNumber(row.abandonedCallRate).toFixed(1)}% abandoned calls.`
      });
    }

    if (normalizeNumber(row.noShowRate) > 0 && normalizeNumber(row.noShowRate) < 4) {
      wins.push({
        severity: "good",
        text: `${entity} is keeping no-show rate low at ${normalizeNumber(row.noShowRate).toFixed(1)}%.`
      });
    }
  });

  if (!wins.length) {
    wins.push({
      severity: "warning",
      text: "No standout wins for the selected period yet."
    });
  }

  container.innerHTML = wins.map((win) => `
    <div class="${win.severity}" style="margin-bottom:8px; padding:12px; border:1px solid #1d435b; border-radius:8px; background:#0a2233;">
      ${win.text}
    </div>
  `).join("");
}

function renderDashboardSnapshot(current, entityScope, compareAgainst) {
  const container = byId("dashboardSnapshot");
  if (!container) return;

  const rows = (current.regions || []).filter((r) => entityScope === "ALL" || r.entity === entityScope);

  if (!rows.length) {
    container.innerHTML = "<p>No saved entities found for the selected period.</p>";
    return;
  }

  if (compareAgainst === "budget") {
    container.innerHTML = `
      <table class="regionTable">
        <thead>
          <tr>
            <th>Entity</th>
            <th>Visits</th>
            <th>Visit Budget</th>
            <th>Visit Var</th>
            <th>New</th>
            <th>New Budget</th>
            <th>New Var</th>
            <th>Calls</th>
            <th>PT Visits</th>
          </tr>
        </thead>
        <tbody>
          ${rows.map((r) => `
            <tr>
              <td>${r.entity}</td>
              <td>${formatWhole(r.visitVolume)}</td>
              <td>${formatWhole(r.visitVolumeBudget)}</td>
              <td>${formatVariance(r.visitVolume, r.visitVolumeBudget)}</td>
              <td>${formatWhole(r.newPatients)}</td>
              <td>${formatWhole(r.newPatientsBudget)}</td>
              <td>${formatVariance(r.newPatients, r.newPatientsBudget)}</td>
              <td>${formatWhole(r.callVolume)}</td>
              <td>${formatWhole(r.pt?.visitsSeen || 0)}</td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `;
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
          <th>PT Visits</th>
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
            <td>${formatWhole(r.visitVolume)}</td>
            <td>${formatWhole(r.callVolume)}</td>
            <td>${formatWhole(r.newPatients)}</td>
            <td>${formatWhole(r.pt?.visitsSeen || 0)}</td>
            <td>${formatPercent(r.noShowRate)}</td>
            <td>${formatPercent(r.cancellationRate)}</td>
            <td>${formatPercent(r.abandonedCallRate)}</td>
            <td>${r.status || "saved"}</td>
          </tr>
        `).join("")}
      </tbody>
    </table>
  `;
}

function renderVisitsChart(weeks, currentData, compareAgainst = "priorPeriod") {
  const ctx = document.getElementById("visitsChart");
  if (!ctx) return;

  const validWeeks = (weeks || []).filter(Boolean);
  const labels = validWeeks.map((w) => {
    const parts = String(w).split("-");
    return parts.length === 3 ? `${parts[1]}/${parts[2]}` : w;
  });

  const totalsByWeek = validWeeks.map((week) => {
    const rows = (currentData.regions || [])
      .flatMap((region) => region.weekEntries || [])
      .filter((entry) => entry.weekEnding === week);

    return {
      visitVolume: rows.reduce((sum, entry) => sum + normalizeNumber(entry.visitVolume), 0),
      callVolume: rows.reduce((sum, entry) => sum + normalizeNumber(entry.callVolume), 0),
      newPatients: rows.reduce((sum, entry) => sum + normalizeNumber(entry.newPatients), 0),
      visitVolumeBudget: rows.reduce((sum, entry) => sum + normalizeNumber(entry.visitVolumeBudget), 0)
    };
  });

  const visitData = totalsByWeek.map((row) => row.visitVolume);
  const callData = totalsByWeek.map((row) => row.callVolume);
  const npData = totalsByWeek.map((row) => row.newPatients);
  const budgetData = totalsByWeek.map((row) => row.visitVolumeBudget);

  if (window.visitsChartInstance) {
    window.visitsChartInstance.destroy();
  }

  window.visitsChartInstance = new Chart(ctx, {
    type: "line",
    data: {
      labels,
      datasets: [
        {
          label: "Visits",
          data: visitData,
          tension: 0.35,
          borderColor: "#6cb6ff",
          backgroundColor: "rgba(108, 182, 255, 0.16)",
          borderWidth: 3,
          pointRadius: 3,
          pointHoverRadius: 5,
          fill: false
        },
        {
          label: "Calls",
          data: callData,
          tension: 0.35,
          borderColor: "#f7c62f",
          backgroundColor: "rgba(247, 198, 47, 0.16)",
          borderWidth: 3,
          pointRadius: 3,
          pointHoverRadius: 5,
          fill: false,
          hidden: true
        },
        {
          label: "New Patients",
          data: npData,
          tension: 0.35,
          borderColor: "#7cfc98",
          backgroundColor: "rgba(124, 252, 152, 0.16)",
          borderWidth: 3,
          pointRadius: 3,
          pointHoverRadius: 5,
          fill: false
        },
        ...(compareAgainst === "budget"
          ? [
              {
                label: "Visit Budget",
                data: budgetData,
                tension: 0.35,
                borderColor: "#ff7d7d",
                backgroundColor: "rgba(255, 125, 125, 0.12)",
                borderDash: [6, 6],
                borderWidth: 2,
                pointRadius: 2,
                pointHoverRadius: 4,
                fill: false
              }
            ]
          : [])
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: {
        mode: "index",
        intersect: false
      },
      plugins: {
        legend: {
          labels: {
            color: "#b8d3e6",
            boxWidth: 14,
            boxHeight: 14
          }
        },
        tooltip: {
          enabled: true
        }
      },
      scales: {
        x: {
          ticks: { color: "#8eb2c9" },
          grid: { color: "rgba(255,255,255,0.06)" }
        },
        y: {
          ticks: { color: "#8eb2c9" },
          grid: { color: "rgba(255,255,255,0.06)" }
        }
      }
    }
  });
}

function renderPtTrendsChart(items = []) {
  const ctx = document.getElementById("trendsPtChart");
  if (!ctx) return;

  const rows = [...items].slice().reverse();
  const labels = rows.map((item) => {
    const parts = String(item.weekEnding || "").split("-");
    return parts.length === 3 ? `${parts[1]}/${parts[2]}` : item.weekEnding || "";
  });

  const visitsSeenData = rows.map((item) => normalizeNumber(item.ptVisitsSeen));
  const unitsData = rows.map((item) => normalizeNumber(item.ptTotalUnitsBilled));
  const scheduledData = rows.map((item) => normalizeNumber(item.ptScheduledVisits));

  if (window.trendsPtChartInstance) {
    window.trendsPtChartInstance.destroy();
  }

  window.trendsPtChartInstance = new Chart(ctx, {
    type: "line",
    data: {
      labels,
      datasets: [
        {
          label: "PT Visits Seen",
          data: visitsSeenData,
          tension: 0.35,
          borderColor: "#b49cff",
          backgroundColor: "rgba(180, 156, 255, 0.16)",
          borderWidth: 3,
          pointRadius: 3,
          pointHoverRadius: 5,
          fill: false
        },
        {
          label: "PT Units Billed",
          data: unitsData,
          tension: 0.35,
          borderColor: "#8f7cff",
          backgroundColor: "rgba(143, 124, 255, 0.14)",
          borderWidth: 3,
          pointRadius: 3,
          pointHoverRadius: 5,
          fill: false
        },
        {
          label: "PT Scheduled Visits",
          data: scheduledData,
          tension: 0.35,
          borderColor: "#d8cbff",
          backgroundColor: "rgba(216, 203, 255, 0.12)",
          borderWidth: 2,
          borderDash: [6, 6],
          pointRadius: 2,
          pointHoverRadius: 4,
          fill: false
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: {
        mode: "index",
        intersect: false
      },
      plugins: {
        legend: {
          labels: {
            color: "#ddd5ff",
            boxWidth: 14,
            boxHeight: 14
          }
        },
        tooltip: {
          enabled: true
        }
      },
      scales: {
        x: {
          ticks: { color: "#cdbfff" },
          grid: { color: "rgba(180,156,255,0.10)" }
        },
        y: {
          ticks: { color: "#cdbfff" },
          grid: { color: "rgba(180,156,255,0.10)" }
        }
      }
    }
  });
}

async function loadDashboardLanding() {
  const compareAgainst = byId("dashboardCompareAgainst")?.value || "priorPeriod";
  const entityScope = byId("dashboardEntityScope")?.value || "ALL";
  const weekSets = buildWeekSets();

  let current;
  let comparison;

  if (compareAgainst === "budget") {
    current = await loadDashboardDataForWeeks(weekSets.primaryWeeks, entityScope, { includeBudget: true });

    comparison = {
      entityCount: current.entityCount,
      totals: {
        visitVolume: current.budgetTotals.visitVolumeBudget,
        callVolume: 0,
        newPatients: current.budgetTotals.newPatientsBudget,
        surgeries: 0
      },
      budgetTotals: current.budgetTotals,
      ptTotals: current.ptTotals,
      ptAverages: current.ptAverages,
      regions: (current.regions || []).map((r) => ({
        entity: r.entity,
        visitVolume: r.visitVolumeBudget,
        callVolume: 0,
        newPatients: r.newPatientsBudget,
        surgeries: 0,
        pt: r.pt || {}
      }))
    };
  } else {
    current = await loadDashboardDataForWeeks(weekSets.primaryWeeks, entityScope, { includeBudget: false });
    comparison = await loadDashboardDataForWeeks(weekSets.comparisonWeeks, entityScope, { includeBudget: false });
  }

  const summaryEl = byId("dashboardSummaryText");
  if (summaryEl) {
    summaryEl.innerHTML = `<div style="font-size:13px; opacity:0.85;">${weekSets.summary}${entityScope !== "ALL" ? ` • Scope: ${entityScope}` : " • Scope: All Entities"}</div>`;
  }

  const noticeEl = byId("dashboardBenchmarkNotice");
  if (noticeEl) {
    if (compareAgainst === "budget") {
      noticeEl.innerHTML = `
        <div class="good" style="padding:10px 12px; border:1px solid #1d435b; border-radius:8px; background:#0a2233;">
          Budget comparison is live for Visit Volume and New Patients. Call Volume and percentage metrics remain actual-only.
        </div>
      `;
    } else {
      noticeEl.innerHTML = "";
    }
  }

  renderDashboardCards(current, comparison, compareAgainst);
  renderDashboardEntities(current, comparison, compareAgainst, entityScope);
  renderDashboardAlerts(current, comparison, entityScope, compareAgainst);
  renderDashboardWins(current, comparison, entityScope, compareAgainst);
  renderDashboardSnapshot(current, entityScope, compareAgainst);
  renderVisitsChart(weekSets.primaryWeeks, current, compareAgainst);

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
    { label: "Saved Regions", value: summary.entityCount || 0, className: "kpi-neutral" },
    { label: "Visit Volume", value: summary.totals?.visitVolume || 0, className: "kpi-neutral" },
    { label: "PT Visits", value: formatWhole(summary.ptTotals?.visitsSeen || 0), className: "kpi-neutral" },
    { label: "Total Surgeries", value: summary.totals?.surgeries || 0, className: "kpi-neutral" },
    { label: "Call Volume", value: summary.totals?.callVolume || 0, className: "kpi-neutral" },
    { label: "New Patients", value: summary.totals?.newPatients || 0, className: "kpi-neutral" },
    { label: "Avg No Show %", value: `${avg("noShowRate").toFixed(1)}%`, className: "kpi-neutral" },
    { label: "Avg Cancel %", value: `${avg("cancellationRate").toFixed(1)}%`, className: "kpi-neutral" }
  ]);

  renderMetricCards("executivePtCards", [
    {
      label: "PT Snapshot",
      value: formatWhole(summary.ptTotals?.visitsSeen || 0),
      meta: `Units ${formatWhole(summary.ptTotals?.totalUnitsBilled || 0)}
Units/Visit ${normalizeNumber(summary.ptAverages?.unitsPerVisit).toFixed(2)}`,
      className: "kpi-neutral"
    },
    {
      label: "PT Scheduling",
      value: formatWhole(summary.ptTotals?.scheduledVisits || 0),
      meta: `No Shows ${formatWhole(summary.ptTotals?.noShows || 0)}
Cancels ${formatWhole(summary.ptTotals?.cancellations || 0)}
Reschedules ${formatWhole(summary.ptTotals?.reschedules || 0)}`,
      className: "kpi-neutral"
    }
  ]);
}

function renderExecutiveRegions(summary) {
  const container = byId("executiveRegions");
  if (!container) return;

  if (!summary.regions || !summary.regions.length) {
    container.innerHTML = "<p>No saved regions found for this week.</p>";
    const ptContainer = byId("executivePtRegions");
    if (ptContainer) ptContainer.innerHTML = "<p>No PT activity found for this week.</p>";
    return;
  }

  const rows = summary.regions.map((r) => `
    <tr>
      <td>${r.entity}</td>
      <td>${r.visitVolume}</td>
      <td>${r.surgeries || 0}</td>
      <td>${r.callVolume}</td>
      <td>${r.newPatients}</td>
      <td>${normalizeNumber(r.noShowRate).toFixed(1)}%</td>
      <td>${normalizeNumber(r.cancellationRate).toFixed(1)}%</td>
      <td>${normalizeNumber(r.abandonedCallRate).toFixed(1)}%</td>
      <td>${r.status || "saved"}</td>
    </tr>
  `).join("");

  container.innerHTML = `
    <table class="regionTable">
      <thead>
        <tr>
          <th>Entity</th>
          <th>Visit</th>
          <th>Surgeries</th>
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

  const ptContainer = byId("executivePtRegions");
  if (!ptContainer) return;

  const ptRows = (summary.regions || []).filter((r) => {
    const pt = r.pt || {};
    return (
      normalizeNumber(pt.scheduledVisits) > 0 ||
      normalizeNumber(pt.visitsSeen) > 0 ||
      normalizeNumber(pt.totalUnitsBilled) > 0 ||
      normalizeNumber(pt.cancellations) > 0 ||
      normalizeNumber(pt.noShows) > 0
    );
  });

  if (!ptRows.length) {
    ptContainer.innerHTML = "<p>No PT activity found for this week.</p>";
    return;
  }

  ptContainer.innerHTML = `
    <table class="regionTable">
      <thead>
        <tr>
          <th>Entity</th>
          <th>PT Seen</th>
          <th>PT Units</th>
          <th>Units/Visit</th>
          <th>No Shows</th>
          <th>Cancels</th>
        </tr>
      </thead>
      <tbody>
        ${ptRows.map((r) => `
          <tr>
            <td>${r.entity}</td>
            <td>${formatWhole(r.pt?.visitsSeen || 0)}</td>
            <td>${formatWhole(r.pt?.totalUnitsBilled || 0)}</td>
            <td>${normalizeNumber(r.pt?.unitsPerVisit).toFixed(2)}</td>
            <td>${formatWhole(r.pt?.noShows || 0)}</td>
            <td>${formatWhole(r.pt?.cancellations || 0)}</td>
          </tr>
        `).join("")}
      </tbody>
    </table>
  `;
}

async function loadExecutiveSummary() {
  const weekEnding = byId("executiveWeekEnding")?.value || getDefaultWeekEnding();
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
    { label: "Weeks Loaded", value: items.length, className: "kpi-neutral" },
    {
      label: "Latest Visit Volume",
      value: latest ? formatDelta(latest.visitVolume, previous?.visitVolume) : "-",
      className: latest ? getTrendClass(latest.visitVolume, previous?.visitVolume) : "kpi-neutral"
    },
    {
      label: "Latest Surgeries",
      value: latest ? formatDelta(latest.surgeries, previous?.surgeries) : "-",
      className: latest ? getTrendClass(latest.surgeries, previous?.surgeries) : "kpi-neutral"
    },
    {
      label: "Latest Call Volume",
      value: latest ? formatDelta(latest.callVolume, previous?.callVolume) : "-",
      className: latest ? getTrendClass(latest.callVolume, previous?.callVolume) : "kpi-neutral"
    },
    {
      label: "Latest New Patients",
      value: latest ? formatDelta(latest.newPatients, previous?.newPatients) : "-",
      className: latest ? getTrendClass(latest.newPatients, previous?.newPatients) : "kpi-neutral"
    }
  ]);

  renderMetricCards("trendsPtCards", [
    {
      label: "Latest PT Seen",
      value: latest ? formatDelta(latest.ptVisitsSeen, previous?.ptVisitsSeen) : "-",
      className: latest ? getTrendClass(latest.ptVisitsSeen, previous?.ptVisitsSeen) : "kpi-neutral"
    },
    {
      label: "Latest PT Units",
      value: latest ? formatDelta(latest.ptTotalUnitsBilled, previous?.ptTotalUnitsBilled) : "-",
      className: latest ? getTrendClass(latest.ptTotalUnitsBilled, previous?.ptTotalUnitsBilled) : "kpi-neutral"
    },
    {
      label: "Latest PT Cancel",
      value: latest ? formatDelta(latest.ptCancellations, previous?.ptCancellations) : "-",
      className: latest ? getTrendClass(latest.ptCancellations, previous?.ptCancellations) : "kpi-neutral"
    },
    {
      label: "Latest PT No Show",
      value: latest ? formatDelta(latest.ptNoShows, previous?.ptNoShows) : "-",
      className: latest ? getTrendClass(latest.ptNoShows, previous?.ptNoShows) : "kpi-neutral"
    },
    {
      label: "Latest PT Units/Visit",
      value: latest ? `${normalizeNumber(latest.ptUnitsPerVisit).toFixed(2)}` : "-",
      className: "kpi-neutral"
    }
  ]);
}

function renderTrendsTable(result) {
  const wrap = byId("trendsTableWrap");
  if (!wrap) return;

  const items = result.items || [];

  if (!items.length) {
    wrap.innerHTML = "<p>No trend data found for this entity.</p>";
    const ptWrap = byId("trendsPtTableWrap");
    if (ptWrap) ptWrap.innerHTML = "<p>No PT trend data found for this entity.</p>";
    return;
  }

  const isAdmin = !!currentUser?.access?.isAdmin;

  wrap.innerHTML = `
    <table class="regionTable">
      <thead>
        <tr>
          <th>Week Ending</th>
          <th>New</th>
          <th>Surgeries</th>
          <th>Established</th>
          <th>No Shows</th>
          <th>Cancelled</th>
          <th>Total Calls</th>
          <th>Abandoned Calls</th>
          <th>Visit Volume</th>
          <th>No Show %</th>
          <th>Cancel %</th>
          <th>Abandoned %</th>
          <th>Status</th>
          ${isAdmin ? "<th>Actions</th>" : ""}
        </tr>
      </thead>
      <tbody>
        ${items.map((item) => `
          <tr>
            <td>${item.weekEnding}</td>
            <td>${normalizeNumber(item.newPatients)}</td>
            <td>${normalizeNumber(item.surgeries)}</td>
            <td>${normalizeNumber(item.established)}</td>
            <td>${normalizeNumber(item.noShows)}</td>
            <td>${normalizeNumber(item.cancelled)}</td>
            <td>${normalizeNumber(item.totalCalls || item.callVolume)}</td>
            <td>${normalizeNumber(item.abandonedCalls)}</td>
            <td>${normalizeNumber(item.visitVolume)}</td>
            <td>${normalizeNumber(item.noShowRate).toFixed(1)}%</td>
            <td>${normalizeNumber(item.cancellationRate).toFixed(1)}%</td>
            <td>${normalizeNumber(item.abandonedCallRate).toFixed(1)}%</td>
            <td>${item.status || "saved"}</td>
            ${isAdmin ? `
              <td>
                <button class="actionBtn" data-action="override" data-entity="${item.entity}" data-week="${item.weekEnding}">Edit</button>
                <button class="actionBtn" data-action="delete" data-entity="${item.entity}" data-week="${item.weekEnding}">Delete</button>
              </td>
            ` : ""}
          </tr>
        `).join("")}
      </tbody>
    </table>
  `;

  const ptWrap = byId("trendsPtTableWrap");
  if (ptWrap) {
    ptWrap.innerHTML = `
      <table class="regionTable">
        <thead>
          <tr>
            <th>Week Ending</th>
            <th>PT Scheduled</th>
            <th>PT Seen</th>
            <th>PT Units</th>
            <th>PT Cancel</th>
            <th>PT No Show</th>
            <th>PT Reschedules</th>
            <th>Units/Visit</th>
            <th>Visits/Day</th>
          </tr>
        </thead>
        <tbody>
          ${items.map((item) => `
            <tr>
              <td>${item.weekEnding}</td>
              <td>${normalizeNumber(item.ptScheduledVisits)}</td>
              <td>${normalizeNumber(item.ptVisitsSeen)}</td>
              <td>${normalizeNumber(item.ptTotalUnitsBilled)}</td>
              <td>${normalizeNumber(item.ptCancellations)}</td>
              <td>${normalizeNumber(item.ptNoShows)}</td>
              <td>${normalizeNumber(item.ptReschedules)}</td>
              <td>${normalizeNumber(item.ptUnitsPerVisit).toFixed(2)}</td>
              <td>${normalizeNumber(item.ptVisitsPerDay).toFixed(2)}</td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `;
  }

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
  const mode = byId("trendsRangeMode")?.value || "recent";
  const weeksWrap = byId("trendsWeeksWrap");
  const startWrap = byId("trendsStartWrap");
  const endWrap = byId("trendsEndWrap");

  if (weeksWrap) weeksWrap.style.display = mode === "recent" ? "" : "none";
  if (startWrap) startWrap.style.display = mode === "dateRange" ? "" : "none";
  if (endWrap) endWrap.style.display = mode === "dateRange" ? "" : "none";
}

async function loadTrends() {
  const entity = getSelectedTrendsEntity();
  const mode = byId("trendsRangeMode")?.value || "recent";
  const weeks = byId("trendsLimit")?.value || "12";
  const startDate = byId("trendsStartDate")?.value || "";
  const endDate = byId("trendsEndDate")?.value || "";

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
  renderPtTrendsChart(result.items || []);
  renderTrendsTable(result);
  setTrendsDebug(result);
}

function renderActivityTable(result) {
  const wrap = byId("activityTableWrap");
  if (!wrap) return;

  const items = result.items || [];

  if (!items.length) {
    wrap.innerHTML = "<p>No activity found for the selected filters.</p>";
    return;
  }

  wrap.innerHTML = `
    <table class="regionTable">
      <thead>
        <tr>
          <th>When</th>
          <th>Action</th>
          <th>Entity</th>
          <th>Week Ending</th>
          <th>User</th>
          <th>Role</th>
          <th>Summary</th>
        </tr>
      </thead>
      <tbody>
        ${items.map((item) => `
          <tr>
            <td>${formatDateTime(item.timestamp)}</td>
            <td>${item.eventType}</td>
            <td>${item.entity}</td>
            <td>${item.weekEnding || ""}</td>
            <td>${item.actorEmail || ""}</td>
            <td>${item.actorRole || ""}</td>
            <td>${item.summary || ""}</td>
          </tr>
        `).join("")}
      </tbody>
    </table>
  `;
}

async function loadActivityLog() {
  const entity = byId("activityEntityFilter")?.value || "";
  const weekEnding = byId("activityWeekFilter")?.value || "";
  const limit = byId("activityLimit")?.value || "50";

  let url = `/api/activity-log?limit=${encodeURIComponent(limit)}`;
  if (entity) url += `&entity=${encodeURIComponent(entity)}`;
  if (weekEnding) url += `&weekEnding=${encodeURIComponent(weekEnding)}`;

  const result = await apiGet(url);

  const summary = byId("activitySummary");
  if (summary) {
    summary.innerHTML = `
      Showing <strong>${result.count || 0}</strong> audit events
      ${entity ? `for <strong>${entity}</strong>` : "across all entities"}
      ${weekEnding ? ` and week ending <strong>${weekEnding}</strong>` : ""}.
    `;
  }

  renderActivityTable(result);
  setActivityDebug(result);
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
    throw new Error("Select a workbook file first");
  }

  setImportStatus("Reading workbook...");
  const fileBase64 = await readFileAsBase64(file);

  const payload = {
    fileName: file.name,
    fileBase64
  };

  setImportStatus("Importing workbook...");
  setImportDebug({
    route: "/api/import-excel",
    fileName: payload.fileName
  });

  const result = await apiPost("/api/import-excel", payload);

  setImportStatus(result.message || "Import completed");
  setImportDebug(result);
}

async function runBudgetImport() {
  throw new Error("Standalone budget import is retired. Use the main workbook import.");
}

(function injectDashboardCardStyles() {
  if (document.getElementById("dashboard-entity-card-style-block")) return;

  const style = document.createElement("style");
  style.id = "dashboard-entity-card-style-block";
  style.textContent = `
    .entityCardGrid {
      display:grid;
      grid-template-columns:repeat(auto-fit,minmax(300px,1fr));
      gap:18px;
    }
    .entityCard {
      background:linear-gradient(180deg, rgba(18,56,81,0.98) 0%, rgba(13,43,63,0.98) 100%);
      border:1px solid rgba(108,182,255,0.12);
      border-radius:18px;
      padding:18px;
      box-shadow:0 12px 24px rgba(0,0,0,0.14);
      transition:transform .18s ease, box-shadow .18s ease, border-color .18s ease;
    }
    .entityCard:hover {
      transform:translateY(-3px);
      box-shadow:0 18px 34px rgba(0,0,0,0.2);
      border-color:rgba(108,182,255,0.24);
    }
    .entityCardHeader {
      display:flex;
      justify-content:space-between;
      gap:12px;
      align-items:flex-start;
      margin-bottom:14px;
    }
    .entityHeaderLeft {
      min-width:0;
      flex:1;
    }
    .entityStatusRow {
      display:flex;
      gap:8px;
      flex-wrap:wrap;
      margin-bottom:10px;
    }
    .entityStatusPill {
      display:inline-flex;
      align-items:center;
      padding:4px 9px;
      border-radius:999px;
      background:rgba(124,252,152,0.14);
      color:#7CFC98;
      font-size:11px;
      font-weight:800;
      text-transform:uppercase;
      letter-spacing:.05em;
    }
    .metricChip {
      display:inline-flex;
      align-items:center;
      padding:4px 9px;
      border-radius:999px;
      background:rgba(255,255,255,0.07);
      color:#dcebf8;
      font-size:11px;
      font-weight:800;
      letter-spacing:.03em;
    }
    .metricChipGood {
      background:rgba(124,252,152,0.12);
      color:#7CFC98;
    }
    .metricChipWarning {
      background:rgba(247,198,47,0.14);
      color:#f7c62f;
    }
    .metricChipBad {
      background:rgba(255,125,125,0.14);
      color:#ff9a9a;
    }
    .entityTitle {
      font-size:22px;
      font-weight:900;
      line-height:1;
      margin-bottom:6px;
    }
    .entitySubtitle {
      font-size:12px;
      color:#b8d3e6;
      line-height:1.45;
    }
    .entityLogoWrap {
      width:88px;
      height:46px;
      background:#fff;
      border-radius:10px;
      padding:6px;
      display:flex;
      align-items:center;
      justify-content:center;
      flex-shrink:0;
    }
    .entityLogo {
      max-width:100%;
      max-height:100%;
      object-fit:contain;
    }
    .entityTopMetrics {
      display:grid;
      grid-template-columns:repeat(3,1fr);
      gap:10px;
      margin-bottom:14px;
    }
    .entityMetricHero {
      background:rgba(255,255,255,0.04);
      border:1px solid rgba(255,255,255,0.06);
      border-radius:14px;
      padding:12px;
    }
    .entityMetricLabel {
      display:block;
      font-size:11px;
      text-transform:uppercase;
      letter-spacing:.06em;
      color:#8eb2c9;
      margin-bottom:6px;
      font-weight:800;
    }
    .entityMetricHero strong {
      font-size:24px;
      font-weight:900;
      line-height:1;
      letter-spacing:-.03em;
    }
    .entityBudgetGrid,
    .entityCompareGrid {
      display:grid;
      gap:10px;
      margin-bottom:14px;
    }
    .entityBudgetGrid {
      grid-template-columns:repeat(2,1fr);
    }
    .entityCompareGrid {
      grid-template-columns:repeat(3,1fr);
    }
    .entityBudgetTile,
    .entityMiniStat {
      background:linear-gradient(180deg, rgba(20,67,97,0.96) 0%, rgba(17,54,79,0.96) 100%);
      border:1px solid rgba(108,182,255,0.12);
      border-radius:14px;
      padding:12px;
    }
    .entityBudgetLabel,
    .entityMiniLabel {
      display:block;
      font-size:11px;
      text-transform:uppercase;
      letter-spacing:.06em;
      color:#8eb2c9;
      margin-bottom:6px;
      font-weight:800;
    }
    .entityBudgetValue,
    .entityMiniStat strong {
      font-size:22px;
      font-weight:900;
      line-height:1;
      margin-bottom:8px;
      display:block;
    }
    .entityBudgetMeta {
      margin-top:6px;
      font-size:12px;
      color:#b8d3e6;
    }
    .entityProgressWrap {
      margin-top:10px;
    }
    .entityProgressLabelRow {
      display:flex;
      justify-content:space-between;
      gap:10px;
      align-items:center;
      font-size:12px;
      color:#b8d3e6;
      margin-bottom:6px;
    }
    .entityProgressLabelRow strong {
      color:#fff;
      font-size:12px;
    }
    .entityProgressTrack {
      width:100%;
      height:10px;
      border-radius:999px;
      background:rgba(255,255,255,0.08);
      overflow:hidden;
      position:relative;
    }
    .entityProgressBar {
      height:100%;
      border-radius:999px;
      transition:width .28s ease;
      background:linear-gradient(90deg, #6cb6ff, #74f0ff);
    }
    .entityProgressGood {
      background:linear-gradient(90deg, #33d17a, #7CFC98);
    }
    .entityProgressWarning {
      background:linear-gradient(90deg, #f1b718, #f7c62f);
    }
    .entityProgressBad {
      background:linear-gradient(90deg, #ff7d7d, #ff4f73);
    }
    .entityHealthRow {
      display:grid;
      grid-template-columns:repeat(3,1fr);
      gap:10px;
      margin-bottom:12px;
    }
    .entityHealthItem {
      display:flex;
      justify-content:space-between;
      gap:10px;
      align-items:center;
      padding:10px 12px;
      border-radius:12px;
      background:rgba(255,255,255,0.03);
      border:1px solid rgba(255,255,255,0.05);
      color:#b8d3e6;
      font-size:13px;
    }
    .entityHealthItem strong {
      color:#fff;
      font-size:14px;
    }
    .entityAccessPanel {
      margin-bottom:12px;
      padding:12px;
      border-radius:14px;
      background:rgba(255,255,255,0.03);
      border:1px solid rgba(255,255,255,0.05);
    }
    .entityDetailDrawer {
      border:1px solid rgba(255,255,255,0.06);
      border-radius:12px;
      background:rgba(7,31,51,0.34);
      overflow:hidden;
    }
    .entityDetailDrawer summary {
      cursor:pointer;
      list-style:none;
      padding:12px 14px;
      font-weight:800;
      font-size:13px;
      color:#dcebf8;
    }
    .entityDetailDrawer summary::-webkit-details-marker {
      display:none;
    }
    .entityDetailGrid {
      display:grid;
      grid-template-columns:repeat(3,1fr);
      gap:10px;
      padding:0 14px 14px;
    }
    .entityDetailTile {
      padding:12px;
      border-radius:12px;
      background:rgba(255,255,255,0.03);
      border:1px solid rgba(255,255,255,0.05);
    }
    .entityDetailLabel {
      font-size:11px;
      text-transform:uppercase;
      letter-spacing:.06em;
      color:#8eb2c9;
      margin-bottom:6px;
      font-weight:800;
    }
    .entityDetailValue {
      font-size:18px;
      font-weight:900;
      line-height:1.1;
    }
    @media (max-width: 820px) {
      .entityTopMetrics,
      .entityCompareGrid,
      .entityHealthRow,
      .entityDetailGrid {
        grid-template-columns:1fr;
      }
      .entityBudgetGrid {
        grid-template-columns:1fr;
      }
    }
  `;
  document.head.appendChild(style);
})();

(function injectActivityStyles() {
  if (document.getElementById("activity-style-block")) return;

  const style = document.createElement("style");
  style.id = "activity-style-block";
  style.textContent = `
    .activity-pill-create,
    .activity-pill-update,
    .activity-pill-delete {
      display:inline-flex;
      align-items:center;
      justify-content:center;
      min-width:72px;
      padding:4px 10px;
      border-radius:999px;
      font-size:11px;
      font-weight:800;
      text-transform:uppercase;
      letter-spacing:.05em;
    }
    .activity-pill-create {
      background:rgba(124,252,152,0.14);
      color:#7CFC98;
    }
    .activity-pill-update {
      background:rgba(247,198,47,0.14);
      color:#f7c62f;
    }
    .activity-pill-delete {
      background:rgba(255,125,125,0.14);
      color:#ff9a9a;
    }
  `;
  document.head.appendChild(style);
})();

(function injectMetricDetailsStyles() {
  if (document.getElementById("metric-details-style-block")) return;

  const style = document.createElement("style");
  style.id = "metric-details-style-block";
  style.textContent = `
    .metricDetails {
      margin-top: 6px;
    }
    .metricDetails summary {
      cursor: pointer;
      font-weight: 700;
      color: #dcebf8;
      outline: none;
    }
    .metricDetails div {
      margin-top: 4px;
      color: #b8d3e6;
    }
  `;
  document.head.appendChild(style);
})();

(async function init() {
  try {
    currentUser = await apiGet("/api/me");

    renderUser(currentUser);
    setupEntityDropdown();
    setupTrendsEntityDropdown();
    renderForm();

    const defaultWeek = getDefaultWeekEnding();

    if (byId("dashboardWeekEnding")) byId("dashboardWeekEnding").value = defaultWeek;
    if (byId("dashboardCustomEnd")) byId("dashboardCustomEnd").value = defaultWeek;
    if (byId("dashboardCustomStart")) byId("dashboardCustomStart").value = getDateWeeksAgo(8, defaultWeek);
    if (byId("weekEnding")) byId("weekEnding").value = defaultWeek;
    if (byId("executiveWeekEnding")) byId("executiveWeekEnding").value = defaultWeek;
    if (byId("trendsStartDate")) byId("trendsStartDate").value = getDateWeeksAgo(12, defaultWeek);
    if (byId("trendsEndDate")) byId("trendsEndDate").value = defaultWeek;

    const dashboardPeriodType = byId("dashboardPeriodType");
    if (dashboardPeriodType) {
      const currentWeekOption = dashboardPeriodType.querySelector('option[value="currentWeek"]');
      if (currentWeekOption) currentWeekOption.remove();
      dashboardPeriodType.value = "lastWeek";
    }

    document.querySelectorAll('.quickPresetPill[data-preset="currentWeek"]').forEach((btn) => btn.remove());

    if (byId("trendsRangeMode")) syncTrendsRangeUi();
    if (byId("dashboardPeriodType")) syncDashboardPeriodUi();

    renderEntityBrand("entryBrandWrap", getSelectedEntity());
    renderEntityBrand("trendsBrandWrap", byId("trendsEntitySelect") ? getSelectedTrendsEntity() : "");
    syncEntryModeVisibility();

    if (!currentUser?.access?.isAdmin) {
      const importBtn = byId("navImportBtn");
      if (importBtn) importBtn.style.display = "none";
    }

    if (byId("entitySelect")) {
      byId("entitySelect").addEventListener("change", async () => {
        renderEntityBrand("entryBrandWrap", getSelectedEntity());
        syncEntryModeVisibility();
        updateDerivedDisplays();
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

    if (byId("navActivityBtn")) {
      byId("navActivityBtn").addEventListener("click", async () => {
        showActivityView();
        try {
          await loadActivityLog();
        } catch (e) {
          setActivityDebug(String(e));
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
        setActiveQuickPreset("");
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
        const mode = byId("dashboardPeriodType")?.value || "lastWeek";
        if (mode !== "custom" && mode !== "mtd") {
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
        if ((byId("dashboardPeriodType")?.value || "lastWeek") === "custom") {
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
        if ((byId("dashboardPeriodType")?.value || "lastWeek") === "custom") {
          try {
            await loadDashboardLanding();
          } catch (e) {
            setDashboardDebug(String(e));
          }
        }
      });
    }

    document.querySelectorAll(".quickPresetPill").forEach((btn) => {
      btn.addEventListener("click", async () => {
        try {
          await applyDashboardPreset(btn.dataset.preset || "");
        } catch (e) {
          setDashboardDebug(String(e));
        }
      });
    });

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
        if ((byId("trendsRangeMode")?.value || "recent") === "recent") {
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
        if ((byId("trendsRangeMode")?.value || "recent") === "dateRange") {
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
        if ((byId("trendsRangeMode")?.value || "recent") === "dateRange") {
          try {
            await loadTrends();
          } catch (e) {
            setTrendsDebug(String(e));
          }
        }
      });
    }

    if (byId("loadActivityBtn")) {
      byId("loadActivityBtn").addEventListener("click", async () => {
        try {
          await loadActivityLog();
        } catch (e) {
          setActivityDebug(String(e));
        }
      });
    }

    if (byId("activityEntityFilter")) {
      byId("activityEntityFilter").addEventListener("change", async () => {
        try {
          await loadActivityLog();
        } catch (e) {
          setActivityDebug(String(e));
        }
      });
    }

    if (byId("activityWeekFilter")) {
      byId("activityWeekFilter").addEventListener("change", async () => {
        try {
          await loadActivityLog();
        } catch (e) {
          setActivityDebug(String(e));
        }
      });
    }

    if (byId("activityLimit")) {
      byId("activityLimit").addEventListener("change", async () => {
        try {
          await loadActivityLog();
        } catch (e) {
          setActivityDebug(String(e));
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

    setActiveQuickPreset("lastWeek");
    showDashboardView();
    await loadDashboardLanding();
  } catch (error) {
    setStatus(error.message || "Failed to load app", true);
    setDebug(String(error));
    setDashboardDebug(String(error));
    setTrendsDebug(String(error));
    setActivityDebug(String(error));
    setImportDebug(String(error));
  }
})();