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

const OPERATIONS_NARRATIVE_PROMPTS = [
  "What went well this week?",
  "What could have made it even better?",
  "What were the major red flags?",
  "Any adjustments or updates for next week?",
  "Any additional insight?"
];

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

  if (data !== null) return data;

  return { ok: true, raw: text };
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

function setActivityDebug(data) {
  const el = byId("activityDebugOutput");
  if (!el) return;
  el.textContent = typeof data === "string" ? data : JSON.stringify(data, null, 2);
}

function setImportStatus(message, isError = false) {
  const el = firstExistingId([
    "importStatusMessage",
    "adminImportStatusMessage",
    "budgetImportStatusMessage"
  ]);
  if (!el) return;
  el.textContent = message;
  el.style.color = isError ? "#ff8a8a" : "#7CFC98";
}

function setImportDebug(data) {
  const el = firstExistingId(["importDebugOutput", "adminImportDebugOutput"]);
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

function clamp(value, min, max) {
  return Math.min(Math.max(value, min), max);
}

function formatWhole(value) {
  return Math.round(normalizeNumber(value)).toLocaleString();
}

function formatCurrency(value, decimals = 0) {
  const n = normalizeNumber(value);
  if (Math.abs(n) >= 1_000_000) {
    const m = n / 1_000_000;
    return "$" + (m % 1 === 0 ? m.toFixed(0) : m.toFixed(1)) + "M";
  }
  if (Math.abs(n) >= 10_000) {
    const k = n / 1_000;
    return "$" + (k % 1 === 0 ? k.toFixed(0) : k.toFixed(1)) + "K";
  }
  return n.toLocaleString(undefined, {
    style: "currency",
    currency: "USD",
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals
  });
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

function formatDateTime(value) {
  if (!value) return "n/a";
  const d = new Date(value);
  if (Number.isNaN(d.getTime())) return value;
  return d.toLocaleString();
}

function getTrendClass(current, comparison) {
  const diff = normalizeNumber(current) - normalizeNumber(comparison);
  if (diff > 0) return "kpi-positive";
  if (diff < 0) return "kpi-negative";
  return "kpi-neutral";
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

function isSpineOneBaseEntity() {
  return getSelectedEntity() === "SpineOne" && !entityHasPtEntry();
}

function getSelectedTrendsEntity() {
  const el = byId("trendsEntitySelect");
  return el ? el.value : "LAOSS";
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
    ptVisitsPerDay,

    ptoDays: normalizeNumber(values.ptoDays),
    cashCollected: normalizeNumber(values.cashCollected),
    piNp: normalizeNumber(values.piNp),
    piCashCollection: normalizeNumber(values.piCashCollection),
    operationsNarrative: String(values.operationsNarrative || "").trim()
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
  if (el) el.innerText = label;
}

function setupEntityDropdown() {
  const select = byId("entitySelect");
  if (!select) return;

  select.innerHTML = "";
  ENTITY_OPTIONS.forEach((entry) => {
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

function renderNarrativePromptBox() {
  return `
    <div class="allEntryField narrativeFieldBlock" style="grid-column:1 / -1;">
      <div class="formSectionBreak">
        <h4>Operator Notes</h4>
      </div>

      <div class="notesShell">
        <div class="notesHeader">
          <div class="notesEyebrow">Weekly Context</div>
          <div class="notesTitle">Add color behind the numbers</div>
          <div class="notesSubtext">Use this section to explain what drove performance, issues, changes, wins, and next-step planning.</div>
        </div>

        <div class="notesPromptGrid">
          ${OPERATIONS_NARRATIVE_PROMPTS.map((q) => `
            <div class="notesPromptCard">
              <div class="notesPromptBullet"></div>
              <div class="notesPromptText">${q}</div>
            </div>
          `).join("")}
        </div>

        <label for="operationsNarrative">Weekly Notes / Operations Narrative</label>
        <textarea id="operationsNarrative" rows="10" placeholder="Enter weekly notes here..."></textarea>
      </div>
    </div>
  `;
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

    <div class="nonPtField">
      <label for="cashCollected">Cash Collected</label>
      <input type="number" id="cashCollected" step="0.01" min="0" />
    </div>

    <div class="allEntryField">
      <label for="ptoDays">PTO Days</label>
      <input type="number" id="ptoDays" step="0.5" min="0" />
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

    <div class="spineOnlyField" style="display:none;">
      <div class="formSectionBreak">
        <h4>SpineOne PI</h4>
      </div>
    </div>

    <div class="spineOnlyField" style="display:none;">
      <label for="piNp">PI NP</label>
      <input type="number" id="piNp" step="1" min="0" />
    </div>

    <div class="spineOnlyField" style="display:none;">
      <label for="piCashCollection">PI Cash Collection</label>
      <input type="number" id="piCashCollection" step="0.01" min="0" />
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

    ${renderNarrativePromptBox()}
  `;

  [
    "newPatients",
    "surgeries",
    "established",
    "noShows",
    "cancelled",
    "totalCalls",
    "abandonedCalls",
    "cashCollected",
    "ptoDays",
    "piNp",
    "piCashCollection",
    "ptScheduledVisits",
    "ptCancellations",
    "ptNoShows",
    "ptReschedules",
    "ptTotalUnitsBilled",
    "ptVisitsSeen",
    "ptWorkingDays"
  ].forEach((id) => {
    const input = byId(id);
    if (input) input.addEventListener("input", updateDerivedDisplays);
  });

  syncEntryModeVisibility();
  updateDerivedDisplays();
}

function syncEntryModeVisibility() {
  const ptMode = entityHasPtEntry();
  const spineMode = isSpineOneBaseEntity();

  document.querySelectorAll(".ptField").forEach((el) => {
    el.style.display = ptMode ? "" : "none";
  });

  document.querySelectorAll(".nonPtField").forEach((el) => {
    el.style.display = ptMode ? "none" : "";
  });

  document.querySelectorAll(".spineOnlyField").forEach((el) => {
    el.style.display = !ptMode && spineMode ? "" : "none";
  });

  document.querySelectorAll(".allEntryField").forEach((el) => {
    el.style.display = "";
  });
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
    cashCollected: values?.cashCollected ?? values?.cashActual ?? "",
    ptoDays: values?.ptoDays ?? "",
    piNp: values?.piNp ?? "",
    piCashCollection: values?.piCashCollection ?? "",
    operationsNarrative: values?.operationsNarrative ?? "",
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
    ...mapped,
    visitVolume: derived.visitVolume,
    noShowRate: derived.noShowRate,
    cancellationRate: derived.cancellationRate,
    abandonedCallRate: derived.abandonedCallRate,
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
    "cashCollected",
    "ptoDays",
    "piNp",
    "piCashCollection",
    "operationsNarrative",
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
    cashCollected: ptMode ? 0 : (byId("cashCollected")?.value || ""),
    ptoDays: byId("ptoDays")?.value || "",
    piNp: ptMode ? 0 : (byId("piNp")?.value || ""),
    piCashCollection: ptMode ? 0 : (byId("piCashCollection")?.value || ""),
    operationsNarrative: byId("operationsNarrative")?.value || "",
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
    cashCollected: derived.cashCollected,
    ptoDays: derived.ptoDays,
    piNp: derived.piNp,
    piCashCollection: derived.piCashCollection,
    operationsNarrative: derived.operationsNarrative,
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
      ${item.movement ? item.movement : ""}
      ${item.meta ? `<div class="kpiMeta" style="white-space:pre-line;">${item.meta}</div>` : ""}
    </div>
  `).join("");
}

function buildKpiMovement(current, comparison, formatFn) {
  const c = normalizeNumber(current);
  const p = normalizeNumber(comparison);
  if (!p && !c) return "";

  const diff = c - p;

  if (diff === 0) {
    return `<div class="kpiMovement kpiMovementFlat"><span class="kpiArrow">—</span> No change</div>`;
  }

  const isPositive = diff > 0;
  const arrow = diff > 0 ? "▲" : "▼";
  const cls = isPositive ? "kpiMovementUp" : "kpiMovementDown";
  const label = formatFn ? formatFn(Math.abs(diff)) : `${Math.abs(((diff / (p || 1)) * 100)).toFixed(1)}%`;

  return `<div class="kpiMovement ${cls}"><span class="kpiArrow">${arrow}</span> ${label}</div>`;
}

function buildVariancePct(current, comparison) {
  const c = normalizeNumber(current);
  const p = normalizeNumber(comparison);
  if (!p) return null;
  return ((c - p) / p) * 100;
}

function getEntityMap(summary) {
  const map = {};
  (summary.regions || []).forEach((r) => {
    map[r.entity] = r;
  });
  return map;
}

function averageMetric(rows, key) {
  if (!rows || !rows.length) return 0;
  return rows.reduce((sum, row) => sum + normalizeNumber(row[key]), 0) / rows.length;
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

  if (visitDelta >= 0 && npDelta >= 0) return { text: "Above budget on visits and NP", tone: "good" };
  if (visitDelta < 0 && npDelta < 0) return { text: "Below budget on visits and NP", tone: "bad" };
  if (visitDelta >= 0 && npDelta < 0) return { text: "Visits strong, NP below budget", tone: "warning" };
  if (visitDelta < 0 && npDelta >= 0) return { text: "NP strong, visits below budget", tone: "warning" };

  return { text: "Mixed performance", tone: "warning" };
}

function renderDashboardCards(current, comparison, compareAgainst, entityScope) {
  const ptVisitsCurrent = normalizeNumber(current.ptTotals?.visitsSeen);
  const ptVisitsComparison = normalizeNumber(comparison.ptTotals?.visitsSeen);

  const avgNoShow = averageMetric(current.regions, "noShowRate");
  const avgCancel = averageMetric(current.regions, "cancellationRate");
  const avgCxnsCombined = avgNoShow + avgCancel;

  const visitCurrent = normalizeNumber(current.totals?.visitVolume);
  const visitComparison = normalizeNumber(comparison.totals?.visitVolume);
  const visitVariance = visitCurrent - visitComparison;

  const isVsBudget = compareAgainst === "budget";

  const cards = [
    {
      label: "Visit Volume",
      value: formatWhole(visitCurrent),
      movement: isVsBudget
        ? buildKpiMovement(visitCurrent, normalizeNumber(current.budgetTotals?.visitVolumeBudget), formatWhole)
        : buildKpiMovement(visitCurrent, visitComparison, formatWhole),
      meta: isVsBudget
        ? `Budget ${formatWhole(current.budgetTotals?.visitVolumeBudget || 0)}`
        : `${visitVariance >= 0 ? "+" : ""}${formatWhole(visitVariance)} vs prior`,
      className: isVsBudget
        ? getTrendClass(visitCurrent, current.budgetTotals?.visitVolumeBudget)
        : getTrendClass(visitCurrent, visitComparison)
    },
    ...(entityHasPtData(entityScope) || entityScope === "ALL" ? [{
      label: "PT Visits",
      value: formatWhole(ptVisitsCurrent),
      movement: buildKpiMovement(ptVisitsCurrent, ptVisitsComparison, formatWhole),
      meta: `${ptVisitsCurrent - ptVisitsComparison >= 0 ? "+" : ""}${formatWhole(ptVisitsCurrent - ptVisitsComparison)} vs prior`,
      className: getTrendClass(ptVisitsCurrent, ptVisitsComparison)
    }] : []),
    {
      label: "Cash Collected",
      value: formatCurrency(current.totals?.cashCollected || 0),
      meta: isVsBudget ? "Actual only" : "Across selected period",
      className: "kpi-neutral"
    },
    {
      label: "PTO Days",
      value: formatWhole(current.totals?.ptoDays || 0),
      meta: "Across selected scope",
      className: "kpi-neutral"
    },
    {
      label: "Total Surgeries",
      value: formatWhole(current.totals?.surgeries || 0),
      meta: isVsBudget ? "Actual only" : "Across selected period",
      className: "kpi-neutral"
    },
    {
      label: "Call Volume",
      value: formatWhole(current.totals?.callVolume || 0),
      meta: isVsBudget ? "Actual only" : "Across selected period",
      className: "kpi-neutral"
    },
    {
      label: "Avg CXNS %",
      value: `${avgCxnsCombined.toFixed(1)}%`,
      meta: `No Show ${avgNoShow.toFixed(1)}%\nCancel ${avgCancel.toFixed(1)}%`,
      className: "kpi-neutral"
    },
    {
      label: "Avg Abandoned %",
      value: `${averageMetric(current.regions, "abandonedCallRate").toFixed(1)}%`,
      meta: "Across selected scope",
      className: "kpi-neutral"
    }
  ];

  if (entityScope === "SpineOne") {
    cards.push({
      label: "SpineOne PI NP",
      value: formatWhole(current.totals?.piNp || 0),
      meta: "SpineOne only",
      className: "kpi-neutral"
    });

    cards.push({
      label: "SpineOne PI Cash",
      value: formatCurrency(current.totals?.piCashCollection || 0),
      meta: "SpineOne only",
      className: "kpi-neutral"
    });
  }

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
          cashCollected: 0,
          ptoDays: 0,
          noShowRate: 0,
          cancellationRate: 0,
          abandonedCallRate: 0,
          visitVolumeBudget: 0,
          newPatientsBudget: 0,
          pt: {
            scheduledVisits: 0,
            cancellations: 0,
            noShows: 0,
            reschedules: 0,
            totalUnitsBilled: 0,
            visitsSeen: 0,
            unitsPerVisit: 0,
            visitsPerDay: 0
          }
        };
        const prior = comparisonMap[entity] || {
          visitVolume: 0,
          callVolume: 0,
          newPatients: 0,
          surgeries: 0,
          pt: { visitsSeen: 0 },
          visitVolumeBudget: 0,
          newPatientsBudget: 0
        };

        const visitPct = buildVariancePct(row.visitVolume, prior.visitVolume);
        const callPct = buildVariancePct(row.callVolume, prior.callVolume);
        const npPct = buildVariancePct(row.newPatients, prior.newPatients);
        const ptPct = buildVariancePct(row.pt?.visitsSeen, prior.pt?.visitsSeen);

        const visitBudgetStatus = getGoalStatus(row.visitVolume, row.visitVolumeBudget);
        const npBudgetStatus = getGoalStatus(row.newPatients, row.newPatientsBudget);
        const callStatus    = getGoalStatus(row.abandonedCallRate, 10, true);
        const noShowStatus  = getGoalStatus(row.noShowRate, 6, true);
        const cancelStatus  = getGoalStatus(row.cancellationRate, 8, true);
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
              ${entityHasPtData(entity) ? `
              <div class="entityMetricHero">
                <span class="entityMetricLabel">PT Visits</span>
                <strong>${formatWhole(row.pt?.visitsSeen || 0)}</strong>
              </div>` : ""}
              <div class="entityMetricHero">
                <span class="entityMetricLabel">Cash</span>
                <strong class="entityCurrencyValue">${formatCurrency(row.cashCollected || 0)}</strong>
              </div>
            </div>

            ${comparisonBlock}

            <div class="entityHealthRow">
              <div class="entityHealthItem">
                <span class="entityHealthLabel">No Show</span>
                <strong class="${noShowStatus === "good" ? "good" : noShowStatus === "warning" ? "warning" : "bad"}">${formatPercent(row.noShowRate)}</strong>
              </div>
              <div class="entityHealthItem">
                <span class="entityHealthLabel">Cancel</span>
                <strong class="${cancelStatus === "good" ? "good" : cancelStatus === "warning" ? "warning" : "bad"}">${formatPercent(row.cancellationRate)}</strong>
              </div>
              <div class="entityHealthItem">
                <span class="entityHealthLabel">Abandoned</span>
                <strong class="${callStatus === "good" ? "good" : callStatus === "warning" ? "warning" : "bad"}">${formatPercent(row.abandonedCallRate)}</strong>
              </div>
            </div>

            <div class="entityAccessPanel">
              <div class="entityAccessLabelRow">
                <span class="entityAccessLabel">Access Health</span>
                <span class="entityAccessValue">${Math.round(accessPct)}%</span>
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
                  <div class="entityDetailLabel">PT Variance</div>
                  <div class="entityDetailValue">${fmtPct(ptPct)}</div>
                </div>
                <div class="entityDetailTile">
                  <div class="entityDetailLabel">PTO Days</div>
                  <div class="entityDetailValue">${formatWhole(row.ptoDays || 0)}</div>
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
    }

    if (compareAgainst === "budget") {
      const visitGap = normalizeNumber(row.visitVolume) - normalizeNumber(row.visitVolumeBudget);
      if (normalizeNumber(row.visitVolumeBudget) > 0 && visitGap < 0) {
        alerts.push({ severity: "warning", text: `${entity} is ${Math.abs(Math.round(visitGap))} visits below budget.` });
      }
    }
  });

  if (!alerts.length) {
    alerts.push({ severity: "good", text: "No major operational alerts for the selected period." });
  }

  container.innerHTML = alerts.map((alert) => {
    const isBad     = alert.severity === "bad";
    const isWarning = alert.severity === "warning";
    const isGood    = alert.severity === "good";
    const icon      = isBad ? "▲" : isWarning ? "◆" : "✓";
    const cls       = isBad ? "" : isWarning ? "alertWarning" : "";
    if (isGood) {
      return `<div class="winItem"><div class="winItemIcon">✓</div><div class="winItemText">${alert.text}</div></div>`;
    }
    return `<div class="alertItem ${cls}"><span class="alertIcon">${icon}</span><span>${alert.text}</span></div>`;
  }).join("");
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
      if (normalizeNumber(row.visitVolumeBudget) > 0 && visitGap > 0) {
        wins.push({ text: `${entity} is ${Math.round(visitGap)} visits above budget.` });
      }
    } else {
      const visitDiff = normalizeNumber(row.visitVolume) - normalizeNumber(prior.visitVolume);
      const ptDiff = normalizeNumber(row.pt?.visitsSeen) - normalizeNumber(prior.pt?.visitsSeen);

      if (visitDiff > 100) {
        wins.push({ text: `${entity} visits improved by ${Math.round(visitDiff)} vs prior period.` });
      }

      if (ptDiff > 20) {
        wins.push({ text: `${entity} PT visits improved by ${Math.round(ptDiff)} vs prior period.` });
      }
    }

    if (normalizeNumber(row.abandonedCallRate) > 0 && normalizeNumber(row.abandonedCallRate) < 5) {
      wins.push({ text: `${entity} has strong call handling with only ${normalizeNumber(row.abandonedCallRate).toFixed(1)}% abandoned calls.` });
    }
  });

  if (!wins.length) {
    container.innerHTML = `<div class="winItemEmpty">No standout wins for the selected period yet.</div>`;
    return;
  }

  container.innerHTML = wins.map((win) => `
    <div class="winItem">
      <div class="winItemIcon">&#10003;</div>
      <div class="winItemText">${win.text}</div>
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
            <th>PT Visits</th>
            <th>Cash</th>
          </tr>
        </thead>
        <tbody>
          ${rows.map((r) => `
            <tr>
              <td>${r.entity}</td>
              <td>${formatWhole(r.visitVolume)}</td>
              <td>${formatWhole(r.visitVolumeBudget)}</td>
              <td>${formatVariance(r.visitVolume, r.visitVolumeBudget)}</td>
              <td>${formatWhole(r.pt?.visitsSeen || 0)}</td>
              <td>${formatCurrency(r.cashCollected || 0)}</td>
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
          <th>PT</th>
          <th>Cash</th>
          <th>PTO</th>
          <th>No Show</th>
          <th>Cancel</th>
          <th>Abandoned</th>
        </tr>
      </thead>
      <tbody>
        ${rows.map((r) => `
          <tr>
            <td>${r.entity}</td>
            <td>${formatWhole(r.visitVolume)}</td>
            <td>${formatWhole(r.pt?.visitsSeen || 0)}</td>
            <td>${formatCurrency(r.cashCollected || 0)}</td>
            <td>${formatWhole(r.ptoDays || 0)}</td>
            <td>${formatPercent(r.noShowRate)}</td>
            <td>${formatPercent(r.cancellationRate)}</td>
            <td>${formatPercent(r.abandonedCallRate)}</td>
          </tr>
        `).join("")}
      </tbody>
    </table>
  `;
}

function renderActivityTable(rows) {
  const wrap = byId("activityTableWrap");
  if (!wrap) return;

  if (!rows.length) {
    wrap.innerHTML = "<p>No activity found for the selected filters.</p>";
    return;
  }

  const statusPillClass = (status) => {
    if (status === "approved") return "activityStatusPill activityStatusApproved";
    if (status === "submitted") return "activityStatusPill activityStatusSubmitted";
    return "activityStatusPill activityStatusSaved";
  };

  wrap.innerHTML = `
    <table class="regionTable">
      <thead>
        <tr>
          <th>Entity</th>
          <th>Week Ending</th>
          <th>Status</th>
          <th>Updated By</th>
          <th>Updated At</th>
          <th>Created By</th>
          <th>Created At</th>
          <th>Visits</th>
          <th>New Pts</th>
        </tr>
      </thead>
      <tbody>
        ${rows.map((item) => `
          <tr>
            <td><strong>${item.entity}</strong></td>
            <td>${item.weekEnding || "—"}</td>
            <td><span class="${statusPillClass(item.status)}">${item.status || "saved"}</span></td>
            <td>${item.updatedBy || "—"}</td>
            <td>
              <div>${item.updatedAt ? formatDateTime(item.updatedAt) : "—"}</div>
            </td>
            <td>${item.createdBy || "—"}</td>
            <td>${item.createdAt ? formatDateTime(item.createdAt) : "—"}</td>
            <td>${formatWhole(item.visitVolume)}</td>
            <td>${formatWhole(item.newPatients)}</td>
          </tr>
        `).join("")}
      </tbody>
    </table>
  `;
}

async function loadActivity() {
  const entityFilter = byId("activityEntityFilter")?.value || "";
  const weekFilter = byId("activityWeekFilter")?.value || "";
  const limit = parseInt(byId("activityLimit")?.value || "50", 10);

  const summaryEl = byId("activitySummary");
  if (summaryEl) summaryEl.textContent = "Loading activity...";

  const entitiesToFetch = entityFilter ? [entityFilter] : ENTITIES;

  const results = await Promise.all(
    entitiesToFetch.map((entity) =>
      apiGet(`/api/trends?entity=${encodeURIComponent(entity)}&mode=recent&weeks=52`)
        .then((r) => r.items || [])
        .catch(() => [])
    )
  );

  let allRows = results.flat();

  if (weekFilter) {
    allRows = allRows.filter((item) => item.weekEnding === weekFilter);
  }

  allRows.sort((a, b) => {
    const ta = a.updatedAt || a.createdAt || "";
    const tb = b.updatedAt || b.createdAt || "";
    return tb.localeCompare(ta);
  });

  const limited = allRows.slice(0, limit);

  if (summaryEl) {
    const scopeLabel = entityFilter || "All Entities";
    summaryEl.textContent = `Showing ${limited.length} of ${allRows.length} entries — ${scopeLabel}${weekFilter ? ` — Week ending ${weekFilter}` : ""}`;
  }

  renderActivityTable(limited);
  setActivityDebug({ fetched: allRows.length, showing: limited.length, entitiesToFetch });
}

function renderVisitsChart(weeks, currentData, compareAgainst = "priorPeriod") {
  const ctx = byId("visitsChart");
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
      ptVisits: rows.reduce((sum, entry) => sum + normalizeNumber(entry.ptVisitsSeen), 0),
      visitVolumeBudget: rows.reduce((sum, entry) => sum + normalizeNumber(entry.visitVolumeBudget), 0)
    };
  });

  const visitData = totalsByWeek.map((row) => row.visitVolume);
  const callData = totalsByWeek.map((row) => row.callVolume);
  const npData = totalsByWeek.map((row) => row.newPatients);
  const ptData = totalsByWeek.map((row) => row.ptVisits);
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
          label: "PT Visits",
          data: ptData,
          tension: 0.35,
          borderColor: "#b49cff",
          backgroundColor: "rgba(180, 156, 255, 0.16)",
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

function renderExecutiveCards(summary) {
  const rows = summary.regions || [];
  const avg = (key) => {
    if (!rows.length) return 0;
    return rows.reduce((sum, r) => sum + normalizeNumber(r[key]), 0) / rows.length;
  };

  renderMetricCards("executiveCards", [
    { label: "Saved Regions", value: summary.entityCount || 0, className: "kpi-neutral" },
    { label: "Visit Volume", value: formatWhole(summary.totals?.visitVolume || 0), className: "kpi-neutral" },
    { label: "PT Visits", value: formatWhole(summary.ptTotals?.visitsSeen || 0), className: "kpi-neutral" },
    { label: "Cash Collected", value: formatCurrency(summary.totals?.cashCollected || 0), className: "kpi-neutral" },
    { label: "PTO Days", value: formatWhole(summary.totals?.ptoDays || 0), className: "kpi-neutral" },
    { label: "Total Surgeries", value: formatWhole(summary.totals?.surgeries || 0), className: "kpi-neutral" },
    { label: "Call Volume", value: formatWhole(summary.totals?.callVolume || 0), className: "kpi-neutral" },
    { label: "New Patients", value: formatWhole(summary.totals?.newPatients || 0), className: "kpi-neutral" },
    { label: "Avg No Show %", value: `${avg("noShowRate").toFixed(1)}%`, className: "kpi-neutral" },
    { label: "Avg Cancel %", value: `${avg("cancellationRate").toFixed(1)}%`, className: "kpi-neutral" }
  ]);

  renderMetricCards("executivePtCards", [
    {
      label: "PT Summary",
      value: formatWhole(summary.ptTotals?.visitsSeen || 0),
      meta: `Units ${formatWhole(summary.ptTotals?.totalUnitsBilled || 0)}\nUnits/Visit ${normalizeNumber(summary.averages?.ptUnitsPerVisit).toFixed(2)}`,
      className: "kpi-neutral"
    },
    {
      label: "SpineOne PI",
      value: formatWhole(summary.totals?.piNp || 0),
      meta: `Cash ${formatCurrency(summary.totals?.piCashCollection || 0)}`,
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

  container.innerHTML = `
    <table class="regionTable">
      <thead>
        <tr>
          <th>Entity</th>
          <th>Visit</th>
          <th>Cash</th>
          <th>PTO</th>
          <th>Calls</th>
          <th>New</th>
          <th>No Show</th>
          <th>Cancel</th>
          <th>Abandoned</th>
        </tr>
      </thead>
      <tbody>
        ${summary.regions.map((r) => `
          <tr>
            <td>${r.entity}</td>
            <td>${formatWhole(r.visitVolume)}</td>
            <td>${formatCurrency(r.cashCollected || 0)}</td>
            <td>${formatWhole(r.ptoDays || 0)}</td>
            <td>${formatWhole(r.callVolume)}</td>
            <td>${formatWhole(r.newPatients)}</td>
            <td>${formatPercent(r.noShowRate)}</td>
            <td>${formatPercent(r.cancellationRate)}</td>
            <td>${formatPercent(r.abandonedCallRate)}</td>
          </tr>
        `).join("")}
      </tbody>
    </table>
  `;

  const ptContainer = byId("executivePtRegions");
  if (!ptContainer) return;

  ptContainer.innerHTML = `
    <table class="regionTable">
      <thead>
        <tr>
          <th>Entity</th>
          <th>PT Seen</th>
          <th>Units</th>
          <th>Units/Visit</th>
          <th>PI NP</th>
          <th>PI Cash</th>
        </tr>
      </thead>
      <tbody>
        ${summary.regions.map((r) => `
          <tr>
            <td>${r.entity}</td>
            <td>${formatWhole(r.pt?.visitsSeen || 0)}</td>
            <td>${formatWhole(r.pt?.totalUnitsBilled || 0)}</td>
            <td>${normalizeNumber(r.pt?.unitsPerVisit).toFixed(2)}</td>
            <td>${formatWhole(r.piNp || 0)}</td>
            <td>${formatCurrency(r.piCashCollection || 0)}</td>
          </tr>
        `).join("")}
      </tbody>
    </table>
  `;
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
      label: "Latest PT Visits",
      value: latest ? formatDelta(latest.ptVisitsSeen, previous?.ptVisitsSeen) : "-",
      className: latest ? getTrendClass(latest.ptVisitsSeen, previous?.ptVisitsSeen) : "kpi-neutral"
    },
    {
      label: "Latest Cash",
      value: latest ? formatCurrency(latest.cashCollected || 0) : "-",
      className: "kpi-neutral"
    },
    {
      label: "Latest Call Volume",
      value: latest ? formatDelta(latest.callVolume, previous?.callVolume) : "-",
      className: latest ? getTrendClass(latest.callVolume, previous?.callVolume) : "kpi-neutral"
    }
  ]);
}

function renderTrendsTable(result) {
  const wrap = byId("trendsTableWrap");
  if (!wrap) return;

  const items = result.items || [];

  if (!items.length) {
    wrap.innerHTML = "<p>No trend data found for this entity.</p>";
    return;
  }

  wrap.innerHTML = `
    <table class="regionTable">
      <thead>
        <tr>
          <th>Week Ending</th>
          <th>Visits</th>
          <th>PT</th>
          <th>Cash</th>
          <th>PTO</th>
          <th>Calls</th>
          <th>No Show %</th>
          <th>Cancel %</th>
          <th>Abandoned %</th>
        </tr>
      </thead>
      <tbody>
        ${items.map((item) => `
          <tr>
            <td>${item.weekEnding}</td>
            <td>${formatWhole(item.visitVolume)}</td>
            <td>${formatWhole(item.ptVisitsSeen || 0)}</td>
            <td>${formatCurrency(item.cashCollected || 0)}</td>
            <td>${formatWhole(item.ptoDays || 0)}</td>
            <td>${formatWhole(item.callVolume)}</td>
            <td>${formatPercent(item.noShowRate)}</td>
            <td>${formatPercent(item.cancellationRate)}</td>
            <td>${formatPercent(item.abandonedCallRate)}</td>
          </tr>
        `).join("")}
      </tbody>
    </table>
  `;
}

function hideAllViews() {
  const ids = ["dashboardView", "entryView", "executiveView", "trendsView", "ptoForecastView", "activityView", "importView", "helpView"];
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
    "navPtoForecastBtn",
    "navActivityBtn",
    "navImportBtn",
    "navHelpBtn"
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
  renderEntityBrand("trendsBrandWrap", getSelectedTrendsEntity());
  syncTrendsRangeUi();
  syncTrendsPtVisibility();
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

function showPtoForecastView() {
  hideAllViews();
  const el = byId("ptoForecastView");
  if (el) el.style.display = "";
  setActiveNav("navPtoForecastBtn");
}

function setPtoForecastDebug(data) {
  const el = byId("ptoForecastDebugOutput");
  if (!el) return;
  el.textContent = typeof data === "string" ? data : JSON.stringify(data, null, 2);
}

function renderPtoForecastEntities(data) {
  const container = byId("ptoForecastEntities");
  if (!container) return;

  const { entities, monthLabels, monthKeys } = data;

  container.innerHTML = entities.map((entityData) => {
    const { entity, rates, months } = entityData;
    const brand = getBranding(entity);

    return `
      <div class="panel sectionPanel ptoEntityPanel" style="border-top:4px solid ${brand.accent};">
        <div class="ptoEntityHeader">
          <div>
            <div class="sectionEyebrow">${brand.fullName}</div>
            <h3 style="margin:4px 0 2px;">${entity}</h3>
            <div style="font-size:12px;color:var(--text-muted);">
              Rates based on last ${rates.weeksUsed} weeks &nbsp;·&nbsp;
              ${rates.clinicalVisitsPerDay} clinical visits/day &nbsp;·&nbsp;
              ${rates.surgeriesPerDay} surgeries/day
            </div>
          </div>
          <div class="entityLogoWrap" style="width:80px;height:40px;">
            <img src="${brand.logo}" alt="${entity}" class="entityLogo" />
          </div>
        </div>

        <div class="ptoMonthGrid">
          ${months.map((month, i) => `
            <div class="ptoMonthCard">
              <div class="ptoMonthLabel">${month.monthLabel}</div>

              <div class="ptoInputRow">
                <div class="ptoInputGroup">
                  <label class="ptoInputLabel" for="pto_clinical_${entity}_${i}">Clinical PTO Days</label>
                  <input
                    type="number"
                    id="pto_clinical_${entity}_${i}"
                    class="ptoInput"
                    data-entity="${entity}"
                    data-monthkey="${month.monthKey}"
                    data-type="clinical"
                    step="0.5"
                    min="0"
                    value="${month.clinicalPtoDays || ""}"
                    placeholder="0"
                  />
                </div>
                <div class="ptoInputGroup">
                  <label class="ptoInputLabel" for="pto_surgical_${entity}_${i}">Surgical PTO Days</label>
                  <input
                    type="number"
                    id="pto_surgical_${entity}_${i}"
                    class="ptoInput"
                    data-entity="${entity}"
                    data-monthkey="${month.monthKey}"
                    data-type="surgical"
                    step="0.5"
                    min="0"
                    value="${month.surgicalPtoDays || ""}"
                    placeholder="0"
                  />
                </div>
              </div>

              <button
                class="ptoSaveBtn"
                data-entity="${entity}"
                data-monthkey="${month.monthKey}"
                data-idx="${i}"
                type="button"
              >Save</button>

              <div class="ptoForecastResult" id="ptoResult_${entity}_${i}">
                ${month.totalMissedVisits > 0 ? `
                  <div class="ptoImpactRow">
                    <span class="ptoImpactLabel">Missed Clinical Visits</span>
                    <strong class="ptoImpactValue">${formatWhole(month.missedClinicalVisits)}</strong>
                  </div>
                  <div class="ptoImpactRow">
                    <span class="ptoImpactLabel">Missed Surgeries</span>
                    <strong class="ptoImpactValue">${formatWhole(month.missedSurgeries)}</strong>
                  </div>
                  <div class="ptoImpactRow ptoImpactTotal">
                    <span class="ptoImpactLabel">Total Forecasted Impact</span>
                    <strong class="ptoImpactValue">${formatWhole(month.totalMissedVisits)} visits</strong>
                  </div>
                  ${month.savedBy ? `<div class="ptoSavedBy">Saved by ${month.savedBy}</div>` : ""}
                ` : `<div class="ptoNoData">Enter PTO days above to see forecast</div>`}
              </div>
            </div>
          `).join("")}
        </div>
      </div>
    `;
  }).join("");

  // Wire up save buttons
  document.querySelectorAll(".ptoSaveBtn").forEach((btn) => {
    btn.addEventListener("click", async () => {
      const entity = btn.getAttribute("data-entity");
      const monthKey = btn.getAttribute("data-monthkey");
      const idx = btn.getAttribute("data-idx");

      const clinicalInput = document.querySelector(
        `.ptoInput[data-entity="${entity}"][data-monthkey="${monthKey}"][data-type="clinical"]`
      );
      const surgicalInput = document.querySelector(
        `.ptoInput[data-entity="${entity}"][data-monthkey="${monthKey}"][data-type="surgical"]`
      );

      const clinicalPtoDays = parseFloat(clinicalInput?.value || "0") || 0;
      const surgicalPtoDays = parseFloat(surgicalInput?.value || "0") || 0;

      btn.textContent = "Saving...";
      btn.disabled = true;

      try {
        const result = await apiPost("/api/pto-forecast", {
          entity,
          monthKey,
          clinicalPtoDays,
          surgicalPtoDays
        });

        const resultEl = byId(`ptoResult_${entity}_${idx}`);
        if (resultEl && result.forecast) {
          const f = result.forecast;
          resultEl.innerHTML = `
            <div class="ptoImpactRow">
              <span class="ptoImpactLabel">Missed Clinical Visits</span>
              <strong class="ptoImpactValue">${formatWhole(f.missedClinicalVisits)}</strong>
            </div>
            <div class="ptoImpactRow">
              <span class="ptoImpactLabel">Missed Surgeries</span>
              <strong class="ptoImpactValue">${formatWhole(f.missedSurgeries)}</strong>
            </div>
            <div class="ptoImpactRow ptoImpactTotal">
              <span class="ptoImpactLabel">Total Forecasted Impact</span>
              <strong class="ptoImpactValue">${formatWhole(f.totalMissedVisits)} visits</strong>
            </div>
            <div class="ptoSavedBy">Saved successfully</div>
          `;
        }

        // Refresh summary
        await refreshPtoSummary();

        btn.textContent = "Saved ✓";
        setTimeout(() => {
          btn.textContent = "Save";
          btn.disabled = false;
        }, 2000);
      } catch (e) {
        btn.textContent = "Error — retry";
        btn.disabled = false;
        setPtoForecastDebug(String(e));
      }
    });
  });
}

function renderPtoForecastSummary(data) {
  const container = byId("ptoForecastSummary");
  if (!container) return;

  const { entities, monthLabels } = data;

  const totalByMonth = monthLabels.map((label, mi) => {
    const total = entities.reduce((sum, e) => {
      return sum + normalizeNumber(e.months[mi]?.totalMissedVisits);
    }, 0);
    return { label, total };
  });

  const quarterlyTotal = totalByMonth.reduce((sum, m) => sum + m.total, 0);

  container.innerHTML = `
    <table class="regionTable">
      <thead>
        <tr>
          <th>Entity</th>
          ${monthLabels.map((l) => `<th>${l}</th>`).join("")}
          <th>Q Total</th>
          <th>Daily Rate</th>
        </tr>
      </thead>
      <tbody>
        ${entities.map((e) => `
          <tr>
            <td><strong>${e.entity}</strong></td>
            ${e.months.map((m) => `
              <td>${m.totalMissedVisits > 0
                ? `<span class="ptoImpactBadge">${formatWhole(m.totalMissedVisits)}</span>`
                : "—"
              }</td>
            `).join("")}
            <td><strong>${formatWhole(e.quarterlyTotalMissed)}</strong></td>
            <td style="font-size:12px;color:var(--text-muted);">${e.rates.visitsPerDay}/day</td>
          </tr>
        `).join("")}
        <tr style="border-top:2px solid rgba(255,255,255,0.14);">
          <td><strong>All Entities</strong></td>
          ${totalByMonth.map((m) => `<td><strong>${m.total > 0 ? formatWhole(m.total) : "—"}</strong></td>`).join("")}
          <td><strong>${formatWhole(quarterlyTotal)}</strong></td>
          <td></td>
        </tr>
      </tbody>
    </table>
    <div style="margin-top:12px;font-size:13px;color:var(--text-muted);">
      Forecasted missed visits = (Clinical PTO Days × clinical visits/day) + (Surgical PTO Days × surgeries/day).
      Rates computed from last ${entities[0]?.rates?.weeksUsed || 12} weeks of actual data per entity.
    </div>
  `;
}

let _lastPtoData = null;

async function loadPtoForecast() {
  const bannerEl = byId("ptoForecastBanner");
  if (bannerEl) bannerEl.textContent = "Loading forecast data...";

  const data = await apiGet("/api/pto-forecast");
  _lastPtoData = data;

  if (bannerEl) {
    const labels = (data.monthLabels || []).join(", ");
    bannerEl.textContent = `Rolling quarter: ${labels} · Enter projected PTO days per entity and month, then save.`;
  }

  renderPtoForecastEntities(data);
  renderPtoForecastSummary(data);
  setPtoForecastDebug(data);
}

async function refreshPtoSummary() {
  try {
    const data = await apiGet("/api/pto-forecast");
    _lastPtoData = data;
    renderPtoForecastSummary(data);
  } catch {
    // non-fatal
  }
}

function showHelpView() {
  hideAllViews();
  const el = byId("helpView");
  if (el) el.style.display = "";
  setActiveNav("navHelpBtn");
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

function syncQuickPresetPills(activePreset) {
  document.querySelectorAll(".quickPresetPill").forEach((pill) => {
    const preset = pill.getAttribute("data-preset");
    pill.classList.toggle("quickPresetActive", preset === activePreset);
  });
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
          status: "saved",
          visitVolume: 0,
          callVolume: 0,
          newPatients: 0,
          surgeries: 0,
          cashCollected: 0,
          ptoDays: 0,
          piNp: 0,
          piCashCollection: 0,
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
      row.cashCollected += normalizeNumber(region.cashCollected);
      row.ptoDays += normalizeNumber(region.ptoDays);
      row.piNp += normalizeNumber(region.piNp);
      row.piCashCollection += normalizeNumber(region.piCashCollection);
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
        ptVisitsSeen: normalizeNumber(region.pt?.visitsSeen),
        cashCollected: normalizeNumber(region.cashCollected),
        ptoDays: normalizeNumber(region.ptoDays),
        noShowRate: normalizeNumber(region.noShowRate),
        cancellationRate: normalizeNumber(region.cancellationRate),
        abandonedCallRate: normalizeNumber(region.abandonedCallRate),
        visitVolumeBudget: normalizeNumber(region.budget?.visitVolumeBudget),
        newPatientsBudget: normalizeNumber(region.budget?.newPatientsBudget)
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
    cashCollected: row.cashCollected,
    ptoDays: row.ptoDays,
    piNp: row.piNp,
    piCashCollection: row.piCashCollection,
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
    surgeries: regions.reduce((sum, r) => sum + normalizeNumber(r.surgeries), 0),
    cashCollected: regions.reduce((sum, r) => sum + normalizeNumber(r.cashCollected), 0),
    ptoDays: regions.reduce((sum, r) => sum + normalizeNumber(r.ptoDays), 0),
    piNp: regions.reduce((sum, r) => sum + normalizeNumber(r.piNp), 0),
    piCashCollection: regions.reduce((sum, r) => sum + normalizeNumber(r.piCashCollection), 0)
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

  const ptAverages = {
    unitsPerVisit:
      ptTotals.visitsSeen > 0
        ? Number((ptTotals.totalUnitsBilled / ptTotals.visitsSeen).toFixed(2))
        : 0,
    visitsPerDay:
      regions.length
        ? Number(
            (
              regions.reduce((sum, r) => sum + normalizeNumber(r.pt?.visitsPerDay), 0) /
              regions.length
            ).toFixed(2)
          )
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
      totals: {
        visitVolume: 0,
        callVolume: 0,
        newPatients: 0,
        surgeries: 0,
        cashCollected: 0,
        ptoDays: 0,
        piNp: 0,
        piCashCollection: 0
      },
      budgetTotals: { visitVolumeBudget: 0, newPatientsBudget: 0 },
      ptTotals: {
        scheduledVisits: 0,
        cancellations: 0,
        noShows: 0,
        reschedules: 0,
        totalUnitsBilled: 0,
        visitsSeen: 0
      },
      ptAverages: { unitsPerVisit: 0, visitsPerDay: 0 },
      regions: []
    };
  }

  const summaries = await Promise.all(validWeeks.map((week) => fetchExecutiveSummaryByWeek(week)));
  return aggregateExecutiveSummaries(summaries, entityScope, options);
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
        surgeries: 0,
        cashCollected: 0,
        ptoDays: 0,
        piNp: 0,
        piCashCollection: 0
      },
      budgetTotals: current.budgetTotals,
      ptTotals: current.ptTotals,
      ptAverages: current.ptAverages,
      regions: current.regions || []
    };
  } else {
    current = await loadDashboardDataForWeeks(weekSets.primaryWeeks, entityScope, { includeBudget: false });
    comparison = await loadDashboardDataForWeeks(weekSets.comparisonWeeks, entityScope, { includeBudget: false });
  }

  const summaryEl = byId("dashboardSummaryText");
  if (summaryEl) {
    summaryEl.innerHTML = `<div style="font-size:13px; opacity:0.85;">${weekSets.summary}${entityScope !== "ALL" ? ` • Scope: ${entityScope}` : " • Scope: All Entities"}</div>`;
  }

  renderDashboardCards(current, comparison, compareAgainst, entityScope);
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

async function loadExecutiveSummary() {
  const weekEnding = byId("executiveWeekEnding")?.value || getDefaultWeekEnding();
  const result = await apiGet(`/api/executive-summary?weekEnding=${encodeURIComponent(weekEnding)}`);
  renderExecutiveCards(result);
  renderExecutiveRegions(result);
  setExecutiveDebug(result);
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

function syncTrendsRangeUi() {
  const mode = byId("trendsRangeMode")?.value || "recent";
  const isDateRange = mode === "dateRange";

  const weeksWrap = byId("trendsWeeksWrap");
  const startWrap = byId("trendsStartWrap");
  const endWrap = byId("trendsEndWrap");

  if (weeksWrap) weeksWrap.style.display = isDateRange ? "none" : "";
  if (startWrap) startWrap.style.display = isDateRange ? "" : "none";
  if (endWrap) endWrap.style.display = isDateRange ? "" : "none";
}

const PT_ENTITIES = ["NES", "SpineOne", "MRO"];

function entityHasPtData(entity) {
  return PT_ENTITIES.includes(entity);
}

function syncTrendsPtVisibility() {
  const entity = getSelectedTrendsEntity();
  const ptSection = byId("trendsView")?.querySelector(".ptSectionPanel");
  if (ptSection) ptSection.style.display = entityHasPtData(entity) ? "" : "none";
}

function isMeaningfulTrendsRow(item) {
  const today = new Date().toISOString().slice(0, 10);
  return item.weekEnding <= today;
}

async function loadTrends() {
  const entity = getSelectedTrendsEntity();
  renderEntityBrand("trendsBrandWrap", entity);

  const mode = byId("trendsRangeMode")?.value || "recent";
  const weeks = byId("trendsLimit")?.value || "12";
  const startDate = byId("trendsStartDate")?.value || "";
  const endDate = byId("trendsEndDate")?.value || "";

  let url;
  if (mode === "dateRange" && startDate && endDate) {
    url = `/api/trends?entity=${encodeURIComponent(entity)}&mode=dateRange&startDate=${encodeURIComponent(startDate)}&endDate=${encodeURIComponent(endDate)}`;
  } else {
    url = `/api/trends?entity=${encodeURIComponent(entity)}&mode=recent&weeks=52`;
  }

  const raw = await apiGet(url);

  // Filter out future placeholder rows, then take the requested count
  const filtered = (raw.items || []).filter(isMeaningfulTrendsRow);
  const limited = mode === "dateRange" ? filtered : filtered.slice(0, parseInt(weeks, 10));

  const result = { ...raw, items: limited, count: limited.length };

  renderTrendsCards(result);
  renderTrendsTable(result);
  syncTrendsPtVisibility();
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
  return firstExistingId(["importFile"]);
}

function getWeeklyImportButton() {
  return firstExistingId(["runImportBtn"]);
}

async function runImport() {
  if (!currentUser?.access?.isAdmin) {
    throw new Error("Admin only");
  }

  const fileInput = getWeeklyImportFileInput();
  const file = fileInput?.files && fileInput.files[0];

  if (!file) throw new Error("Select a workbook file first");

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

function injectUiPolishStyles() {
  if (document.getElementById("ops-dashboard-ui-polish")) return;

  const style = document.createElement("style");
  style.id = "ops-dashboard-ui-polish";
  style.textContent = `
    textarea {
      width: 100%;
      padding: 11px 12px;
      box-sizing: border-box;
      border: 1px solid rgba(255,255,255,0.12);
      border-radius: 12px;
      background: rgba(255,255,255,0.06);
      color: #ffffff;
      resize: vertical;
      line-height: 1.5;
    }

    .narrativeFieldBlock { margin-top: 10px; }

    .notesShell {
      padding: 18px;
      border-radius: 18px;
      background: linear-gradient(180deg, rgba(17,63,93,0.42) 0%, rgba(11,42,63,0.42) 100%);
      border: 1px solid rgba(108,182,255,0.16);
      box-shadow: inset 0 1px 0 rgba(255,255,255,0.03);
    }

    .notesHeader { margin-bottom: 14px; }

    .notesEyebrow {
      font-size: 11px;
      text-transform: uppercase;
      letter-spacing: 0.12em;
      color: #6cb6ff;
      font-weight: 800;
      margin-bottom: 6px;
    }

    .notesTitle {
      font-size: 22px;
      font-weight: 900;
      line-height: 1.05;
      margin-bottom: 6px;
      color: #f4f8fc;
    }

    .notesSubtext {
      font-size: 13px;
      color: #b8d3e6;
      line-height: 1.5;
    }

    .notesPromptGrid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
      gap: 10px;
      margin-bottom: 16px;
    }

    .notesPromptCard {
      display: flex;
      gap: 10px;
      align-items: flex-start;
      padding: 12px;
      border-radius: 14px;
      background: rgba(255,255,255,0.04);
      border: 1px solid rgba(255,255,255,0.06);
    }

    .notesPromptBullet {
      width: 10px;
      height: 10px;
      border-radius: 999px;
      background: linear-gradient(180deg, #6cb6ff 0%, #74f0ff 100%);
      margin-top: 5px;
      flex-shrink: 0;
      box-shadow: 0 0 0 4px rgba(108,182,255,0.12);
    }

    .notesPromptText {
      color: #dcebf8;
      font-size: 13px;
      line-height: 1.45;
      font-weight: 700;
    }

    .summaryCard { min-width: 0; overflow: hidden; }

    .summaryCard .value {
      font-size: clamp(1.55rem, 2.2vw, 2.5rem) !important;
      line-height: 1.02 !important;
      letter-spacing: -0.03em;
      overflow-wrap: anywhere;
      word-break: break-word;
      max-width: 100%;
      display: block;
    }

    .summaryCard h3,
    .summaryCard div:not(.value):not(h3) {
      min-width: 0;
      overflow-wrap: anywhere;
      word-break: break-word;
    }

    .entityCurrencyValue {
      font-size: clamp(1.05rem, 1.5vw, 1.55rem) !important;
      line-height: 1.05 !important;
      overflow-wrap: anywhere;
      word-break: break-word;
    }

    .entityCardGrid {
      display:grid;
      grid-template-columns:repeat(auto-fit,minmax(290px,1fr));
      gap:14px;
    }

    .entityCard {
      background:linear-gradient(160deg, rgba(11,26,42,0.98) 0%, rgba(8,20,32,0.98) 100%);
      border:1px solid rgba(255,255,255,0.08);
      border-radius:16px;
      padding:16px;
      box-shadow:0 8px 24px rgba(0,0,0,0.2);
      transition:transform .18s ease, box-shadow .22s ease, border-color .18s ease;
    }

    .entityCard:hover {
      transform:translateY(-2px);
      box-shadow:0 16px 36px rgba(0,0,0,0.28);
      border-color:rgba(91,168,255,0.2);
    }

    .entityCardHeader {
      display:flex;
      justify-content:space-between;
      gap:10px;
      align-items:flex-start;
      margin-bottom:12px;
    }

    .entityHeaderLeft { min-width:0; flex:1; }

    .entityStatusRow {
      display:flex;
      gap:6px;
      flex-wrap:wrap;
      align-items:center;
      margin-bottom:8px;
    }

    .entityStatusPill {
      display:inline-flex;
      align-items:center;
      padding:3px 8px;
      border-radius:4px;
      background:rgba(77,217,138,0.12);
      color:#4dd98a;
      font-family:'DM Mono','Fira Mono',monospace;
      font-size:10px;
      font-weight:500;
      text-transform:uppercase;
      letter-spacing:.08em;
    }

    .metricChip {
      display:inline-flex;
      align-items:center;
      padding:3px 8px;
      border-radius:4px;
      background:rgba(255,255,255,0.06);
      color:#93b4cc;
      font-family:'DM Mono','Fira Mono',monospace;
      font-size:10px;
      font-weight:500;
      letter-spacing:.04em;
    }

    .metricChipGood    { background:rgba(77,217,138,0.1);   color:#4dd98a; }
    .metricChipWarning { background:rgba(247,198,47,0.12);  color:#f7c62f; }
    .metricChipBad     { background:rgba(240,100,112,0.12); color:#f06470; }

    .entityTitle    { font-size:20px; font-weight:800; line-height:1.05; margin-bottom:4px; letter-spacing:-0.02em; }
    .entitySubtitle { font-size:11.5px; color:#506a7e; line-height:1.4; }

    .entityLogoWrap {
      width:76px; height:40px; background:#fff; border-radius:8px; padding:5px;
      display:flex; align-items:center; justify-content:center; flex-shrink:0;
    }

    .entityLogo { max-width:100%; max-height:100%; object-fit:contain; }

    .entityTopMetrics {
      display:grid;
      grid-template-columns:repeat(3,1fr);
      gap:8px;
      margin-bottom:12px;
    }

    .entityMetricHero {
      padding:10px 11px;
      border-bottom:1px solid rgba(255,255,255,0.06);
      min-width:0;
    }

    .entityMetricLabel {
      display:block;
      font-size:9.5px;
      text-transform:uppercase;
      letter-spacing:.1em;
      color:#506a7e;
      margin-bottom:4px;
      font-weight:700;
    }

    .entityMetricHero strong {
      font-family:'DM Mono','Fira Mono',monospace;
      font-size:21px;
      font-weight:500;
      line-height:1;
      letter-spacing:-.02em;
      display:block;
      min-width:0;
      color:#eaf1f8;
    }

    .entityBudgetGrid, .entityCompareGrid { display:grid; gap:8px; margin-bottom:12px; }
    .entityBudgetGrid  { grid-template-columns:repeat(2,1fr); }
    .entityCompareGrid { grid-template-columns:repeat(3,1fr); }

    .entityBudgetTile, .entityMiniStat {
      background:rgba(91,168,255,0.06);
      border:1px solid rgba(91,168,255,0.1);
      border-radius:10px;
      padding:10px 11px;
    }

    .entityBudgetLabel, .entityMiniLabel {
      display:block;
      font-size:9.5px;
      text-transform:uppercase;
      letter-spacing:.1em;
      color:#506a7e;
      margin-bottom:5px;
      font-weight:700;
    }

    .entityBudgetValue, .entityMiniStat strong {
      font-family:'DM Mono','Fira Mono',monospace;
      font-size:20px; font-weight:500; line-height:1; margin-bottom:6px; display:block;
      letter-spacing:-.02em;
    }

    .entityBudgetMeta { margin-top:5px; font-size:11.5px; color:#506a7e; }
    .entityProgressWrap { margin-top:8px; }

    .entityProgressLabelRow {
      display:flex; justify-content:space-between; gap:8px; align-items:center;
      font-size:11.5px; color:#506a7e; margin-bottom:5px;
    }

    .entityProgressLabelRow strong { color:#93b4cc; font-size:11.5px; font-family:'DM Mono','Fira Mono',monospace; }

    .entityProgressTrack {
      width:100%; height:5px; border-radius:999px;
      background:rgba(255,255,255,0.07); overflow:hidden; position:relative;
    }

    .entityProgressBar {
      height:100%; border-radius:999px; transition:width .32s ease;
      background:rgba(91,168,255,0.6);
    }

    .entityProgressGood    { background:#4dd98a; }
    .entityProgressWarning { background:#f7c62f; }
    .entityProgressBad     { background:#f06470; }

    /* Health stats — three inline pills in a row */
    .entityHealthRow {
      display:flex;
      gap:6px;
      margin-bottom:10px;
      flex-wrap:wrap;
    }

    .entityHealthItem {
      flex:1;
      display:flex;
      flex-direction:column;
      gap:2px;
      padding:8px 10px;
      border-radius:8px;
      background:rgba(255,255,255,0.03);
      border:1px solid rgba(255,255,255,0.06);
      min-width:0;
    }

    .entityHealthLabel {
      font-size:9.5px;
      text-transform:uppercase;
      letter-spacing:.08em;
      color:#506a7e;
      font-weight:700;
      white-space:nowrap;
    }

    .entityHealthItem strong {
      font-family:'DM Mono','Fira Mono',monospace;
      font-size:15px;
      font-weight:500;
      line-height:1;
      color:#eaf1f8;
    }

    /* Access health bar */
    .entityAccessPanel {
      margin-bottom:10px;
      padding:10px 12px;
      border-radius:10px;
      background:rgba(255,255,255,0.025);
      border:1px solid rgba(255,255,255,0.05);
    }

    .entityAccessLabelRow {
      display:flex;
      justify-content:space-between;
      align-items:center;
      margin-bottom:6px;
    }

    .entityAccessLabel {
      font-size:9.5px;
      text-transform:uppercase;
      letter-spacing:.1em;
      color:#506a7e;
      font-weight:700;
    }

    .entityAccessValue {
      font-family:'DM Mono','Fira Mono',monospace;
      font-size:13px;
      font-weight:500;
      color:#eaf1f8;
    }

    .entityDetailDrawer {
      border:1px solid rgba(255,255,255,0.05); border-radius:10px;
      background:rgba(4,9,15,0.4); overflow:hidden;
    }

    .entityDetailDrawer summary {
      cursor:pointer; list-style:none; padding:10px 13px;
      font-weight:600; font-size:12.5px; color:#93b4cc;
      display:flex; align-items:center; gap:6px;
    }

    .entityDetailDrawer summary::before {
      content:"›";
      font-size:14px;
      transition:transform .15s ease;
      display:inline-block;
    }

    details[open] .entityDetailDrawer summary::before { transform:rotate(90deg); }
    .entityDetailDrawer summary::-webkit-details-marker { display:none; }

    .entityDetailGrid {
      display:grid; grid-template-columns:repeat(3,1fr); gap:8px; padding:0 13px 13px;
    }

    .entityDetailTile {
      padding:10px; border-radius:9px;
      background:rgba(255,255,255,0.025); border:1px solid rgba(255,255,255,0.05);
    }

    .entityDetailLabel {
      font-size:9.5px; text-transform:uppercase; letter-spacing:.08em;
      color:#506a7e; margin-bottom:4px; font-weight:700;
    }

    .entityDetailValue {
      font-family:'DM Mono','Fira Mono',monospace;
      font-size:17px; font-weight:500; line-height:1.1;
      letter-spacing:-.01em;
    }

    @media (max-width: 820px) {
      .entityTopMetrics, .entityCompareGrid, .entityDetailGrid { grid-template-columns:1fr; }
      .entityBudgetGrid { grid-template-columns:1fr; }
    }

    @media (max-width: 768px) {
      .notesTitle { font-size: 18px; }
      .notesPromptGrid { grid-template-columns: 1fr; }
    }
  `;
  document.head.appendChild(style);
}

(async function init() {
  try {
    currentUser = await apiGet("/api/me");

    renderUser(currentUser);
    setupEntityDropdown();
    setupTrendsEntityDropdown();
    renderForm();
    injectUiPolishStyles();

    const defaultWeek = getDefaultWeekEnding();

    if (byId("dashboardWeekEnding")) byId("dashboardWeekEnding").value = defaultWeek;
    if (byId("dashboardCustomEnd")) byId("dashboardCustomEnd").value = defaultWeek;
    if (byId("dashboardCustomStart")) byId("dashboardCustomStart").value = getDateWeeksAgo(8, defaultWeek);
    if (byId("weekEnding")) byId("weekEnding").value = defaultWeek;
    if (byId("executiveWeekEnding")) byId("executiveWeekEnding").value = defaultWeek;

    const dashboardPeriodType = byId("dashboardPeriodType");
    if (dashboardPeriodType) dashboardPeriodType.value = "lastWeek";
    syncQuickPresetPills("lastWeek");

    renderEntityBrand("entryBrandWrap", getSelectedEntity());

    // Quick preset pills
    document.querySelectorAll(".quickPresetPill").forEach((pill) => {
      pill.addEventListener("click", async () => {
        const preset = pill.getAttribute("data-preset");
        if (!preset) return;
        const periodSelect = byId("dashboardPeriodType");
        if (periodSelect) periodSelect.value = preset;
        syncDashboardPeriodUi();
        syncQuickPresetPills(preset);
        try {
          await loadDashboardLanding();
        } catch (e) {
          setDashboardDebug(String(e));
        }
      });
    });

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
          await loadActivity();
        } catch (e) {
          setActivityDebug(String(e));
        }
      });
    }

    if (byId("navPtoForecastBtn")) {
      byId("navPtoForecastBtn").addEventListener("click", async () => {
        showPtoForecastView();
        try {
          await loadPtoForecast();
        } catch (e) {
          setPtoForecastDebug(String(e));
        }
      });
    }

    if (byId("navImportBtn")) {
      byId("navImportBtn").addEventListener("click", showImportView);
    }

    if (byId("navHelpBtn")) {
      byId("navHelpBtn").addEventListener("click", showHelpView);
    }

    if (byId("loadDashboardBtn")) {
      byId("loadDashboardBtn").addEventListener("click", async () => {
        const preset = byId("dashboardPeriodType")?.value;
        syncQuickPresetPills(preset || "lastWeek");
        try {
          await loadDashboardLanding();
        } catch (e) {
          setDashboardDebug(String(e));
        }
      });
    }

    if (byId("dashboardPeriodType")) {
      byId("dashboardPeriodType").addEventListener("change", async () => {
        const preset = byId("dashboardPeriodType").value;
        syncDashboardPeriodUi();
        syncQuickPresetPills(preset);
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

    if (byId("loadExecutiveBtn")) {
      byId("loadExecutiveBtn").addEventListener("click", async () => {
        try {
          await loadExecutiveSummary();
        } catch (e) {
          setExecutiveDebug(String(e));
        }
      });
    }

    if (byId("trendsEntitySelect")) {
      byId("trendsEntitySelect").addEventListener("change", async () => {
        renderEntityBrand("trendsBrandWrap", getSelectedTrendsEntity());
        syncTrendsPtVisibility();
        try {
          await loadTrends();
        } catch (e) {
          setTrendsDebug(String(e));
        }
      });
    }

    if (byId("trendsRangeMode")) {
      byId("trendsRangeMode").addEventListener("change", async () => {
        syncTrendsRangeUi();
        // Only auto-load on switch back to recent; date range waits for both dates
        if (byId("trendsRangeMode").value === "recent") {
          try {
            await loadTrends();
          } catch (e) {
            setTrendsDebug(String(e));
          }
        }
      });
    }

    if (byId("trendsLimit")) {
      byId("trendsLimit").addEventListener("change", async () => {
        try {
          await loadTrends();
        } catch (e) {
          setTrendsDebug(String(e));
        }
      });
    }

    if (byId("trendsStartDate")) {
      byId("trendsStartDate").addEventListener("change", async () => {
        if (byId("trendsEndDate")?.value) {
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
        if (byId("trendsStartDate")?.value) {
          try {
            await loadTrends();
          } catch (e) {
            setTrendsDebug(String(e));
          }
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

    if (byId("loadActivityBtn")) {
      byId("loadActivityBtn").addEventListener("click", async () => {
        try {
          await loadActivity();
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

    showDashboardView();
    await loadDashboardLanding();
  } catch (error) {
    setStatus(error.message || "Failed to load app", true);
    setDebug(String(error));
    setDashboardDebug(String(error));
    setExecutiveDebug(String(error));
    setTrendsDebug(String(error));
    setImportDebug(String(error));
  }
})();
