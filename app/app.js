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

// ── Entity working days per week ───────────────────────────────
const ENTITY_WORKING_DAYS = {
  LAOSS:    5,
  NES:      4.5,
  SpineOne: 5,
  MRO:      5
};

function getEntityWorkingDays(entity) {
  return ENTITY_WORKING_DAYS[entity] ?? 5;
}

// ── PT Scope helpers ───────────────────────────────────────────
const PT_SCOPE_MAP = {
  "NES-PT":      "NES",
  "SpineOne-PT": "SpineOne",
  "MRO-PT":      "MRO"
};

function isPtScope(entityScope) {
  return entityScope === "ALL-PT" || entityScope in PT_SCOPE_MAP;
}

function ptScopeBaseEntity(entityScope) {
  return PT_SCOPE_MAP[entityScope] || null;
}

function ptScopeEntities(entityScope) {
  if (entityScope === "ALL-PT") return ["NES", "SpineOne", "MRO"];
  const base = ptScopeBaseEntity(entityScope);
  return base ? [base] : [];
}

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

function formHasUnsavedData() {
  // Check if any visible, editable form field has a value
  // compared to what was last loaded from the server
  const loaded = currentWeekData?.values || currentWeekData?.data || {};
  const fieldIds = [
    "newPatients", "surgeries", "established", "noShows", "cancelled",
    "totalCalls", "abandonedCalls", "cashCollected", "ptoDays",
    "reschedules", "piNp", "piCashCollection",
    "ptScheduledVisits", "ptCancellations", "ptNoShows",
    "ptReschedules", "ptTotalUnitsBilled", "ptVisitsSeen"
  ];
  for (const id of fieldIds) {
    const el = byId(id);
    if (!el || el.style.display === "none" || el.closest("[style*='display:none']")) continue;
    const current = String(el.value || "").trim();
    const stored  = String(loaded[id] ?? "").trim();
    // If user typed something and it differs from what was loaded, flag it
    if (current !== "" && current !== "0" && current !== stored) return true;
  }
  // Also check narrative
  const narr = byId("operationsNarrative");
  if (narr && narr.value.trim() && narr.value.trim() !== String(loaded.operationsNarrative || "").trim()) return true;
  return false;
}

// ── Alert Thresholds (configurable, stored in localStorage) ───
const THRESHOLD_DEFAULTS = {
  noShowRate:        6,
  cancellationRate:  8,
  abandonedCallRate: 10,
  visitDropVsPrior:  100
};

function getThreshold(key, fallback) {
  const stored = localStorage.getItem("threshold_" + key);
  if (stored !== null) {
    const n = parseFloat(stored);
    if (Number.isFinite(n)) return n;
  }
  return fallback ?? THRESHOLD_DEFAULTS[key] ?? 0;
}

function setThreshold(key, value) {
  localStorage.setItem("threshold_" + key, String(value));
}

function resetThresholds() {
  Object.keys(THRESHOLD_DEFAULTS).forEach(k => localStorage.removeItem("threshold_" + k));
}

// ── View mode ─────────────────────────────────────────────────
// The app only ships one theme now — a light SaaS console inspired by
// Ramp / Linear / Retool. The legacy Command and Default (dark) themes
// have been retired. The `lightView` class is applied on every load so
// every `body.lightView` CSS rule acts as the base style.
function setViewMode(mode) {
  // Legacy signature kept for any callers; unconditionally applies light.
  document.body.classList.remove("commandView");
  document.body.classList.add("lightView");
  localStorage.setItem("viewMode", "light");
}

function toggleLightView() { setViewMode("light"); }
window.toggleLightView = toggleLightView;

function initCommandView() {
  // Ensure light is applied even when localStorage has a legacy value.
  setViewMode("light");
  const lvBtn = byId("lightViewToggle");
  if (lvBtn) lvBtn.addEventListener("click", toggleLightView);
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
  // Always return the most recent past Friday (or today if today is Friday)
  const d = new Date();
  const day = d.getDay(); // 0=Sun, 5=Fri
  const daysBack = day >= 5 ? day - 5 : day + 2; // how many days back to reach Friday
  d.setDate(d.getDate() - daysBack);
  return d.toISOString().slice(0, 10);
}

// Friday of the week that CONTAINS today. Used for the live phones strip —
// which always shows current week data regardless of the dashboard's
// selected (usually retrospective) period.
function getCurrentWeekEnding() {
  const d = new Date();
  const day = d.getDay(); // 0=Sun, 5=Fri, 6=Sat
  // Days forward to reach this week's Friday. Sat rolls to next week's Fri
  // (matches how weekEnding is stored as the following Friday).
  const daysForward = day <= 5 ? 5 - day : 6;
  d.setDate(d.getDate() + daysForward);
  return d.toISOString().slice(0, 10);
}

function nearestPriorFriday(isoDate) {
  const d = new Date(isoDate + "T12:00:00Z");
  const day = d.getUTCDay();
  if (day === 5) return isoDate;
  const daysBack = day < 5 ? day + 2 : day - 5;
  d.setUTCDate(d.getUTCDate() - daysBack);
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

// Escape any string before interpolating into innerHTML templates. Prevents
// stored XSS via notes fields, usernames, statuses, etc.
function escapeHtml(value) {
  if (value == null) return "";
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;")
    .replace(/`/g, "&#96;");
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
  const n = normalizeNumber(value);
  if (isNaN(n)) return "—";
  return `${n.toFixed(digits)}%`;
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

// ── Entry form field registry ──────────────────────────────────────────
//
// Single source of truth for every field on the Weekly Entry form. Every
// add/rename/retype of a field is a one-line change here — the form HTML,
// save/load path, input listener wiring, and derived-metric recompute all
// read from this array. Historically each field had to be added in 7+
// places which is where bugs like the imaging+reschedules "not sticking"
// came from.
//
// Field shape:
//   key             — DOM id and storage key (required for input fields)
//   label           — user-visible label (required for input fields)
//   section         — "team" | "pt" | "spineOnePI" | "calculated" | "ptCalculated"
//                     (used to emit section headers)
//   sectionHeader   — true on entries that are ONLY a header, no input
//   mode            — "ortho" | "pt" | "all"  (controls show/hide)
//   entities        — "all" (default) | "spineone" | "non-laoss"
//   step, min       — passed through to the <input>
//   defaultValue    — initial value if none stored
//   readonly        — true for derived/calculated fields
//   derived         — true if value is computed by calculateDerivedMetrics
//                     (save/load reads it as output, not user input)
//   aliases         — alternate storage keys to read from on load
//                     (for legacy records; e.g. visitVolume had "totalVisits")
//   preserveOnOppositeMode — default true; false for ptoDays which is shared
//
const ENTRY_FIELDS = [
  { section: "team", sectionHeader: true, mode: "ortho", label: "Team Inputs" },

  { key: "newPatients",     label: "New Patients",        mode: "ortho", step: "1",    min: "0", aliases: ["npActual"] },
  { key: "surgeries",       label: "Surgeries",           mode: "ortho", step: "1",    min: "0", aliases: ["surgeryActual"] },
  { key: "imaging",         label: "Imaging",             mode: "ortho", step: "1",    min: "0", entities: "spineone" },
  { key: "established",     label: "Established",         mode: "ortho", step: "1",    min: "0", aliases: ["establishedActual"] },
  { key: "noShows",         label: "No Shows",            mode: "ortho", step: "1",    min: "0" },
  { key: "cancelled",       label: "Cancelled",           mode: "ortho", step: "1",    min: "0" },
  { key: "reschedules",     label: "Rescheduled",         mode: "ortho", step: "1",    min: "0", entities: "non-laoss" },
  { key: "totalCalls",      label: "Total Calls",         mode: "ortho", step: "1",    min: "0", aliases: ["callVolume", "totalCallsActual"] },
  { key: "abandonedCalls",  label: "Abandoned Calls",     mode: "ortho", step: "1",    min: "0" },
  { key: "cashCollected",   label: "Cash Collected",      mode: "ortho", step: "0.01", min: "0", aliases: ["cashActual"] },
  { key: "ptoDays",         label: "PTO Days",            mode: "all",   step: "0.5",  min: "0", preserveOnOppositeMode: false },

  { section: "pt", sectionHeader: true, mode: "pt", label: "PT Inputs" },

  { key: "ptScheduledVisits",  label: "PT Scheduled Visits",   mode: "pt", step: "1", min: "0" },
  { key: "ptCancellations",    label: "PT Cancellations",      mode: "pt", step: "1", min: "0" },
  { key: "ptNoShows",          label: "PT No Shows",           mode: "pt", step: "1", min: "0" },
  { key: "ptReschedules",      label: "PT Reschedules",        mode: "pt", step: "1", min: "0" },
  { key: "ptTotalUnitsBilled", label: "PT Total Units Billed", mode: "pt", step: "1", min: "0" },
  { key: "ptVisitsSeen",       label: "PT Visits Seen (wk)",   mode: "pt", step: "1", min: "0" },
  { key: "ptWorkingDays",      label: "PT Working Days",       mode: "pt", step: "1", min: "1", defaultValue: () => String(getEntityWorkingDays(getSelectedEntity())) },

  { section: "spineOnePI", sectionHeader: true, mode: "ortho", entities: "spineone", label: "SpineOne PI" },

  { key: "piNp",             label: "PI NP",              mode: "ortho", step: "1",    min: "0", entities: "spineone" },
  { key: "piCashCollection", label: "PI Cash Collection", mode: "ortho", step: "0.01", min: "0", entities: "spineone" },

  { section: "calculated", sectionHeader: true, mode: "ortho", label: "Calculated" },

  { key: "visitVolume",       label: "Visit Volume",     mode: "ortho", step: "1",   readonly: true, derived: true, aliases: ["totalVisits"] },
  { key: "noShowRate",        label: "No Show %",        mode: "ortho", step: "0.1", readonly: true, derived: true },
  { key: "cancellationRate",  label: "Cancellation %",   mode: "ortho", step: "0.1", readonly: true, derived: true },
  { key: "abandonedCallRate", label: "Abandoned Call %", mode: "ortho", step: "0.1", readonly: true, derived: true, aliases: ["abandonmentRate"] },

  { section: "ptCalculated", sectionHeader: true, mode: "pt", label: "PT Calculated" },

  { key: "ptUnitsPerVisit", label: "PT Units/Visit",     mode: "pt", step: "0.01", readonly: true, derived: true },
  { key: "ptVisitsPerDay",  label: "PT Visits/Day (wk)", mode: "pt", step: "0.01", readonly: true, derived: true }
];

// Fields the user actually types into (not headers, not derived).
const ENTRY_INPUT_FIELDS = ENTRY_FIELDS.filter((f) => f.key && !f.derived && !f.sectionHeader);

// All fields with a storage key (input + derived, but not headers).
const ENTRY_STORED_FIELDS = ENTRY_FIELDS.filter((f) => f.key && !f.sectionHeader);

function fieldWrapperClass(f) {
  if (f.mode === "all") return "allEntryField";
  if (f.mode === "pt") return "ptField";
  // ortho mode
  if (f.entities === "spineone") return "spineOnlyField";
  if (f.entities === "non-laoss") return "nonPtField nonLaossField";
  return "nonPtField";
}

// Read a stored value from a weekly record, honoring the field's aliases.
function readStoredValue(values, field) {
  if (!values) return undefined;
  if (values[field.key] !== undefined) return values[field.key];
  if (field.aliases) {
    for (const a of field.aliases) {
      if (values[a] !== undefined) return values[a];
    }
  }
  return undefined;
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
    reschedules: normalizeNumber(values.reschedules),
    ptTotalUnitsBilled,
    ptVisitsSeen,
    ptWorkingDays,
    ptUnitsPerVisit,
    ptVisitsPerDay,

    ptoDays: normalizeNumber(values.ptoDays),
    cashCollected: normalizeNumber(values.cashCollected),
    piNp: normalizeNumber(values.piNp),
    piCashCollection: normalizeNumber(values.piCashCollection),
    imaging: normalizeNumber(values.imaging),
    operationsNarrative: String(values.operationsNarrative || "").trim()
  };
}

function updateDerivedDisplays() {
  // Read every input field's current value and recompute derived metrics.
  const currentValues = {};
  ENTRY_INPUT_FIELDS.forEach((f) => {
    const input = byId(f.key);
    currentValues[f.key] = input ? (input.value || "") : "";
  });
  if (!currentValues.ptWorkingDays) currentValues.ptWorkingDays = "5";

  const derived = calculateDerivedMetrics(currentValues);

  // Write derived values back to their read-only inputs.
  const writeDerived = (key, value) => {
    const el = byId(key);
    if (el) el.value = value;
  };
  writeDerived("visitVolume", derived.visitVolume ? String(derived.visitVolume) : "0");
  writeDerived("noShowRate", derived.noShowRate.toFixed(1));
  writeDerived("cancellationRate", derived.cancellationRate.toFixed(1));
  writeDerived("abandonedCallRate", derived.abandonedCallRate.toFixed(1));
  writeDerived("ptUnitsPerVisit", derived.ptUnitsPerVisit.toFixed(2));
  writeDerived("ptVisitsPerDay", derived.ptVisitsPerDay.toFixed(2));
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

  // Render from ENTRY_FIELDS registry. Every field on this form (HTML input,
  // save/load mapping, derived-metric recompute) reads from the same list so
  // adding a field is a single-line change to the registry.
  let html = "";
  for (const f of ENTRY_FIELDS) {
    const cls = fieldWrapperClass(f);
    const hidden = (f.mode !== "all") ? ' style="display:none;"' : "";
    if (f.sectionHeader) {
      html += `
    <div class="${cls}"${hidden}>
      <div class="formSectionBreak">
        <h4>${escapeHtml(f.label)}</h4>
      </div>
    </div>`;
      continue;
    }
    const readonlyAttr = f.readonly ? " readonly" : "";
    const stepAttr = f.step ? ` step="${escapeHtml(f.step)}"` : "";
    const minAttr = f.min !== undefined ? ` min="${escapeHtml(f.min)}"` : "";
    const defaultAttr = typeof f.defaultValue === "string" ? ` value="${escapeHtml(f.defaultValue)}"` : "";
    html += `
    <div class="${cls}"${hidden}>
      <label for="${escapeHtml(f.key)}">${escapeHtml(f.label)}</label>
      <input type="number" id="${escapeHtml(f.key)}"${stepAttr}${minAttr}${readonlyAttr}${defaultAttr} />
    </div>`;
  }
  html += `\n    ${renderNarrativePromptBox()}`;
  container.innerHTML = html;

  // Wire the input listeners for every editable field: recompute derived
  // values on change, and schedule a localStorage draft autosave.
  ENTRY_INPUT_FIELDS.forEach((f) => {
    const input = byId(f.key);
    if (!input) return;
    input.addEventListener("input", () => {
      updateDerivedDisplays();
      scheduleDraftAutosave();
    });
  });
  // Ops narrative is outside the registry but we still want to autosave it.
  const notesEl = byId("operationsNarrative");
  if (notesEl) notesEl.addEventListener("input", scheduleDraftAutosave);

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

  // Reschedules: show for NES, SpineOne, MRO — not LAOSS
  const baseEntity = getCurrentEntityOption().baseEntity;
  const hasReschedules = !ptMode && baseEntity !== "LAOSS";
  document.querySelectorAll(".nonLaossField").forEach((el) => {
    el.style.display = hasReschedules ? "" : "none";
  });

  // Show working days hint for entity
  const workingDaysHint = byId("workingDaysHint");
  if (workingDaysHint) {
    const days = getEntityWorkingDays(baseEntity);
    workingDaysHint.textContent = `${baseEntity} operates ${days} days/week — rates calculated accordingly.`;
    workingDaysHint.style.display = days !== 5 ? "" : "none";
  }
}

// Map a stored Azure Table record into the shape the form expects.
// Honors each field's legacy aliases (e.g. `totalVisits` → `visitVolume`),
// then overwrites derived fields with freshly-computed values.
function mapWeeklyValuesToFormData(values) {
  const mapped = {};
  ENTRY_STORED_FIELDS.forEach((f) => {
    const stored = readStoredValue(values, f);
    if (stored !== undefined) {
      mapped[f.key] = stored;
    } else if (f.defaultValue !== undefined) {
      mapped[f.key] = typeof f.defaultValue === "function" ? f.defaultValue() : f.defaultValue;
    } else {
      mapped[f.key] = "";
    }
  });
  mapped.operationsNarrative = values?.operationsNarrative ?? "";

  // Always recompute derived fields so stored stale values can't override the
  // canonical formulas (e.g. visitVolume = NP + Surgeries + Established).
  const derived = calculateDerivedMetrics(mapped);
  ENTRY_STORED_FIELDS.filter((f) => f.derived).forEach((f) => {
    if (derived[f.key] !== undefined) mapped[f.key] = derived[f.key];
  });
  return mapped;
}

// Populate all form inputs from a stored record. Both input and derived
// fields get written so the user sees the full form state.
function setFormValues(data) {
  const mapped = mapWeeklyValuesToFormData(data || {});
  ENTRY_STORED_FIELDS.forEach((f) => {
    const input = byId(f.key);
    if (!input) return;
    input.value = mapped[f.key] !== null && mapped[f.key] !== undefined ? mapped[f.key] : "";
  });

  // Ops narrative lives outside the registry (textarea, not number input).
  const notesEl = byId("operationsNarrative");
  if (notesEl) notesEl.value = data?.operationsNarrative ?? "";

  syncEntryModeVisibility();
  updateDerivedDisplays();
}

// Collect values from the form + preserve the opposite mode's stored values.
// Critical invariant: saving in PT mode MUST NOT wipe out the ortho fields
// (and vice versa). Breaking this caused the 4/17 PT zero-out incident.
function getFormValues() {
  const ptMode = entityHasPtEntry();
  // Azure Table records come back flat at top level (no `values` wrapper).
  // Fall back to the record itself if the wrapped shape is absent.
  // A PT-only record has visitVolume == null but PT fields populated; the old
  // guard dropped it on the floor, letting the next ortho save wipe PT to 0.
  // Trust any flat record marked `found: true`.
  const existing = currentWeekData?.values || currentWeekData?.data ||
    (currentWeekData?.found ? currentWeekData : {});

  const raw = {};
  ENTRY_INPUT_FIELDS.forEach((f) => {
    // Is this field live in the current mode?
    const fieldLive = f.mode === "all" || (f.mode === "pt" && ptMode) || (f.mode === "ortho" && !ptMode);
    const preserve = f.preserveOnOppositeMode !== false;

    if (fieldLive) {
      // Read from form. Fall back to default if empty and a default is defined.
      const inputVal = byId(f.key)?.value;
      if (inputVal !== undefined && inputVal !== "") {
        raw[f.key] = inputVal;
      } else if (f.defaultValue !== undefined) {
        raw[f.key] = typeof f.defaultValue === "function" ? f.defaultValue() : f.defaultValue;
      } else {
        raw[f.key] = "";
      }
    } else if (preserve) {
      // Opposite mode — keep whatever's on the existing record so we don't wipe it.
      // If no stored value, fall back to the field's defaultValue (e.g. ptWorkingDays)
      // or 0 so downstream calculations don't divide by zero unexpectedly.
      const stored = readStoredValue(existing, f);
      if (stored !== undefined) {
        raw[f.key] = stored;
      } else if (f.defaultValue !== undefined) {
        raw[f.key] = typeof f.defaultValue === "function" ? f.defaultValue() : f.defaultValue;
      } else {
        raw[f.key] = 0;
      }
    } else {
      // Field is mode=all but not live (shouldn't happen) — fall through to form read.
      raw[f.key] = byId(f.key)?.value || "";
    }
  });

  raw.operationsNarrative = byId("operationsNarrative")?.value || "";

  const derived = calculateDerivedMetrics(raw);

  // Build the save payload: every stored field, with derived values recomputed
  // and rates rounded to 2 decimals.
  const payload = {};
  ENTRY_STORED_FIELDS.forEach((f) => {
    const v = derived[f.key] !== undefined ? derived[f.key] : raw[f.key];
    if (f.derived && typeof v === "number" && !Number.isInteger(v)) {
      payload[f.key] = Number(v.toFixed(2));
    } else {
      payload[f.key] = v;
    }
  });
  payload.operationsNarrative = derived.operationsNarrative;
  payload.callVolume = derived.callVolume;
  return payload;
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
    <strong>Created by:</strong> ${escapeHtml(data.createdBy || "n/a")}<br />
    <strong>Created at:</strong> ${escapeHtml(formatDateTime(data.createdAt))}<br />
    <strong>Last updated by:</strong> ${escapeHtml(data.updatedBy || "n/a")}<br />
    <strong>Last updated at:</strong> ${escapeHtml(formatDateTime(data.updatedAt))}<br />
    <strong>Status:</strong> ${escapeHtml(data.status || "saved")}
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
    <div class="entryBrandCard">
      <div class="entryBrandLogo">
        <img src="${brand.logo}" alt="${brand.label}" />
      </div>
      <div>
        <div class="entryBrandLabel" style="color:${brand.accent};">${brand.label}</div>
        <div class="entryBrandFullName">${brand.fullName}</div>
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

function renderDataCompletenessBanner(current, entityScope, weekSets) {
  const container = byId("dashboardCompletenessBanner");
  if (!container) return;

  // PT scopes map to base entities for completeness check
  const entities = isPtScope(entityScope)
    ? ptScopeEntities(entityScope)
    : entityScope === "ALL" ? ENTITIES : [entityScope];
  const missing = [];
  const hasData = [];

  (current.regions || []).forEach(r => {
    if (entities.includes(r.entity)) {
      if (r.status === "missing" || normalizeNumber(r.visitVolume) === 0) {
        missing.push(r.entity);
      } else {
        hasData.push(r.entity);
      }
    }
  });

  // Check entities not in regions at all
  entities.forEach(e => {
    const found = (current.regions || []).find(r => r.entity === e);
    if (!found && !missing.includes(e)) missing.push(e);
  });

  const total = entities.length;
  const entered = total - missing.length;
  const weekLabel = weekSets?.primaryWeeks?.[0]
    ? new Date(weekSets.primaryWeeks[0] + "T12:00:00Z").toLocaleDateString("en-US", { month: "short", day: "numeric" })
    : "selected week";

  const entityWord = total === 1 ? "entity" : "entities";
  const scopeLabel = isPtScope(entityScope) ? " PT programs" : ` ${entityWord}`;

  if (missing.length === 0) {
    container.innerHTML = `
      <div class="completenessBanner completenessGood">
        <span class="completenessIcon">✓</span>
        <span class="completenessText">All <strong>${total}</strong>${scopeLabel} have data entered for the week of ${weekLabel}.</span>
      </div>`;
  } else {
    container.innerHTML = `
      <div class="completenessBanner completenessWarn">
        <span class="completenessIcon">◆</span>
        <span class="completenessText">
          <strong>${entered} of ${total}</strong>${scopeLabel} have data for the week of ${weekLabel}.
          Missing: <strong>${missing.join(", ")}</strong>.
          <span class="completenessAction" onclick="showEntryViewForMissing('${missing[0]}')">Enter data →</span>
        </span>
      </div>`;
  }
}

function showEntryViewForMissing(entity) {
  // Switch to weekly entry with the missing entity pre-selected
  const select = byId("entitySelect");
  if (select) {
    select.value = entity;
    select.dispatchEvent(new Event("change"));
  }
  showEntryView();
}

function renderDashboardCards(current, comparison, compareAgainst, entityScope, weekSets) {
  const ptVisitsCurrent = normalizeNumber(current.ptTotals?.visitsSeen);
  const ptVisitsComparison = normalizeNumber(comparison.ptTotals?.visitsSeen);

  const avgNoShow = averageMetric(current.regions, "noShowRate");
  const avgCancel = averageMetric(current.regions, "cancellationRate");
  const avgCxnsCombined = avgNoShow + avgCancel;

  // For entities/scopes with PT, include PT visits seen in the combined visit volume
  const hasPt = entityHasPtData(entityScope) || entityScope === "ALL";
  const orthoVisitCurrent    = normalizeNumber(current.totals?.visitVolume);
  const orthoVisitComparison = normalizeNumber(comparison.totals?.visitVolume);
  const visitCurrent    = hasPt ? orthoVisitCurrent    + normalizeNumber(current.ptTotals?.visitsSeen)    : orthoVisitCurrent;
  const visitComparison = hasPt ? orthoVisitComparison + normalizeNumber(comparison.ptTotals?.visitsSeen) : orthoVisitComparison;
  const visitVariance = visitCurrent - visitComparison;
  const numWeeks = weekSets?.primaryWeeks?.length || 1;
  const daysPerWeek = entityScope === "NES" ? 4.5 : 5;
  const totalDays = numWeeks * daysPerWeek;
  const visitPerDay = totalDays > 0 ? (visitCurrent / totalDays).toFixed(1) : "—";

  const isVsBudget = compareAgainst === "budget";

  const cards = [
    {
      // Budget comparison uses totalVisitsBudget (ortho + PT + imaging) so it
      // matches the PT-lumped actuals above. Falls back to visitVolumeBudget
      // for pre-2026-xlsx data that only has the ortho total.
      label: "Visit Volume",
      value: formatWhole(visitCurrent),
      movement: isVsBudget
        ? buildKpiMovement(visitCurrent, normalizeNumber(current.budgetTotals?.totalVisitsBudget || current.budgetTotals?.visitVolumeBudget), formatWhole)
        : buildKpiMovement(visitCurrent, visitComparison, formatWhole),
      meta: isVsBudget
        ? `Budget ${formatWhole(current.budgetTotals?.totalVisitsBudget || current.budgetTotals?.visitVolumeBudget || 0)} · ${visitPerDay}/day avg`
        : `${visitVariance >= 0 ? "+" : ""}${formatWhole(visitVariance)} vs prior · ${visitPerDay}/day avg`,
      className: isVsBudget
        ? getTrendClass(visitCurrent, current.budgetTotals?.totalVisitsBudget || current.budgetTotals?.visitVolumeBudget)
        : getTrendClass(visitCurrent, visitComparison)
    },
    {
      label: "New Patients",
      value: formatWhole(current.totals?.newPatients || 0),
      movement: isVsBudget
        ? buildKpiMovement(normalizeNumber(current.totals?.newPatients), normalizeNumber(current.budgetTotals?.newPatientsBudget), formatWhole)
        : buildKpiMovement(normalizeNumber(current.totals?.newPatients), normalizeNumber(comparison.totals?.newPatients), formatWhole),
      meta: isVsBudget
        ? `Budget ${formatWhole(current.budgetTotals?.newPatientsBudget || 0)}`
        : `${((current.totals?.newPatients||0)-(comparison.totals?.newPatients||0)) >= 0 ? "+" : ""}${formatWhole((current.totals?.newPatients||0)-(comparison.totals?.newPatients||0))} vs prior`,
      className: isVsBudget
        ? getTrendClass(normalizeNumber(current.totals?.newPatients), normalizeNumber(current.budgetTotals?.newPatientsBudget))
        : getTrendClass(normalizeNumber(current.totals?.newPatients), normalizeNumber(comparison.totals?.newPatients))
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
    // PTO Days removed 2026-04-23: PTO is tracked via PTO Forecast (not weekly entry).
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
      label: "Imaging",
      value: formatWhole(current.totals?.imaging || 0),
      meta: "SpineOne only · across selected period",
      className: "kpi-neutral"
    });

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
          imaging: 0,
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

        // Tessa: lump PT visits into total visits for entities that have a PT line
        // (NES, SpineOne, MRO). Entity cards previously showed ortho-only, which
        // didn't match the top-line KPI card that already combines both.
        const hasPt = entityHasPtData(entity);
        const ptVisits = hasPt ? normalizeNumber(row.pt?.visitsSeen) : 0;
        const priorPtVisits = hasPt ? normalizeNumber(prior.pt?.visitsSeen) : 0;
        const totalVisits = normalizeNumber(row.visitVolume) + ptVisits;
        const priorTotalVisits = normalizeNumber(prior.visitVolume) + priorPtVisits;
        // Budget now carries PT + Surgery + Imaging + ortho totals (2026 xlsx ingest).
        // totalVisitsBudget = ortho + PT + imaging; fall back to ortho + PT if the
        // row somehow predates the xlsx ingest.
        const ptVisitsBudget = normalizeNumber(row.ptVisitsBudget);
        const totalVisitsBudget = normalizeNumber(row.totalVisitsBudget) ||
          (normalizeNumber(row.visitVolumeBudget) + ptVisitsBudget + normalizeNumber(row.imagingBudget));

        const visitPct = buildVariancePct(totalVisits, priorTotalVisits);
        const callPct = buildVariancePct(row.callVolume, prior.callVolume);
        const npPct = buildVariancePct(row.newPatients, prior.newPatients);
        const ptPct = buildVariancePct(ptVisits, priorPtVisits);
        const ptBudgetPct = buildVariancePct(ptVisits, ptVisitsBudget);

        const visitBudgetStatus = getGoalStatus(totalVisits, totalVisitsBudget);
        const ptBudgetStatus = ptVisitsBudget > 0 ? getGoalStatus(ptVisits, ptVisitsBudget) : "neutral";
        const npBudgetStatus = getGoalStatus(row.newPatients, row.newPatientsBudget);
        const callStatus    = getGoalStatus(row.abandonedCallRate, getThreshold("abandonedCallRate", 10), true);
        const noShowStatus  = getGoalStatus(row.noShowRate, getThreshold("noShowRate", 6), true);
        const cancelStatus  = getGoalStatus(row.cancellationRate, getThreshold("cancellationRate", 8), true);
        const summary = getPerformanceSummary(row, compareAgainst);

        const visitGoalPct = progressPercent(totalVisits, totalVisitsBudget);
        const npGoalPct = progressPercent(row.newPatients, row.newPatientsBudget);
        const accessPct = accessPercentFromAbandoned(row.abandonedCallRate, 10);

        const fmtPct = (value) => value === null ? "n/a" : `${value >= 0 ? "+" : ""}${value.toFixed(1)}%`;

        const comparisonBlock = compareAgainst === "budget"
          ? `
            <div class="entityBudgetGrid">
              <div class="entityBudgetTile">
                <div class="entityBudgetLabel">Visit Budget</div>
                <div class="entityBudgetValue">${formatWhole(totalVisitsBudget)}</div>
                <div class="${getMetricChipClass(visitBudgetStatus)}">
                  ${visitBudgetStatus === "good" ? "Above Goal" : visitBudgetStatus === "warning" ? "Near Goal" : "Below Goal"}
                </div>
                <div class="entityBudgetMeta">Variance ${formatVariance(totalVisits, totalVisitsBudget)}</div>
                <div class="entityBudgetMeta">To Goal ${formatToGoal(totalVisits, totalVisitsBudget)}</div>
                <div class="entityProgressWrap">
                  <div class="entityProgressLabelRow">
                    <span>Goal Progress</span>
                    <strong>${formatToGoal(totalVisits, totalVisitsBudget)}</strong>
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
                <span class="entityMetricLabel">Visits${hasPt ? " (incl. PT)" : ""}</span>
                <strong>${formatWhole(totalVisits)}</strong>
                ${hasPt ? `<span class="entityMetricSub">Ortho ${formatWhole(row.visitVolume)} · PT ${formatWhole(ptVisits)}</span>` : ""}
              </div>
              ${hasPt ? `
              <div class="entityMetricHero">
                <span class="entityMetricLabel">PT Visits</span>
                <strong>${formatWhole(ptVisits)}</strong>
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

            <div class="entityLivePhones" id="livePhones_${entity}" data-entity="${entity}">
              <div class="entityLivePhonesHeader">
                <span class="entityLivePhonesLabel">Live Phones (this week)</span>
                <span class="entityLivePhonesState">loading…</span>
              </div>
              <div class="entityLivePhonesGrid"></div>
            </div>

            <details class="entityDetailDrawer">
              <summary>More detail</summary>
              <div class="entityDetailGrid">
                <div class="entityDetailTile">
                  <div class="entityDetailLabel">Visit Variance</div>
                  <div class="entityDetailValue">${compareAgainst === "budget" ? formatVariance(totalVisits, totalVisitsBudget) : fmtPct(visitPct)}</div>
                </div>
                ${hasPt ? `
                <div class="entityDetailTile">
                  <div class="entityDetailLabel">PT Variance ${compareAgainst === "budget" ? "(vs Budget)" : "(vs Prior)"}</div>
                  <div class="entityDetailValue">${compareAgainst === "budget" ? (ptVisitsBudget > 0 ? formatVariance(ptVisits, ptVisitsBudget) : "n/a") : fmtPct(ptPct)}</div>
                </div>` : ""}
                <div class="entityDetailTile">
                  <div class="entityDetailLabel">PTO</div>
                  <div class="entityDetailValue"><a href="#" onclick="showPtoForecastView();return false;" class="entityDetailLink">View Forecast →</a></div>
                </div>
              </div>
              <!-- Ops narrative is now rendered in the dashboard-wide Ops Notes
                   panel below the entity grid, not duplicated per-card. -->
              ${""}
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

    const noShowLimit    = getThreshold("noShowRate", 6);
    const cancelLimit    = getThreshold("cancellationRate", 8);
    const abandonedLimit = getThreshold("abandonedCallRate", 10);

    if (normalizeNumber(row.noShowRate) >= noShowLimit) {
      alerts.push({ severity: "bad", text: `${entity} no show rate is elevated at ${normalizeNumber(row.noShowRate).toFixed(1)}% (threshold: ${noShowLimit}%).` });
    }

    if (normalizeNumber(row.cancellationRate) >= cancelLimit) {
      alerts.push({ severity: "bad", text: `${entity} cancellation rate is elevated at ${normalizeNumber(row.cancellationRate).toFixed(1)}% (threshold: ${cancelLimit}%).` });
    }

    if (normalizeNumber(row.abandonedCallRate) >= abandonedLimit) {
      alerts.push({ severity: "bad", text: `${entity} abandoned call rate is elevated at ${normalizeNumber(row.abandonedCallRate).toFixed(1)}% (threshold: ${abandonedLimit}%).` });
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

// Show each entity's weekly narrative as a stand-alone panel on the dashboard.
// The ops narrative is the highest-signal qualitative field in the system
// (regional leads explain *why* the numbers moved) and used to be buried
// behind an entity card drawer — promoting it makes cross-entity context
// easy to skim like a Slack channel.
function renderDashboardOpsNotes(current, entityScope, weekSets) {
  const container = byId("dashboardOpsNotes");
  const panel = byId("dashboardOpsNotesPanel");
  if (!container) return;

  const entities = entityScope === "ALL" ? ENTITIES : (isPtScope(entityScope) ? [] : [entityScope]);
  if (!entities.length) {
    if (panel) panel.style.display = "none";
    return;
  }
  if (panel) panel.style.display = "";

  const currentMap = getEntityMap(current);
  const weekLabel = weekSets?.primaryWeeks?.length === 1
    ? `week ending ${weekSets.primaryWeeks[0]}`
    : `selected period`;

  const cards = entities.map((entity) => {
    const row = currentMap[entity];
    const brand = getBranding(entity);
    const rawNarrative = String(row?.operationsNarrative || "").trim();
    const hasNote = !!rawNarrative;

    const body = hasNote
      ? `<div class="opsNotesBody">${escapeHtml(rawNarrative).replace(/\n/g, "<br />")}</div>`
      : `<div class="opsNotesEmpty">
          No note for ${escapeHtml(weekLabel)}.
          <span class="opsNotesAction" onclick="showEntryViewForMissing('${escapeHtml(entity)}')">Add one →</span>
        </div>`;

    return `
      <div class="opsNotesCard" style="border-top:3px solid ${brand.accent};">
        <div class="opsNotesHeader">
          <div class="opsNotesEntity">${escapeHtml(entity)}</div>
          <div class="opsNotesSubtitle">${escapeHtml(brand.fullName)}</div>
        </div>
        ${body}
      </div>
    `;
  }).join("");

  container.innerHTML = `<div class="opsNotesGrid">${cards}</div>`;
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
            <td><strong>${escapeHtml(item.entity)}</strong></td>
            <td>${escapeHtml(item.weekEnding || "—")}</td>
            <td><span class="${statusPillClass(item.status)}">${escapeHtml(item.status || "saved")}</span></td>
            <td>${escapeHtml(item.updatedBy || "—")}</td>
            <td>
              <div>${escapeHtml(item.updatedAt ? formatDateTime(item.updatedAt) : "—")}</div>
            </td>
            <td>${escapeHtml(item.createdBy || "—")}</td>
            <td>${escapeHtml(item.createdAt ? formatDateTime(item.createdAt) : "—")}</td>
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

  // Series palette matches design tokens: visits=blue, PT=green, calls=amber, NP=purple
  window.visitsChartInstance = new Chart(ctx, {
    type: "line",
    data: {
      labels,
      datasets: [
        {
          label: "Visits",
          data: visitData,
          tension: 0.3,
          borderColor: "#2D4BA3",
          backgroundColor: "rgba(45,75,163,0.04)",
          borderWidth: 1.8,
          pointRadius: 2,
          pointHoverRadius: 4,
          fill: false
        },
        {
          label: "PT Visits",
          data: ptData,
          tension: 0.3,
          borderColor: "#1E6B4F",
          backgroundColor: "rgba(30,107,79,0.04)",
          borderWidth: 1.8,
          pointRadius: 2,
          pointHoverRadius: 4,
          fill: false
        },
        {
          label: "Calls",
          data: callData,
          tension: 0.3,
          borderColor: "#B8893D",
          backgroundColor: "rgba(184,137,61,0.04)",
          borderWidth: 1.8,
          pointRadius: 2,
          pointHoverRadius: 4,
          fill: false,
          hidden: true
        },
        {
          label: "New Patients",
          data: npData,
          tension: 0.3,
          borderColor: "#8B5CF6",
          backgroundColor: "rgba(139,92,246,0.04)",
          borderWidth: 1.8,
          pointRadius: 2,
          pointHoverRadius: 4,
          fill: false
        },
        ...(compareAgainst === "budget"
          ? [
              {
                label: "Visit Budget",
                data: budgetData,
                tension: 0.3,
                borderColor: "#B23A3A",
                backgroundColor: "rgba(178,58,58,0.04)",
                borderDash: [6, 6],
                borderWidth: 1.6,
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
      interaction: { mode: "index", intersect: false },
      plugins: {
        legend: {
          labels: {
            color: "#6B7280",
            boxWidth: 10,
            boxHeight: 10,
            font: { size: 11.5, family: "Geist, 'Inter Tight', sans-serif" }
          }
        }
      },
      scales: {
        x: {
          ticks: { color: "#6B7280", font: { size: 10, family: "Geist Mono, monospace" } },
          grid: { color: "#ECEAE3", drawBorder: false }
        },
        y: {
          ticks: { color: "#6B7280", font: { size: 10, family: "Geist Mono, monospace" } },
          grid: { color: "#ECEAE3", drawBorder: false }
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
    // PTO Days removed 2026-04-23: tracked in PTO Forecast, not weekly entry.
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
    },
    ...(normalizeNumber(summary.totals?.imaging) > 0 ? [{
      label: "Imaging",
      value: formatWhole(summary.totals?.imaging || 0),
      meta: "SpineOne only",
      className: "kpi-neutral"
    }] : [])
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

  const entity = getSelectedTrendsEntity();
  const baseCards = [
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
  ];

  if (entity === "SpineOne") {
    baseCards.push({
      label: "PI New Patients",
      value: latest ? formatDelta(latest.piNp || 0, previous?.piNp || 0) : "-",
      className: latest ? getTrendClass(latest.piNp, previous?.piNp) : "kpi-neutral"
    });
    baseCards.push({
      label: "PI Cash Collection",
      value: latest ? formatCurrency(latest.piCashCollection || 0) : "-",
      meta: previous ? `Prior: ${formatCurrency(previous.piCashCollection || 0)}` : "",
      className: latest ? getTrendClass(latest.piCashCollection, previous?.piCashCollection) : "kpi-neutral"
    });
  }

  renderMetricCards("trendsCards", baseCards);
}

function renderTrendsTable(result) {
  const wrap = byId("trendsTableWrap");
  if (!wrap) return;

  const entity = getSelectedTrendsEntity();
  const items = result.items || [];

  if (!items.length) {
    wrap.innerHTML = "<p>No trend data found for this entity.</p>";
    return;
  }

  // Show/hide SpineOne PI section and render chart
  const piSection = byId("trendsPiSection");
  if (piSection) piSection.style.display = entity === "SpineOne" ? "" : "none";
  if (entity === "SpineOne") renderSpineOnePiChart(items);

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
          ${entity === "SpineOne" ? "<th>PI NP</th><th>PI Cash</th>" : ""}
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
            ${entity === "SpineOne" ? `<td>${formatWhole(item.piNp || 0)}</td><td>${formatCurrency(item.piCashCollection || 0)}</td>` : ""}
          </tr>
        `).join("")}
      </tbody>
    </table>
  `;
}

function hideAllViews() {
  const ids = ["dashboardView", "entryView", "executiveView", "trendsView", "ptoForecastView", "activityView", "importView", "providerSettingsView", "changelogView", "goalSettingsView", "thresholdSettingsView", "helpView"];
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
    "navProviderSettingsBtn",
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

  // Default to most recent Friday if not already set
  const filter = byId("activityWeekFilter");
  if (filter && !filter.value) {
    filter.value = getDefaultWeekEnding();
  }
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


// ── PTO Weekly Breakdown Helpers ─────────────────────────────
function getFridaysInMonth(monthKey) {
  // Returns array of {date, label} for each Friday in the month
  const [yyyy, mm] = monthKey.split("-").map(Number);
  const fridays = [];
  const d = new Date(Date.UTC(yyyy, mm - 1, 1));
  // Advance to first Friday
  while (d.getUTCDay() !== 5) d.setUTCDate(d.getUTCDate() + 1);
  let week = 1;
  while (d.getUTCMonth() === mm - 1) {
    const iso = d.toISOString().slice(0, 10);
    const label = `Wk ${week} — ${d.toLocaleDateString("en-US", { month: "short", day: "numeric", timeZone: "UTC" })}`;
    fridays.push({ date: iso, label, week });
    d.setUTCDate(d.getUTCDate() + 7);
    week++;
  }
  return fridays;
}

function getPtoWeeklyMode(entity, monthKey) {
  return localStorage.getItem(`ptoWeeklyMode_${entity}_${monthKey}`) === "1";
}

function setPtoWeeklyMode(entity, monthKey, on) {
  localStorage.setItem(`ptoWeeklyMode_${entity}_${monthKey}`, on ? "1" : "0");
}

function getPtoWeekValue(entity, monthKey, date, type) {
  const key = `ptoWeek_${entity}_${monthKey}_${date}_${type}`;
  const v = localStorage.getItem(key);
  return v !== null ? parseFloat(v) || 0 : 0;
}

function setPtoWeekValue(entity, monthKey, date, type, value) {
  localStorage.setItem(`ptoWeek_${entity}_${monthKey}_${date}_${type}`, String(value));
}

function sumPtoWeekly(entity, monthKey, type) {
  const fridays = getFridaysInMonth(monthKey);
  return fridays.reduce((sum, f) => sum + getPtoWeekValue(entity, monthKey, f.date, type), 0);
}

function renderPtoWeeklyRows(entity, monthKey, hasPt, canEdit) {
  const fridays = getFridaysInMonth(monthKey);
  const disabledAttr = canEdit ? "" : "readonly disabled";

  return `
    <div class="ptoWeeklyTable">
      <div class="ptoWeeklyHeader">
        <span class="ptoWeeklyCol ptoWeeklyColLabel">Week Ending</span>
        <span class="ptoWeeklyCol">MD Days</span>
        <span class="ptoWeeklyCol">PA Days</span>
        ${hasPt ? `<span class="ptoWeeklyCol">PT Days</span>` : ""}
      </div>
      ${fridays.map(f => `
        <div class="ptoWeeklyRow" data-date="${f.date}">
          <span class="ptoWeeklyCol ptoWeeklyColLabel">${f.label}</span>
          <input type="number" class="ptoWeekInput ptoWeekInputMd"
            data-entity="${entity}" data-monthkey="${monthKey}" data-date="${f.date}" data-type="md"
            step="0.5" min="0" placeholder="0"
            value="${getPtoWeekValue(entity, monthKey, f.date, "md") || ""}"
            ${disabledAttr} />
          <input type="number" class="ptoWeekInput ptoWeekInputPa"
            data-entity="${entity}" data-monthkey="${monthKey}" data-date="${f.date}" data-type="pa"
            step="0.5" min="0" placeholder="0"
            value="${getPtoWeekValue(entity, monthKey, f.date, "pa") || ""}"
            ${disabledAttr} />
          ${hasPt ? `
          <input type="number" class="ptoWeekInput ptoWeekInputPt"
            data-entity="${entity}" data-monthkey="${monthKey}" data-date="${f.date}" data-type="pt"
            step="0.5" min="0" placeholder="0"
            value="${getPtoWeekValue(entity, monthKey, f.date, "pt") || ""}"
            ${disabledAttr} />` : ""}
        </div>
      `).join("")}
      <div class="ptoWeeklyTotalsRow">
        <span class="ptoWeeklyCol ptoWeeklyColLabel">Month Total</span>
        <span class="ptoWeeklyTotal" id="ptoWeekTotal_md_${entity}_${monthKey}">
          ${sumPtoWeekly(entity, monthKey, "md") || "—"}
        </span>
        <span class="ptoWeeklyTotal" id="ptoWeekTotal_pa_${entity}_${monthKey}">
          ${sumPtoWeekly(entity, monthKey, "pa") || "—"}
        </span>
        ${hasPt ? `
        <span class="ptoWeeklyTotal" id="ptoWeekTotal_pt_${entity}_${monthKey}">
          ${sumPtoWeekly(entity, monthKey, "pt") || "—"}
        </span>` : ""}
      </div>
    </div>
  `;
}

// ── PTO Provider Breakdown Helpers ───────────────────────────
function syncPtoTotalsFromProviders(entity, monthKey) {
  // Sum provider days by type and update the aggregate inputs
  const rows = document.querySelectorAll(`.ptoProviderName[data-entity="${entity}"][data-monthkey="${monthKey}"]`);
  let mdTotal = 0, paTotal = 0, ptTotal = 0;
  rows.forEach(nameInput => {
    const container = nameInput.closest(".ptoProviderRow");
    if (!container) return;
    const type = container.querySelector(".ptoProviderType")?.value;
    const days = parseFloat(container.querySelector(".ptoProviderDays")?.value || "0") || 0;
    if (type === "MD") mdTotal += days;
    else if (type === "PA") paTotal += days;
    else if (type === "PT") ptTotal += days;
  });

  // Find aggregate inputs for this entity/month and update
  const mdInput = document.querySelector(`.ptoInput[data-entity="${entity}"][data-monthkey="${monthKey}"][data-type="md"]`);
  const paInput = document.querySelector(`.ptoInput[data-entity="${entity}"][data-monthkey="${monthKey}"][data-type="pa"]`);
  const ptInput = document.querySelector(`.ptoInput[data-entity="${entity}"][data-monthkey="${monthKey}"][data-type="pt"]`);
  if (mdInput && mdTotal > 0) mdInput.value = mdTotal;
  if (paInput && paTotal > 0) paInput.value = paTotal;
  if (ptInput && ptTotal > 0) ptInput.value = ptTotal;
}

function collectProviderBreakdown(entity, monthKey) {
  const rows = document.querySelectorAll(`.ptoProviderName[data-entity="${entity}"][data-monthkey="${monthKey}"]`);
  const breakdown = [];
  rows.forEach(nameInput => {
    const container = nameInput.closest(".ptoProviderRow");
    if (!container) return;
    const name = nameInput.value.trim();
    const type = container.querySelector(".ptoProviderType")?.value || "MD";
    const days = parseFloat(container.querySelector(".ptoProviderDays")?.value || "0") || 0;
    if (name || days > 0) breakdown.push({ name, type, days });
  });
  return breakdown;
}

function renderPtoForecastEntities(data) {
  const container = byId("ptoForecastEntities");
  if (!container) return;

  const entities = data?.entities || [];
  const monthLabels = data?.monthLabels || [];
  const monthKeys = data?.monthKeys || [];

  if (!entities.length) {
    container.innerHTML = `<div class="subtleBanner" style="margin:0;">No entity data available. Run a weekly import first.</div>`;
    return;
  }

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
              Rates based on last ${entityData.rates.weeksUsed} weeks &nbsp;·&nbsp;
              ${entityData.rates.clinicalVisitsPerDay} clinical visits/day &nbsp;·&nbsp;
              ${entityData.providerSettings?.mdCount || 0} MDs &nbsp;·&nbsp;
              ${entityData.providerSettings?.paCount || 0} PAs &nbsp;·&nbsp;
              ${entityData.providerSettings?.ptCount || 0} PTs
            </div>
          </div>
          <div class="entityLogoWrap" style="width:80px;height:40px;">
            <img src="${brand.logo}" alt="${entity}" class="entityLogo" />
          </div>
        </div>

        <div class="ptoMonthGrid">
          ${months.map((month, i) => `
            <div class="ptoMonthCard" id="ptoMonthCard_${entity}_${i}">
              <div class="ptoMonthCardHeader">
                <div class="ptoMonthLabel">${month.monthLabel}</div>
                ${(() => {
                  const isAdmin = currentUser?.access?.isAdmin;
                  const ue = currentUser?.access?.entity;
                  const canEdit = isAdmin || !ue || ue === entity;
                  if (!canEdit) return "";
                  const isWeekly = getPtoWeeklyMode(entity, month.monthKey);
                  return `
                    <div class="ptoModeToggle">
                      <button type="button" class="ptoModeBtn ${!isWeekly ? "ptoModeBtnActive" : ""}"
                        data-entity="${entity}" data-monthkey="${month.monthKey}" data-mode="monthly">
                        Monthly
                      </button>
                      <button type="button" class="ptoModeBtn ${isWeekly ? "ptoModeBtnActive" : ""}"
                        data-entity="${entity}" data-monthkey="${month.monthKey}" data-mode="weekly">
                        By Week
                      </button>
                    </div>`;
                })()}
              </div>

              ${(() => {
                const isAdmin = currentUser?.access?.isAdmin;
                const ue = currentUser?.access?.entity;
                const canEdit = isAdmin || !ue || ue === entity;
                const hasPt = entityData.providerSettings?.ptCount > 0;
                const isWeekly = getPtoWeeklyMode(entity, month.monthKey);
                const disabledAttr = canEdit ? "" : "readonly disabled";

                if (isWeekly) {
                  return renderPtoWeeklyRows(entity, month.monthKey, hasPt, canEdit);
                }

                return `
                <div class="ptoInputRow">
                  <div class="ptoInputGroup">
                    <label class="ptoInputLabel" for="pto_md_${entity}_${i}">MD PTO Days</label>
                    <input type="number" id="pto_md_${entity}_${i}" class="ptoInput"
                      data-entity="${entity}" data-monthkey="${month.monthKey}" data-type="md"
                      step="0.5" min="0" value="${month.mdPtoDays || ""}" placeholder="0"
                      ${disabledAttr} />
                  </div>
                  <div class="ptoInputGroup">
                    <label class="ptoInputLabel" for="pto_pa_${entity}_${i}">PA PTO Days</label>
                    <input type="number" id="pto_pa_${entity}_${i}" class="ptoInput"
                      data-entity="${entity}" data-monthkey="${month.monthKey}" data-type="pa"
                      step="0.5" min="0" value="${month.paPtoDays || ""}" placeholder="0"
                      ${disabledAttr} />
                  </div>
                  ${hasPt ? `
                  <div class="ptoInputGroup">
                    <label class="ptoInputLabel" for="pto_pt_${entity}_${i}">PT PTO Days</label>
                    <input type="number" id="pto_pt_${entity}_${i}" class="ptoInput"
                      data-entity="${entity}" data-monthkey="${month.monthKey}" data-type="pt"
                      step="0.5" min="0" value="${month.ptPtoDays || ""}" placeholder="0"
                      ${disabledAttr} />
                  </div>` : ""}
                </div>`;
              })()}

              ${(month.providerBreakdown?.length > 0) ? `
              <details class="ptoProviderBreakdown">
                <summary class="ptoProviderBreakdownToggle">Provider breakdown &nbsp;·&nbsp; ${month.providerBreakdown.length} logged</summary>
                <div class="ptoProviderList">
                  ${month.providerBreakdown.map((p, pi) => `
                    <div class="ptoProviderRow">
                      <input type="text" class="ptoProviderName" placeholder="Provider name"
                        data-entity="${entity}" data-monthkey="${month.monthKey}" data-pi="${pi}"
                        value="${p.name || ""}" />
                      <select class="ptoProviderType"
                        data-entity="${entity}" data-monthkey="${month.monthKey}" data-pi="${pi}">
                        <option value="MD" ${p.type === "MD" ? "selected" : ""}>MD</option>
                        <option value="PA" ${p.type === "PA" ? "selected" : ""}>PA</option>
                        <option value="PT" ${p.type === "PT" ? "selected" : ""}>PT</option>
                      </select>
                      <input type="number" class="ptoProviderDays" placeholder="Days"
                        data-entity="${entity}" data-monthkey="${month.monthKey}" data-pi="${pi}"
                        step="0.5" min="0" value="${p.days || ""}" />
                    </div>
                  `).join("")}
                  <button class="ptoAddProviderBtn" type="button"
                    data-entity="${entity}" data-monthkey="${month.monthKey}" data-idx="${i}">
                    + Add Provider
                  </button>
                </div>
              </details>` : `
              <details class="ptoProviderBreakdown">
                <summary class="ptoProviderBreakdownToggle">+ Add provider-level breakdown</summary>
                <div class="ptoProviderList" id="ptoProviderList_${entity}_${i}">
                  <button class="ptoAddProviderBtn" type="button"
                    data-entity="${entity}" data-monthkey="${month.monthKey}" data-idx="${i}">
                    + Add Provider
                  </button>
                </div>
              </details>`}

              ${(() => {
                const isAdmin = currentUser?.access?.isAdmin;
                const userEntity = currentUser?.access?.entity;
                const canEdit = isAdmin || !userEntity || userEntity === entity;
                return canEdit
                  ? `<button class="ptoSaveBtn" data-entity="${entity}" data-monthkey="${month.monthKey}" data-idx="${i}" type="button">Save</button>`
                  : `<div class="ptoReadonlyNotice">Read only — contact an admin to update ${entity}</div>`;
              })()}

              <div class="ptoForecastResult" id="ptoResult_${entity}_${i}">
                ${month.totalMissedVisits > 0 ? `
                  <div class="ptoImpactRow">
                    <span class="ptoImpactLabel">MD Impact</span>
                    <strong class="ptoImpactValue">${formatWhole(month.missedMdVisits)} visits</strong>
                  </div>
                  <div class="ptoImpactRow">
                    <span class="ptoImpactLabel">PA Impact</span>
                    <strong class="ptoImpactValue">${formatWhole(month.missedPaVisits)} visits</strong>
                  </div>
                  ${month.missedPtVisits > 0 ? `
                  <div class="ptoImpactRow">
                    <span class="ptoImpactLabel">PT Impact</span>
                    <strong class="ptoImpactValue">${formatWhole(month.missedPtVisits)} visits</strong>
                  </div>` : ""}
                  <div class="ptoImpactRow ptoImpactTotal">
                    <span class="ptoImpactLabel">Total Forecasted Impact</span>
                    <strong class="ptoImpactValue">${formatWhole(month.totalMissedVisits)} visits</strong>
                  </div>
                  ${month.savedBy ? `<div class="ptoSavedBy">Saved by ${escapeHtml(month.savedBy)}</div>` : ""}
                ` : `<div class="ptoNoData">Enter PTO days above to see forecast</div>`}
              </div>
            </div>
          `).join("")}
        </div>
      </div>
    `;
  }).join("");

  // Wire up mode toggle buttons (Monthly / By Week)
  document.querySelectorAll(".ptoModeBtn").forEach((btn) => {
    btn.addEventListener("click", () => {
      const entity = btn.dataset.entity;
      const monthKey = btn.dataset.monthkey;
      const mode = btn.dataset.mode;
      const isWeekly = mode === "weekly";

      setPtoWeeklyMode(entity, monthKey, isWeekly);

      // Find this month card's index and re-render the entity panel
      // Easiest: reload the whole PTO view
      loadPtoForecast();
    });
  });

  // Wire weekly inputs to update totals live and sync to monthly aggregate fields
  document.querySelectorAll(".ptoWeekInput").forEach((input) => {
    input.addEventListener("input", () => {
      const entity   = input.dataset.entity;
      const monthKey = input.dataset.monthkey;
      const date     = input.dataset.date;
      const type     = input.dataset.type;
      const val      = parseFloat(input.value) || 0;

      // Persist to localStorage
      setPtoWeekValue(entity, monthKey, date, type, val);

      // Update totals display
      const total = sumPtoWeekly(entity, monthKey, type);
      const totalEl = byId(`ptoWeekTotal_${type}_${entity}_${monthKey}`);
      if (totalEl) totalEl.textContent = total > 0 ? total : "—";

      // Sync to the aggregate monthly input (used by save button)
      const aggInput = document.querySelector(
        `.ptoInput[data-entity="${entity}"][data-monthkey="${monthKey}"][data-type="${type}"]`
      );
      if (aggInput) {
        aggInput.value = total > 0 ? total : "";
      } else {
        // Store in a hidden data attr on the card for save to pick up
        const card = byId(`ptoMonthCard_${entity}_${getMonthIdx(entity, monthKey)}`);
        if (card) card.dataset[`ptoTotal_${type}`] = total;
      }
    });
  });

  function getMonthIdx(entity, monthKey) {
    // Find idx by scanning ptoMonthCard IDs
    const cards = document.querySelectorAll(`[id^="ptoMonthCard_${entity}_"]`);
    for (const card of cards) {
      const parts = card.id.split("_");
      const idx = parts[parts.length - 1];
      // Check if month inputs match
      const md = document.querySelector(`.ptoInput[data-entity="${entity}"][data-monthkey="${monthKey}"]`);
      if (md) return idx;
    }
    return 0;
  }

  // Wire up add provider buttons
  document.querySelectorAll(".ptoAddProviderBtn").forEach((btn) => {
    btn.addEventListener("click", () => {
      const entity = btn.dataset.entity;
      const monthKey = btn.dataset.monthkey;
      const idx = btn.dataset.idx;
      const list = btn.closest(".ptoProviderList");
      if (!list) return;

      const row = document.createElement("div");
      row.className = "ptoProviderRow";
      row.innerHTML = `
        <input type="text" class="ptoProviderName" placeholder="Provider name"
          data-entity="${entity}" data-monthkey="${monthKey}" data-pi="new" />
        <select class="ptoProviderType" data-entity="${entity}" data-monthkey="${monthKey}" data-pi="new">
          <option value="MD">MD</option>
          <option value="PA">PA</option>
          <option value="PT">PT</option>
        </select>
        <input type="number" class="ptoProviderDays" placeholder="Days"
          data-entity="${entity}" data-monthkey="${monthKey}" data-pi="new" step="0.5" min="0" />
        <button type="button" class="ptoRemoveProviderBtn" title="Remove">✕</button>
      `;
      list.insertBefore(row, btn);
      row.querySelector(".ptoRemoveProviderBtn").addEventListener("click", () => row.remove());

      // Sync totals when days change
      row.querySelector(".ptoProviderDays").addEventListener("input", () => {
        syncPtoTotalsFromProviders(entity, monthKey);
      });
      row.querySelector(".ptoProviderType").addEventListener("change", () => {
        syncPtoTotalsFromProviders(entity, monthKey);
      });
    });
  });

  // Wire auto-sync for existing provider rows
  document.querySelectorAll(".ptoProviderDays, .ptoProviderType").forEach(input => {
    input.addEventListener("input", () => {
      const entity = input.dataset.entity;
      const monthKey = input.dataset.monthkey;
      syncPtoTotalsFromProviders(entity, monthKey);
    });
  });

  // Wire up save buttons
  document.querySelectorAll(".ptoSaveBtn").forEach((btn) => {
    btn.addEventListener("click", async () => {
      const entity = btn.getAttribute("data-entity");
      const monthKey = btn.getAttribute("data-monthkey");
      const idx = btn.getAttribute("data-idx");

      // In weekly mode, sum from localStorage; in monthly mode, read the input
      const isWeekly = getPtoWeeklyMode(entity, monthKey);

      let mdPtoDays, paPtoDays, ptPtoDays;
      if (isWeekly) {
        mdPtoDays = sumPtoWeekly(entity, monthKey, "md");
        paPtoDays = sumPtoWeekly(entity, monthKey, "pa");
        ptPtoDays = sumPtoWeekly(entity, monthKey, "pt");
      } else {
        const mdInput = document.querySelector(`.ptoInput[data-entity="${entity}"][data-monthkey="${monthKey}"][data-type="md"]`);
        const paInput = document.querySelector(`.ptoInput[data-entity="${entity}"][data-monthkey="${monthKey}"][data-type="pa"]`);
        const ptInput = document.querySelector(`.ptoInput[data-entity="${entity}"][data-monthkey="${monthKey}"][data-type="pt"]`);
        mdPtoDays = parseFloat(mdInput?.value || "0") || 0;
        paPtoDays = parseFloat(paInput?.value || "0") || 0;
        ptPtoDays = parseFloat(ptInput?.value || "0") || 0;
      }

      btn.textContent = "Saving...";
      btn.disabled = true;

      try {
        // Collect provider breakdown if present
        const providerBreakdown = collectProviderBreakdown(entity, monthKey);

        const result = await apiPost("/api/pto-forecast", {
          entity, monthKey, mdPtoDays, paPtoDays, ptPtoDays, providerBreakdown
        });

        const resultEl = byId(`ptoResult_${entity}_${idx}`);
        if (resultEl && result.forecast) {
          const f = result.forecast;
          resultEl.innerHTML = `
            <div class="ptoImpactRow">
              <span class="ptoImpactLabel">MD Impact</span>
              <strong class="ptoImpactValue">${formatWhole(f.missedMdVisits)} visits</strong>
            </div>
            <div class="ptoImpactRow">
              <span class="ptoImpactLabel">PA Impact</span>
              <strong class="ptoImpactValue">${formatWhole(f.missedPaVisits)} visits</strong>
            </div>
            ${f.missedPtVisits > 0 ? `
            <div class="ptoImpactRow">
              <span class="ptoImpactLabel">PT Impact</span>
              <strong class="ptoImpactValue">${formatWhole(f.missedPtVisits)} visits</strong>
            </div>` : ""}
            <div class="ptoImpactRow ptoImpactTotal">
              <span class="ptoImpactLabel">Total Forecasted Impact</span>
              <strong class="ptoImpactValue">${formatWhole(f.totalMissedVisits)} visits</strong>
            </div>
            <div class="ptoSavedBy">Saved · ${f.perProviderClinicalRate}/day per provider</div>
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

  const entities = data?.entities || [];
  const monthLabels = data?.monthLabels || [];

  if (!entities.length) {
    container.innerHTML = "";
    return;
  }

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
      Forecasted missed visits = (MD PTO Days + PA PTO Days) × per-provider clinical rate + (PT PTO Days × per-PT rate). Rates computed from last 12 weeks of actuals. Update provider counts in Provider Settings.
      Rates computed from last ${entities[0]?.rates?.weeksUsed || 12} weeks of actual data per entity.
    </div>
  `;
}

let _lastPtoData = null;

async function loadPtoForecast() {
  const bannerEl = byId("ptoForecastBanner");
  if (bannerEl) bannerEl.textContent = "Loading forecast data...";

  try {
    const data = await apiGet("/api/pto-forecast");
    _lastPtoData = data;

    if (!data || !data.ok) {
      if (bannerEl) bannerEl.textContent = `Error: ${data?.error || "Failed to load forecast data."}`;
      setPtoForecastDebug(data);
      return;
    }

    if (!Array.isArray(data.entities) || data.entities.length === 0) {
      const isAdmin = currentUser?.access?.isAdmin;
      const userEntity = currentUser?.access?.entity;
      if (bannerEl) {
        if (!isAdmin && userEntity && userEntity !== "admin") {
          bannerEl.textContent = `No PTO data found for ${userEntity} yet. Enter days in the form below once it loads, or check that weekly data has been imported first.`;
        } else {
          bannerEl.textContent = "No entity data returned. Check that weekly data has been imported.";
        }
      }
      setPtoForecastDebug(data);
      return;
    }

    if (bannerEl) {
      const labels = (data.monthLabels || []).join(", ");
      const isAdmin = currentUser?.access?.isAdmin;
      if (isAdmin) {
        bannerEl.textContent = `Rolling quarter: ${labels} · Enter projected PTO days per entity and month, then save.`;
        bannerEl.className = "subtleBanner";
      } else {
        const userEntity = currentUser?.access?.entity || "your entity";
        bannerEl.innerHTML = `Rolling quarter: ${labels} · You are entering PTO data for <strong>${userEntity}</strong>. Enter days below and hit Save.`;
        bannerEl.className = "subtleBanner";
      }
    }

    renderPtoForecastEntities(data);
    renderPtoForecastSummary(data);
    setPtoForecastDebug(data);
  } catch (e) {
    const msg = e?.message || String(e);
    if (bannerEl) bannerEl.textContent = `API Error: ${msg}`;
    setPtoForecastDebug({ error: msg, hint: "Check that /api/pto-forecast is deployed and the PTOForecastData table can be created in Azure." });
  }
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

function showProviderSettingsView() {
  hideAllViews();
  const el = byId("providerSettingsView");
  if (el) el.style.display = "";
  setActiveNav("navProviderSettingsBtn");
}

function setProviderSettingsDebug(data) {
  const el = byId("providerSettingsDebug");
  if (!el) return;
  el.textContent = typeof data === "string" ? data : JSON.stringify(data, null, 2);
}

async function loadProviderSettings() {
  const bannerEl = byId("providerSettingsBanner");
  if (bannerEl) bannerEl.textContent = "Loading provider settings...";

  try {
    const data = await apiGet("/api/provider-settings");
    if (!data?.ok || !Array.isArray(data.entities)) {
      if (bannerEl) bannerEl.textContent = `Error: ${data?.error || "Failed to load settings."}`;
      setProviderSettingsDebug(data);
      return;
    }

    if (bannerEl) bannerEl.textContent = "Provider counts are used by the PTO Forecast to calculate per-provider daily visit rates. Podiatrists count as MDs. PRN staff should not be included.";

    renderProviderSettingsTable(data.entities);
    setProviderSettingsDebug(data);
  } catch (e) {
    if (bannerEl) bannerEl.textContent = `Error: ${e?.message || String(e)}`;
    setProviderSettingsDebug(String(e));
  }
}

function renderProviderSettingsTable(entities) {
  const container = byId("providerSettingsTable");
  if (!container) return;

  container.innerHTML = `
    <table class="regionTable">
      <thead>
        <tr>
          <th>Entity</th>
          <th>MDs (incl. Pods)</th>
          <th>PAs</th>
          <th>PTs</th>
          <th>Total Clinical</th>
          <th>Last Updated</th>
          <th></th>
        </tr>
      </thead>
      <tbody>
        ${entities.map((e) => `
          <tr id="provRow_${e.entity}">
            <td><strong>${e.entity}</strong></td>
            <td>
              <input type="number" min="0" step="1"
                id="prov_md_${e.entity}"
                value="${e.mdCount}"
                style="width:70px;padding:6px 8px;font-family:var(--font-mono);font-size:13px;"
              />
            </td>
            <td>
              <input type="number" min="0" step="1"
                id="prov_pa_${e.entity}"
                value="${e.paCount}"
                style="width:70px;padding:6px 8px;font-family:var(--font-mono);font-size:13px;"
              />
            </td>
            <td>
              <input type="number" min="0" step="1"
                id="prov_pt_${e.entity}"
                value="${e.ptCount}"
                style="width:70px;padding:6px 8px;font-family:var(--font-mono);font-size:13px;"
              />
            </td>
            <td style="font-family:var(--font-mono);" id="prov_total_${e.entity}">
              ${e.mdCount + e.paCount + e.ptCount}
            </td>
            <td style="font-size:11.5px;color:var(--text-muted);">
              ${e.updatedAt ? escapeHtml(new Date(e.updatedAt).toLocaleDateString() + (e.updatedBy ? ` · ${e.updatedBy}` : "")) : "Default"}
            </td>
            <td>
              <button type="button"
                onclick="saveProviderSettings('${e.entity}')"
                style="padding:6px 12px;font-size:12px;box-shadow:none;"
              >Save</button>
            </td>
          </tr>
        `).join("")}
      </tbody>
    </table>
  `;

  // Live-update total column on input change
  entities.forEach((e) => {
    ["md", "pa", "pt"].forEach((type) => {
      const input = byId(`prov_${type}_${e.entity}`);
      if (input) {
        input.addEventListener("input", () => {
          const md = parseInt(byId(`prov_md_${e.entity}`)?.value || "0", 10) || 0;
          const pa = parseInt(byId(`prov_pa_${e.entity}`)?.value || "0", 10) || 0;
          const pt = parseInt(byId(`prov_pt_${e.entity}`)?.value || "0", 10) || 0;
          const totalEl = byId(`prov_total_${e.entity}`);
          if (totalEl) totalEl.textContent = md + pa + pt;
        });
      }
    });
  });
}

async function saveProviderSettings(entity) {
  const mdCount = parseInt(byId(`prov_md_${entity}`)?.value || "0", 10) || 0;
  const paCount = parseInt(byId(`prov_pa_${entity}`)?.value || "0", 10) || 0;
  const ptCount = parseInt(byId(`prov_pt_${entity}`)?.value || "0", 10) || 0;

  try {
    const result = await apiPost("/api/provider-settings", { entity, mdCount, paCount, ptCount });
    if (result?.ok) {
      const bannerEl = byId("providerSettingsBanner");
      if (bannerEl) {
        bannerEl.textContent = `✓ ${entity} settings saved — ${mdCount} MDs, ${paCount} PAs, ${ptCount} PTs`;
        setTimeout(() => {
          bannerEl.textContent = "Provider counts are used by the PTO Forecast to calculate per-provider daily visit rates.";
        }, 3000);
      }
    }
    setProviderSettingsDebug(result);
  } catch (e) {
    setProviderSettingsDebug(String(e));
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
    // The anchor week IS the selected data week; prior period is the week before it
    return {
      primaryWeeks: [anchorWeek],
      comparisonWeeks: [getPreviousWeekEnding(anchorWeek)],
      summary: `Viewing week ending ${anchorWeek}`
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

  if (periodType === "ytd") {
    const now = new Date();
    const yearStart = new Date(Date.UTC(now.getUTCFullYear(), 0, 1));
    const today = now.toISOString().slice(0, 10);
    const primaryWeeks = [];
    const walker = new Date(yearStart);
    while (walker.toISOString().slice(0, 10) <= today) {
      if (walker.getUTCDay() === 5) primaryWeeks.push(walker.toISOString().slice(0, 10));
      walker.setUTCDate(walker.getUTCDate() + 1);
    }
    const comparisonWeeks = primaryWeeks.map((w) => addDays(w, -364));
    const yearLabel = now.getUTCFullYear();
    return {
      primaryWeeks,
      comparisonWeeks,
      summary: `Year to Date ${yearLabel} (${primaryWeeks.length} weeks)`
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
  const hideAnchor = periodType === "mtd" || periodType === "lastMonth" || periodType === "ytd";

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
          imaging: 0,
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
          emEstBudget: 0,
          surgeryBudget: 0,
          ptVisitsBudget: 0,
          imagingBudget: 0,
          totalVisitsBudget: 0,
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
      row.imaging += normalizeNumber(region.imaging);
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
        row.emEstBudget += normalizeNumber(region.budget?.emEstBudget);
        row.surgeryBudget += normalizeNumber(region.budget?.surgeryBudget);
        row.ptVisitsBudget += normalizeNumber(region.budget?.ptVisitsBudget);
        row.imagingBudget += normalizeNumber(region.budget?.imagingBudget);
        row.totalVisitsBudget += normalizeNumber(region.budget?.totalVisitsBudget);
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
    imaging: row.imaging,
    cashCollected: row.cashCollected,
    ptoDays: row.ptoDays,
    piNp: row.piNp,
    piCashCollection: row.piCashCollection,
    noShowRate: row.weekCount ? row.noShowRateTotal / row.weekCount : 0,
    cancellationRate: row.weekCount ? row.cancellationRateTotal / row.weekCount : 0,
    abandonedCallRate: row.weekCount ? row.abandonedCallRateTotal / row.weekCount : 0,
    visitVolumeBudget: row.visitVolumeBudget,
    newPatientsBudget: row.newPatientsBudget,
    emEstBudget: row.emEstBudget,
    surgeryBudget: row.surgeryBudget,
    ptVisitsBudget: row.ptVisitsBudget,
    imagingBudget: row.imagingBudget,
    totalVisitsBudget: row.totalVisitsBudget,
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
    imaging: regions.reduce((sum, r) => sum + normalizeNumber(r.imaging), 0),
    cashCollected: regions.reduce((sum, r) => sum + normalizeNumber(r.cashCollected), 0),
    ptoDays: regions.reduce((sum, r) => sum + normalizeNumber(r.ptoDays), 0),
    piNp: regions.reduce((sum, r) => sum + normalizeNumber(r.piNp), 0),
    piCashCollection: regions.reduce((sum, r) => sum + normalizeNumber(r.piCashCollection), 0)
  };

  const sumReg = (k) => regions.reduce((sum, r) => sum + normalizeNumber(r[k]), 0);
  const budgetTotals = includeBudget
    ? {
        visitVolumeBudget: sumReg("visitVolumeBudget"),
        newPatientsBudget: sumReg("newPatientsBudget"),
        emEstBudget: sumReg("emEstBudget"),
        surgeryBudget: sumReg("surgeryBudget"),
        ptVisitsBudget: sumReg("ptVisitsBudget"),
        imagingBudget: sumReg("imagingBudget"),
        totalVisitsBudget: sumReg("totalVisitsBudget")
      }
    : {
        visitVolumeBudget: 0,
        newPatientsBudget: 0,
        emEstBudget: 0,
        surgeryBudget: 0,
        ptVisitsBudget: 0,
        imagingBudget: 0,
        totalVisitsBudget: 0
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

  // PT scopes use same data but render differently
  if (isPtScope(entityScope)) {
    await loadDashboardLandingPt(entityScope, compareAgainst, weekSets);
    return;
  }

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
    const scopeLabel = entityScope === "ALL" ? "All Entities"
      : entityScope === "ALL-PT" ? "All PT Programs"
      : entityScope;
    summaryEl.innerHTML = `<div style="font-size:13px; opacity:0.85;">${weekSets.summary} • Scope: ${scopeLabel}</div>`;
  }

  renderDashboardCards(current, comparison, compareAgainst, entityScope, weekSets);
  renderDataCompletenessBanner(current, entityScope, weekSets);
  renderDashboardEntities(current, comparison, compareAgainst, entityScope);
  renderDashboardOpsNotes(current, entityScope, weekSets);
  renderDashboardAlerts(current, comparison, entityScope, compareAgainst);
  renderDashboardWins(current, comparison, entityScope, compareAgainst);
  renderDashboardSnapshot(current, entityScope, compareAgainst);
  renderVisitsChart(weekSets.primaryWeeks, current, compareAgainst);

  // Live call activity per entity card — pulled from Landis (LAOSS/NES)
  // and RingCentral (MRO/SpineOne). Follows the dashboard's selected week
  // so the numbers change when the user picks a different period. Fires
  // in parallel after the main render.
  const livePhonesWeek = weekSets?.primaryWeeks?.[weekSets.primaryWeeks.length - 1] || getCurrentWeekEnding();
  const livePhonesEntities = entityScope === "ALL" ? ENTITIES : [entityScope];
  livePhonesEntities.forEach((e) => { populateEntityLivePhones(e, livePhonesWeek); });

  // Store for digest generator
  window._lastDashboardCurrent        = current;
  window._lastDashboardComparison     = comparison;
  window._lastDashboardCompareAgainst = compareAgainst;
  window._lastDashboardWeekSets       = weekSets;

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

  // Pause autosave while we overwrite the form with the server record so
  // we don't immediately write our loaded values back on top of any
  // newer local draft we're about to offer the user.
  _draftAutosaveSuspended = true;

  const result = await apiGet(
    `/api/weekly?weekEnding=${encodeURIComponent(weekEnding)}&entity=${encodeURIComponent(entity)}`
  );

  currentWeekData = result;

  // API may return data nested under result.values / result.data (manual entry path)
  // OR flat at the top level (Azure Table row returned directly).
  // A PT-only record has visitVolume == null; trust any flat record that says found.
  const formData = result.values || result.data ||
    (result.found ? result : {});
  setFormValues(formData);
  renderEntryAuditSummary(result);

  // If there's a local autosaved draft newer than the server record,
  // offer to restore it. Draft is cleared on successful save so this
  // only triggers when the user left the form mid-entry.
  offerDraftRestoreIfNewer(entity, weekEnding, result);
  _draftAutosaveSuspended = false;

  // Auto-populate call metrics from phone system if not already entered
  const existingValues = result.values || result.data || {};
  const hasCallData = normalizeNumber(existingValues.totalCalls || existingValues.callVolume) > 0;
  if (!hasCallData) {
    loadAndPopulateCallData(entity, weekEnding);
  } else {
    // Show call data badge regardless (so team can see it's synced)
    loadCallDataBadge(entity, weekEnding);
  }

  setStatus(`Loaded ${selectedLabel} for ${weekEnding} (${result.status || "saved"})`);
  setDebug({
    selectedEntry: selectedLabel,
    baseEntity: entity,
    result
  });
}

async function loadAndPopulateCallData(entity, weekEnding) {
  const banner = byId("callDataBanner");
  if (banner) { banner.textContent = "Checking phone system data..."; banner.style.display = ""; banner.className = "callDataBanner callDataBannerLoading"; }

  try {
    const callData = await apiGet(
      `/api/call-data?entity=${encodeURIComponent(entity)}&weekEnding=${encodeURIComponent(weekEnding)}`
    );

    if (!callData.hasData) {
      if (banner) { banner.textContent = "No phone system data found for this week yet — enter manually."; banner.className = "callDataBanner callDataBannerMissing"; }
      return;
    }

    // Auto-fill call fields
    const totalCallsEl   = byId("totalCalls");
    const abandonedEl    = byId("abandonedCalls");

    if (totalCallsEl && !totalCallsEl.value) {
      totalCallsEl.value = callData.totalCalls;
    }
    if (abandonedEl && !abandonedEl.value) {
      abandonedEl.value = callData.abandonedCalls;
    }

    // Show rich banner
    if (banner) {
      banner.className = "callDataBanner callDataBannerSynced";
      banner.innerHTML = `
        <span class="callDataBannerIcon">📞</span>
        <span class="callDataBannerText">
          <strong>Phone data synced from ${callData.source === "ringcentral" ? "RingCentral" : "Landis"}</strong>
          &nbsp;·&nbsp; ${formatWhole(callData.totalCalls)} calls
          &nbsp;·&nbsp; ${callData.abandonedCalls} abandoned (${callData.abandonedRate}%)
          &nbsp;·&nbsp; Avg wait ${callData.avgWaitFormatted}
          &nbsp;·&nbsp; Avg talk ${callData.avgTalkFormatted}
          &nbsp;·&nbsp; Answer rate ${callData.answerRate}%
        </span>
        <span class="callDataBannerSource">${callData.datesFound.length} of 5 days synced</span>
      `;
    }

    updateDerivedDisplays();
  } catch (e) {
    if (banner) { banner.textContent = "Could not load phone system data."; banner.className = "callDataBanner callDataBannerMissing"; }
  }
}

// Populate the live phones strip on a dashboard entity card with data pulled
// directly from Landis/RingCentral (via /api/call-data). Runs after the main
// dashboard render so the primary KPIs show immediately and the phones strip
// fills in progressively.
async function populateEntityLivePhones(entity, weekEnding) {
  const container = byId(`livePhones_${entity}`);
  if (!container) return;

  const stateEl = container.querySelector(".entityLivePhonesState");
  const gridEl  = container.querySelector(".entityLivePhonesGrid");

  try {
    const callData = await apiGet(
      `/api/call-data?entity=${encodeURIComponent(entity)}&weekEnding=${encodeURIComponent(weekEnding)}`
    );

    if (!callData?.ok || !callData.hasData) {
      if (stateEl) stateEl.textContent = callData?.source ? "no data this week" : "not connected";
      if (gridEl) gridEl.innerHTML = "";
      return;
    }

    const sourceLabel = callData.source === "ringcentral" ? "RingCentral" : callData.source === "landis" ? "Landis" : callData.source;
    const daysLabel = callData.datesFound?.length ? `${callData.datesFound.length}/5 days` : "";
    if (stateEl) stateEl.textContent = [sourceLabel, daysLabel].filter(Boolean).join(" · ");

    const abandonedBad = normalizeNumber(callData.abandonedRate) >= getThreshold("abandonedCallRate", 10);
    const waitBad = normalizeNumber(callData.avgWaitSeconds) >= 60;

    if (gridEl) {
      gridEl.innerHTML = `
        <div class="entityLivePhonesTile">
          <span class="entityLivePhonesTileLabel">Calls</span>
          <strong class="entityLivePhonesTileValue">${formatWhole(callData.totalCalls)}</strong>
        </div>
        <div class="entityLivePhonesTile">
          <span class="entityLivePhonesTileLabel">Answered</span>
          <strong class="entityLivePhonesTileValue">${formatWhole(callData.answeredCalls)}<span class="entityLivePhonesPct">&nbsp;${callData.answerRate}%</span></strong>
        </div>
        <div class="entityLivePhonesTile ${abandonedBad ? "entityLivePhonesTileBad" : ""}">
          <span class="entityLivePhonesTileLabel">Abandoned</span>
          <strong class="entityLivePhonesTileValue">${formatWhole(callData.abandonedCalls)}<span class="entityLivePhonesPct">&nbsp;${callData.abandonedRate}%</span></strong>
        </div>
        <div class="entityLivePhonesTile ${waitBad ? "entityLivePhonesTileBad" : ""}">
          <span class="entityLivePhonesTileLabel">Avg Wait</span>
          <strong class="entityLivePhonesTileValue">${escapeHtml(callData.avgWaitFormatted || "0:00")}</strong>
        </div>
      `;
    }
  } catch {
    if (stateEl) stateEl.textContent = "error";
    if (gridEl) gridEl.innerHTML = "";
  }
}

async function loadCallDataBadge(entity, weekEnding) {
  try {
    const callData = await apiGet(
      `/api/call-data?entity=${encodeURIComponent(entity)}&weekEnding=${encodeURIComponent(weekEnding)}`
    );
    const banner = byId("callDataBanner");
    if (banner && callData.hasData) {
      banner.style.display = "";
      banner.className = "callDataBanner callDataBannerSynced";
      banner.innerHTML = `
        <span class="callDataBannerIcon">📞</span>
        <span class="callDataBannerText">
          <strong>${callData.source === "ringcentral" ? "RingCentral" : "Landis"} data available</strong>
          &nbsp;·&nbsp; ${formatWhole(callData.totalCalls)} calls
          &nbsp;·&nbsp; ${callData.abandonedRate}% abandoned
          &nbsp;·&nbsp; ${callData.answerRate}% answer rate
          &nbsp;·&nbsp; Avg wait ${callData.avgWaitFormatted}
        </span>
      `;
    }
  } catch (e) { /* silent */ }
}

// ── Entry-form draft autosave + destructive-change confirmation ────────
//
// Two safety rails that would have caught the 4/17 PT zero-out before it
// hit the database:
//   1. Every change to the form is autosaved to localStorage every 2s so
//      closing the tab mid-entry doesn't lose work.
//   2. On save, we diff the form values against the stored record and
//      block the submit if any non-zero field is being set to zero unless
//      the user explicitly confirms.
//
const DRAFT_KEY = (entity, weekEnding) => `entryDraft_${entity}_${weekEnding}`;
const DRAFT_AUTOSAVE_MS = 2000;
let _draftAutosaveTimer = null;
let _draftAutosaveSuspended = false;

function scheduleDraftAutosave() {
  if (_draftAutosaveSuspended) return;
  clearTimeout(_draftAutosaveTimer);
  _draftAutosaveTimer = setTimeout(() => {
    persistEntryDraft();
    _draftAutosaveTimer = null;
  }, DRAFT_AUTOSAVE_MS);
}

function persistEntryDraft() {
  try {
    const entity = getSelectedEntity();
    const weekEnding = byId("weekEnding")?.value;
    if (!entity || !weekEnding) return;
    const values = getFormValues();
    localStorage.setItem(DRAFT_KEY(entity, weekEnding), JSON.stringify({
      values,
      savedAt: new Date().toISOString()
    }));
    const indicator = byId("draftAutosaveIndicator");
    if (indicator) {
      indicator.textContent = `Draft autosaved ${new Date().toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" })}`;
      indicator.classList.remove("draftIndicatorError");
    }
  } catch (e) {
    const indicator = byId("draftAutosaveIndicator");
    if (indicator) {
      indicator.textContent = "Draft autosave failed (storage full?)";
      indicator.classList.add("draftIndicatorError");
    }
  }
}

function loadEntryDraft(entity, weekEnding) {
  try {
    const raw = localStorage.getItem(DRAFT_KEY(entity, weekEnding));
    return raw ? JSON.parse(raw) : null;
  } catch { return null; }
}

function clearEntryDraft(entity, weekEnding) {
  try { localStorage.removeItem(DRAFT_KEY(entity, weekEnding)); } catch {}
}

// After loading a server record, check for a local autosaved draft that
// post-dates the server. If one exists, render an inline banner offering
// to restore or discard — don't auto-apply in case the draft is stale.
function offerDraftRestoreIfNewer(entity, weekEnding, serverRecord) {
  const existing = byId("draftRestoreBanner");
  if (existing) existing.remove();
  const draft = loadEntryDraft(entity, weekEnding);
  if (!draft?.savedAt) return;

  const serverUpdatedAt = serverRecord?.updatedAt || serverRecord?.createdAt;
  if (serverUpdatedAt && new Date(draft.savedAt) <= new Date(serverUpdatedAt)) {
    // Server record is newer than the draft — draft is stale, drop it.
    clearEntryDraft(entity, weekEnding);
    return;
  }

  const container = byId("callDataBanner")?.parentElement || byId("kpiForm")?.parentElement;
  if (!container) return;

  const banner = document.createElement("div");
  banner.id = "draftRestoreBanner";
  banner.className = "draftRestoreBanner";
  banner.innerHTML = `
    <span class="draftRestoreIcon">✎</span>
    <span class="draftRestoreText">
      You have unsaved changes for this week (autosaved ${escapeHtml(new Date(draft.savedAt).toLocaleString())}).
    </span>
    <button type="button" class="draftRestoreBtn" id="draftRestoreApplyBtn">Restore</button>
    <button type="button" class="draftDiscardBtn" id="draftRestoreDiscardBtn">Discard</button>
  `;
  const after = byId("callDataBanner");
  if (after) after.insertAdjacentElement("afterend", banner);
  else container.insertBefore(banner, container.firstChild);

  byId("draftRestoreApplyBtn").onclick = () => {
    _draftAutosaveSuspended = true;
    setFormValues(draft.values);
    _draftAutosaveSuspended = false;
    banner.remove();
    setStatus("Draft restored — review and save to commit.");
  };
  byId("draftRestoreDiscardBtn").onclick = () => {
    clearEntryDraft(entity, weekEnding);
    banner.remove();
  };
}

// Compare stored record against the new values we're about to save.
// Returns a list of changes; each flagged destructive if it zeros out a
// previously non-zero numeric field or clears a non-empty narrative.
function computeEntrySaveDiff(existing, nextValues) {
  const changes = [];
  for (const f of ENTRY_STORED_FIELDS) {
    const prevRaw = readStoredValue(existing, f);
    const prev = normalizeNumber(prevRaw);
    const next = normalizeNumber(nextValues[f.key]);
    if (prev === next) continue;
    const destructive = prev > 0 && next === 0;
    changes.push({ key: f.key, label: f.label, prev, next, destructive });
  }
  const prevNotes = String(existing?.operationsNarrative || "").trim();
  const nextNotes = String(nextValues?.operationsNarrative || "").trim();
  if (prevNotes !== nextNotes) {
    const destructive = prevNotes.length > 0 && nextNotes.length === 0;
    changes.push({
      key: "operationsNarrative", label: "Operator Notes",
      prev: prevNotes, next: nextNotes, destructive, isText: true
    });
  }
  return changes;
}

function confirmDestructiveSave(changes) {
  const destructive = changes.filter((c) => c.destructive);
  if (destructive.length === 0) return Promise.resolve(true);

  return new Promise((resolve) => {
    const existing = byId("diffConfirmModal");
    if (existing) existing.remove();

    const modal = document.createElement("div");
    modal.id = "diffConfirmModal";
    modal.className = "diffConfirmOverlay";
    modal.innerHTML = `
      <div class="diffConfirmBox">
        <div class="diffConfirmHeader">
          <span class="diffConfirmTag">Confirm save</span>
          <h2>${destructive.length} non-zero field${destructive.length === 1 ? "" : "s"} will be cleared</h2>
          <p>Review the changes below. If anything looks wrong, cancel and re-enter — saving will overwrite the existing record.</p>
        </div>
        <div class="diffConfirmBody">
          <table class="diffTable">
            <thead>
              <tr><th>Field</th><th>Before</th><th></th><th>After</th></tr>
            </thead>
            <tbody>
              ${destructive.map((c) => `
                <tr class="diffRowDestructive">
                  <td><strong>${escapeHtml(c.label)}</strong></td>
                  <td class="diffFrom">${escapeHtml(c.isText ? (c.prev.length > 60 ? c.prev.slice(0, 60) + "…" : c.prev) : String(c.prev))}</td>
                  <td class="diffArrow">→</td>
                  <td class="diffTo">${c.isText ? "(cleared)" : "0"}</td>
                </tr>
              `).join("")}
            </tbody>
          </table>
        </div>
        <div class="diffConfirmFooter">
          <button type="button" class="diffConfirmCancel" id="diffConfirmCancelBtn">Cancel — let me review</button>
          <button type="button" class="diffConfirmOk" id="diffConfirmOkBtn">Save anyway</button>
        </div>
      </div>
    `;
    document.body.appendChild(modal);
    modal.style.display = "flex";

    const cleanup = (answer) => {
      modal.remove();
      document.body.style.overflow = "";
      resolve(answer);
    };
    byId("diffConfirmCancelBtn").onclick = () => cleanup(false);
    byId("diffConfirmOkBtn").onclick = () => cleanup(true);
    modal.onclick = (e) => { if (e.target === modal) cleanup(false); };
    document.body.style.overflow = "hidden";
  });
}

async function saveWeek() {
  const weekEnding = byId("weekEnding")?.value || "";
  const entity = getSelectedEntity();
  const nextValues = getFormValues();

  // Diff against the record currently loaded in the form and require
  // explicit confirmation for any destructive changes.
  // PT-only records have visitVolume == null; fall back to any flat record
  // flagged found so the diff captures PT-side changes too.
  const existing = currentWeekData?.values || currentWeekData?.data ||
    (currentWeekData?.found ? currentWeekData : null);
  if (existing) {
    const diff = computeEntrySaveDiff(existing, nextValues);
    const ok = await confirmDestructiveSave(diff);
    if (!ok) {
      setStatus("Save cancelled — no changes written.");
      return;
    }
  }

  const payload = { weekEnding, entity, data: nextValues };

  setStatus("Saving...");
  setDebug({
    selectedEntry: getSelectedEntryLabel(),
    baseEntity: entity,
    payload
  });

  const result = await apiPost("/api/weekly-save", payload);

  setStatus(result.message || "Saved successfully");
  setDebug({
    selectedEntry: getSelectedEntryLabel(),
    baseEntity: entity,
    result
  });

  // Save succeeded — clear the autosaved draft so we don't nag the user on reload.
  clearEntryDraft(entity, weekEnding);

  await loadWeek();

  // Check if all entities now have data — if so prompt digest
  checkAndPromptDigest(weekEnding);
}

async function checkAndPromptDigest(weekEnding) {
  if (!weekEnding) return;
  try {
    // Reload summary for all entities for this week
    const summaries = await Promise.all(
      ENTITIES.map(e => apiGet(`/api/weekly?weekEnding=${encodeURIComponent(weekEnding)}&entity=${encodeURIComponent(e)}`))
    );
    const allEntered = summaries.every(s => {
      const vv = s.values?.visitVolume ?? s.data?.visitVolume ?? (s.found ? s.visitVolume : null);
      return s.found && normalizeNumber(vv) > 0;
    });
    if (allEntered) {
      showAllDataEnteredPrompt(weekEnding);
    }
  } catch (e) {
    // Silent — don't disrupt the save flow
  }
}

function showAllDataEnteredPrompt(weekEnding) {
  const existing = byId("allDataPrompt");
  if (existing) return; // Already showing

  const weekLabel = new Date(weekEnding + "T12:00:00Z").toLocaleDateString("en-US", {
    weekday: "long", month: "long", day: "numeric"
  });

  const toast = document.createElement("div");
  toast.id = "allDataPrompt";
  toast.className = "allDataToast";
  toast.innerHTML = `
    <div class="allDataToastIcon">✓</div>
    <div class="allDataToastContent">
      <div class="allDataToastTitle">All entities entered for ${weekLabel}</div>
      <div class="allDataToastSub">Ready to send the weekly digest.</div>
    </div>
    <div class="allDataToastActions">
      <button type="button" class="allDataToastBtn" id="allDataDigestBtn">Generate Digest</button>
      <button type="button" class="allDataToastDismiss" id="allDataDismissBtn">✕</button>
    </div>
  `;
  document.body.appendChild(toast);

  // Animate in
  requestAnimationFrame(() => toast.classList.add("allDataToastVisible"));

  byId("allDataDismissBtn").onclick = () => {
    toast.classList.remove("allDataToastVisible");
    setTimeout(() => toast.remove(), 300);
  };

  byId("allDataDigestBtn").onclick = () => {
    toast.classList.remove("allDataToastVisible");
    setTimeout(() => toast.remove(), 300);
    // Load dashboard data for this week then show digest
    loadDashboardDataForWeeks([weekEnding], "ALL", { includeBudget: true }).then(current => {
      const prev = addDays(weekEnding, -7);
      loadDashboardDataForWeeks([prev], "ALL", {}).then(comparison => {
        showDigestModal(current, comparison, "priorPeriod", { primaryWeeks: [weekEnding] });
      });
    });
  };

  // Auto-dismiss after 12 seconds
  setTimeout(() => {
    if (byId("allDataPrompt")) {
      toast.classList.remove("allDataToastVisible");
      setTimeout(() => toast.remove(), 300);
    }
  }, 12000);
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


// ── SpineOne PI Trend Chart ───────────────────────────────────
function renderSpineOnePiChart(items) {
  const ctx = byId("trendsPiChart");
  if (!ctx) return;

  const labels = [...items].reverse().map(i => i.weekEnding);
  const piNpData = [...items].reverse().map(i => normalizeNumber(i.piNp));
  const piCashData = [...items].reverse().map(i => normalizeNumber(i.piCashCollection));

  if (window.piChartInstance) {
    window.piChartInstance.destroy();
  }

  window.piChartInstance = new Chart(ctx, {
    type: "bar",
    data: {
      labels,
      datasets: [
        {
          label: "PI New Patients",
          data: piNpData,
          backgroundColor: "rgba(139,92,246,0.5)",
          borderColor: "#8B5CF6",
          borderWidth: 1,
          borderRadius: 3,
          yAxisID: "yNp"
        },
        {
          label: "PI Cash ($)",
          data: piCashData,
          type: "line",
          borderColor: "#1E6B4F",
          backgroundColor: "rgba(30,107,79,0.05)",
          borderWidth: 1.8,
          pointRadius: 2,
          tension: 0.3,
          fill: true,
          yAxisID: "yCash"
        }
      ]
    },
    options: {
      responsive: true,
      interaction: { mode: "index", intersect: false },
      plugins: {
        legend: { labels: { color: "#6B7280", font: { size: 11, family: "Geist, sans-serif" } } },
        tooltip: {
          callbacks: {
            label: (ctx) => {
              if (ctx.dataset.yAxisID === "yCash") return ` PI Cash: ${formatCurrency(ctx.raw)}`;
              return ` PI NP: ${formatWhole(ctx.raw)}`;
            }
          }
        }
      },
      scales: {
        x: { ticks: { color: "#6B7280", font: { size: 10, family: "Geist Mono, monospace" } }, grid: { color: "#ECEAE3", drawBorder: false } },
        yNp:   { position: "left",  ticks: { color: "#8B5CF6", font: { size: 10, family: "Geist Mono, monospace" } }, grid: { color: "#ECEAE3", drawBorder: false }, title: { display: true, text: "PI NP", color: "#8B5CF6", font: { size: 10 } } },
        yCash: { position: "right", ticks: { color: "#1E6B4F", font: { size: 10, family: "Geist Mono, monospace" }, callback: v => formatCurrency(v) }, grid: { display: false }, title: { display: true, text: "PI Cash", color: "#1E6B4F", font: { size: 10 } } }
      }
    }
  });
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
      padding: 10px 12px;
      box-sizing: border-box;
      border: 1px solid var(--border-mid);
      border-radius: var(--radius-sm);
      background: var(--bg-panel);
      color: var(--text);
      resize: vertical;
      line-height: 1.5;
      font-family: var(--font);
      font-size: 13px;
      transition: border-color var(--transition-fast), box-shadow var(--transition-fast);
    }
    textarea:focus { outline: none; border-color: var(--accent); box-shadow: 0 0 0 3px var(--accent-soft); }

    .narrativeFieldBlock { margin-top: 10px; }

    .notesShell {
      padding: 20px;
      border-radius: var(--radius-lg);
      background: var(--bg-panel);
      border: 1px solid var(--border-soft);
      box-shadow: var(--shadow-sm);
    }

    .notesHeader { margin-bottom: 16px; }

    .notesEyebrow {
      font-size: 10.5px;
      text-transform: uppercase;
      letter-spacing: 0.12em;
      color: var(--accent);
      font-weight: 600;
      margin-bottom: 6px;
    }

    .notesTitle {
      font-size: 18px;
      font-weight: 600;
      line-height: 1.15;
      margin-bottom: 6px;
      color: var(--text);
      letter-spacing: -0.01em;
    }

    .notesSubtext {
      font-size: 13px;
      color: var(--text-soft);
      line-height: 1.55;
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
      padding: 12px 14px;
      border-radius: var(--radius-sm);
      background: var(--bg-panel-soft);
      border: 1px solid var(--border-soft);
    }

    .notesPromptBullet {
      width: 8px;
      height: 8px;
      border-radius: 999px;
      background: var(--accent);
      margin-top: 6px;
      flex-shrink: 0;
      box-shadow: 0 0 0 3px var(--accent-soft);
    }

    .notesPromptText {
      color: var(--text-soft);
      font-size: 12.5px;
      line-height: 1.5;
      font-weight: 500;
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
      background: var(--bg-panel);
      border: 1px solid var(--border-soft);
      border-radius: var(--radius-lg);
      padding: 14px 16px;
      box-shadow: var(--shadow-sm);
      transition: box-shadow .18s ease, border-color .15s ease;
    }

    .entityCard:hover {
      box-shadow: var(--shadow-md);
      border-color: var(--border-mid);
    }

    .entityCardHeader {
      display:flex;
      justify-content:space-between;
      gap:8px;
      align-items:flex-start;
      margin-bottom:10px;
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
      background: var(--good-soft);
      color: var(--good);
      font-family: var(--font-mono);
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
      background: var(--bg-panel-soft);
      color: var(--text-muted);
      font-family: var(--font-mono);
      font-size:10px;
      font-weight:500;
      letter-spacing:.04em;
    }

    .metricChipGood    { background: var(--good-soft);    color: var(--good); }
    .metricChipWarning { background: var(--warning-soft); color: var(--warning); }
    .metricChipBad     { background: var(--bad-soft);     color: var(--bad); }

    .entityTitle    { font-size:16px; font-weight:600; line-height:1.1; margin-bottom:3px; letter-spacing:-0.01em; color: var(--text); }
    .entitySubtitle { font-size:11px; color: var(--text-muted); line-height:1.35; }

    .entityLogoWrap {
      width:64px; height:34px; background:#fff; border-radius:6px; padding:4px;
      display:flex; align-items:center; justify-content:center; flex-shrink:0;
      border: 1px solid var(--border-soft);
    }

    .entityLogo { max-width:100%; max-height:100%; object-fit:contain; }

    .entityTopMetrics {
      display:grid;
      grid-template-columns:repeat(3,1fr);
      gap: 1px;
      margin-bottom:10px;
      border:1px solid var(--border-soft);
      border-radius: var(--radius-sm);
      overflow:hidden;
      background: var(--border-soft);
    }

    .entityMetricHero {
      padding:9px 10px;
      min-width:0;
      background: var(--bg-panel);
    }

    .entityMetricLabel {
      display:block;
      font-size:9px;
      text-transform:uppercase;
      letter-spacing:.1em;
      color: var(--text-muted);
      margin-bottom:3px;
      font-weight:600;
    }

    .entityMetricHero strong {
      font-family: var(--font);
      font-variant-numeric: tabular-nums;
      font-size:16px;
      font-weight:600;
      line-height:1.1;
      letter-spacing:-.01em;
      display:block;
      min-width:0;
      color: var(--text);
    }

    .entityBudgetGrid, .entityCompareGrid { display:grid; gap:8px; margin-bottom:12px; }
    .entityBudgetGrid  { grid-template-columns:repeat(2,1fr); }
    .entityCompareGrid { grid-template-columns:repeat(3,1fr); }

    .entityBudgetTile, .entityMiniStat {
      background: var(--bg-panel-soft);
      border:1px solid var(--border-soft);
      border-radius: var(--radius-sm);
      padding:10px 12px;
    }

    .entityBudgetLabel, .entityMiniLabel {
      display:block;
      font-size:9.5px;
      text-transform:uppercase;
      letter-spacing:.1em;
      color: var(--text-muted);
      margin-bottom:5px;
      font-weight:600;
    }

    .entityBudgetValue, .entityMiniStat strong {
      font-family: var(--font);
      font-variant-numeric: tabular-nums;
      font-size:18px; font-weight:600; line-height:1.1; margin-bottom:6px; display:block;
      letter-spacing:-.01em;
      color: var(--text);
    }

    .entityBudgetMeta { margin-top:5px; font-size:11.5px; color: var(--text-muted); }
    .entityProgressWrap { margin-top:8px; }

    .entityProgressLabelRow {
      display:flex; justify-content:space-between; gap:8px; align-items:center;
      font-size:11.5px; color: var(--text-muted); margin-bottom:5px;
    }

    .entityProgressLabelRow strong { color: var(--text-soft); font-size:11.5px; font-family: var(--font); font-variant-numeric: tabular-nums; font-weight: 600; }

    .entityProgressTrack {
      width:100%; height:6px; border-radius:999px;
      background: var(--border-soft); overflow:hidden; position:relative;
    }

    .entityProgressBar {
      height:100%; border-radius:999px; transition:width .32s ease;
      background: var(--accent);
    }

    .entityProgressGood    { background: var(--good); }
    .entityProgressWarning { background: var(--warning); }
    .entityProgressBad     { background: var(--bad); }

    /* Health stats — three inline pills in a row */
    .entityHealthRow {
      display:flex;
      gap: 1px;
      margin-bottom:8px;
      border:1px solid var(--border-soft);
      border-radius: var(--radius-sm);
      overflow:hidden;
      background: var(--border-soft);
    }

    .entityHealthItem {
      flex:1;
      display:flex;
      flex-direction:column;
      gap:2px;
      padding:7px 9px;
      background: var(--bg-panel);
      min-width:0;
    }

    .entityHealthLabel {
      font-size:9px;
      text-transform:uppercase;
      letter-spacing:.1em;
      color: var(--text-muted);
      font-weight:600;
      white-space:nowrap;
    }

    .entityHealthItem strong {
      font-family: var(--font);
      font-variant-numeric: tabular-nums;
      font-size:13px;
      font-weight:600;
      line-height:1.1;
      color: var(--text);
    }

    .entityHealthItem strong.good    { color: var(--good); }
    .entityHealthItem strong.warning { color: var(--warning); }
    .entityHealthItem strong.bad     { color: var(--bad); }

    /* Access health bar */
    .entityAccessPanel {
      margin-bottom:8px;
      padding:8px 10px;
      border-radius: var(--radius-sm);
      background: var(--bg-panel-soft);
      border:1px solid var(--border-soft);
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
      color: var(--text-muted);
      font-weight:600;
    }

    .entityAccessValue {
      font-family: var(--font);
      font-variant-numeric: tabular-nums;
      font-size:13px;
      font-weight:600;
      color: var(--text);
    }

    .entityDetailDrawer {
      border:1px solid var(--border-soft);
      border-radius: var(--radius-sm);
      background: var(--bg-panel);
      overflow:hidden;
    }

    .entityDetailDrawer summary {
      cursor:pointer; list-style:none; padding:10px 13px;
      font-weight:600; font-size:12.5px; color: var(--text-soft);
      display:flex; align-items:center; gap:6px;
      background: var(--bg-panel-soft);
    }

    .entityDetailDrawer summary::before {
      content:"›";
      font-size:14px;
      transition:transform .15s ease;
      display:inline-block;
      color: var(--text-muted);
    }

    details[open] .entityDetailDrawer summary::before { transform:rotate(90deg); }
    .entityDetailDrawer summary::-webkit-details-marker { display:none; }

    .entityDetailGrid {
      display:grid; grid-template-columns:repeat(3,1fr); gap:8px; padding: 10px 13px 13px;
    }

    .entityDetailTile {
      padding:10px; border-radius: var(--radius-sm);
      background: var(--bg-panel-soft); border:1px solid var(--border-soft);
    }

    .entityDetailLabel {
      font-size:9.5px; text-transform:uppercase; letter-spacing:.08em;
      color: var(--text-muted); margin-bottom:4px; font-weight:600;
    }

    .entityDetailValue {
      font-family: var(--font);
      font-variant-numeric: tabular-nums;
      font-size:16px; font-weight:600; line-height:1.1;
      letter-spacing:-.01em; color: var(--text);
    }

    /* Entry-page brand/entity context card */
    .entryBrandCard {
      display: flex;
      align-items: center;
      gap: 14px;
      padding: 12px 16px;
      border: 1px solid var(--border-soft);
      border-radius: var(--radius-lg);
      background: var(--bg-panel);
      box-shadow: var(--shadow-sm);
    }
    .entryBrandLogo {
      width: 108px; height: 48px; background: var(--bg-panel-soft);
      border: 1px solid var(--border-soft); border-radius: var(--radius-sm);
      padding: 6px; display: flex; align-items: center; justify-content: center; flex-shrink: 0;
    }
    .entryBrandLogo img { max-width: 100%; max-height: 100%; object-fit: contain; }
    .entryBrandLabel { font-weight: 600; font-size: 13px; letter-spacing: 0.02em; color: var(--accent); text-transform: uppercase; }
    .entryBrandFullName { font-size: 12.5px; color: var(--text-soft); margin-top: 2px; }

    @media (max-width: 820px) {
      .entityTopMetrics, .entityCompareGrid, .entityDetailGrid { grid-template-columns:1fr; }
      .entityBudgetGrid { grid-template-columns:1fr; }
    }

    @media (max-width: 768px) {
      .notesTitle { font-size: 16px; }
      .notesPromptGrid { grid-template-columns: 1fr; }
    }
  `;
  document.head.appendChild(style);
}

// ══════════════════════════════════════════════════════════════
// PT SCOPE DASHBOARD
// ══════════════════════════════════════════════════════════════

async function loadDashboardLandingPt(entityScope, compareAgainst, weekSets) {
  const entities = ptScopeEntities(entityScope);

  const current    = await loadDashboardDataForWeeks(weekSets.primaryWeeks, "ALL", {});
  const comparison = await loadDashboardDataForWeeks(weekSets.comparisonWeeks, "ALL", {});

  // Filter regions to PT entities only AND recompute ptTotals from filtered set
  const filterPt = (data) => {
    const regions = (data.regions || []).filter(r => entities.includes(r.entity));
    const ptTotals = {
      scheduledVisits:  regions.reduce((s, r) => s + normalizeNumber(r.pt?.scheduledVisits), 0),
      cancellations:    regions.reduce((s, r) => s + normalizeNumber(r.pt?.cancellations), 0),
      noShows:          regions.reduce((s, r) => s + normalizeNumber(r.pt?.noShows), 0),
      reschedules:      regions.reduce((s, r) => s + normalizeNumber(r.pt?.reschedules), 0),
      totalUnitsBilled: regions.reduce((s, r) => s + normalizeNumber(r.pt?.totalUnitsBilled), 0),
      visitsSeen:       regions.reduce((s, r) => s + normalizeNumber(r.pt?.visitsSeen), 0)
    };
    const totals = {
      visitVolume:   regions.reduce((s, r) => s + normalizeNumber(r.visitVolume), 0),
      newPatients:   regions.reduce((s, r) => s + normalizeNumber(r.newPatients), 0),
      cashCollected: regions.reduce((s, r) => s + normalizeNumber(r.cashCollected), 0),
      callVolume:    regions.reduce((s, r) => s + normalizeNumber(r.callVolume), 0),
      surgeries:     regions.reduce((s, r) => s + normalizeNumber(r.surgeries), 0),
      ptoDays:       regions.reduce((s, r) => s + normalizeNumber(r.ptoDays), 0)
    };
    return { ...data, regions, ptTotals, totals };
  };

  const currentPt    = filterPt(current);
  const comparisonPt = filterPt(comparison);

  // Clear stale ortho cards before rendering PT view
  const dashCards = byId("dashboardCards");
  if (dashCards) dashCards.innerHTML = "";
  const dashEntities = byId("dashboardEntities");
  if (dashEntities) dashEntities.innerHTML = "";

  const summaryEl = byId("dashboardSummaryText");
  if (summaryEl) {
    const label = entityScope === "ALL-PT" ? "All PT Programs" : `${entityScope} Program`;
    summaryEl.innerHTML = `<div style="font-size:13px;opacity:0.85;">${weekSets.summary} &nbsp;·&nbsp; Scope: ${label} (PT only)</div>`;
  }

  renderDashboardCardsPt(currentPt, comparisonPt, entities, compareAgainst);
  renderDataCompletenessBanner(currentPt, entityScope, weekSets);
  renderDashboardEntitiesPt(currentPt, comparisonPt, entities, compareAgainst);
  renderDashboardAlerts(currentPt, comparisonPt, entityScope, compareAgainst);
  renderDashboardWins(currentPt, comparisonPt, entityScope, compareAgainst);
  renderDashboardSnapshot(currentPt, entityScope, compareAgainst);
  renderVisitsChart(weekSets.primaryWeeks, currentPt, compareAgainst);

  window._lastDashboardCurrent        = currentPt;
  window._lastDashboardComparison     = comparisonPt;
  window._lastDashboardCompareAgainst = compareAgainst;
  window._lastDashboardWeekSets       = weekSets;

  setDashboardDebug({ entityScope, compareAgainst, weekSets, current: currentPt, comparison: comparisonPt });
}

function renderDashboardCardsPt(current, comparison, entities, compareAgainst) {
  // PT-focused top-line cards
  const ptCurrent    = current.ptTotals    || {};
  const ptPrior      = comparison.ptTotals || {};

  const visitsSeen      = normalizeNumber(ptCurrent.visitsSeen);
  const visitsSeenPrior = normalizeNumber(ptPrior.visitsSeen);
  const visitsDelta     = visitsSeen - visitsSeenPrior;

  const scheduled      = normalizeNumber(ptCurrent.scheduledVisits);
  const scheduledPrior = normalizeNumber(ptPrior.scheduledVisits);

  const cancellations  = normalizeNumber(ptCurrent.cancellations);
  const noShows        = normalizeNumber(ptCurrent.noShows);
  const reschedules    = normalizeNumber(ptCurrent.reschedules);
  const unitsBilled    = normalizeNumber(ptCurrent.totalUnitsBilled);
  const unitsBilledPrior = normalizeNumber(ptPrior.totalUnitsBilled);

  const cancelRate  = scheduled > 0 ? Number(((cancellations / scheduled) * 100).toFixed(1)) : 0;
  const noShowRate  = scheduled > 0 ? Number(((noShows / scheduled) * 100).toFixed(1)) : 0;
  const utilRate    = scheduled > 0 ? Number(((visitsSeen / scheduled) * 100).toFixed(1)) : 0;

  const unitsPerVisit = visitsSeen > 0
    ? Number((unitsBilled / visitsSeen).toFixed(1))
    : 0;
  const unitsPerVisitPrior = visitsSeenPrior > 0
    ? Number((unitsBilledPrior / visitsSeenPrior).toFixed(1))
    : 0;

  const cards = [
    {
      label: "PT Visits Seen",
      value: formatWhole(visitsSeen),
      movement: buildKpiMovement(visitsSeen, visitsSeenPrior, formatWhole),
      meta: `${visitsDelta >= 0 ? "+" : ""}${formatWhole(visitsDelta)} vs prior`,
      className: getTrendClass(visitsSeen, visitsSeenPrior)
    },
    {
      label: "Scheduled Visits",
      value: formatWhole(scheduled),
      movement: buildKpiMovement(scheduled, scheduledPrior, formatWhole),
      meta: `${(scheduled - scheduledPrior) >= 0 ? "+" : ""}${formatWhole(scheduled - scheduledPrior)} vs prior`,
      className: getTrendClass(scheduled, scheduledPrior)
    },
    {
      label: "Utilization Rate",
      value: `${utilRate}%`,
      meta: `${visitsSeen} seen of ${scheduled} scheduled`,
      className: utilRate >= 85 ? "kpi-positive" : utilRate >= 70 ? "kpi-neutral" : "kpi-negative"
    },
    {
      label: "Units Billed",
      value: formatWhole(unitsBilled),
      movement: buildKpiMovement(unitsBilled, unitsBilledPrior, formatWhole),
      meta: `${unitsPerVisit} units/visit avg`,
      className: getTrendClass(unitsBilled, unitsBilledPrior)
    },
    {
      label: "Cancellations",
      value: formatWhole(cancellations),
      meta: `${cancelRate}% cancel rate`,
      className: cancelRate >= 10 ? "kpi-negative" : cancelRate >= 6 ? "kpi-neutral" : "kpi-positive"
    },
    {
      label: "No Shows",
      value: formatWhole(noShows),
      meta: `${noShowRate}% no-show rate`,
      className: noShowRate >= 8 ? "kpi-negative" : noShowRate >= 5 ? "kpi-neutral" : "kpi-positive"
    },
    {
      label: "Reschedules",
      value: formatWhole(reschedules),
      meta: "Across selected scope",
      className: "kpi-neutral"
    },
    {
      label: "Units / Visit",
      value: unitsPerVisit.toFixed(1),
      movement: buildKpiMovement(unitsPerVisit, unitsPerVisitPrior, (v) => v.toFixed(1)),
      meta: `Prior: ${unitsPerVisitPrior.toFixed(1)}`,
      className: getTrendClass(unitsPerVisit, unitsPerVisitPrior)
    }
  ];

  renderMetricCards("dashboardCards", cards);
}

function renderDashboardEntitiesPt(current, comparison, entities, compareAgainst) {
  const container = byId("dashboardEntities");
  if (!container) return;

  const currentMap    = getEntityMap(current);
  const comparisonMap = getEntityMap(comparison);

  container.innerHTML = `
    <div class="entityCardGrid">
      ${entities.map(entity => {
        const brand = getBranding(entity);
        const row   = currentMap[entity];
        const prior = comparisonMap[entity] || {};
        if (!row) return `
          <div class="entityCard" style="border-top:4px solid ${brand.accent};">
            <div class="entityCardHeader">
              <div class="entityHeaderLeft">
                <div class="entityStatusRow"><span class="entityStatusPill">missing</span></div>
                <div class="entityTitle">${entity} PT</div>
                <div class="entitySubtitle">${brand.fullName}</div>
              </div>
              <div class="entityLogoWrap"><img src="${brand.logo}" alt="${entity}" class="entityLogo" /></div>
            </div>
            <div style="padding:12px;font-size:12px;color:var(--text-muted);">No PT data for this period.</div>
          </div>`;

        const pt      = row.pt      || {};
        const ptPrior = prior.pt    || {};

        const visitsSeen      = normalizeNumber(pt.visitsSeen);
        const visitsSeenPrior = normalizeNumber(ptPrior.visitsSeen);
        const scheduled       = normalizeNumber(pt.scheduledVisits);
        const cancellations   = normalizeNumber(pt.cancellations);
        const noShows         = normalizeNumber(pt.noShows);
        const reschedules     = normalizeNumber(pt.reschedules);
        const unitsBilled     = normalizeNumber(pt.totalUnitsBilled);
        const unitsPerVisit   = normalizeNumber(pt.unitsPerVisit);
        const visitsPerDay    = normalizeNumber(pt.visitsPerDay);

        const cancelRate  = scheduled > 0 ? Number(((cancellations / scheduled) * 100).toFixed(1)) : 0;
        const noShowRate  = scheduled > 0 ? Number(((noShows / scheduled) * 100).toFixed(1)) : 0;
        const utilRate    = scheduled > 0 ? Number(((visitsSeen / scheduled) * 100).toFixed(1)) : 0;

        const visitsDelta = visitsSeen - visitsSeenPrior;
        const fmtDelta = (n) => `${n >= 0 ? "+" : ""}${formatWhole(n)}`;

        const noShowCls  = noShowRate  >= 8 ? "bad" : noShowRate  >= 5 ? "warning" : "good";
        const cancelCls  = cancelRate  >= 10 ? "bad" : cancelRate  >= 6 ? "warning" : "good";
        const utilCls    = utilRate    >= 85 ? "good" : utilRate   >= 70 ? "warning" : "bad";

        return `
          <div class="entityCard" style="border-top:4px solid ${brand.accent};">
            <div class="entityCardHeader">
              <div class="entityHeaderLeft">
                <div class="entityStatusRow">
                  <span class="entityStatusPill">PT Program</span>
                  <span class="metricChip metricChip${utilRate >= 85 ? "Good" : utilRate >= 70 ? "Warning" : "Bad"}">${utilRate}% utilization</span>
                </div>
                <div class="entityTitle">${entity} — PT</div>
                <div class="entitySubtitle">${brand.fullName}</div>
              </div>
              <div class="entityLogoWrap"><img src="${brand.logo}" alt="${entity}" class="entityLogo" /></div>
            </div>

            <div class="entityTopMetrics">
              <div class="entityMetricHero">
                <span class="entityMetricLabel">Visits Seen</span>
                <strong>${formatWhole(visitsSeen)}</strong>
              </div>
              <div class="entityMetricHero">
                <span class="entityMetricLabel">Scheduled</span>
                <strong>${formatWhole(scheduled)}</strong>
              </div>
              <div class="entityMetricHero">
                <span class="entityMetricLabel">Units Billed</span>
                <strong>${formatWhole(unitsBilled)}</strong>
              </div>
            </div>

            <div class="entityCompareGrid">
              <div class="entityMiniStat">
                <span class="entityMiniLabel">vs Prior</span>
                <strong class="${visitsDelta >= 0 ? "good" : "bad"}">${fmtDelta(visitsDelta)}</strong>
              </div>
              <div class="entityMiniStat">
                <span class="entityMiniLabel">Units/Visit</span>
                <strong>${unitsPerVisit.toFixed(1)}</strong>
              </div>
              <div class="entityMiniStat">
                <span class="entityMiniLabel">Visits/Day</span>
                <strong>${visitsPerDay.toFixed(1)}</strong>
              </div>
            </div>

            <div class="entityHealthRow">
              <div class="entityHealthItem">
                <span class="entityHealthLabel">No Show</span>
                <strong class="${noShowCls}">${noShowRate}%</strong>
              </div>
              <div class="entityHealthItem">
                <span class="entityHealthLabel">Cancel</span>
                <strong class="${cancelCls}">${cancelRate}%</strong>
              </div>
              <div class="entityHealthItem">
                <span class="entityHealthLabel">Reschedules</span>
                <strong>${formatWhole(reschedules)}</strong>
              </div>
            </div>

            <div class="entityAccessPanel">
              <div class="entityAccessLabelRow">
                <span class="entityAccessLabel">Utilization</span>
                <span class="entityAccessValue">${utilRate}%</span>
              </div>
              <div class="entityProgressTrack">
                <div class="entityProgressBar ${utilRate >= 85 ? "entityProgressGood" : utilRate >= 70 ? "entityProgressWarning" : "entityProgressBad"}"
                  style="width:${Math.min(utilRate, 100)}%"></div>
              </div>
            </div>

            <details class="entityDetailDrawer">
              <summary>More detail</summary>
              <div class="entityDetailGrid">
                <div class="entityDetailTile">
                  <div class="entityDetailLabel">No Shows</div>
                  <div class="entityDetailValue">${formatWhole(noShows)}</div>
                </div>
                <div class="entityDetailTile">
                  <div class="entityDetailLabel">Cancellations</div>
                  <div class="entityDetailValue">${formatWhole(cancellations)}</div>
                </div>
                <div class="entityDetailTile">
                  <div class="entityDetailLabel">Reschedules</div>
                  <div class="entityDetailValue">${formatWhole(reschedules)}</div>
                </div>
              </div>
            </details>
          </div>
        `;
      }).join("")}
    </div>
  `;
}

(async function init() {
  // Inject design tokens/polish before any async work so the shell renders
  // correctly even if /api/me is slow or fails.
  injectUiPolishStyles();
  initCommandView();

  try {
    currentUser = await apiGet("/api/me");

    renderUser(currentUser);
    setupEntityDropdown();
    setupTrendsEntityDropdown();
    renderForm();
    syncEntryModeVisibility();

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
        // Auto-set anchor week for period types that use it
        const anchorInput = byId("dashboardWeekEnding");
        if (anchorInput) {
          const pt = byId("dashboardPeriodType")?.value;
          if (pt === "lastMonth" || pt === "ytd") {
            anchorInput.value = "";
          } else if (!anchorInput.value) {
            anchorInput.value = getDefaultWeekEnding();
          }
        }
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
        const input = byId("weekEnding");
        const warning = byId("weekEndingWarning");

        // Snap to nearest Friday if not already a Friday
        if (input?.value) {
          const d = new Date(input.value + "T12:00:00Z");
          const day = d.getUTCDay(); // 0=Sun, 5=Fri
          if (day !== 5) {
            // Calculate how many days to shift to nearest prior Friday
            const daysToFriday = day < 5 ? -(day + 2) : -(day - 5);
            d.setUTCDate(d.getUTCDate() + daysToFriday);
            const corrected = d.toISOString().slice(0, 10);

            if (warning) {
              warning.textContent = `⚠ ${input.value} is not a Friday — auto-corrected to ${corrected}.`;
              warning.style.display = "";
              warning.className = "weekEndingWarning weekEndingCorrected";
              setTimeout(() => { if (warning) warning.style.display = "none"; }, 5000);
            }
            input.value = corrected;
          } else {
            if (warning) warning.style.display = "none";
          }
        }

        // If the user has unsaved data in the form, warn before wiping
        const hasUnsavedData = formHasUnsavedData();
        if (hasUnsavedData) {
          const confirmed = confirm(
            "You have unsaved data in the form. Changing the week will load the new date and your current entries will be lost.\n\nContinue?"
          );
          if (!confirmed) {
            const prev = currentWeekData?.weekEnding || getDefaultWeekEnding();
            byId("weekEnding").value = prev;
            return;
          }
        }
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

    if (byId("navProviderSettingsBtn")) {
      byId("navProviderSettingsBtn").addEventListener("click", async () => {
        showProviderSettingsView();
        try { await loadProviderSettings(); } catch (e) { setProviderSettingsDebug(String(e)); }
      });
    }

    if (byId("navHelpBtn")) {
      byId("navHelpBtn").addEventListener("click", showHelpView);
    }

    if (byId("navGoalSettingsBtn")) {
      byId("navGoalSettingsBtn").addEventListener("click", showGoalSettingsView);
    }

    if (byId("navThresholdSettingsBtn")) {
      byId("navThresholdSettingsBtn").addEventListener("click", showThresholdSettingsView);
    }

    if (byId("digestBtn")) {
      byId("digestBtn").addEventListener("click", () => {
        showDigestModal(
          window._lastDashboardCurrent,
          window._lastDashboardComparison,
          window._lastDashboardCompareAgainst,
          window._lastDashboardWeekSets
        );
      });
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


// ══════════════════════════════════════════════════════════════
// CHANGELOG DATA & LOGIC
// ══════════════════════════════════════════════════════════════

const CHANGELOG = [
  {
    version: "2.1.0",
    date: "2026-04-22",
    label: "Latest",
    sections: [
      {
        type: "feature",
        title: "New Features",
        items: [
          "Live phone activity strip on every dashboard entity card — pulls real-time call volume, answer rate, abandoned calls, and average wait time directly from Landis (LAOSS + NES) and RingCentral (MRO + SpineOne). Follows the selected dashboard week so numbers change when you change the period.",
          "Ops Notes panel — the weekly narrative each regional lead writes is now a first-class section on the dashboard (was previously buried inside an entity card drawer). One card per entity, pre-scoped 'Add one →' link for entities without a note yet.",
          "Autosave for the Weekly Entry form — every change is saved to your browser every two seconds with a timestamp indicator next to the Save button. Close the tab mid-entry and you'll be prompted to restore on reload.",
          "Diff-before-save confirmation — saving a record that would overwrite any non-zero field with zero (or clear a non-empty Operator Notes) now triggers a modal listing each change before committing. Intended to prevent the kind of accidental wipe we saw on 4/17.",
          "RingCentral server-to-server JWT integration for MRO and SpineOne — replaces the deprecated password-grant flow and pulls Call Log hourly into the dashboard's CallData store.",
          "New Azure Function App 'ops-unitymsk-timers' handles the hourly RingCentral sync; Landis webhook now lands events directly in the SWA Managed API."
        ]
      },
      {
        type: "fix",
        title: "Bug Fixes",
        items: [
          "Imaging field not sticking on reload — the SpineOne imaging input was missing from the form's populate list so saved values never reappeared.",
          "SpineOne total visit count discrepancy — the executive-summary backend was reading a stale '{entity}-PT' partition and clobbering the correct PT values on the base partition.",
          "PTO provider breakdown disappearing for Annette/NES — the breakdown array was sent to the backend but never persisted. Now stored as JSON on the forecast record.",
          "Reschedules field not populating on load — same root cause as the imaging issue; both were missed when the fields were first added.",
          "Landis webhook entity resolver now matches actual Landis naming ('Landis-LAOSS-*', 'Landis-NES-*', agent UPNs at laorthos.com / nespecialists.com) with word-boundary safety so short tokens like 'NES' can't accidentally match inside unrelated words.",
          "Phone-data strip now reflects the selected dashboard period instead of always showing the current week — fixes the 'numbers don't change when I pick a different week' complaint.",
          "Aggregate imaging field initialized to zero so totals can't start as NaN."
        ]
      },
      {
        type: "improvement",
        title: "Improvements",
        items: [
          "Weekly Entry form refactored to a single field registry — adding or renaming a tracked field is now a one-line change instead of touching 7+ places. Long-term, this kills the class of bugs that caused the imaging + reschedules + PT zero-out incidents.",
          "Executive Summary is now properly scoped to the caller's entity — regional leads see only their own data; admins still see everything.",
          "Command View theme retired — removed ~200 lines of alternate-theme CSS that nobody used. Default (dark) and Light View remain. Anyone previously on Command View migrates back to default automatically.",
          "staticwebapp.config.json hardened with X-Frame-Options, nosniff, Referrer-Policy, Permissions-Policy, and HSTS headers; Landis webhook route explicitly opened for anonymous POSTs."
        ]
      },
      {
        type: "security",
        title: "Security",
        items: [
          "/api/call-data endpoint now requires Azure AD auth, validates the entity + weekEnding inputs, and blocks cross-tenant reads (regionals can only query their own entity).",
          "Landis webhook fails closed when the shared secret env var is missing, uses a timing-safe comparison, and no longer logs full event payloads (which could include caller identifiers).",
          "Stored XSS protections — escapeHtml helper now applied to every free-text field rendered via innerHTML (ops narrative, status, savedBy, updatedBy, created-by, etc.). Weekly digest textarea populated via .value instead of template interpolation so a stored '</textarea>' payload can no longer break out.",
          "/api/pto-forecast POST whitelists the entity against the known four and validates monthKey as YYYY-MM before writing.",
          "RingCentral integration moved from the deprecated password-grant flow to per-tenant JWT so user credentials are no longer stored in app settings."
        ]
      }
    ]
  },
  {
    version: "2.0.0",
    date: "2026-04-14",
    sections: [
      {
        type: "feature",
        title: "New Features",
        items: [
          "Visit Volume KPI card now shows per-day average — displays alongside the vs prior or budget comparison (e.g. 539.8/day avg), period and entity-aware",
          "Year to Date (YTD) period option added to dashboard — pulls all Fridays from Jan 1 to today and compares against the same period last year",
          "Anchor Week field now hides automatically when Last Month or YTD is selected — clears cleanly so it does not imply a selection that has no effect",
          "PT scope dashboard — Entity Scope dropdown now includes NES-PT, SpineOne-PT, MRO-PT, and All PT Programs as dedicated views",
          "PT-specific KPI cards — visits seen, scheduled visits, utilization rate, units billed, cancellations, no shows, reschedules, and units per visit",
          "PT entity cards — each PT entity gets a dedicated card with utilization progress bar and health stats",
          "Utilization rate metric — visits seen ÷ scheduled, color coded green above 85%, yellow above 70%, red below",
          "Command View — dark terminal-style alternate theme with Syne + JetBrains Mono and scan-line overlay",
          "Light View — warm minimalist theme with Inter + IBM Plex Mono and shadow-depth cards",
          "New Patients KPI card added to top-line dashboard metrics",
          "Rescheduled field added to weekly entry for NES, SpineOne, and MRO",
          "Data completeness banner — shows missing entities with a direct link to enter data",
          "Inline ops narrative — weekly notes visible on entity cards under More Detail",
          "Alert Thresholds editor — configure alert triggers from Admin without code changes",
          "Goal Settings — set weekly targets per entity from Admin",
          "Weekly Digest generator — copy-ready email and Teams summary from the dashboard",
          "Automated digest prompt — toast fires when all entities have entered data for the week",
          "Mobile responsive layout — sidebar, cards, and filters adapt to phone and tablet",
          "SpineOne PI trend chart — PI NP and cash charted over time in Trends view",
          "Provider-level PTO breakdown — add individual providers by name and days in PTO Forecast",
          "Phone system integration foundation — Landis webhook (LAOSS + NES) and RingCentral polling (MRO + SpineOne)",
          "Activity Log now defaults to most recent Friday on open"
        ]
      },
      {
        type: "fix",
        title: "Bug Fixes",
        items: [
          "PT scope dashboard not rendering — PT functions were scoped inside the init IIFE and inaccessible; moved to module level",
          "PT KPI cards showing no color — card class names corrected from kpi-good/bad/warning to kpi-positive/negative/neutral",
          "PT scope showing stale ortho cards — dashboard now clears previous cards before rendering PT view",
          "PT totals showing all-entity numbers instead of filtered scope — ptTotals now recomputed from filtered regions",
          "PT entry wiping ortho data — saving in PT mode no longer zeros ortho fields and vice versa",
          "Anchor week off-by-one — selecting 4/10 now loads the 4/10 record as expected",
          "Budget comparison always showing zero — daysInPeriod now uses entity-specific defaults",
          "Reschedules field not visible on initial load — syncEntryModeVisibility now called on first render",
          "No Show, Cancel, Abandoned stats blank on entity cards — CSS specificity conflict resolved",
          "Week ending date change wiping unsaved form data — confirmation dialog added before reloading",
          "Week ending auto-defaulting to next Friday instead of most recent past Friday",
          "Non-Friday week ending dates auto-correct to nearest prior Friday with a visible notice",
          "LAOSS 4/12 record migrated to correct date 4/10"
        ]
      },
      {
        type: "improvement",
        title: "Improvements",
        items: [
          "PTO Forecast now role-aware on the frontend — non-admins see only their entity, inputs disabled for other entities, access notice shown in banner",
          "Visit Volume per-day avg is period-aware — Last Week divides by 5 days, Rolling 4 divides by 20, NES uses 4.5 days/week automatically",
          "Period dropdown now has 6 options: Last Week, MTD, Last Month, YTD, Rolling 4 Weeks, Custom Range",
          "NES working days updated to 4.5 per week — budget proration, PTO forecast rates, and PT calculations all reflect Portland's actual schedule",
          "Entity-specific working days config — each entity has its own default; NES is 4.5, all others 5",
          "Working days hint banner appears in weekly entry when an entity operates fewer than 5 days",
          "Week ending field shows 'must be a Friday' label and helper text with an example date",
          "Completeness banner grammar fixed — '1 entities' now correctly reads '1 entity'",
          "formatPercent now returns '—' instead of NaN% for missing values",
          "View mode toggles are mutually exclusive and persist across sessions via localStorage"
        ]
      }
    ]
  },
  {
    version: "1.6.0",
    date: "2026-04-13",
    sections: [
      {
        type: "feature",
        title: "New Features",
        items: [
          "Command View and Light View theme toggles",
          "New Patients KPI card on dashboard",
          "Rescheduled field for NES, SpineOne, MRO",
          "PTO Forecast How It Works panel",
          "What's New modal and Changelog sidebar view"
        ]
      },
      {
        type: "fix",
        title: "Bug Fixes",
        items: [
          "PT entry wiping ortho data on save",
          "Anchor week showing previous week's data",
          "Budget comparison returning zero",
          "Reschedules field hidden on initial load"
        ]
      }
    ]
  },
  {
    version: "1.5.0",
    date: "2026-04-10",
    sections: [
      {
        type: "feature",
        title: "New Features",
        items: [
          "PTO Forecast — MD, PA, and PT PTO entry per entity and month with missed visit impact calculation",
          "Provider Settings admin panel — configure provider counts without code changes",
          "Budget vs Actual comparison mode — dashboard compares actuals against prorated monthly targets",
          "New Patients card in executive summary view"
        ]
      },
      {
        type: "improvement",
        title: "Improvements",
        items: [
          "Smart currency abbreviation — values display as $4.5M or $914K",
          "Sidebar reorganized into Views, Planning, Admin, and Help sections",
          "Entity card layout refresh — tighter padding, unified metric rows, dividers",
          "Health stats now stacked with color-coded thresholds",
          "Alert items use left-border severity pattern with red and gold bars",
          "Status pills updated to rectangular monospace chips"
        ]
      }
    ]
  },
  {
    version: "1.4.0",
    date: "2026-03-28",
    sections: [
      {
        type: "feature",
        title: "New Features",
        items: [
          "Executive Summary view with entity-level KPI cards and region drill-down",
          "Trends view with rolling visit volume and PT charts",
          "Activity Log — historical week-over-week entry audit trail",
          "Excel import for bulk weekly data entry via XLSX upload"
        ]
      },
      {
        type: "fix",
        title: "Bug Fixes",
        items: [
          "Dashboard entity bar chart rendering on Safari",
          "Rolling 4-week period including incorrect week boundaries",
          "Cash collected displaying as revenue in executive summary"
        ]
      }
    ]
  },
  {
    version: "1.3.0",
    date: "2026-03-14",
    sections: [
      {
        type: "feature",
        title: "New Features",
        items: [
          "Multi-entity dashboard with LAOSS, NES, SpineOne, and MRO",
          "Prior period comparison across all dashboard metrics",
          "No-show and cancellation rate thresholds with color-coded health indicators",
          "Weekly entry form with ortho and PT sub-entity modes"
        ]
      }
    ]
  }
];

const CHANGELOG_VERSION_KEY = "changelogLastSeen";

function getLatestVersion() {
  return CHANGELOG[0]?.version || "1.0.0";
}

function hasUnseenChangelog() {
  const lastSeen = localStorage.getItem(CHANGELOG_VERSION_KEY);
  return lastSeen !== getLatestVersion();
}

function markChangelogSeen() {
  localStorage.setItem(CHANGELOG_VERSION_KEY, getLatestVersion());
  const badge = byId("changelogBadge");
  if (badge) badge.style.display = "none";
}

function formatChangelogDate(dateStr) {
  const d = new Date(dateStr + "T12:00:00Z");
  return d.toLocaleDateString("en-US", { month: "long", day: "numeric", year: "numeric" });
}

function renderChangelogEntries(entries, container) {
  const typeConfig = {
    feature:     { label: "New",      cls: "clTagFeature"  },
    fix:         { label: "Fixed",    cls: "clTagFix"      },
    improvement: { label: "Improved", cls: "clTagImprove"  },
    security:    { label: "Security", cls: "clTagSecurity" }
  };

  container.innerHTML = entries.map((release, idx) => `
    <div class="clRelease${idx === 0 ? " clReleaseCurrent" : ""}">
      <div class="clReleaseHeader">
        <div class="clReleaseLeft">
          <span class="clVersion">v${release.version}</span>
          ${release.label ? `<span class="clLatestBadge">${release.label}</span>` : ""}
        </div>
        <span class="clDate">${formatChangelogDate(release.date)}</span>
      </div>
      ${release.sections.map(sec => `
        <div class="clSection">
          <div class="clSectionHeader">
            <span class="clTag ${typeConfig[sec.type]?.cls || ""}">${typeConfig[sec.type]?.label || sec.title}</span>
            <span class="clSectionTitle">${sec.title}</span>
          </div>
          <ul class="clItems">
            ${sec.items.map(item => `<li class="clItem">${item}</li>`).join("")}
          </ul>
        </div>
      `).join("")}
    </div>
  `).join("");
}

function showChangelogModal() {
  const modal = byId("changelogModal");
  if (!modal) return;
  const latest = CHANGELOG[0];
  if (!latest) return;

  const vEl = byId("changelogModalVersion");
  if (vEl) vEl.textContent = "v" + latest.version;

  const body = byId("changelogModalBody");
  if (body) renderChangelogEntries([latest], body);

  modal.style.display = "flex";
  document.body.style.overflow = "hidden";

  const dismissBtn = byId("changelogDismissBtn");
  if (dismissBtn) dismissBtn.onclick = dismissChangelogModal;

  const viewAllBtn = byId("changelogViewAllBtn");
  if (viewAllBtn) viewAllBtn.onclick = () => {
    dismissChangelogModal();
    showChangelogView();
  };

  modal.onclick = (e) => { if (e.target === modal) dismissChangelogModal(); };
}

function dismissChangelogModal() {
  const modal = byId("changelogModal");
  if (modal) {
    modal.style.opacity = "0";
    setTimeout(() => {
      modal.style.display = "none";
      modal.style.opacity = "";
      document.body.style.overflow = "";
    }, 200);
  }
  markChangelogSeen();
}

function showChangelogView() {
  hideAllViews();
  const el = byId("changelogView");
  if (el) el.style.display = "";
  setActiveNav("navChangelogBtn");
  markChangelogSeen();
  const container = byId("changelogViewEntries");
  if (container) renderChangelogEntries(CHANGELOG, container);
}

function initChangelog() {
  if (hasUnseenChangelog()) {
    const badge = byId("changelogBadge");
    if (badge) { badge.textContent = "New"; badge.style.display = ""; }
  }
  const navBtn = byId("navChangelogBtn");
  if (navBtn) navBtn.addEventListener("click", showChangelogView);
  if (hasUnseenChangelog()) {
    setTimeout(showChangelogModal, 900);
  }
}


// ── Entity Goals (per-entity, per-metric monthly targets) ─────
function getEntityGoal(entity, metric) {
  const key = `goal_${entity}_${metric}`;
  const stored = localStorage.getItem(key);
  if (stored !== null) {
    const n = parseFloat(stored);
    if (Number.isFinite(n)) return n;
  }
  return null;
}

function setEntityGoal(entity, metric, value) {
  localStorage.setItem(`goal_${entity}_${metric}`, String(value));
}

function renderGoalSettingsView() {
  const container = byId("goalSettingsContent");
  if (!container) return;

  const metrics = [
    { key: "visitVolume",   label: "Weekly Visit Target" },
    { key: "newPatients",   label: "Weekly NP Target"    },
    { key: "cashCollected", label: "Weekly Cash Target"  }
  ];

  container.innerHTML = `
    <div class="goalGrid">
      ${ENTITIES.map(entity => `
        <div class="goalEntityBlock">
          <div class="goalEntityTitle">${entity}</div>
          ${metrics.map(m => {
            const val = getEntityGoal(entity, m.key);
            return `
              <div class="goalRow">
                <label class="goalLabel">${m.label}</label>
                <input
                  type="number"
                  class="goalInput"
                  data-entity="${entity}"
                  data-metric="${m.key}"
                  value="${val !== null ? val : ""}"
                  placeholder="Not set"
                  min="0"
                  step="1"
                />
              </div>`;
          }).join("")}
        </div>
      `).join("")}
    </div>
    <div class="goalActions">
      <button type="button" class="primaryBtn" id="saveGoalsBtn">Save Goals</button>
      <button type="button" class="secondaryBtn" id="clearGoalsBtn">Clear All</button>
    </div>
  `;

  byId("saveGoalsBtn").addEventListener("click", () => {
    container.querySelectorAll(".goalInput").forEach(input => {
      const entity = input.dataset.entity;
      const metric = input.dataset.metric;
      const val = parseFloat(input.value);
      if (Number.isFinite(val) && val > 0) {
        setEntityGoal(entity, metric, val);
      } else {
        localStorage.removeItem(`goal_${entity}_${metric}`);
      }
    });
    showGoalBanner("Goals saved.", true);
  });

  byId("clearGoalsBtn").addEventListener("click", () => {
    ENTITIES.forEach(e => ["visitVolume","newPatients","cashCollected"].forEach(m => {
      localStorage.removeItem(`goal_${e}_${m}`);
    }));
    renderGoalSettingsView();
    showGoalBanner("All goals cleared.");
  });
}

function showGoalBanner(msg, good = false) {
  const el = byId("goalStatusBanner");
  if (!el) return;
  el.textContent = msg;
  el.className = "goalStatusBanner " + (good ? "goalStatusGood" : "goalStatusNeutral");
  el.style.display = "";
  setTimeout(() => { el.style.display = "none"; }, 3000);
}

function showGoalSettingsView() {
  hideAllViews();
  const el = byId("goalSettingsView");
  if (el) el.style.display = "";
  setActiveNav("navGoalSettingsBtn");
  renderGoalSettingsView();
}


// ── Weekly Digest Generator ───────────────────────────────────
function generateWeeklyDigest(current, comparison, compareAgainst, weekSets) {
  const weekLabel = weekSets?.primaryWeeks?.[0]
    ? new Date(weekSets.primaryWeeks[0] + "T12:00:00Z").toLocaleDateString("en-US", { weekday: "long", month: "long", day: "numeric", year: "numeric" })
    : "Selected Week";

  const lines = [];
  lines.push(`📊 UnityMSK Weekly Ops Summary — Week Ending ${weekLabel}`);
  lines.push("─".repeat(52));

  ENTITIES.forEach(entity => {
    const row = (current.regions || []).find(r => r.entity === entity);
    const prior = (comparison.regions || []).find(r => r.entity === entity);
    if (!row || normalizeNumber(row.visitVolume) === 0) {
      lines.push(`
${entity}: No data entered`);
      return;
    }

    const visitDelta = prior ? normalizeNumber(row.visitVolume) - normalizeNumber(prior.visitVolume) : null;
    const npDelta    = prior ? normalizeNumber(row.newPatients) - normalizeNumber(prior.newPatients) : null;
    const sign = (n) => n >= 0 ? "+" : "";

    lines.push(`
${entity}`);
    lines.push(`  Visits:     ${formatWhole(row.visitVolume)}${visitDelta !== null ? "  (" + sign(visitDelta) + visitDelta + " vs prior)" : ""}`);
    lines.push(`  New Pts:    ${formatWhole(row.newPatients)}${npDelta !== null ? "  (" + sign(npDelta) + npDelta + " vs prior)" : ""}`);
    lines.push(`  Cash:       ${formatCurrency(row.cashCollected || 0)}`);
    lines.push(`  No Show:    ${normalizeNumber(row.noShowRate).toFixed(1)}%  |  Cancel: ${normalizeNumber(row.cancellationRate).toFixed(1)}%  |  Abandoned: ${normalizeNumber(row.abandonedCallRate).toFixed(1)}%`);
    if (row.ptoDays > 0) lines.push(`  PTO Days:   ${row.ptoDays}`);
    if (row.operationsNarrative) lines.push(`  Notes:      ${row.operationsNarrative}`);
  });

  lines.push("\n" + "─".repeat(52));
  lines.push(`Generated by UnityMSK Ops Dashboard · ${new Date().toLocaleDateString()}`);
  return lines.join("\n");
}

function showDigestModal(current, comparison, compareAgainst, weekSets) {
  const existing = byId("digestModal");
  if (existing) existing.remove();

  const text = generateWeeklyDigest(current, comparison, compareAgainst, weekSets);

  const modal = document.createElement("div");
  modal.id = "digestModal";
  modal.className = "changelogModalOverlay";
  modal.innerHTML = `
    <div class="changelogModalBox" style="max-width:620px;">
      <div class="changelogModalHeader">
        <div class="changelogModalMeta">
          <span class="changelogModalTag">Weekly Summary</span>
        </div>
        <h2 class="changelogModalTitle">Email / Teams Digest</h2>
        <p class="changelogModalSub">Copy and paste into an email or Teams message to share with leadership.</p>
      </div>
      <div class="changelogModalBody" style="padding:16px 24px;">
        <textarea id="digestTextarea" class="digestTextarea" readonly></textarea>
      </div>
      <div class="changelogModalFooter">
        <button type="button" class="changelogViewAllBtn" id="digestCopyBtn">Copy to Clipboard</button>
        <button type="button" class="changelogDismissBtn" id="digestCloseBtn">Close</button>
      </div>
    </div>
  `;

  document.body.appendChild(modal);
  // Assign via .value rather than interpolating into innerHTML — prevents
  // a stored </textarea><script>... payload in any source field from breaking out.
  const digestTa = byId("digestTextarea");
  if (digestTa) digestTa.value = text;
  modal.style.display = "flex";
  document.body.style.overflow = "hidden";

  byId("digestCloseBtn").onclick = () => {
    modal.remove();
    document.body.style.overflow = "";
  };
  modal.onclick = (e) => { if (e.target === modal) { modal.remove(); document.body.style.overflow = ""; } };

  byId("digestCopyBtn").onclick = () => {
    const ta = byId("digestTextarea");
    ta.select();
    document.execCommand("copy");
    byId("digestCopyBtn").textContent = "Copied ✓";
    setTimeout(() => { byId("digestCopyBtn").textContent = "Copy to Clipboard"; }, 2000);
  };
}


// ── Threshold Settings View ───────────────────────────────────
function renderThresholdSettingsView() {
  const container = byId("thresholdSettingsContent");
  if (!container) return;

  const thresholds = [
    { key: "noShowRate",        label: "No-Show Rate Alert (%)",        unit: "%", desc: "Alert fires when no-show rate exceeds this value" },
    { key: "cancellationRate",  label: "Cancellation Rate Alert (%)",   unit: "%", desc: "Alert fires when cancellation rate exceeds this value" },
    { key: "abandonedCallRate", label: "Abandoned Call Rate Alert (%)", unit: "%", desc: "Alert fires when abandoned call rate exceeds this value" },
    { key: "visitDropVsPrior",  label: "Visit Drop Alert (visits)",     unit: "",  desc: "Alert fires when visits drop more than this vs prior period" }
  ];

  container.innerHTML = `
    <div class="thresholdGrid">
      ${thresholds.map(t => `
        <div class="thresholdRow">
          <div class="thresholdInfo">
            <div class="thresholdLabel">${t.label}</div>
            <div class="thresholdDesc">${t.desc}</div>
          </div>
          <div class="thresholdInputWrap">
            <input
              type="number"
              class="thresholdInput"
              data-key="${t.key}"
              value="${getThreshold(t.key, THRESHOLD_DEFAULTS[t.key])}"
              min="0"
              step="0.5"
            />
            ${t.unit ? `<span class="thresholdUnit">${t.unit}</span>` : ""}
          </div>
        </div>
      `).join("")}
    </div>
    <div class="goalActions">
      <button type="button" class="primaryBtn" id="saveThresholdsBtn">Save Thresholds</button>
      <button type="button" class="secondaryBtn" id="resetThresholdsBtn">Reset to Defaults</button>
    </div>
    <div id="thresholdStatusBanner" class="goalStatusBanner" style="display:none;"></div>
  `;

  byId("saveThresholdsBtn").addEventListener("click", () => {
    container.querySelectorAll(".thresholdInput").forEach(input => {
      const key = input.dataset.key;
      const val = parseFloat(input.value);
      if (Number.isFinite(val)) setThreshold(key, val);
    });
    const banner = byId("thresholdStatusBanner");
    if (banner) { banner.textContent = "Thresholds saved."; banner.className = "goalStatusBanner goalStatusGood"; banner.style.display = ""; setTimeout(() => banner.style.display = "none", 3000); }
  });

  byId("resetThresholdsBtn").addEventListener("click", () => {
    resetThresholds();
    renderThresholdSettingsView();
    const banner = byId("thresholdStatusBanner");
    if (banner) { banner.textContent = "Reset to defaults."; banner.className = "goalStatusBanner goalStatusNeutral"; banner.style.display = ""; setTimeout(() => banner.style.display = "none", 3000); }
  });
}

function showThresholdSettingsView() {
  hideAllViews();
  const el = byId("thresholdSettingsView");
  if (el) el.style.display = "";
  setActiveNav("navThresholdSettingsBtn");
  renderThresholdSettingsView();
}

    initChangelog();
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
