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

function formatWhole(value) {
  return Math.round(normalizeNumber(value)).toLocaleString();
}

function formatCurrency(value) {
  return normalizeNumber(value).toLocaleString(undefined, {
    style: "currency",
    currency: "USD",
    maximumFractionDigits: 0
  });
}

function formatPercent(value, digits = 1) {
  return `${normalizeNumber(value).toFixed(digits)}%`;
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

function isSpineOneBaseEntity() {
  return getSelectedEntity() === "SpineOne" && !entityHasPtEntry();
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
  if (el) {
    el.innerText = label;
  }
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

function getSelectedTrendsEntity() {
  const el = byId("trendsEntitySelect");
  return el ? el.value : "LAOSS";
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
        <textarea id="operationsNarrative" rows="10" placeholder="Example:

What went well this week:
Call volume improved and scheduling stabilized.

What could have made it even better:
Additional provider availability and lower cancellation volume.

What were the major red flags:
NP volume came in light and referral lag impacted throughput.

Any adjustments or updates for next week:
Reinforce outreach and review scheduling templates.

Any additional insight:
One high-value case closed and cash collections were strong."></textarea>
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
    if (input) {
      input.addEventListener("input", updateDerivedDisplays);
    }
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

function hideAllViews() {
  const ids = ["dashboardView", "entryView", "executiveView", "trendsView", "importView"];
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

function renderDashboardCards(current, comparison, compareAgainst, entityScope) {
  const ptVisitsCurrent = normalizeNumber(current.ptTotals?.visitsSeen);
  const ptVisitsComparison = normalizeNumber(comparison.ptTotals?.visitsSeen);

  const avgNoShow = averageMetric(current.regions, "noShowRate");
  const avgCancel = averageMetric(current.regions, "cancellationRate");
  const avgCxnsCombined = avgNoShow + avgCancel;

  const visitVariance = normalizeNumber(current.totals?.visitVolume) - normalizeNumber(comparison.totals?.visitVolume);

  const cards = [
    {
      label: "Visit Volume",
      value: formatWhole(current.totals?.visitVolume || 0),
      meta: compareAgainst === "budget"
        ? `Budget ${formatWhole(current.budgetTotals?.visitVolumeBudget || 0)}`
        : `${visitVariance >= 0 ? "+" : ""}${formatWhole(visitVariance)} vs prior`,
      className:
        compareAgainst === "budget"
          ? getTrendClass(current.totals?.visitVolume, current.budgetTotals?.visitVolumeBudget)
          : getTrendClass(current.totals?.visitVolume, comparison.totals?.visitVolume)
    },
    {
      label: "PT Visits",
      value: formatWhole(ptVisitsCurrent),
      meta: `${ptVisitsCurrent - ptVisitsComparison >= 0 ? "+" : ""}${formatWhole(ptVisitsCurrent - ptVisitsComparison)} vs prior`,
      className: getTrendClass(ptVisitsCurrent, ptVisitsComparison)
    },
    {
      label: "Cash Collected",
      value: formatCurrency(current.totals?.cashCollected || 0),
      meta: compareAgainst === "budget" ? "Actual only" : "Across selected period",
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
      meta: compareAgainst === "budget" ? "Actual only" : "Across selected period",
      className: "kpi-neutral"
    },
    {
      label: "Call Volume",
      value: formatWhole(current.totals?.callVolume || 0),
      meta: compareAgainst === "budget" ? "Actual only" : "Across selected period",
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

async function loadExecutiveSummary() {
  const weekEnding = byId("executiveWeekEnding")?.value || getDefaultWeekEnding();
  const result = await apiGet(`/api/executive-summary?weekEnding=${encodeURIComponent(weekEnding)}`);
  renderExecutiveCards(result);
  renderExecutiveRegions(result);
  setExecutiveDebug(result);
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
  setDashboardDebug({
    compareAgainst,
    entityScope,
    weekSets,
    current,
    comparison
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

    .narrativeFieldBlock {
      margin-top: 10px;
    }

    .notesShell {
      padding: 18px;
      border-radius: 18px;
      background: linear-gradient(180deg, rgba(17,63,93,0.42) 0%, rgba(11,42,63,0.42) 100%);
      border: 1px solid rgba(108,182,255,0.16);
      box-shadow: inset 0 1px 0 rgba(255,255,255,0.03);
    }

    .notesHeader {
      margin-bottom: 14px;
    }

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

    .summaryCard {
      min-width: 0;
      overflow: hidden;
    }

    .summaryCard .value {
      font-size: clamp(1.7rem, 2.6vw, 2.5rem) !important;
      line-height: 1.05 !important;
      letter-spacing: -0.03em;
      overflow-wrap: anywhere;
      word-break: break-word;
      max-width: 100%;
      display: block;
    }

    .summaryCard h3 {
      min-width: 0;
      overflow-wrap: anywhere;
      word-break: break-word;
    }

    .summaryCard div:not(.value):not(h3) {
      min-width: 0;
      overflow-wrap: anywhere;
      word-break: break-word;
    }

    @media (max-width: 768px) {
      .notesTitle {
        font-size: 18px;
      }

      .notesPromptGrid {
        grid-template-columns: 1fr;
      }
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

    const defaultWeek = getDefaultWeekEnding();

    if (byId("dashboardWeekEnding")) byId("dashboardWeekEnding").value = defaultWeek;
    if (byId("dashboardCustomEnd")) byId("dashboardCustomEnd").value = defaultWeek;
    if (byId("dashboardCustomStart")) byId("dashboardCustomStart").value = getDateWeeksAgo(8, defaultWeek);
    if (byId("weekEnding")) byId("weekEnding").value = defaultWeek;
    if (byId("executiveWeekEnding")) byId("executiveWeekEnding").value = defaultWeek;

    const dashboardPeriodType = byId("dashboardPeriodType");
    if (dashboardPeriodType) {
      const currentWeekOption = dashboardPeriodType.querySelector('option[value="currentWeek"]');
      if (currentWeekOption) currentWeekOption.remove();
      dashboardPeriodType.value = "lastWeek";
    }

    renderEntityBrand("entryBrandWrap", getSelectedEntity());
    injectUiPolishStyles();

    if (byId("entitySelect")) {
      byId("entitySelect").addEventListener("change", async () => {
        renderEntityBrand("entryBrandWrap", getSelectedEntity());
        syncEntryModeVisibility();
        updateDerivedDisplays();
        await loadWeek();
      });
    }

    if (byId("weekEnding")) {
      byId("weekEnding").addEventListener("change", async () => {
        await loadWeek();
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
        await loadDashboardLanding();
      });
    }

    if (byId("navEntryBtn")) {
      byId("navEntryBtn").addEventListener("click", showEntryView);
    }

    if (byId("navExecutiveBtn")) {
      byId("navExecutiveBtn").addEventListener("click", async () => {
        showExecutiveView();
        await loadExecutiveSummary();
      });
    }

    if (byId("navTrendsBtn")) {
      byId("navTrendsBtn").addEventListener("click", showTrendsView);
    }

    if (byId("navImportBtn")) {
      byId("navImportBtn").addEventListener("click", showImportView);
    }

    if (byId("loadDashboardBtn")) {
      byId("loadDashboardBtn").addEventListener("click", async () => {
        await loadDashboardLanding();
      });
    }

    if (byId("dashboardPeriodType")) {
      byId("dashboardPeriodType").addEventListener("change", async () => {
        syncDashboardPeriodUi();
        await loadDashboardLanding();
      });
    }

    if (byId("dashboardCompareAgainst")) {
      byId("dashboardCompareAgainst").addEventListener("change", async () => {
        await loadDashboardLanding();
      });
    }

    if (byId("dashboardEntityScope")) {
      byId("dashboardEntityScope").addEventListener("change", async () => {
        await loadDashboardLanding();
      });
    }

    if (byId("loadExecutiveBtn")) {
      byId("loadExecutiveBtn").addEventListener("click", async () => {
        await loadExecutiveSummary();
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
    setImportDebug(String(error));
  }
})();
