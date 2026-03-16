import {
  safeApiGet,
  apiPost,
  getDefaultWeekEnding,
  formatDate,
  escapeHtml,
  escapeAttr,
  handleFatalError,
  comparisonClass,
  formatChange,
  renderSummaryMiniCard,
  collectRegionFormValues,
  collectSharedFormValues
} from "./helpers.js";

// ... the rest of your app.js code starts here
import {
  getRegionSections,
  // ... etc
```javascript
import {
  getRegionSections,
  getAllMetricKeysForEntity,
  getSharedPageDefinition,
  getAllMetricKeysForSharedPage,
  ENTITY_LABELS
} from "./definitions.js";
import {
  calculateRegionSummaries,
  calculateSharedSummaries,
  getRegionCalculatedValues,
  getSharedCalculatedValues,
  formatByType
} from "./calculations.js";
import {
  renderAdminEditorShell,
  renderTargetsEditor,
  renderThresholdsEditor,
  renderHolidaysEditor,
  renderBudgetEditor,
  renderSubmissionTracker,
  renderAuditViewer,
  renderImportTool,
  collectAdminRows
} from "./admin.js";

const state = {
  me: null,
  currentRoute: "executive",
  currentEntity: null,
  currentSharedPage: null,
  currentWeekEnding: getDefaultWeekEnding(),
  pageData: null,
  activeRegionSectionKey: null,
  activeSharedSectionKey: null,
  activeAdminTab: "targets",
  activeAdminEntity: "LAOSS",
  activeAdminYear: String(new Date().getFullYear()),
  activeAdminAuditEntity: "",
  lastImportResult: null
};

const els = {
  pageTitle: document.getElementById("pageTitle"),
  pageSubtitle: document.getElementById("pageSubtitle"),
  pageContent: document.getElementById("pageContent"),
  dashboardCards: document.getElementById("dashboardCards"),
  userDisplayName: document.getElementById("userDisplayName"),
  assignedEntityText: document.getElementById("assignedEntityText"),
  roleText: document.getElementById("roleText"),
  weekEndingSelect: document.getElementById("weekEndingSelect"),
  selectedWeekText: document.getElementById("selectedWeekText"),
  submissionStatusText: document.getElementById("submissionStatusText"),
  sidebarNav: document.getElementById("sidebarNav"),
  adminNavBtn: document.getElementById("adminNavBtn"),
  saveBtn: document.getElementById("saveBtn"),
  submitBtn: document.getElementById("submitBtn"),
  goAssignedRegionBtn: document.getElementById("goAssignedRegionBtn"),
  goExecutiveBtn: document.getElementById("goExecutiveBtn")
};

bootstrap();

function bootstrap() {
  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", () => {
      init().catch(handleFatalError);
    }, { once: true });
    return;
  }

  init().catch(handleFatalError);
}

async function init() {
  if (els.weekEndingSelect) {
    els.weekEndingSelect.value = state.currentWeekEnding;
  }

  if (els.selectedWeekText) {
    els.selectedWeekText.textContent = formatDate(state.currentWeekEnding);
  }

  renderKpiCards([]);
  bindEvents();

  const rawMe = await safeApiGet("/api/me", {
    authenticated: false,
    userDetails: "",
    roles: ["anonymous"]
  });

  state.me = normalizeMe(rawMe);
  applyUserContext();

  if (state.me.isAdmin) {
    await setRoute("executive");
    return;
  }

  await setRoute("region", state.me.entity || "LAOSS");
}

function normalizeMe(raw) {
  const roles = Array.isArray(raw?.roles) ? raw.roles : [];
  const userDetails = raw?.userDetails || "";
  const inferredEntity = inferEntityFromEmail(userDetails);

  return {
    authenticated: Boolean(raw?.authenticated),
    email: userDetails || "",
    displayName: userDetails || "Unknown User",
    entity: inferredEntity,
    roles,
    isAdmin: roles.includes("admin")
  };
}

function inferEntityFromEmail(email) {
  const lower = String(email || "").toLowerCase();

  if (!lower) return "LAOSS";
  if (lower.includes("nes")) return "NES";
  if (lower.includes("spine")) return "SpineOne";
  if (lower.includes("mro")) return "MRO";

  return "LAOSS";
}

function bindEvents() {
  if (els.weekEndingSelect) {
    els.weekEndingSelect.addEventListener("change", async (e) => {
      state.currentWeekEnding = e.target.value || getDefaultWeekEnding();

      if (els.selectedWeekText) {
        els.selectedWeekText.textContent = formatDate(state.currentWeekEnding);
      }

      await loadCurrentRoute();
    });
  }

  if (els.sidebarNav) {
    els.sidebarNav.addEventListener("click", async (e) => {
      const btn = e.target.closest(".nav-link");
      if (!btn) return;

      const route = btn.dataset.route;
      const entity = btn.dataset.entity || null;
      const page = btn.dataset.page || null;

      if (route === "region") {
        await setRoute("region", entity);
        return;
      }

      if (route === "shared") {
        await setRoute("shared", null, page);
        return;
      }

      if (route === "admin") {
        await setRoute("admin");
        return;
      }

      await setRoute("executive");
    });
  }

  document.addEventListener("click", async (e) => {
    const regionTab = e.target.closest(".section-tab.region-tab");
    if (regionTab) {
      state.activeRegionSectionKey = regionTab.dataset.sectionKey;
      if (state.currentRoute === "region" && state.currentEntity) {
        await renderRegion(state.currentEntity);
      }
      return;
    }

    const sharedTab = e.target.closest(".section-tab.shared-tab");
    if (sharedTab) {
      state.activeSharedSectionKey = sharedTab.dataset.sectionKey;
      if (state.currentRoute === "shared" && state.currentSharedPage) {
        await renderSharedPage(state.currentSharedPage);
      }
      return;
    }

    const adminTab = e.target.closest(".admin-editor-tab");
    if (adminTab) {
      state.activeAdminTab = adminTab.dataset.adminTab;
      if (state.currentRoute === "admin") {
        await renderAdmin();
      }
      return;
    }

    if (e.target.closest("#runWorkbookImportBtn")) {
      const fileInput = document.getElementById("workbookUploadInput");
      const file = fileInput?.files?.[0];

      if (!file) {
        alert("Select a workbook first.");
        return;
      }

      const fileBase64 = await fileToBase64(file);
      const result = await apiPost("/api/import-excel", {
        fileName: file.name,
        fileBase64
      });

      state.lastImportResult = result;
      alert("Workbook import completed.");
      await renderAdmin();
      return;
    }

    if (e.target.closest("#saveTargetsBtn")) {
      const rows = collectAdminRows("target");
      await apiPost("/api/admin-reference-save", {
        entity: state.activeAdminEntity,
        kind: "targets",
        rows
      });
      alert("Targets saved.");
      await renderAdmin();
      return;
    }

    if (e.target.closest("#saveThresholdsBtn")) {
      const rows = collectAdminRows("threshold");
      await apiPost("/api/admin-reference-save", {
        entity: state.activeAdminEntity,
        kind: "thresholds",
        rows
      });
      alert("Thresholds saved.");
      await renderAdmin();
      return;
    }

    if (e.target.closest("#saveHolidaysBtn")) {
      const rows = collectAdminRows("holiday");
      await apiPost("/api/admin-reference-save", {
        kind: "holidays",
        year: state.activeAdminYear,
        rows
      });
      alert("Holidays saved.");
      await renderAdmin();
      return;
    }

    if (e.target.closest("#saveBudgetBtn")) {
      const rows = collectAdminRows("budget");
      await apiPost("/api/admin-reference-save", {
        entity: state.activeAdminEntity,
        kind: "budget",
        rows
      });
      alert("Budget saved.");
      await renderAdmin();
      return;
    }

    if (e.target.closest("#printMeetingSummaryBtn")) {
      window.print();
      return;
    }

    const approveBtn = e.target.closest(".approve-week-btn");
    if (approveBtn) {
      const entity = approveBtn.dataset.entity;
      const weekEnding = approveBtn.dataset.weekEnding;
      await apiPost("/api/approve-week", { entity, weekEnding });
      alert(`${entity} approved for ${weekEnding}.`);
      await renderAdmin();
    }
  });

  document.addEventListener("change", async (e) => {
    if (e.target.id === "adminEntityFilter") {
      state.activeAdminEntity = e.target.value;
      if (state.currentRoute === "admin") await renderAdmin();
      return;
    }

    if (e.target.id === "adminYearFilter") {
      state.activeAdminYear = e.target.value;
      if (state.currentRoute === "admin") await renderAdmin();
      return;
    }

    if (e.target.id === "adminAuditEntityFilter") {
      state.activeAdminAuditEntity = e.target.value;
      if (state.currentRoute === "admin") await renderAdmin();
    }
  });

  if (els.goAssignedRegionBtn) {
    els.goAssignedRegionBtn.addEventListener("click", async () => {
      if (!state.me) return;

      if (state.me.isAdmin) {
        await setRoute("executive");
        return;
      }

      await setRoute("region", state.me.entity || "LAOSS");
    });
  }

  if (els.goExecutiveBtn) {
    els.goExecutiveBtn.addEventListener("click", async () => {
      await setRoute("executive");
    });
  }

  if (els.saveBtn) {
    els.saveBtn.addEventListener("click", async () => {
      if (state.currentRoute === "region") {
        const payload = collectRegionFormValues();
        const res = await apiPost("/api/weekly-save", payload);

        if (els.submissionStatusText) {
          els.submissionStatusText.textContent = res.status || "Draft";
        }

        alert("Saved successfully.");
        await loadCurrentRoute();
        return;
      }

      if (state.currentRoute === "shared") {
        const payload = collectSharedFormValues();
        const res = await apiPost("/api/shared-save", payload);

        if (els.submissionStatusText) {
          els.submissionStatusText.textContent = res.status || "Draft";
        }

        alert("Shared page saved successfully.");
        await loadCurrentRoute();
        return;
      }

      alert("Save is only active on region and shared pages.");
    });
  }

  if (els.submitBtn) {
    els.submitBtn.addEventListener("click", async () => {
      if (state.currentRoute !== "region") {
        alert("Submit is only active on region pages.");
        return;
      }

      const res = await apiPost("/api/submit-week", {
        entity: state.currentEntity,
        weekEnding: state.currentWeekEnding
      });

      if (els.submissionStatusText) {
        els.submissionStatusText.textContent = res.status || "Submitted";
      }

      alert("Week submitted.");
      await loadCurrentRoute();
    });
  }
}

function applyUserContext() {
  if (els.userDisplayName) {
    els.userDisplayName.textContent = state.me.displayName || state.me.email || "Unknown User";
  }

  if (els.assignedEntityText) {
    els.assignedEntityText.textContent = state.me.entity || "None";
  }

  if (els.roleText) {
    els.roleText.textContent = state.me.isAdmin ? "Admin" : "Editor";
  }

  if (els.adminNavBtn) {
    if (state.me.isAdmin) {
      els.adminNavBtn.classList.remove("hidden");
    } else {
      els.adminNavBtn.classList.add("hidden");
    }
  }
}

async function setRoute(route, entity = null, sharedPage = null) {
  state.currentRoute = route;
  state.currentEntity = entity;
  state.currentSharedPage = sharedPage;

  document.querySelectorAll(".nav-link").forEach((btn) => btn.classList.remove("active"));

  const selector = route === "region"
    ? `.nav-link[data-route="region"][data-entity="${cssEscape(entity)}"]`
    : route === "shared"
      ? `.nav-link[data-route="shared"][data-page="${cssEscape(sharedPage)}"]`
      : `.nav-link[data-route="${cssEscape(route)}"]`;

  const activeBtn = document.querySelector(selector);
  if (activeBtn) {
    activeBtn.classList.add("active");
  }

  if (route === "region" && entity) {
    const sections = getRegionSections(entity);
    state.activeRegionSectionKey = sections[0]?.key || null;
  }

  if (route === "shared" && sharedPage) {
    const def = getSharedPageDefinition(sharedPage);
    state.activeSharedSectionKey = def?.sections?.[0]?.key || null;
  }

  await loadCurrentRoute();
}

async function loadCurrentRoute() {
  if (state.currentRoute === "executive") {
    await renderExecutive();
    return;
  }

  if (state.currentRoute === "region") {
    await renderRegion(state.currentEntity);
    return;
  }

  if (state.currentRoute === "shared") {
    await renderSharedPage(state.currentSharedPage);
    return;
  }

  if (state.currentRoute === "admin") {
    await renderAdmin();
  }
}

async function renderExecutive() {
  if (els.pageTitle) {
    els.pageTitle.textContent = "Executive Summary";
  }

  if (els.pageSubtitle) {
    els.pageSubtitle.textContent =
      "Weekly companywide KPI overview with trends, target comparisons, and submission visibility.";
  }

  const data = await safeApiGet(`/api/dashboard?weekEnding=${encodeURIComponent(state.currentWeekEnding)}`, {
    weekEnding: state.currentWeekEnding,
    previousWeekEnding: "",
    kpis: [],
    entities: [],
    comparison: [],
    riskMetrics: [],
    commentaryRollup: []
  });

  state.pageData = data;

  renderKpiCards(data.kpis || []);
  if (els.submissionStatusText) {
    els.submissionStatusText.textContent = "Summary View";
  }

  if (els.pageContent) {
    els.pageContent.innerHTML = `
      <div class="section-head meeting-summary-header">
        <div class="meeting-summary-top">
          <div>
            <h3>UnityMSK Executive Summary</h3>
            <p class="section-copy">Week-over-week operational view across LAOSS, NES, SpineOne, and MRO.</p>
          </div>
          <div class="meeting-summary-actions no-print">
            <button id="printMeetingSummaryBtn" class="btn btn-primary" type="button">Print / Save PDF</button>
          </div>
        </div>
      </div>

      <div class="exec-strip print-keep">
        <div class="exec-card"><span class="exec-label">Current Week</span><strong class="exec-value">${escapeHtml(formatDate(data.weekEnding || state.currentWeekEnding))}</strong></div>
        <div class="exec-card"><span class="exec-label">Previous Week</span><strong class="exec-value">${escapeHtml(formatDate(data.previousWeekEnding || ""))}</strong></div>
        <div class="exec-card"><span class="exec-label">Entities Reporting</span><strong class="exec-value">${escapeHtml(String((data.entities || []).length))}</strong></div>
      </div>

      <section class="section-block print-keep" style="margin-bottom:18px;">
        <div class="section-head">
          <h3>Week-over-Week Comparison</h3>
          <p class="section-copy">High-level KPI movement compared with the prior reporting week.</p>
        </div>
        <div class="comparison-grid">
          ${(data.comparison || []).map((item) => `
            <div class="comparison-card">
              <span class="comparison-label">${escapeHtml(item.label)}</span>
              <strong class="comparison-current">${escapeHtml(formatByType(item.current, item.format))}</strong>
              <span class="comparison-meta">Prior: ${escapeHtml(formatByType(item.previous, item.format))}</span>
              <span class="comparison-change ${comparisonClass(item.change, item.key)}">${escapeHtml(formatChange(item.change, item.format))}</span>
            </div>
          `).join("") || `<div class="note-panel"><p>No comparison data yet.</p></div>`}
        </div>
      </section>

      <div class="split-grid print-stack">
        <section class="section-block print-keep">
          <div class="section-head">
            <h3>Region Comparison</h3>
            <p class="section-copy">Current-week region performance snapshot.</p>
          </div>
          <div class="table-wrap">
            <table>
              <thead><tr><th>Entity</th><th>Visit Volume</th><th>Call Volume</th><th>No Show Rate</th><th>Cancellation Rate</th><th>Abandoned Call Rate</th><th>Status</th></tr></thead>
              <tbody>
                ${(data.entities || []).map((row) => `
                  <tr>
                    <td>${escapeHtml(row.entity)}</td>
                    <td>${escapeHtml(String(row.visitVolume ?? "—"))}</td>
                    <td>${escapeHtml(String(row.callVolume ?? "—"))}</td>
                    <td>${escapeHtml(String(row.noShowRate ?? "—"))}</td>
                    <td>${escapeHtml(String(row.cancellationRate ?? "—"))}</td>
                    <td>${escapeHtml(String(row.abandonedCallRate ?? "—"))}</td>
                    <td>${escapeHtml(row.status ?? "Draft")}</td>
                  </tr>
                `).join("") || `<tr><td colspan="7">No regional data yet.</td></tr>`}
              </tbody>
            </table>
          </div>
        </section>

        <section class="section-block print-keep">
          <div class="section-head">
            <h3>Top Risk Metrics</h3>
            <p class="section-copy">Fast view of the riskiest metrics across regions.</p>
          </div>
          <div class="risk-list">
            ${(data.riskMetrics || []).map((risk) => `
              <div class="risk-item">
                <div>
                  <span class="risk-entity">${escapeHtml(risk.entity)}</span>
                  <span class="risk-label">${escapeHtml(risk.label)}</span>
                </div>
                <div class="risk-right">
                  <strong>${escapeHtml(risk.formattedValue)}</strong>
                  <span class="risk-badge ${escapeHtml(risk.statusColor)}">${escapeHtml(risk.statusColor)}</span>
                </div>
              </div>
            `).join("") || `<div class="note-panel"><p>No risk metrics yet.</p></div>`}
          </div>
        </section>
      </div>

      <section class="section-block print-keep" style="margin-top:18px;">
        <div class="section-head">
          <h3>Regional Commentary Rollup</h3>
          <p class="section-copy">Meeting-ready notes pulled from regional submissions.</p>
        </div>
        <div class="commentary-grid">
          ${(data.commentaryRollup || []).map((item) => `
            <div class="commentary-card">
              <h4>${escapeHtml(item.entity)}</h4>
              <p><strong>Commentary:</strong> ${escapeHtml(item.commentary || "—")}</p>
              <p><strong>Blockers:</strong> ${escapeHtml(item.blockers || "—")}</p>
              <p><strong>Opportunities:</strong> ${escapeHtml(item.opportunities || "—")}</p>
            </div>
          `).join("") || `<div class="note-panel"><p>No commentary submitted yet.</p></div>`}
        </div>
      </section>
    `;
  }
}

async function renderRegion(entity) {
  const friendlyName = ENTITY_LABELS[entity] || entity;

  if (els.pageTitle) {
    els.pageTitle.textContent = `${entity} Regional Dashboard`;
  }

  if (els.pageSubtitle) {
    els.pageSubtitle.textContent =
      `Weekly data entry, KPI visibility, narratives, and workflow tracking for ${friendlyName}.`;
  }

  const data = await safeApiGet(
    `/api/weekly?entity=${encodeURIComponent(entity)}&weekEnding=${encodeURIComponent(state.currentWeekEnding)}`,
    {
      entity,
      weekEnding: state.currentWeekEnding,
      status: "Draft",
      kpis: [],
      inputs: {},
      narrative: {
        commentary: "",
        blockers: "",
        opportunities: ""
      }
    }
  );

  state.pageData = data;

  renderKpiCards(data.kpis || []);
  if (els.submissionStatusText) {
    els.submissionStatusText.textContent = data.status || "Draft";
  }

  const isApproved = data.status === "Approved";
  const canEdit = Boolean((state.me?.isAdmin || state.me?.entity === entity) && !isApproved);
  const sections = getRegionSections(entity);
  const activeSection = sections.find((section) => section.key === state.activeRegionSectionKey) || sections[0];
  state.activeRegionSectionKey = activeSection?.key || null;

  const summaries = calculateRegionSummaries(data.inputs || {});
  const calculatedValues = getRegionCalculatedValues(data.inputs || {});

  if (els.pageContent) {
    els.pageContent.innerHTML = `
      <div class="section-head">
        <h3>${friendlyName}</h3>
        <p class="section-copy">This regional view now supports more workbook-style metric groups and read-only derived outputs.</p>
      </div>

      <div class="summary-mini-grid">${summaries.map(renderSummaryMiniCard).join("")}</div>

      <div class="mor-strip">
        <div class="mor-card"><span class="mor-label">Region</span><strong class="mor-value">${escapeHtml(entity)}</strong></div>
        <div class="mor-card"><span class="mor-label">Week Ending</span><strong class="mor-value">${escapeHtml(formatDate(state.currentWeekEnding))}</strong></div>
        <div class="mor-card"><span class="mor-label">Status</span><strong class="mor-value">${escapeHtml(data.status || "Draft")}</strong></div>
      </div>

      ${isApproved ? `<div class="note-panel" style="margin-bottom:18px;"><h4>Week Locked</h4><p>This week has been approved and is now locked for editing.</p></div>` : ""}

      <div class="section-tabs">
        ${sections.map((section) => `
          <button class="section-tab region-tab ${section.key === activeSection.key ? "active" : ""}" data-section-key="${escapeAttr(section.key)}" type="button">
            ${escapeHtml(section.title)}
          </button>
        `).join("")}
      </div>

      ${renderSectionBlock(activeSection, data.inputs || {}, canEdit, calculatedValues)}

      <div class="split-grid" style="margin-top:18px;">
        <div class="note-panel">
          <h4>Regional Commentary</h4>
          <div class="field">
            <label for="commentary">Regional Commentary</label>
            <textarea id="commentary" ${canEdit ? "" : "disabled"}>${escapeHtml(data.narrative?.commentary ?? "")}</textarea>
          </div>
        </div>

        <div class="note-panel">
          <h4>Blockers and Opportunities</h4>
          <div class="field">
            <label for="blockers">Blockers</label>
            <textarea id="blockers" ${canEdit ? "" : "disabled"}>${escapeHtml(data.narrative?.blockers ?? "")}</textarea>
          </div>
          <div class="field">
            <label for="opportunities">Opportunities</label>
            <textarea id="opportunities" ${canEdit ? "" : "disabled"}>${escapeHtml(data.narrative?.opportunities ?? "")}</textarea>
          </div>
        </div>
      </div>
    `;
  }
}

async function renderSharedPage(pageName) {
  const def = getSharedPageDefinition(pageName);

  if (!def) {
    if (els.pageTitle) els.pageTitle.textContent = pageName || "Shared Page";
    if (els.pageSubtitle) els.pageSubtitle.textContent = "Shared page definition not found.";
    if (els.pageContent) {
      els.pageContent.innerHTML = `<div class="note-panel"><h4>Missing Definition</h4><p>No definition exists for this shared page yet.</p></div>`;
    }
    return;
  }

  if (els.pageTitle) els.pageTitle.textContent = def.title;
  if (els.pageSubtitle) els.pageSubtitle.textContent = def.description;

  const data = await safeApiGet(
    `/api/shared?page=${encodeURIComponent(pageName)}&weekEnding=${encodeURIComponent(state.currentWeekEnding)}`,
    {
      page: pageName,
      weekEnding: state.currentWeekEnding,
      status: "Draft",
      kpis: [],
      inputs: {}
    }
  );

  state.pageData = data;

  const activeSection = def.sections.find((section) => section.key === state.activeSharedSectionKey) || def.sections[0];
  state.activeSharedSectionKey = activeSection?.key || null;

  const summaries = calculateSharedSummaries(pageName, data.inputs || {});
  const calculatedValues = getSharedCalculatedValues(pageName, data.inputs || {});

  renderKpiCards((data.kpis || []).length
    ? data.kpis
    : summaries.map((s) => ({
        label: s.label,
        value: s.value,
        meta: s.meta,
        status: "Tracking",
        statusColor: "yellow"
      }))
  );

  if (els.submissionStatusText) {
    els.submissionStatusText.textContent = data.status || "Draft";
  }

  if (els.pageContent) {
    els.pageContent.innerHTML = `
      <div class="section-head">
        <h3>${escapeHtml(def.title)}</h3>
        <p class="section-copy">${escapeHtml(def.description)}</p>
      </div>

      <div class="summary-mini-grid">${summaries.map(renderSummaryMiniCard).join("")}</div>

      <div class="section-tabs">
        ${def.sections.map((section) => `
          <button class="section-tab shared-tab ${section.key === activeSection.key ? "active" : ""}" data-section-key="${escapeAttr(section.key)}" type="button">
            ${escapeHtml(section.title)}
          </button>
        `).join("")}
      </div>

      ${renderSectionBlock(activeSection, data.inputs || {}, true, calculatedValues)}

      <div class="note-panel" style="margin-top:18px;">
        <h4>Shared Page Notes</h4>
        <p>Shared pages now support richer workbook-style groupings and read-only calculated outputs.</p>
      </div>
    `;
  }
}

function renderSectionBlock(section, inputs, canEdit, calculatedValues = {}) {
  return `
    <section class="section-block">
      <div class="section-head">
        <h3>${escapeHtml(section.title)}</h3>
        <p class="section-copy">${escapeHtml(section.description || "")}</p>
      </div>

      <div class="form-grid">
        ${section.fields.map((field) => renderField(field, inputs[field.key], canEdit)).join("")}
      </div>

      ${(section.calculatedFields && section.calculatedFields.length)
        ? `<div class="computed-grid">
            ${section.calculatedFields.map((field) => renderCalculatedField(field, calculatedValues[field.key])).join("")}
          </div>`
        : ""}
    </section>
  `;
}

function renderField(field, value, canEdit) {
  const type = field.type || "text";
  const step = field.step ? `step="${escapeAttr(field.step)}"` : "";
  const placeholder = field.placeholder ? `placeholder="${escapeAttr(field.placeholder)}"` : "";
  const disabled = canEdit ? "" : "disabled";

  return `
    <div class="field">
      <label for="${escapeAttr(field.key)}">${escapeHtml(field.label)}</label>
      <input
        id="${escapeAttr(field.key)}"
        data-metric-key="${escapeAttr(field.key)}"
        type="${escapeAttr(type)}"
        ${step}
        ${placeholder}
        value="${escapeAttr(value ?? "")}"
        ${disabled}
      />
    </div>
  `;
}

function renderCalculatedField(field, value) {
  return `
    <div class="computed-card">
      <span class="computed-label">${escapeHtml(field.label)}</span>
      <strong class="computed-value">${escapeHtml(formatByType(value, field.format || "decimal2"))}</strong>
    </div>
  `;
}

async function renderAdmin() {
  if (els.pageTitle) els.pageTitle.textContent = "Admin";
  if (els.pageSubtitle) {
    els.pageSubtitle.textContent =
      "Manage references, monitor submissions, review audit history, and import workbook data.";
  }

  renderKpiCards([
    { label: "Admin Module", value: "Live", statusColor: "green", meta: "Reference editors enabled", status: "Ready" },
    { label: "Entity", value: state.activeAdminEntity, statusColor: "yellow", meta: "Current selection", status: "Active" },
    { label: "Editor", value: state.activeAdminTab, statusColor: "yellow", meta: "Active admin tab", status: "Tracking" },
    { label: "Week", value: state.currentWeekEnding, statusColor: "green", meta: "Tracking period", status: "Current" }
  ]);

  if (els.submissionStatusText) {
    els.submissionStatusText.textContent = "Admin View";
  }

  if (els.pageContent) {
    els.pageContent.innerHTML = renderAdminEditorShell();
  }

  const adminContent = document.getElementById("adminEditorContent");
  const entityFilter = document.getElementById("adminEntityFilter");
  const yearFilter = document.getElementById("adminYearFilter");
  const auditEntityFilter = document.getElementById("adminAuditEntityFilter");
  const entityWrap = document.getElementById("adminEntityFilterWrap");
  const yearWrap = document.getElementById("adminYearFilterWrap");
  const auditEntityWrap = document.getElementById("adminAuditEntityFilterWrap");

  if (entityFilter) entityFilter.value = state.activeAdminEntity;
  if (yearFilter) yearFilter.value = state.activeAdminYear;
  if (auditEntityFilter) auditEntityFilter.value = state.activeAdminAuditEntity;

  document.querySelectorAll(".admin-editor-tab").forEach((el) => el.classList.remove("active"));
  const activeTab = document.querySelector(`.admin-editor-tab[data-admin-tab="${cssEscape(state.activeAdminTab)}"]`);
  if (activeTab) activeTab.classList.add("active");

  entityWrap?.classList.add("hidden");
  yearWrap?.classList.add("hidden");
  auditEntityWrap?.classList.add("hidden");

  if (state.activeAdminTab === "holidays") {
    yearWrap?.classList.remove("hidden");
    const data = await safeApiGet(`/api/admin-reference?kind=holidays&year=${encodeURIComponent(state.activeAdminYear)}`, { rows: [] });
    if (adminContent) adminContent.innerHTML = renderHolidaysEditor(state.activeAdminYear, data.rows || []);
    return;
  }

  if (state.activeAdminTab === "targets") {
    entityWrap?.classList.remove("hidden");
    const data = await safeApiGet(`/api/admin-reference?entity=${encodeURIComponent(state.activeAdminEntity)}&kind=targets`, { rows: [] });
    if (adminContent) adminContent.innerHTML = renderTargetsEditor(state.activeAdminEntity, data.rows || []);
    return;
  }

  if (state.activeAdminTab === "thresholds") {
    entityWrap?.classList.remove("hidden");
    const data = await safeApiGet(`/api/admin-reference?entity=${encodeURIComponent(state.activeAdminEntity)}&kind=thresholds`, { rows: [] });
    if (adminContent) adminContent.innerHTML = renderThresholdsEditor(state.activeAdminEntity, data.rows || []);
    return;
  }

  if (state.activeAdminTab === "budget") {
    entityWrap?.classList.remove("hidden");
    const data = await safeApiGet(`/api/admin-reference?entity=${encodeURIComponent(state.activeAdminEntity)}&kind=budget`, { rows: [] });
    if (adminContent) adminContent.innerHTML = renderBudgetEditor(state.activeAdminEntity, data.rows || []);
    return;
  }

  if (state.activeAdminTab === "submissions") {
    const data = await safeApiGet(`/api/submissionsfeed?weekEnding=${encodeURIComponent(state.currentWeekEnding)}`, { items: [] });
    const mapped = mapSubmissionFeedToTracker(data);
    if (adminContent) adminContent.innerHTML = renderSubmissionTracker(mapped);
    return;
  }

  if (state.activeAdminTab === "audit") {
    auditEntityWrap?.classList.remove("hidden");
    const qs = new URLSearchParams({ weekEnding: state.currentWeekEnding });
    if (state.activeAdminAuditEntity) qs.set("entity", state.activeAdminAuditEntity);
    const data = await safeApiGet(`/api/admin-audit?${qs.toString()}`, { rows: [] });
    if (adminContent) adminContent.innerHTML = renderAuditViewer(data);
    return;
  }

  if (state.activeAdminTab === "import") {
    if (adminContent) adminContent.innerHTML = renderImportTool(state.lastImportResult);
  }
}

function mapSubmissionFeedToTracker(data) {
  const items = Array.isArray(data?.items) ? data.items : [];
  const knownEntities = ["LAOSS", "NES", "SpineOne", "MRO", "PT", "CXNS", "Capacity", "Productivity Builder"];

  const byEntity = new Map();

  for (const item of items) {
    const entity = item.entityId || item.market || item.location || "Unknown";
    byEntity.set(entity, {
      entity,
      weekEnding: item.weekEnding || state.currentWeekEnding,
      status: item.status || "Draft",
      updatedBy: item.submittedByEmail || item.submittedBy || "—",
      updatedAt: item.updatedAt || item.submittedAt || "—",
      submittedBy: item.submittedBy || "—",
      submittedAt: item.submittedAt || "—",
      approvedBy: "—",
      approvedAt: "—",
      inputCount: item.payload && typeof item.payload === "object" ? Object.keys(item.payload).length : 0,
      hasNarrative: Boolean(item.payload?.narrative)
    });
  }

  const rows = knownEntities.map((entity) => {
    return byEntity.get(entity) || {
      entity,
      weekEnding: state.currentWeekEnding,
      status: "Missing",
      updatedBy: "—",
      updatedAt: "—",
      submittedBy: "—",
      submittedAt: "—",
      approvedBy: "—",
      approvedAt: "—",
      inputCount: 0,
      hasNarrative: false
    };
  });

  const submittedCount = rows.filter((r) => ["Submitted", "Approved", "Draft"].includes(r.status) && r.status !== "Missing").length;
  const missingEntities = rows.filter((r) => r.status === "Missing").map((r) => r.entity);

  return {
    weekEnding: state.currentWeekEnding,
    summary: {
      totalEntities: rows.length,
      submittedCount,
      missingCount: missingEntities.length,
      missingEntities
    },
    rows
  };
}

async function fileToBase64(file) {
  const arrayBuffer = await file.arrayBuffer();
  let binary = "";
  const bytes = new Uint8Array(arrayBuffer);
  const chunkSize = 0x8000;

  for (let i = 0; i < bytes.length; i += chunkSize) {
    const chunk = bytes.subarray(i, i + chunkSize);
    binary += String.fromCharCode.apply(null, chunk);
  }

  return btoa(binary);
}

function renderKpiCards(kpis) {
  const safe = Array.isArray(kpis) && kpis.length
    ? kpis
    : [
        { label: "Visit Volume", value: "
