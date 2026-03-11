import {
  getRegionSections,
  getAllMetricKeysForEntity,
  getSharedPageDefinition,
  getAllMetricKeysForSharedPage
} from "./definitions.js";
import { calculateRegionSummaries, calculateSharedSummaries } from "./calculations.js";

const state = {
  me: null,
  currentRoute: "executive",
  currentEntity: null,
  currentSharedPage: null,
  currentWeekEnding: getDefaultWeekEnding(),
  pageData: null,
  activeRegionSectionKey: null,
  activeSharedSectionKey: null
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

init().catch(handleFatalError);

async function init() {
  els.weekEndingSelect.value = state.currentWeekEnding;
  els.selectedWeekText.textContent = formatDate(state.currentWeekEnding);

  bindEvents();

  const me = await apiGet("/api/me");
  state.me = me;

  applyUserContext();

  if (me.isAdmin) {
    setRoute("executive");
  } else {
    setRoute("region", me.entity);
  }
}

function bindEvents() {
  els.weekEndingSelect.addEventListener("change", async (e) => {
    state.currentWeekEnding = e.target.value;
    els.selectedWeekText.textContent = formatDate(state.currentWeekEnding);
    await loadCurrentRoute();
  });

  els.sidebarNav.addEventListener("click", async (e) => {
    const btn = e.target.closest(".nav-link");
    if (!btn) return;

    const route = btn.dataset.route;
    const entity = btn.dataset.entity || null;
    const page = btn.dataset.page || null;

    if (route === "region") return setRoute("region", entity);
    if (route === "shared") return setRoute("shared", null, page);
    if (route === "admin") return setRoute("admin");

    return setRoute("executive");
  });

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
    }
  });

  els.goAssignedRegionBtn.addEventListener("click", () => {
    if (!state.me) return;
    if (state.me.isAdmin) return setRoute("executive");
    return setRoute("region", state.me.entity);
  });

  els.goExecutiveBtn.addEventListener("click", () => {
    setRoute("executive");
  });

  els.saveBtn.addEventListener("click", async () => {
    if (state.currentRoute === "region") {
      const payload = collectRegionFormValues();
      const res = await apiPost("/api/weekly-save", payload);
      els.submissionStatusText.textContent = res.status || "Draft";
      alert("Saved successfully.");
      await loadCurrentRoute();
      return;
    }

    if (state.currentRoute === "shared") {
      const payload = collectSharedFormValues();
      const res = await apiPost("/api/shared-save", payload);
      els.submissionStatusText.textContent = res.status || "Draft";
      alert("Shared page saved successfully.");
      await loadCurrentRoute();
      return;
    }

    alert("Save is only active on region and shared pages.");
  });

  els.submitBtn.addEventListener("click", async () => {
    if (state.currentRoute !== "region") {
      alert("Submit is only active on region pages.");
      return;
    }

    const res = await apiPost("/api/submit-week", {
      entity: state.currentEntity,
      weekEnding: state.currentWeekEnding
    });

    els.submissionStatusText.textContent = res.status || "Submitted";
    alert("Week submitted.");
    await loadCurrentRoute();
  });
}

function applyUserContext() {
  els.userDisplayName.textContent = state.me.displayName || state.me.email || "Unknown User";
  els.assignedEntityText.textContent = state.me.entity || "None";
  els.roleText.textContent = state.me.isAdmin ? "Admin" : "Editor";

  if (state.me.isAdmin) {
    els.adminNavBtn.classList.remove("hidden");
  } else {
    els.adminNavBtn.classList.add("hidden");
  }
}

async function setRoute(route, entity = null, sharedPage = null) {
  state.currentRoute = route;
  state.currentEntity = entity;
  state.currentSharedPage = sharedPage;

  document.querySelectorAll(".nav-link").forEach((btn) => btn.classList.remove("active"));

  const selector = route === "region"
    ? `.nav-link[data-route="region"][data-entity="${entity}"]`
    : route === "shared"
      ? `.nav-link[data-route="shared"][data-page="${sharedPage}"]`
      : `.nav-link[data-route="${route}"]`;

  const activeBtn = document.querySelector(selector);
  if (activeBtn) activeBtn.classList.add("active");

  if (route === "region") {
    const sections = getRegionSections(entity);
    state.activeRegionSectionKey = sections[0]?.key || null;
  }

  if (route === "shared") {
    const def = getSharedPageDefinition(sharedPage);
    state.activeSharedSectionKey = def?.sections?.[0]?.key || null;
  }

  await loadCurrentRoute();
}

async function loadCurrentRoute() {
  if (state.currentRoute === "executive") return renderExecutive();
  if (state.currentRoute === "region") return renderRegion(state.currentEntity);
  if (state.currentRoute === "shared") return renderSharedPage(state.currentSharedPage);
  if (state.currentRoute === "admin") return renderAdmin();
}

async function renderExecutive() {
  els.pageTitle.textContent = "Executive Summary";
  els.pageSubtitle.textContent = "Weekly companywide KPI overview with trends, target comparisons, and submission visibility.";

  const data = await apiGet(`/api/dashboard?weekEnding=${encodeURIComponent(state.currentWeekEnding)}`);
  state.pageData = data;

  renderKpiCards(data.kpis || []);
  els.submissionStatusText.textContent = "Summary View";

  els.pageContent.innerHTML = `
    <div class="section-head">
      <h3>Companywide Overview</h3>
      <p class="section-copy">Live summary across LAOSS, NES, SpineOne, and MRO based on saved weekly inputs.</p>
    </div>

    <div class="split-grid">
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>Entity</th>
              <th>Visit Volume</th>
              <th>Call Volume</th>
              <th>No Show Rate</th>
              <th>Cancellation Rate</th>
              <th>Abandoned Call Rate</th>
              <th>Status</th>
            </tr>
          </thead>
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
            `).join("")}
          </tbody>
        </table>
      </div>

      <div class="note-panel">
        <h4>Executive Notes</h4>
        <p>
          KPI cards are now threshold-ready. As you seed the reference tables, card colors and target variance will update automatically.
        </p>
      </div>
    </div>
  `;
}

async function renderRegion(entity) {
  els.pageTitle.textContent = `${entity} Regional Dashboard`;
  els.pageSubtitle.textContent = `Weekly data entry, KPI visibility, narratives, and workflow tracking for ${entity}.`;

  const data = await apiGet(`/api/weekly?entity=${encodeURIComponent(entity)}&weekEnding=${encodeURIComponent(state.currentWeekEnding)}`);
  state.pageData = data;

  renderKpiCards(data.kpis || []);
  els.submissionStatusText.textContent = data.status || "Draft";

  const canEdit = Boolean(state.me?.isAdmin || state.me?.entity === entity);
  const sections = getRegionSections(entity);
  const activeSection = sections.find((section) => section.key === state.activeRegionSectionKey) || sections[0];
  state.activeRegionSectionKey = activeSection?.key || null;

  const summaries = calculateRegionSummaries(data.inputs || {});

  els.pageContent.innerHTML = `
    <div class="section-head">
      <h3>${entity} Weekly Inputs</h3>
      <p class="section-copy">
        This page is definition-driven and tabbed. New fields can be added centrally without rewriting the whole page.
      </p>
    </div>

    <div class="summary-mini-grid">
      ${summaries.map(renderSummaryMiniCard).join("")}
    </div>

    <div class="section-tabs">
      ${sections.map((section) => `
        <button
          class="section-tab region-tab ${section.key === activeSection.key ? "active" : ""}"
          data-section-key="${escapeAttr(section.key)}"
        >
          ${escapeHtml(section.title)}
        </button>
      `).join("")}
    </div>

    ${renderSectionBlock(activeSection, data.inputs || {}, canEdit)}

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

async function renderSharedPage(pageName) {
  const def = getSharedPageDefinition(pageName);
  if (!def) {
    els.pageTitle.textContent = pageName;
    els.pageSubtitle.textContent = "Shared page definition not found.";
    els.pageContent.innerHTML = `<div class="note-panel"><h4>Missing Definition</h4><p>No definition exists for this shared page yet.</p></div>`;
    return;
  }

  els.pageTitle.textContent = def.title;
  els.pageSubtitle.textContent = def.description;

  const data = await apiGet(`/api/shared?page=${encodeURIComponent(pageName)}&weekEnding=${encodeURIComponent(state.currentWeekEnding)}`);
  state.pageData = data;

  const activeSection = def.sections.find((section) => section.key === state.activeSharedSectionKey) || def.sections[0];
  state.activeSharedSectionKey = activeSection?.key || null;

  const summaries = calculateSharedSummaries(pageName, data.inputs || {});
  renderKpiCards((data.kpis || []).length ? data.kpis : summaries.map((s) => ({
    label: s.label,
    value: s.value,
    meta: s.meta,
    status: "Tracking",
    statusColor: "yellow"
  })));

  els.submissionStatusText.textContent = data.status || "Draft";

  els.pageContent.innerHTML = `
    <div class="section-head">
      <h3>${escapeHtml(def.title)}</h3>
      <p class="section-copy">${escapeHtml(def.description)}</p>
    </div>

    <div class="summary-mini-grid">
      ${summaries.map(renderSummaryMiniCard).join("")}
    </div>

    <div class="section-tabs">
      ${def.sections.map((section) => `
        <button
          class="section-tab shared-tab ${section.key === activeSection.key ? "active" : ""}"
          data-section-key="${escapeAttr(section.key)}"
        >
          ${escapeHtml(section.title)}
        </button>
      `).join("")}
    </div>

    ${renderSectionBlock(activeSection, data.inputs || {}, true)}

    <div class="note-panel" style="margin-top:18px;">
      <h4>Shared Page Notes</h4>
      <p>
        Shared pages are now ready to use threshold and target tables as soon as those rows are seeded.
      </p>
    </div>
  `;
}

function renderSectionBlock(section, inputs, canEdit) {
  return `
    <section class="section-block">
      <div class="section-head">
        <h3>${escapeHtml(section.title)}</h3>
        <p class="section-copy">${escapeHtml(section.description || "")}</p>
      </div>

      <div class="form-grid">
        ${section.fields.map((field) => renderField(field, inputs[field.key], canEdit)).join("")}
      </div>
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

async function renderAdmin() {
  els.pageTitle.textContent = "Admin";
  els.pageSubtitle.textContent = "User access, thresholds, holidays, targets, budget references, and missing submissions.";

  renderKpiCards([
    { label: "Active Users", value: "6", statusColor: "green", meta: "Mapped users" },
    { label: "Managed Entities", value: "4", statusColor: "green", meta: "Regional dashboards" },
    { label: "Threshold Sets", value: "Ready", statusColor: "yellow", meta: "To be configured" },
    { label: "Missing Submissions", value: "0", statusColor: "green", meta: "Sample" }
  ]);

  els.submissionStatusText.textContent = "Admin View";

  const data = await apiGet("/api/admin-users");

  els.pageContent.innerHTML = `
    <div class="section-head">
      <h3>Admin Controls</h3>
      <p class="section-copy">Final admin area structure. Wire forms and inline editing next.</p>
    </div>

    <div class="table-wrap">
      <table>
        <thead>
          <tr>
            <th>User</th>
            <th>Email</th>
            <th>Role</th>
            <th>Entity</th>
          </tr>
        </thead>
        <tbody>
          ${(data.users || []).map((u) => `
            <tr>
              <td>${escapeHtml(u.displayName)}</td>
              <td>${escapeHtml(u.email)}</td>
              <td>${escapeHtml(u.role)}</td>
              <td>${escapeHtml(u.entity)}</td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    </div>
  `;
}

function renderKpiCards(kpis) {
  const safe = Array.isArray(kpis) && kpis.length
    ? kpis
    : [
        { label: "Visit Volume", value: "—", statusColor: "yellow", meta: "No data yet" },
        { label: "Call Volume", value: "—", statusColor: "yellow", meta: "No data yet" },
        { label: "No Show Rate", value: "—", statusColor: "yellow", meta: "No data yet" },
        { label: "Abandoned Call Rate", value: "—", statusColor: "yellow", meta: "No data yet" }
      ];

  els.dashboardCards.innerHTML = safe.map((kpi, i) => `
    <div class="dashboard-card ${i === 0 ? "highlight" : ""}">
      <span class="card-label">${escapeHtml(kpi.label || "KPI")}</span>
      <h3>${escapeHtml(kpi.title || kpi.label || "Metric")}</h3>
      <div class="kpi-value">${escapeHtml(String(kpi.value ?? "—"))}</div>
      <div class="kpi-meta">${escapeHtml(kpi.meta || "")}</div>
      <div class="kpi-status ${escapeHtml(kpi.statusColor || "yellow")}">${escapeHtml(kpi.status || "Tracking")}</div>
    </div>
  `).join("");
}

function renderSummaryMiniCard(item) {
  return `
    <div class="summary-mini-card">
      <span class="summary-mini-label">${escapeHtml(item.label)}</span>
      <strong class="summary-mini-value">${escapeHtml(String(item.value))}</strong>
      <span class="summary-mini-meta">${escapeHtml(item.meta || "")}</span>
    </div>
  `;
}

function collectRegionFormValues() {
  const metricKeys = getAllMetricKeysForEntity(state.currentEntity);
  const inputs = {};

  metricKeys.forEach((key) => {
    const el = document.querySelector(`[data-metric-key="${key}"]`);
    inputs[key] = numberOrNull(el?.value);
  });

  return {
    entity: state.currentEntity,
    weekEnding: state.currentWeekEnding,
    inputs,
    narrative: {
      commentary: document.getElementById("commentary")?.value?.trim() || "",
      blockers: document.getElementById("blockers")?.value?.trim() || "",
      opportunities: document.getElementById("opportunities")?.value?.trim() || ""
    }
  };
}

function collectSharedFormValues() {
  const metricKeys = getAllMetricKeysForSharedPage(state.currentSharedPage);
  const inputs = {};

  metricKeys.forEach((key) => {
    const el = document.querySelector(`[data-metric-key="${key}"]`);
    inputs[key] = numberOrNull(el?.value);
  });

  return {
    page: state.currentSharedPage,
    weekEnding: state.currentWeekEnding,
    inputs
  };
}

async function apiGet(url) {
  const res = await fetch(url, { headers: { Accept: "application/json" } });
  if (!res.ok) throw new Error(`GET ${url} failed with status ${res.status}`);
  return res.json();
}

async function apiPost(url, body) {
  const res = await fetch(url, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Accept: "application/json"
    },
    body: JSON.stringify(body)
  });

  if (!res.ok) {
    const text = await res.text();
    throw new Error(`POST ${url} failed: ${text}`);
  }

  return res.json();
}

function getDefaultWeekEnding() {
  return new Date().toISOString().slice(0, 10);
}

function formatDate(isoDate) {
  if (!isoDate) return "Not selected";
  return new Date(`${isoDate}T12:00:00`).toLocaleDateString();
}

function numberOrNull(v) {
  if (v === undefined || v === null || v === "") return null;
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function escapeAttr(value) {
  return escapeHtml(value ?? "");
}

function handleFatalError(err) {
  console.error(err);
  els.pageContent.innerHTML = `
    <div class="note-panel">
      <h4>Application Error</h4>
      <p>${escapeHtml(err.message || "Unknown error")}</p>
    </div>
  `;
}
