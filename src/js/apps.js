const state = {
  me: null,
  currentRoute: "executive",
  currentEntity: null,
  currentSharedPage: null,
  currentWeekEnding: getDefaultWeekEnding(),
  pageData: null
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

    if (route === "region") {
      setRoute("region", entity);
      return;
    }

    if (route === "shared") {
      setRoute("shared", null, page);
      return;
    }

    if (route === "admin") {
      setRoute("admin");
      return;
    }

    setRoute("executive");
  });

  els.goAssignedRegionBtn.addEventListener("click", () => {
    if (!state.me) return;
    if (state.me.isAdmin) {
      setRoute("executive");
    } else {
      setRoute("region", state.me.entity);
    }
  });

  els.goExecutiveBtn.addEventListener("click", () => {
    setRoute("executive");
  });

  els.saveBtn.addEventListener("click", async () => {
    if (state.currentRoute !== "region") {
      alert("Save is only active on region pages right now.");
      return;
    }

    const payload = collectRegionFormValues();
    const res = await apiPost("/api/weekly-save", payload);
    els.submissionStatusText.textContent = res.status || "Draft";
    alert("Saved successfully.");
    await loadCurrentRoute();
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
      <p class="section-copy">
        Meeting-ready summary across LAOSS, NES, SpineOne, and MRO.
      </p>
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
                <td>${escapeHtml(String(row.visitVolume ?? "-"))}</td>
                <td>${escapeHtml(String(row.callVolume ?? "-"))}</td>
                <td>${escapeHtml(String(row.noShowRate ?? "-"))}</td>
                <td>${escapeHtml(String(row.cancellationRate ?? "-"))}</td>
                <td>${escapeHtml(String(row.abandonedCallRate ?? "-"))}</td>
                <td>${escapeHtml(row.status ?? "Draft")}</td>
              </tr>
            `).join("")}
          </tbody>
        </table>
      </div>

      <div class="note-panel">
        <h4>Executive Notes</h4>
        <p>
          This foundation is ready for workbook-driven metrics. The structure is final.
          The formulas and tab-specific logic plug in next.
        </p>
      </div>
    </div>
  `;
}

async function renderRegion(entity) {
  els.pageTitle.textContent = `${entity} Regional Dashboard`;
  els.pageSubtitle.textContent = `Weekly data entry, KPI visibility, narrative inputs, and audit-ready tracking for ${entity}.`;

  const data = await apiGet(`/api/weekly?entity=${encodeURIComponent(entity)}&weekEnding=${encodeURIComponent(state.currentWeekEnding)}`);
  state.pageData = data;

  renderKpiCards(data.kpis || []);
  els.submissionStatusText.textContent = data.status || "Draft";

  const canEdit = Boolean(state.me?.isAdmin || state.me?.entity === entity);

  els.pageContent.innerHTML = `
    <div class="section-head">
      <h3>${entity} Weekly Inputs</h3>
      <p class="section-copy">
        Production-shaped starter form. Replace these fields with workbook-mapped fields.
      </p>
    </div>

    <div class="split-grid">
      <div class="section-block" style="padding:18px;">
        <div class="form-grid">
          <div class="field">
            <label for="visitVolume">Visit Volume</label>
            <input id="visitVolume" type="number" value="${escapeAttr(data.inputs?.visitVolume ?? "")}" ${canEdit ? "" : "disabled"} />
          </div>

          <div class="field">
            <label for="callVolume">Call Volume</label>
            <input id="callVolume" type="number" value="${escapeAttr(data.inputs?.callVolume ?? "")}" ${canEdit ? "" : "disabled"} />
          </div>

          <div class="field">
            <label for="noShowRate">No Show Rate (%)</label>
            <input id="noShowRate" type="number" step="0.01" value="${escapeAttr(data.inputs?.noShowRate ?? "")}" ${canEdit ? "" : "disabled"} />
          </div>

          <div class="field">
            <label for="cancellationRate">Cancellation Rate (%)</label>
            <input id="cancellationRate" type="number" step="0.01" value="${escapeAttr(data.inputs?.cancellationRate ?? "")}" ${canEdit ? "" : "disabled"} />
          </div>

          <div class="field">
            <label for="abandonedCallRate">Abandoned Call Rate (%)</label>
            <input id="abandonedCallRate" type="number" step="0.01" value="${escapeAttr(data.inputs?.abandonedCallRate ?? "")}" ${canEdit ? "" : "disabled"} />
          </div>

          <div class="field">
            <label for="newPatients">New Patients</label>
            <input id="newPatients" type="number" value="${escapeAttr(data.inputs?.newPatients ?? "")}" ${canEdit ? "" : "disabled"} />
          </div>
        </div>
      </div>

      <div class="note-panel">
        <h4>Weekly Commentary</h4>
        <div class="field">
          <label for="commentary">Regional Commentary</label>
          <textarea id="commentary" ${canEdit ? "" : "disabled"}>${escapeHtml(data.narrative?.commentary ?? "")}</textarea>
        </div>
        <div class="field">
          <label for="blockers">Blockers</label>
          <textarea id="blockers" ${canEdit ? "" : "disabled"}>${escapeHtml(data.narrative?.blockers ?? "")}</textarea>
        </div>
      </div>
    </div>
  `;
}

function renderSharedPage(pageName) {
  els.pageTitle.textContent = pageName;
  els.pageSubtitle.textContent = `${pageName} shared KPI section scaffold.`;

  renderKpiCards([
    { label: `${pageName} KPI 1`, value: "—", statusColor: "yellow", meta: "Awaiting workbook mapping" },
    { label: `${pageName} KPI 2`, value: "—", statusColor: "yellow", meta: "Awaiting workbook mapping" },
    { label: `${pageName} KPI 3`, value: "—", statusColor: "yellow", meta: "Awaiting workbook mapping" },
    { label: `${pageName} KPI 4`, value: "—", statusColor: "yellow", meta: "Awaiting workbook mapping" }
  ]);

  els.submissionStatusText.textContent = "Shared Section";

  els.pageContent.innerHTML = `
    <div class="section-head">
      <h3>${pageName}</h3>
      <p class="section-copy">
        This section is scaffolded and ready for workbook logic and inputs to be mapped.
      </p>
    </div>
    <div class="note-panel">
      <h4>Workbook Mapping</h4>
      <p>
        Your workbook already has this page. The next step is translating its cells and formulas into stable metric keys and calculation functions.
      </p>
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
      <p class="section-copy">
        Final admin area structure. Wire forms and inline editing next.
      </p>
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

function collectRegionFormValues() {
  return {
    entity: state.currentEntity,
    weekEnding: state.currentWeekEnding,
    inputs: {
      visitVolume: numberOrNull(document.getElementById("visitVolume")?.value),
      callVolume: numberOrNull(document.getElementById("callVolume")?.value),
      noShowRate: numberOrNull(document.getElementById("noShowRate")?.value),
      cancellationRate: numberOrNull(document.getElementById("cancellationRate")?.value),
      abandonedCallRate: numberOrNull(document.getElementById("abandonedCallRate")?.value),
      newPatients: numberOrNull(document.getElementById("newPatients")?.value)
    },
    narrative: {
      commentary: document.getElementById("commentary")?.value?.trim() || "",
      blockers: document.getElementById("blockers")?.value?.trim() || ""
    }
  };
}

async function apiGet(url) {
  const res = await fetch(url, { headers: { "Accept": "application/json" } });
  if (!res.ok) throw new Error(`GET ${url} failed with status ${res.status}`);
  return res.json();
}

async function apiPost(url, body) {
  const res = await fetch(url, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "Accept": "application/json"
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
  const now = new Date();
  return now.toISOString().slice(0, 10);
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
