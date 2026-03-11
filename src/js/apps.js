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
  activeAdminYear: String(new Date().getFullYear())
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
      return;
    }

    const adminTab = e.target.closest(".admin-editor-tab");
    if (adminTab) {
      state.activeAdminTab = adminTab.dataset.adminTab;
      if (state.currentRoute === "admin") await renderAdmin();
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

/* keep your current renderExecutive, renderRegion, renderSharedPage, and helper functions exactly as-is from the last working version */
/* only replace renderAdmin below */

async function renderAdmin() {
  els.pageTitle.textContent = "Admin";
  els.pageSubtitle.textContent = "Manage targets, thresholds, holidays, and budget reference data.";

  renderKpiCards([
    { label: "Admin Module", value: "Live", statusColor: "green", meta: "Reference editors enabled" },
    { label: "Entity", value: state.activeAdminEntity, statusColor: "yellow", meta: "Current selection" },
    { label: "Editor", value: state.activeAdminTab, statusColor: "yellow", meta: "Active admin tab" },
    { label: "Reference Areas", value: "4", statusColor: "green", meta: "Targets, thresholds, holidays, budget" }
  ]);

  els.submissionStatusText.textContent = "Admin View";
  els.pageContent.innerHTML = renderAdminEditorShell();

  const adminContent = document.getElementById("adminEditorContent");
  const entityFilter = document.getElementById("adminEntityFilter");
  const yearFilter = document.getElementById("adminYearFilter");
  const entityWrap = document.getElementById("adminEntityFilterWrap");
  const yearWrap = document.getElementById("adminYearFilterWrap");

  if (entityFilter) entityFilter.value = state.activeAdminEntity;
  if (yearFilter) yearFilter.value = state.activeAdminYear;

  document.querySelectorAll(".admin-editor-tab").forEach((el) => el.classList.remove("active"));
  const activeTab = document.querySelector(`.admin-editor-tab[data-admin-tab="${state.activeAdminTab}"]`);
  if (activeTab) activeTab.classList.add("active");

  if (state.activeAdminTab === "holidays") {
    entityWrap?.classList.add("hidden");
    yearWrap?.classList.remove("hidden");
    const data = await apiGet(`/api/admin-reference?kind=holidays&year=${encodeURIComponent(state.activeAdminYear)}`);
    adminContent.innerHTML = renderHolidaysEditor(state.activeAdminYear, data.rows || []);
    return;
  }

  yearWrap?.classList.add("hidden");
  entityWrap?.classList.remove("hidden");

  if (state.activeAdminTab === "targets") {
    const data = await apiGet(`/api/admin-reference?entity=${encodeURIComponent(state.activeAdminEntity)}&kind=targets`);
    adminContent.innerHTML = renderTargetsEditor(state.activeAdminEntity, data.rows || []);
    return;
  }

  if (state.activeAdminTab === "thresholds") {
    const data = await apiGet(`/api/admin-reference?entity=${encodeURIComponent(state.activeAdminEntity)}&kind=thresholds`);
    adminContent.innerHTML = renderThresholdsEditor(state.activeAdminEntity, data.rows || []);
    return;
  }

  if (state.activeAdminTab === "budget") {
    const data = await apiGet(`/api/admin-reference?entity=${encodeURIComponent(state.activeAdminEntity)}&kind=budget`);
    adminContent.innerHTML = renderBudgetEditor(state.activeAdminEntity, data.rows || []);
  }
}

/* KEEP THE REST OF YOUR EXISTING FUNCTIONS BELOW THIS LINE FROM YOUR LAST WORKING FILE:
   renderExecutive
   renderRegion
   renderSharedPage
   renderSectionBlock
   renderField
   renderCalculatedField
   renderKpiCards
   renderSummaryMiniCard
   comparisonClass
   formatChange
   collectRegionFormValues
   collectSharedFormValues
   apiGet
   apiPost
   getDefaultWeekEnding
   formatDate
   numberOrNull
   escapeHtml
   escapeAttr
   handleFatalError
*/