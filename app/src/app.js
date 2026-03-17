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

// ... remainder of your app.js unchanged (renderExecutive, renderRegion, renderSharedPage, renderSectionBlock, renderField, renderCalculatedField, renderAdmin, etc.)
