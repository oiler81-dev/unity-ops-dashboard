import {
  safeApiGet,
  apiPost,
  getDefaultWeekEnding,
  formatDate,
  escapeHtml,
  escapeAttr,
  handleFatalError,
  renderSummaryMiniCard,
  collectRegionFormValues,
  collectSharedFormValues,
  renderKpiCards
} from "./helpers.js";

import {
  getRegionSections,
  getSharedPageDefinition,
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
  goExecutiveBtn: document.getElementById("goExecutiveBtn"),
  authBtn: document.getElementById("authBtn")
};

bootstrap();

function bootstrap() {
  if (document.readyState === "loading") {
    document.addEventListener(
      "DOMContentLoaded",
      () => {
        init().catch(handleFatalError);
      },
      { once: true }
    );
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
    roles: ["anonymous"],
    entity: null,
    isAdmin: false
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
  const isAdmin = Boolean(raw?.isAdmin) || roles.includes("admin");
  const inferredEntity = isAdmin ? null : (raw?.entity || inferEntityFromEmail(userDetails));

  return {
    authenticated: Boolean(raw
