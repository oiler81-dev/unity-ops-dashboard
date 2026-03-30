import {
  safeApiGet,
  getDefaultWeekEnding,
  formatDate,
  renderKpiCards,
  handleFatalError
} from "./helpers.js";

/* ---------------- STATE ---------------- */

const state = {
  weekEnding: getDefaultWeekEnding(),
  compareMode: "priorPeriod",
  periodType: "currentWeek"
};

/* ---------------- HELPERS ---------------- */

function $(id) {
  return document.getElementById(id);
}

function getKpiClass(change) {
  if (change > 0) return "kpi-positive";
  if (change < 0) return "kpi-negative";
  return "kpi-neutral";
}

/* ---------------- INIT ---------------- */

document.addEventListener("DOMContentLoaded", init);

async function init() {
  try {
    initControls();
    initQuickPresets();
    await loadDashboard();
  } catch (err) {
    handleFatalError(err);
  }
}

/* ---------------- CONTROLS ---------------- */

function initControls() {
  $("loadDashboardBtn").addEventListener("click", loadDashboard);

  $("dashboardCompareAgainst").addEventListener("change", (e) => {
    state.compareMode = e.target.value;
  });

  $("dashboardPeriodType").addEventListener("change", (e) => {
    state.periodType = e.target.value;
  });
}

/* ---------------- QUICK PRESETS ---------------- */

function initQuickPresets() {
  document.querySelectorAll(".quickPresetPill").forEach((pill) => {
    pill.addEventListener("click", async () => {
      const text = pill.innerText.toLowerCase();

      if (text.includes("current")) state.periodType = "currentWeek";
      if (text.includes("last week")) state.periodType = "lastWeek";
      if (text.includes("mtd")) state.periodType = "mtd";
      if (text.includes("rolling")) state.periodType = "rolling4";

      $("dashboardPeriodType").value = state.periodType;

      await loadDashboard();
    });
  });
}

/* ---------------- DASHBOARD LOAD ---------------- */

async function loadDashboard() {
  const week = $("dashboardWeekEnding").value || state.weekEnding;

  $("dashboardSummaryText").innerText = "Loading dashboard...";

  const data = await safeApiGet(
    `/api/executive-summary?weekEnding=${week}`,
    {}
  );

  renderTopCards(data);
  renderEntityCards(data.regions || []);
  renderCharts(data);
  renderWins(data.regions || []);

  $("dashboardSummaryText").innerText = `Viewing ${formatDate(week)}`;
}

/* ---------------- TOP KPI CARDS ---------------- */

function renderTopCards(data) {
  const totals = data.totals || {};
  const budget = data.budgetTotals || {};
  const variance = data.variances || {};

  const cards = [
    {
      label: "Visit Volume",
      value: totals.visitVolume || 0,
      change: variance.visitVolumeVariance || 0
    },
    {
      label: "Call Volume",
      value: totals.callVolume || 0,
      change: 0
    },
    {
      label: "New Patients",
      value: totals.newPatients || 0,
      change: variance.newPatientsVariance || 0
    }
  ];

  const container = $("dashboardCards");

  container.innerHTML = cards
    .map(
      (c) => `
      <div class="summaryCard ${getKpiClass(c.change)}">
        <h3>${c.label}</h3>
        <div class="value">${c.value.toLocaleString()}</div>
        <div>${c.change >= 0 ? "+" : ""}${c.change}</div>
      </div>
    `
    )
    .join("");
}

/* ---------------- ENTITY CARDS ---------------- */

function renderEntityCards(regions) {
  const container = $("dashboardEntities");

  container.innerHTML = regions
    .map((r) => {
      const variance =
        (r.visitVolume || 0) - (r.budget?.visitVolumeBudget || 0);

      return `
      <div class="entityCard ${getKpiClass(variance)}">
        <h3>${r.entity}</h3>
        <div>Visits: ${r.visitVolume}</div>
        <div>NP: ${r.newPatients}</div>
        <div>Calls: ${r.callVolume}</div>
        <div>Variance: ${variance}</div>
      </div>
    `;
    })
    .join("");
}

/* ---------------- CHARTS ---------------- */

function renderCharts(data) {
  renderBudgetChart(data.regions || []);
  renderTrendChart(data.regions || []);
}

function renderBudgetChart(regions) {
  const canvas = document.getElementById("budgetChart");
  if (!canvas) return;

  const ctx = canvas.getContext("2d");

  const labels = regions.map((r) => r.entity);
  const actual = regions.map((r) => r.visitVolume || 0);
  const budget = regions.map((r) => r.budget?.visitVolumeBudget || 0);

  if (window.budgetChart) window.budgetChart.destroy();

  window.budgetChart = new Chart(ctx, {
    type: "bar",
    data: {
      labels,
      datasets: [
        { label: "Actual", data: actual },
        { label: "Budget", data: budget }
      ]
    }
  });
}

function renderTrendChart(regions) {
  const canvas = document.getElementById("trendChart");
  if (!canvas) return;

  const ctx = canvas.getContext("2d");

  const labels = regions.map((r) => r.entity);
  const visits = regions.map((r) => r.visitVolume || 0);

  if (window.trendChart) window.trendChart.destroy();

  window.trendChart = new Chart(ctx, {
    type: "line",
    data: {
      labels,
      datasets: [
        {
          label: "Visits",
          data: visits
        }
      ]
    }
  });
}

/* ---------------- WINS SECTION ---------------- */

function renderWins(regions) {
  const container = $("dashboardAlerts");

  const wins = regions.filter(
    (r) => r.visitVolume > (r.budget?.visitVolumeBudget || 0)
  );

  container.innerHTML = `
    <h4>Wins</h4>
    ${wins
      .map(
        (w) => `
      <div class="winItem">
        ${w.entity} above budget
      </div>
    `
      )
      .join("")}
  `;
}
