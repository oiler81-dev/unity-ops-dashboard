import {
  safeApiGet,
  getDefaultWeekEnding,
  formatDate,
  renderKpiCards,
  handleFatalError
} from "./helpers.js";

/* ---------------- STATE ---------------- */

const state = {
  weekEnding: getDefaultWeekEnding()
};

/* ---------------- INIT ---------------- */

document.addEventListener("DOMContentLoaded", init);

async function init() {
  try {
    initControls();
    await loadDashboard();
  } catch (err) {
    handleFatalError(err);
  }
}

/* ---------------- HELPERS ---------------- */

function $(id) {
  return document.getElementById(id);
}

function getKpiClass(change) {
  if (change > 0) return "kpi-positive";
  if (change < 0) return "kpi-negative";
  return "kpi-neutral";
}

/* ---------------- CONTROLS ---------------- */

function initControls() {
  const btn = $("loadDashboardBtn");
  if (btn) {
    btn.addEventListener("click", loadDashboard);
  }
}

/* ---------------- DASHBOARD ---------------- */

async function loadDashboard() {
  try {
    const week =
      $("dashboardWeekEnding")?.value || state.weekEnding;

    const data = await safeApiGet(
      `/api/executive-summary?weekEnding=${week}`,
      {}
    );

    renderTopCards(data);
    renderEntities(data.regions || []);
    renderWins(data.regions || []);

    const banner = $("dashboardSummaryText");
    if (banner) {
      banner.innerText = `Viewing ${formatDate(week)}`;
    }

    // charts are optional — won’t break if missing
    safeRenderCharts(data);

  } catch (err) {
    console.error(err);
  }
}

/* ---------------- TOP KPI ---------------- */

function renderTopCards(data) {
  const container = $("dashboardCards");
  if (!container) return;

  const totals = data.totals || {};
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

  container.innerHTML = cards
    .map(
      (c) => `
      <div class="summaryCard ${getKpiClass(c.change)}">
        <h3>${c.label}</h3>
        <div class="value">${Number(c.value).toLocaleString()}</div>
        <div>${c.change >= 0 ? "+" : ""}${c.change}</div>
      </div>
    `
    )
    .join("");
}

/* ---------------- ENTITY PERFORMANCE ---------------- */

function renderEntities(regions) {
  const container = $("dashboardEntities");
  if (!container) return;

  container.innerHTML = regions
    .map((r) => {
      const budget = r.budget?.visitVolumeBudget || 0;
      const variance = (r.visitVolume || 0) - budget;

      return `
      <div class="panel innerPanel ${getKpiClass(variance)}">
        <h3>${r.entity}</h3>
        <div>Visits: ${r.visitVolume || 0}</div>
        <div>New Patients: ${r.newPatients || 0}</div>
        <div>Calls: ${r.callVolume || 0}</div>
        <div>Variance: ${variance}</div>
      </div>
    `;
    })
    .join("");
}

/* ---------------- WINS ---------------- */

function renderWins(regions) {
  const container = $("dashboardAlerts");
  if (!container) return;

  const wins = regions.filter(
    (r) => r.visitVolume > (r.budget?.visitVolumeBudget || 0)
  );

  container.innerHTML = `
    <h4>Wins</h4>
    ${
      wins.length
        ? wins
            .map(
              (w) =>
                `<div>${w.entity} performing above budget</div>`
            )
            .join("")
        : "<div>No standout wins this period</div>"
    }
  `;
}

/* ---------------- SAFE CHARTS ---------------- */

function safeRenderCharts(data) {
  if (typeof Chart === "undefined") return;

  try {
    renderBudgetChart(data.regions || []);
    renderTrendChart(data.regions || []);
  } catch (e) {
    console.warn("Chart render skipped:", e);
  }
}

function renderBudgetChart(regions) {
  const canvas = document.getElementById("budgetChart");
  if (!canvas) return;

  const ctx = canvas.getContext("2d");

  if (window.budgetChart) {
    window.budgetChart.destroy();
  }

  window.budgetChart = new Chart(ctx, {
    type: "bar",
    data: {
      labels: regions.map((r) => r.entity),
      datasets: [
        {
          label: "Actual",
          data: regions.map((r) => r.visitVolume || 0)
        },
        {
          label: "Budget",
          data: regions.map(
            (r) => r.budget?.visitVolumeBudget || 0
          )
        }
      ]
    }
  });
}

function renderTrendChart(regions) {
  const canvas = document.getElementById("trendChart");
  if (!canvas) return;

  const ctx = canvas.getContext("2d");

  if (window.trendChart) {
    window.trendChart.destroy();
  }

  window.trendChart = new Chart(ctx, {
    type: "line",
    data: {
      labels: regions.map((r) => r.entity),
      datasets: [
        {
          label: "Visits",
          data: regions.map((r) => r.visitVolume || 0)
        }
      ]
    }
  });
}
