/**
 * Safe GET wrapper
 * @param {string} url
 * @param {any} defaultValue
 */
export async function safeApiGet(url, defaultValue) {
  try {
    const response = await fetch(url);
    if (!response.ok) {
      console.error(`API GET failed for ${url}: Status ${response.status}`);
      return defaultValue;
    }
    return await response.json();
  } catch (err) {
    console.error(`Network error during API GET for ${url}:`, err);
    return defaultValue;
  }
}

/**
 * Safe POST wrapper
 * @param {string} url
 * @param {object} body
 */
export async function apiPost(url, body) {
  try {
    const response = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body)
    });

    if (!response.ok) {
      console.error(`API POST failed for ${url}: Status ${response.status}`);
      const errBody = await response.json().catch(() => ({ error: "Request failed" }));
      return errBody;
    }

    return await response.json();
  } catch (err) {
    console.error(`Network error during API POST for ${url}:`, err);
    return { error: "Network request failed", details: err.message };
  }
}

/**
 * Get most recent Sunday as YYYY-MM-DD
 */
export function getDefaultWeekEnding() {
  const today = new Date();
  const dayOfWeek = today.getDay();
  const date = new Date(today);
  date.setDate(today.getDate() - dayOfWeek);
  return date.toISOString().split("T")[0];
}

/**
 * Format YYYY-MM-DD to readable string
 * @param {string} dateString
 */
export function formatDate(dateString) {
  if (!dateString || !/^\d{4}-\d{2}-\d{2}$/.test(dateString)) return "N/A";
  const date = new Date(`${dateString}T12:00:00Z`);
  return date.toLocaleDateString("en-US", {
    year: "numeric",
    month: "long",
    day: "numeric"
  });
}

/**
 * Escape HTML to prevent XSS
 * @param {string} input
 */
export function escapeHtml(input) {
  const s = String(input ?? "");
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

export function escapeAttr(str) {
  return escapeHtml(str);
}

export function cssEscape(value) {
  return String(value ?? "").replace(/["'\\]/g, "\\$&");
}

export function renderKpiCards(kpis) {
  const container = document.getElementById("dashboardCards");
  if (!container) return;

  const list = Array.isArray(kpis) && kpis.length
    ? kpis
    : [
        { label: "Total Visits", value: "—", statusColor: "yellow", meta: "" },
        { label: "Visits / Day", value: "—", statusColor: "yellow", meta: "" },
        { label: "Abandonment Rate", value: "—", statusColor: "yellow", meta: "" }
      ];

  container.innerHTML = list.map((kpi) => `
    <div class="kpi-card">
      <div class="kpi-label">${escapeHtml(kpi.label)}</div>
      <div class="kpi-value">${escapeHtml(String(kpi.value ?? "—"))}</div>
      ${kpi.meta ? `<div class="kpi-meta">${escapeHtml(kpi.meta)}</div>` : ""}
      ${kpi.status ? `<div class="kpi-status ${escapeAttr(kpi.statusColor || "")}">${escapeHtml(kpi.status)}</div>` : ""}
    </div>
  `).join("");
}

export function handleFatalError(err) {
  console.error("FATAL ERROR:", err);
  const body = document.querySelector("body");
  if (body) {
    body.innerHTML = `<div class="fatal-error">
      <h1>Application Error</h1>
      <p>A critical error occurred and the application cannot continue.</p>
      <pre>${escapeHtml(err && err.stack ? err.stack : String(err))}</pre>
    </div>`;
  }
}

export function comparisonClass(change) {
  if (typeof change === "number") {
    if (change > 0) return "positive";
    if (change < 0) return "negative";
  }
  return "neutral";
}

export function formatChange(change, format) {
  const sign = typeof change === "number" && change > 0 ? "+" : "";
  if (change == null || change === "") return "—";

  if (format === "percent1") {
    return `${sign}${Number(change).toFixed(1)}%`;
  }

  if (format === "whole") {
    return `${sign}${Math.round(Number(change)).toLocaleString()}`;
  }

  return `${sign}${Number(change).toFixed(2)}`;
}

export function renderSummaryMiniCard(summary) {
  return `
    <div class="summary-mini-card">
      <span class="summary-mini-label">${escapeHtml(summary.label)}</span>
      <strong class="summary-mini-value">${escapeHtml(summary.value)}</strong>
      ${summary.meta ? `<span class="summary-mini-meta">${escapeHtml(summary.meta)}</span>` : ""}
    </div>
  `;
}

function parseNumericInputValue(value) {
  if (value == null || value === "") return "";
  const n = Number(String(value).replace(/,/g, "").trim());
  return Number.isFinite(n) ? n : value;
}

function collectFieldValues() {
  const fields = Array.from(document.querySelectorAll("[data-key]"));
  const values = {};

  for (const el of fields) {
    if (!el.matches("input, textarea, select")) continue;

    const key = el.dataset.key;
    if (!key) continue;

    values[key] = parseNumericInputValue(el.value);
  }

  return values;
}

function detectCurrentEntity() {
  const assigned = document.getElementById("assignedEntityText")?.textContent?.trim();
  if (assigned) return assigned;

  const activeRegion = document.querySelector('.nav-link.active[data-route="region"]');
  return activeRegion?.dataset?.entity || "LAOSS";
}

function detectCurrentSharedPage() {
  const activeShared = document.querySelector('.nav-link.active[data-route="shared"]');
  return activeShared?.dataset?.page || "PT";
}

export function collectRegionFormValues() {
  const weekEnding =
    document.getElementById("weekEndingSelect")?.value || getDefaultWeekEnding();

  return {
    entity: detectCurrentEntity(),
    weekEnding,
    values: collectFieldValues()
  };
}

export function collectSharedFormValues() {
  const weekEnding =
    document.getElementById("weekEndingSelect")?.value || getDefaultWeekEnding();

  return {
    page: detectCurrentSharedPage(),
    weekEnding,
    values: collectFieldValues()
  };
}
