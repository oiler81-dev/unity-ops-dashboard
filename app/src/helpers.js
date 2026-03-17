// helpers.js - full file to copy/paste

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
      body: JSON.stringify(body),
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
  const dayOfWeek = today.getDay(); // Sunday = 0
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
    day: "numeric",
  });
}

/**
 * Escape HTML to prevent XSS
 * @param {string} input
 */
export function escapeHtml(input) {
  const s = String(input ?? "");
  return s
    .replace(/&/g, "&")
    .replace(/</g, "<")
    .replace(/>/g, ">")
    .replace(/"/g, """)
    .replace(/'/g, "'");
}

/**
 * Escape attributes (same for our usage)
 * @param {string} str
 */
export function escapeAttr(str) {
  return escapeHtml(str);
}

/**
 * Fatal error handler that shows a simple error page
 * @param {Error} err
 */
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

/* Utility placeholders used by the app */

/**
 * Return a CSS class for change values
 */
export function comparisonClass(change, key) {
  if (typeof change === "number") {
    if (change > 0) return "positive";
    if (change < 0) return "negative";
  }
  return "neutral";
}

/**
 * Simple formatter helper for change values (calls formatByType if available)
 */
export function formatChange(change, format) {
  const sign = typeof change === "number" && change > 0 ? "+" : "";
  // If formatByType exists globally or will be imported elsewhere, this will still work.
  const str = (typeof formatByType === "function") ? formatByType(change, format) : String(change ?? "0");
  return sign + str;
}

/**
 * Minimal render for summary mini card
 */
export function renderSummaryMiniCard(summary) {
  return `
    <div class="summary-mini-card">
      <span class="summary-mini-label">${escapeHtml(summary.label)}</span>
      <strong class="summary-mini-value">${escapeHtml(summary.value)}</strong>
      ${summary.meta ? `<span class="summary-mini-meta">${escapeHtml(summary.meta)}</span>` : ""}
    </div>
  `;
}

/**
 * Collect region form values (placeholder - app expects shape)
 */
export function collectRegionFormValues() {
  console.log("Collecting Region Form Values (placeholder)...");
  return {
    entity: "LAOSS",
    weekEnding: getDefaultWeekEnding(),
    inputs: {},
    narrative: {}
  };
}

/**
 * Collect shared form values (placeholder)
 */
export function collectSharedFormValues() {
  console.log("Collecting Shared Form Values (placeholder)...");
  return {
    page: "Capacity",
    weekEnding: getDefaultWeekEnding(),
    inputs: {}
  };
}

/* Export compatibility: if formatByType is needed by helpers, we can provide a passthrough default.
   In your real app this is likely implemented in calculations.js and imported there; providing a safe
   fallback prevents runtime errors if someone calls it before that module is loaded. */
export function formatByType(value, fmt) {
  if (value == null) return "—";
  if (fmt === "percent") return `${Number(value).toFixed(1)}%`;
  if (fmt === "integer") return String(Math.round(Number(value) || 0));
  if (fmt === "decimal2") return Number(value).toFixed(2);
  return String(value);
}
