/**
 * A safe wrapper for the Fetch API for GET requests.
 * @param {string} url The API endpoint to call.
 * @param {any} defaultValue The value to return if the fetch fails.
 * @returns {Promise<any>}
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
 * A wrapper for the Fetch API for POST requests.
 * @param {string} url The API endpoint to call.
 * @param {object} body The JSON payload to send.
 * @returns {Promise<any>}
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
      // Try to return error details from the API if possible
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
 * Gets the most recent Sunday as a 'YYYY-MM-DD' string.
 * @returns {string}
 */
export function getDefaultWeekEnding() {
  const today = new Date();
  const dayOfWeek = today.getDay(); // Sunday is 0, Monday is 1, etc.
  const date = new Date(today);
  date.setDate(today.getDate() - dayOfWeek);
  return date.toISOString().split("T")[0];
}

/**
 * Formats a 'YYYY-MM-DD' string into a more readable 'Month Day, Year' format.
 * @param {string} dateString
 * @returns {string}
 */
export function formatDate(dateString) {
  if (!dateString || !/^\d{4}-\d{2}-\d{2}$/.test(dateString)) return "N/A";
  const date = new Date(`${dateString}T12:00:00Z`); // Use noon UTC to avoid timezone issues
  return date.toLocaleDateString("en-US", {
    year: "numeric",
    month: "long",
    day: "numeric",
  });
}

/**
 * Escapes HTML to prevent XSS attacks.
 * @param {string} str
 * @returns {string}
 */
export function escapeHtml(str) {
  const s = String(str || "");
  return s
    .replace(/&/g, "&")
    .replace(/</g, "<")
    .replace(/>/g, ">")
    .replace(/"/g, """)
    .replace(/'/g, "'");
}

/**
 * Escapes attributes for safe use in HTML tags.
 * @param {string} str
 * @returns {string}
 */
export function escapeAttr(str) {
  return escapeHtml(str); // For this app's purposes, the same escaping is sufficient.
}

/**
 * A placeholder for a fatal error handler.
 * @param {Error} err
 */
export function handleFatalError(err) {
  console.error("FATAL ERROR:", err);
  const body = document.querySelector("body");
  if (body) {
    body.innerHTML = `<div class="fatal-error">
      <h1>Application Error</h1>
      <p>A critical error occurred and the application cannot continue.</p>
      <pre>${escapeHtml(err.stack)}</pre>
    </div>`;
  }
}

// --- Placeholders for other missing functions to prevent errors ---

export function comparisonClass(change, key) {
  if (change > 0) return "positive";
  if (change < 0) return "negative";
  return "neutral";
}

export function formatChange(change, format) {
  const sign = change > 0 ? "+" : "";
  return sign + formatByType(change, format);
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

export function collectRegionFormValues() {
  // In a real implementation, this would gather all form inputs.
  console.log("Collecting Region Form Values...");
  return {
    entity: "LAOSS",
    weekEnding: "2026-03-15",
    inputs: {},
    narrative: {}
  };
}

export function collectSharedFormValues() {
  console.log("Collecting Shared Form Values...");
  return {
    page: "Capacity",
    weekEnding: "2026-03-15",
    inputs: {}
  };
}
