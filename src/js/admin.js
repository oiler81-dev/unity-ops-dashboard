import { ENTITIES, KPI_METRICS } from "./definitions.js";

export function renderAdminEditorShell() {
  return `
    <div class="admin-editor-wrap">
      <div class="admin-editor-tabs">
        <button class="admin-editor-tab active" data-admin-tab="targets">Targets</button>
        <button class="admin-editor-tab" data-admin-tab="thresholds">Thresholds</button>
        <button class="admin-editor-tab" data-admin-tab="holidays">Holidays</button>
        <button class="admin-editor-tab" data-admin-tab="budget">Budget</button>
        <button class="admin-editor-tab" data-admin-tab="submissions">Submissions</button>
        <button class="admin-editor-tab" data-admin-tab="audit">Audit</button>
      </div>

      <div class="admin-filter-bar">
        <div class="field" id="adminEntityFilterWrap">
          <label for="adminEntityFilter">Entity</label>
          <select id="adminEntityFilter">
            ${ENTITIES.map(entity => `<option value="${entity}">${entity}</option>`).join("")}
          </select>
        </div>

        <div class="field hidden" id="adminYearFilterWrap">
          <label for="adminYearFilter">Year</label>
          <input id="adminYearFilter" type="number" min="2024" max="2035" value="${new Date().getFullYear()}" />
        </div>

        <div class="field hidden" id="adminAuditEntityFilterWrap">
          <label for="adminAuditEntityFilter">Audit Entity</label>
          <select id="adminAuditEntityFilter">
            <option value="">All</option>
            <option value="LAOSS">LAOSS</option>
            <option value="NES">NES</option>
            <option value="SpineOne">SpineOne</option>
            <option value="MRO">MRO</option>
            <option value="PT">PT</option>
            <option value="CXNS">CXNS</option>
            <option value="Capacity">Capacity</option>
            <option value="Productivity Builder">Productivity Builder</option>
          </select>
        </div>
      </div>

      <div id="adminEditorContent"></div>
    </div>
  `;
}

export function renderTargetsEditor(entity, rows = []) {
  const rowMap = new Map(rows.map((r) => [r.metricKey, r]));

  return `
    <div class="section-head">
      <h3>${entity} Targets</h3>
      <p class="section-copy">Manage KPI targets for dashboard cards and variance logic.</p>
    </div>

    <div class="table-wrap">
      <table>
        <thead>
          <tr>
            <th>Metric</th>
            <th>Label</th>
            <th>Target Value</th>
          </tr>
        </thead>
        <tbody>
          ${KPI_METRICS.map((metric) => {
            const row = rowMap.get(metric.key) || {};
            return `
              <tr>
                <td>${metric.key}</td>
                <td>
                  <input class="admin-input" data-admin-kind="target" data-metric-key="${metric.key}" data-field="label" value="${escapeAttr(row.label ?? metric.label)}" />
                </td>
                <td>
                  <input class="admin-input" type="number" step="0.01" data-admin-kind="target" data-metric-key="${metric.key}" data-field="targetValue" value="${escapeAttr(row.targetValue ?? "")}" />
                </td>
              </tr>
            `;
          }).join("")}
        </tbody>
      </table>
    </div>

    <div class="admin-save-row">
      <button id="saveTargetsBtn" class="btn btn-primary">Save Targets</button>
    </div>
  `;
}

export function renderThresholdsEditor(entity, rows = []) {
  const rowMap = new Map(rows.map((r) => [r.metricKey, r]));

  return `
    <div class="section-head">
      <h3>${entity} Thresholds</h3>
      <p class="section-copy">Manage KPI red/yellow/green logic by entity and metric.</p>
    </div>

    <div class="table-wrap">
      <table>
        <thead>
          <tr>
            <th>Metric</th>
            <th>Type</th>
            <th>Green Min</th>
            <th>Yellow Min</th>
            <th>Green Max</th>
            <th>Yellow Max</th>
          </tr>
        </thead>
        <tbody>
          ${KPI_METRICS.map((metric) => {
            const row = rowMap.get(metric.key) || {};
            return `
              <tr>
                <td>${metric.key}</td>
                <td>
                  <select class="admin-input" data-admin-kind="threshold" data-metric-key="${metric.key}" data-field="comparisonType">
                    <option value="higher_better" ${row.comparisonType === "higher_better" ? "selected" : ""}>higher_better</option>
                    <option value="lower_better" ${row.comparisonType === "lower_better" ? "selected" : ""}>lower_better</option>
                  </select>
                </td>
                <td><input class="admin-input" type="number" step="0.01" data-admin-kind="threshold" data-metric-key="${metric.key}" data-field="greenMin" value="${escapeAttr(row.greenMin ?? "")}" /></td>
                <td><input class="admin-input" type="number" step="0.01" data-admin-kind="threshold" data-metric-key="${metric.key}" data-field="yellowMin" value="${escapeAttr(row.yellowMin ?? "")}" /></td>
                <td><input class="admin-input" type="number" step="0.01" data-admin-kind="threshold" data-metric-key="${metric.key}" data-field="greenMax" value="${escapeAttr(row.greenMax ?? "")}" /></td>
                <td><input class="admin-input" type="number" step="0.01" data-admin-kind="threshold" data-metric-key="${metric.key}" data-field="yellowMax" value="${escapeAttr(row.yellowMax ?? "")}" /></td>
              </tr>
            `;
          }).join("")}
        </tbody>
      </table>
    </div>

    <div class="admin-save-row">
      <button id="saveThresholdsBtn" class="btn btn-primary">Save Thresholds</button>
    </div>
  `;
}

export function renderHolidaysEditor(year, rows = []) {
  const normalized = [...rows].sort((a, b) => String(a.date || "").localeCompare(String(b.date || "")));

  return `
    <div class="section-head">
      <h3>${year} Holidays</h3>
      <p class="section-copy">Manage holiday dates used for working-day and planning logic.</p>
    </div>

    <div class="table-wrap">
      <table>
        <thead>
          <tr>
            <th>Date</th>
            <th>Holiday Name</th>
          </tr>
        </thead>
        <tbody id="holidayRows">
          ${normalized.map((row, index) => `
            <tr>
              <td><input class="admin-input" type="date" data-admin-kind="holiday" data-row-index="${index}" data-field="date" value="${escapeAttr(row.date ?? "")}" /></td>
              <td><input class="admin-input" data-admin-kind="holiday" data-row-index="${index}" data-field="holidayName" value="${escapeAttr(row.holidayName ?? "")}" /></td>
            </tr>
          `).join("")}
          ${renderEmptyHolidayRows(normalized.length, 8)}
        </tbody>
      </table>
    </div>

    <div class="admin-save-row">
      <button id="saveHolidaysBtn" class="btn btn-primary">Save Holidays</button>
    </div>
  `;
}

export function renderBudgetEditor(entity, rows = []) {
  const rowMap = new Map(rows.map((r) => [r.monthKey, r]));
  const months = ["01","02","03","04","05","06","07","08","09","10","11","12"];

  return `
    <div class="section-head">
      <h3>${entity} Budget</h3>
      <p class="section-copy">Manage monthly budget reference values for visits and revenue.</p>
    </div>

    <div class="table-wrap">
      <table>
        <thead>
          <tr>
            <th>Month</th>
            <th>Budget Visits</th>
            <th>Budget Revenue</th>
          </tr>
        </thead>
        <tbody>
          ${months.map((month) => {
            const row = rowMap.get(month) || {};
            return `
              <tr>
                <td>${month}</td>
                <td><input class="admin-input" type="number" step="0.01" data-admin-kind="budget" data-month-key="${month}" data-field="budgetVisits" value="${escapeAttr(row.budgetVisits ?? "")}" /></td>
                <td><input class="admin-input" type="number" step="0.01" data-admin-kind="budget" data-month-key="${month}" data-field="budgetRevenue" value="${escapeAttr(row.budgetRevenue ?? "")}" /></td>
              </tr>
            `;
          }).join("")}
        </tbody>
      </table>
    </div>

    <div class="admin-save-row">
      <button id="saveBudgetBtn" class="btn btn-primary">Save Budget</button>
    </div>
  `;
}

export function renderSubmissionTracker(data) {
  return `
    <div class="section-head">
      <h3>Submission Tracking</h3>
      <p class="section-copy">Accountability view for status, updates, and missing weekly submissions.</p>
    </div>

    <div class="summary-mini-grid">
      <div class="summary-mini-card">
        <span class="summary-mini-label">Total Entities</span>
        <strong class="summary-mini-value">${data.summary?.totalEntities ?? 0}</strong>
      </div>
      <div class="summary-mini-card">
        <span class="summary-mini-label">Submitted</span>
        <strong class="summary-mini-value">${data.summary?.submittedCount ?? 0}</strong>
      </div>
      <div class="summary-mini-card">
        <span class="summary-mini-label">Missing</span>
        <strong class="summary-mini-value">${data.summary?.missingCount ?? 0}</strong>
      </div>
      <div class="summary-mini-card">
        <span class="summary-mini-label">Week Ending</span>
        <strong class="summary-mini-value">${escapeHtml(data.weekEnding || "")}</strong>
      </div>
    </div>

    <div class="table-wrap">
      <table>
        <thead>
          <tr>
            <th>Entity</th>
            <th>Status</th>
            <th>Updated By</th>
            <th>Updated At</th>
            <th>Submitted By</th>
            <th>Submitted At</th>
            <th>Inputs</th>
            <th>Narrative</th>
          </tr>
        </thead>
        <tbody>
          ${(data.rows || []).map((row) => `
            <tr>
              <td>${escapeHtml(row.entity)}</td>
              <td>${escapeHtml(row.status)}</td>
              <td>${escapeHtml(row.updatedBy || "—")}</td>
              <td>${escapeHtml(row.updatedAt || "—")}</td>
              <td>${escapeHtml(row.submittedBy || "—")}</td>
              <td>${escapeHtml(row.submittedAt || "—")}</td>
              <td>${escapeHtml(String(row.inputCount ?? 0))}</td>
              <td>${row.hasNarrative ? "Yes" : "No"}</td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    </div>

    <div class="note-panel" style="margin-top:18px;">
      <h4>Missing Entities</h4>
      <p>${escapeHtml((data.summary?.missingEntities || []).join(", ") || "None")}</p>
    </div>
  `;
}

export function renderAuditViewer(data) {
  return `
    <div class="section-head">
      <h3>Audit Log</h3>
      <p class="section-copy">Recent changes for the selected week and entity scope.</p>
    </div>

    <div class="table-wrap">
      <table>
        <thead>
          <tr>
            <th>Entity</th>
            <th>Section</th>
            <th>Metric</th>
            <th>Old Value</th>
            <th>New Value</th>
            <th>Changed By</th>
            <th>Changed At</th>
            <th>Type</th>
          </tr>
        </thead>
        <tbody>
          ${(data.rows || []).map((row) => `
            <tr>
              <td>${escapeHtml(row.entity || "—")}</td>
              <td>${escapeHtml(row.section || "—")}</td>
              <td>${escapeHtml(row.metricKey || "—")}</td>
              <td>${escapeHtml(String(row.oldValue ?? "—"))}</td>
              <td>${escapeHtml(String(row.newValue ?? "—"))}</td>
              <td>${escapeHtml(row.changedBy || "—")}</td>
              <td>${escapeHtml(row.changedAt || "—")}</td>
              <td>${escapeHtml(row.changeType || "—")}</td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    </div>
  `;
}

export function collectAdminRows(kind) {
  if (kind === "holiday") {
    const map = new Map();
    document.querySelectorAll(`.admin-input[data-admin-kind="holiday"]`).forEach((el) => {
      const rowIndex = el.dataset.rowIndex;
      const field = el.dataset.field;
      if (!map.has(rowIndex)) map.set(rowIndex, {});
      const row = map.get(rowIndex);
      row[field] = el.value === "" ? null : el.value;
    });
    return Array.from(map.values()).filter((row) => row.date || row.holidayName);
  }

  if (kind === "budget") {
    const map = new Map();
    document.querySelectorAll(`.admin-input[data-admin-kind="budget"]`).forEach((el) => {
      const monthKey = el.dataset.monthKey;
      const field = el.dataset.field;
      if (!map.has(monthKey)) map.set(monthKey, { monthKey });
      const row = map.get(monthKey);
      row[field] = el.value === "" ? null : el.value;
    });
    return Array.from(map.values());
  }

  const map = new Map();
  document.querySelectorAll(`.admin-input[data-admin-kind="${kind}"]`).forEach((el) => {
    const metricKey = el.dataset.metricKey;
    const field = el.dataset.field;
    if (!map.has(metricKey)) map.set(metricKey, { metricKey });
    const row = map.get(metricKey);
    row[field] = el.value === "" ? null : el.value;
  });

  return Array.from(map.values());
}

function renderEmptyHolidayRows(existingCount, totalRows) {
  const rowsNeeded = Math.max(totalRows - existingCount, 0);
  return Array.from({ length: rowsNeeded }).map((_, i) => {
    const rowIndex = existingCount + i;
    return `
      <tr>
        <td><input class="admin-input" type="date" data-admin-kind="holiday" data-row-index="${rowIndex}" data-field="date" value="" /></td>
        <td><input class="admin-input" data-admin-kind="holiday" data-row-index="${rowIndex}" data-field="holidayName" value="" /></td>
      </tr>
    `;
  }).join("");
}

function escapeAttr(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}