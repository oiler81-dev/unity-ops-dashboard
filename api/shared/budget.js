function normalizeText(value) {
  return String(value || "").trim();
}

function pad2(value) {
  return String(value).padStart(2, "0");
}

function safeNumber(value) {
  if (value == null || value === "") return null;

  if (typeof value === "number") {
    return Number.isFinite(value) ? value : null;
  }

  const cleaned = String(value)
    .replace(/[$,%(),]/g, "")
    .replace(/,/g, "")
    .trim();

  if (!cleaned || /^n\/a$/i.test(cleaned)) {
    return null;
  }

  const n = Number(cleaned);
  return Number.isFinite(n) ? n : null;
}

function normalizeMonthLabel(value) {
  const raw = normalizeText(value).toLowerCase();

  const map = {
    jan: "Jan",
    january: "Jan",
    feb: "Feb",
    february: "Feb",
    mar: "Mar",
    march: "Mar",
    apr: "Apr",
    april: "Apr",
    may: "May",
    jun: "Jun",
    june: "Jun",
    jul: "Jul",
    july: "Jul",
    aug: "Aug",
    august: "Aug",
    sep: "Sep",
    sept: "Sep",
    september: "Sep",
    oct: "Oct",
    october: "Oct",
    nov: "Nov",
    november: "Nov",
    dec: "Dec",
    december: "Dec"
  };

  return map[raw] || "";
}

function monthLabelToMonthNumber(monthLabel) {
  const normalized = normalizeMonthLabel(monthLabel);

  const map = {
    Jan: 1,
    Feb: 2,
    Mar: 3,
    Apr: 4,
    May: 5,
    Jun: 6,
    Jul: 7,
    Aug: 8,
    Sep: 9,
    Oct: 10,
    Nov: 11,
    Dec: 12
  };

  return map[normalized] || null;
}

function monthLabelToMonthKey(monthLabel, year = 2026) {
  const monthNumber = monthLabelToMonthNumber(monthLabel);
  if (!monthNumber) return "";
  return `${year}-${pad2(monthNumber)}`;
}

function getMonthKeyFromDate(dateValue) {
  const date = new Date(`${dateValue}T12:00:00Z`);
  if (Number.isNaN(date.getTime())) return "";
  return `${date.getUTCFullYear()}-${pad2(date.getUTCMonth() + 1)}`;
}

function getMonthDateRange(monthKey) {
  const [yearText, monthText] = String(monthKey || "").split("-");
  const year = Number(yearText);
  const month = Number(monthText);

  const start = new Date(Date.UTC(year, month - 1, 1));
  const end = new Date(Date.UTC(year, month, 0));

  return {
    monthKey,
    startDate: `${start.getUTCFullYear()}-${pad2(start.getUTCMonth() + 1)}-${pad2(start.getUTCDate())}`,
    endDate: `${end.getUTCFullYear()}-${pad2(end.getUTCMonth() + 1)}-${pad2(end.getUTCDate())}`
  };
}

function getPreviousMonthDateRange(dateValue) {
  const date = new Date(`${dateValue}T12:00:00Z`);
  const year = date.getUTCFullYear();
  const month = date.getUTCMonth();

  const start = new Date(Date.UTC(year, month - 1, 1));
  const end = new Date(Date.UTC(year, month, 0));

  return {
    monthKey: `${start.getUTCFullYear()}-${pad2(start.getUTCMonth() + 1)}`,
    startDate: `${start.getUTCFullYear()}-${pad2(start.getUTCMonth() + 1)}-${pad2(start.getUTCDate())}`,
    endDate: `${end.getUTCFullYear()}-${pad2(end.getUTCMonth() + 1)}-${pad2(end.getUTCDate())}`
  };
}

function getWeekRangeFromWeekEnding(weekEnding) {
  const end = new Date(`${weekEnding}T12:00:00Z`);
  const start = new Date(end);
  start.setUTCDate(end.getUTCDate() - 6);

  return {
    startDate: `${start.getUTCFullYear()}-${pad2(start.getUTCMonth() + 1)}-${pad2(start.getUTCDate())}`,
    endDate: `${end.getUTCFullYear()}-${pad2(end.getUTCMonth() + 1)}-${pad2(end.getUTCDate())}`
  };
}

function getWorkingDaysBetween(startDate, endDate) {
  const start = new Date(`${startDate}T12:00:00Z`);
  const end = new Date(`${endDate}T12:00:00Z`);

  if (Number.isNaN(start.getTime()) || Number.isNaN(end.getTime()) || start > end) {
    return 0;
  }

  let count = 0;
  const current = new Date(start);

  while (current <= end) {
    const day = current.getUTCDay();
    if (day !== 0 && day !== 6) {
      count += 1;
    }
    current.setUTCDate(current.getUTCDate() + 1);
  }

  return count;
}

function getWorkingDaysForMonth(monthKey) {
  const range = getMonthDateRange(monthKey);
  return getWorkingDaysBetween(range.startDate, range.endDate);
}

function prorateMonthlyValue(monthlyValue, workingDaysUsed, workingDaysInMonth) {
  const monthly = safeNumber(monthlyValue) || 0;
  const used = safeNumber(workingDaysUsed) || 0;
  const total = safeNumber(workingDaysInMonth) || 0;

  if (!total) return 0;
  return monthly * (used / total);
}

function normalizeEntityLabel(value) {
  const raw = normalizeText(value);
  const upper = raw.toUpperCase();

  if (upper === "LAOSS") return "LAOSS";
  if (upper === "LA") return "LAOSS";
  if (upper === "NES") return "NES";
  if (upper === "PORTLAND") return "NES";
  if (upper === "CHICAGO") return "MRO";
  if (upper === "MRO") return "MRO";
  if (upper === "SO") return "SpineOne";
  if (upper === "SPINEONE") return "SpineOne";
  if (upper === "SPINE ONE") return "SpineOne";
  if (upper === "DENVER") return "SpineOne";

  return "";
}

function isoDate(value) {
  const d = new Date(`${value}T12:00:00Z`);
  if (Number.isNaN(d.getTime())) return "";
  return `${d.getUTCFullYear()}-${pad2(d.getUTCMonth() + 1)}-${pad2(d.getUTCDate())}`;
}

function listMonthKeysBetween(startDate, endDate) {
  const start = new Date(`${startDate}T12:00:00Z`);
  const end = new Date(`${endDate}T12:00:00Z`);
  const keys = [];

  if (Number.isNaN(start.getTime()) || Number.isNaN(end.getTime()) || start > end) {
    return keys;
  }

  const cursor = new Date(Date.UTC(start.getUTCFullYear(), start.getUTCMonth(), 1));
  const endCursor = new Date(Date.UTC(end.getUTCFullYear(), end.getUTCMonth(), 1));

  while (cursor <= endCursor) {
    keys.push(`${cursor.getUTCFullYear()}-${pad2(cursor.getUTCMonth() + 1)}`);
    cursor.setUTCMonth(cursor.getUTCMonth() + 1);
  }

  return keys;
}

function dateMax(a, b) {
  return a > b ? a : b;
}

function dateMin(a, b) {
  return a < b ? a : b;
}

async function getBudgetRecord(table, entity, monthKey) {
  try {
    return await table.getEntity(entity, monthKey);
  } catch (error) {
    if (error.statusCode === 404) {
      return null;
    }
    throw error;
  }
}

async function getProratedBudgetForRange(table, entity, startDate, endDate) {
  const normalizedEntity = normalizeEntityLabel(entity);
  const months = listMonthKeysBetween(startDate, endDate);

  const totals = {
    visitBudget: 0,
    newPatientsBudget: 0,
    workingDaysUsed: 0,
    workingDaysInScopeMonths: 0,
    monthBreakdown: []
  };

  for (const monthKey of months) {
    const record = await getBudgetRecord(table, normalizedEntity, monthKey);
    const monthRange = getMonthDateRange(monthKey);

    const overlapStart = dateMax(startDate, monthRange.startDate);
    const overlapEnd = dateMin(endDate, monthRange.endDate);

    if (overlapStart > overlapEnd) {
      continue;
    }

    const workingDaysUsed = getWorkingDaysBetween(overlapStart, overlapEnd);
    const workingDaysInMonth =
      safeNumber(record?.workingDaysInMonth) || getWorkingDaysForMonth(monthKey);

    const visitBudgetMonthly = safeNumber(record?.visitBudgetMonthly) || 0;
    const newPatientsBudgetMonthly = safeNumber(record?.newPatientsBudgetMonthly) || 0;

    const visitBudgetProrated = prorateMonthlyValue(
      visitBudgetMonthly,
      workingDaysUsed,
      workingDaysInMonth
    );

    const newPatientsBudgetProrated = prorateMonthlyValue(
      newPatientsBudgetMonthly,
      workingDaysUsed,
      workingDaysInMonth
    );

    totals.visitBudget += visitBudgetProrated;
    totals.newPatientsBudget += newPatientsBudgetProrated;
    totals.workingDaysUsed += workingDaysUsed;
    totals.workingDaysInScopeMonths += workingDaysInMonth;

    totals.monthBreakdown.push({
      monthKey,
      overlapStart,
      overlapEnd,
      workingDaysUsed,
      workingDaysInMonth,
      visitBudgetMonthly,
      newPatientsBudgetMonthly,
      visitBudgetProrated,
      newPatientsBudgetProrated
    });
  }

  return totals;
}

module.exports = {
  normalizeMonthLabel,
  monthLabelToMonthKey,
  getMonthKeyFromDate,
  getMonthDateRange,
  getPreviousMonthDateRange,
  getWeekRangeFromWeekEnding,
  getWorkingDaysBetween,
  getWorkingDaysForMonth,
  prorateMonthlyValue,
  normalizeEntityLabel,
  safeNumber,
  isoDate,
  listMonthKeysBetween,
  getProratedBudgetForRange
};
