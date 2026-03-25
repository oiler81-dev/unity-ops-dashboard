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
  if (upper === "NES") return "NES";
  if (upper === "CHICAGO") return "MRO";
  if (upper === "MRO") return "MRO";
  if (upper === "SO") return "SpineOne";
  if (upper === "SPINEONE") return "SpineOne";
  if (upper === "SPINE ONE") return "SpineOne";

  return "";
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
  safeNumber
};
