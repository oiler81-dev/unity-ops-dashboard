const { getTableClient } = require("../shared/table");
const {
  getMonthKeyFromDate,
  getMonthDateRange,
  getPreviousMonthDateRange,
  getWorkingDaysBetween,
  prorateMonthlyValue,
  normalizeEntityLabel,
  safeNumber
} = require("../shared/budget");

const BUDGET_TABLE = "BudgetData";

module.exports = async function (context, req) {
  try {
    const entity = normalizeEntityLabel(req.query.entity);
    const weekEnding = String(req.query.weekEnding || "").trim();
    const period = String(req.query.period || "currentWeek").trim();
    const daysInPeriod = safeNumber(req.query.daysInPeriod);

    if (!entity || !weekEnding) {
      context.res = {
        status: 400,
        body: {
          ok: false,
          error: "Missing entity or weekEnding"
        }
      };
      return;
    }

    const table = getTableClient(BUDGET_TABLE);

    let targetMonthKey = getMonthKeyFromDate(weekEnding);
    let range = getMonthDateRange(targetMonthKey);
    let workingDaysUsed = 0;

    if (period === "lastMonth") {
      const previous = getPreviousMonthDateRange(weekEnding);
      targetMonthKey = previous.monthKey;
      range = previous;
    }

    let record = null;
    try {
      record = await table.getEntity(entity, targetMonthKey);
    } catch (error) {
      if (error.statusCode !== 404) {
        throw error;
      }
    }

    const workingDaysInMonth =
      safeNumber(record?.workingDaysInMonth) ||
      getWorkingDaysBetween(range.startDate, range.endDate);

    if (period === "currentWeek") {
      workingDaysUsed = Number.isFinite(daysInPeriod)
        ? daysInPeriod
        : getWorkingDaysBetween(
            range.weekStartDate || weekEnding,
            weekEnding
          );
    } else if (period === "mtd") {
      workingDaysUsed = getWorkingDaysBetween(range.startDate, weekEnding);
    } else if (period === "lastMonth") {
      workingDaysUsed = workingDaysInMonth;
    } else {
      workingDaysUsed = workingDaysInMonth;
    }

    const visitBudgetMonthly = safeNumber(record?.visitBudgetMonthly) || 0;
    const newPatientsBudgetMonthly = safeNumber(record?.newPatientsBudgetMonthly) || 0;

    context.res = {
      status: 200,
      headers: { "Content-Type": "application/json" },
      body: {
        ok: true,
        entity,
        period,
        weekEnding,
        monthKey: targetMonthKey,
        workingDaysInMonth,
        workingDaysUsed,
        visitBudgetMonthly,
        newPatientsBudgetMonthly,
        visitBudgetProrated: prorateMonthlyValue(visitBudgetMonthly, workingDaysUsed, workingDaysInMonth),
        newPatientsBudgetProrated: prorateMonthlyValue(newPatientsBudgetMonthly, workingDaysUsed, workingDaysInMonth)
      }
    };
  } catch (error) {
    context.log.error("budget GET failed", error);
    context.res = {
      status: 500,
      body: {
        ok: false,
        error: "Failed to load budget data",
        details: error.message
      }
    };
  }
};
