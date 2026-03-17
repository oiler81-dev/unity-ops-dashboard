export function toNumber(value) {
  const n = Number(value);
  return Number.isFinite(n) ? n : null;
}

export function formatPercent(value) {
  return value === null || value === undefined || value === ""
    ? "—"
    : `${Number(value).toFixed(1)}%`;
}

export function formatWhole(value) {
  return value === null || value === undefined || value === ""
    ? "—"
    : Number(value).toLocaleString();
}

export function formatDecimal(value, digits = 2) {
  return value === null || value === undefined || value === ""
    ? "—"
    : Number(value).toFixed(digits);
}

export function formatByType(value, format) {
  switch (format) {
    case "percent1":
      return formatPercent(value);
    case "whole":
      return formatWhole(value);
    case "decimal1":
      return formatDecimal(value, 1);
    case "decimal2":
    default:
      return formatDecimal(value, 2);
  }
}

function safeDivide(numerator, denominator, multiplier = 1) {
  const n = toNumber(numerator);
  const d = toNumber(denominator);

  if (n === null || d === null || d === 0) return null;
  return (n / d) * multiplier;
}

function ceilQuarter(value) {
  const v = toNumber(value);
  if (v === null) return null;
  return Math.ceil(v / 0.25) * 0.25;
}

/* ---------------------------------
   REGION — workbook aligned formulas
   --------------------------------- */

export function getRegionCalculatedValues(inputs = {}) {
  const daysInPeriod = toNumber(inputs.daysInPeriod);
  const workingDaysInMonth = toNumber(inputs.workingDaysInMonth);
  const npActual = toNumber(inputs.npActual);
  const establishedActual = toNumber(inputs.establishedActual);
  const surgeryActual = toNumber(inputs.surgeryActual);
  const totalCalls = toNumber(inputs.totalCalls);
  const abandonedCalls = toNumber(inputs.abandonedCalls);
  const ptVisitsSeen = toNumber(inputs.ptVisitsSeen);
  const monthlyBudgetNp = toNumber(inputs.monthlyBudgetNp);
  const monthlyBudgetEstablished = toNumber(inputs.monthlyBudgetEstablished);
  const monthlyBudgetSurgery = toNumber(inputs.monthlyBudgetSurgery);

  const baseVisits =
    (npActual || 0) +
    (establishedActual || 0) +
    (surgeryActual || 0);

  const totalVisits = baseVisits + (ptVisitsSeen || 0);

  const visitsPerDay = safeDivide(totalVisits, daysInPeriod);
  const npPerDay = safeDivide(npActual, daysInPeriod);
  const abandonmentRate = safeDivide(abandonedCalls, totalCalls, 100);
  const answeredCallRate =
    totalCalls !== null && abandonedCalls !== null && totalCalls > 0
      ? ((totalCalls - abandonedCalls) / totalCalls) * 100
      : null;

  const npToEstablishedConversion = safeDivide(establishedActual, npActual, 100);
  const npToSurgeryConversion = safeDivide(surgeryActual, npActual, 100);

  const npVariance =
    npActual !== null && monthlyBudgetNp !== null && workingDaysInMonth
      ? npActual - ((monthlyBudgetNp / workingDaysInMonth) * (daysInPeriod || 0))
      : null;

  const establishedVariance =
    establishedActual !== null && monthlyBudgetEstablished !== null && workingDaysInMonth
      ? establishedActual - ((monthlyBudgetEstablished / workingDaysInMonth) * (daysInPeriod || 0))
      : null;

  const surgeryVariance =
    surgeryActual !== null && monthlyBudgetSurgery !== null && workingDaysInMonth
      ? surgeryActual - ((monthlyBudgetSurgery / workingDaysInMonth) * (daysInPeriod || 0))
      : null;

  return {
    totalVisits,
    visitsPerDay,
    npPerDay,
    abandonmentRate,
    answeredCallRate,
    npToEstablishedConversion,
    npToSurgeryConversion,
    npVariance,
    establishedVariance,
    surgeryVariance
  };
}

export function calculateRegionSummaries(inputs = {}) {
  const calc = getRegionCalculatedValues(inputs);

  return [
    {
      label: "Total Visits",
      value: formatWhole(calc.totalVisits),
      meta: "NP + Established + Surgery + PT Visits Seen"
    },
    {
      label: "Visits / Day",
      value: formatDecimal(calc.visitsPerDay, 2),
      meta: "Total Visits / Days in Period"
    },
    {
      label: "NP / Day",
      value: formatDecimal(calc.npPerDay, 2),
      meta: "NP Actual / Days in Period"
    },
    {
      label: "Abandonment Rate",
      value: formatPercent(calc.abandonmentRate),
      meta: "Abandoned Calls / Total Calls"
    },
    {
      label: "NP → Established",
      value: formatPercent(calc.npToEstablishedConversion),
      meta: "Established Actual / NP Actual"
    },
    {
      label: "NP → Surgery",
      value: formatPercent(calc.npToSurgeryConversion),
      meta: "Surgery Actual / NP Actual"
    }
  ];
}

/* ----------------------------
   SHARED — PT workbook formulas
   ---------------------------- */

function getPtCalculatedValues(inputs = {}) {
  const workingDaysInWeek = toNumber(inputs.workingDaysInWeek);
  const ptScheduledVisits = toNumber(inputs.ptScheduledVisits);
  const ptCancellations = toNumber(inputs.ptCancellations);
  const ptNoShows = toNumber(inputs.ptNoShows);
  const ptReschedules = toNumber(inputs.ptReschedules);
  const totalUnitsBilled = toNumber(inputs.totalUnitsBilled);

  const ptVisitsSeen =
    ptScheduledVisits !== null
      ? ptScheduledVisits - (ptCancellations || 0) - (ptNoShows || 0)
      : null;

  const unitsPerVisit = safeDivide(totalUnitsBilled, ptVisitsSeen);
  const ptVisitsPerDay = safeDivide(ptVisitsSeen, workingDaysInWeek);
  const ptCxnsRate =
    ptScheduledVisits !== null
      ? safeDivide((ptCancellations || 0) + (ptNoShows || 0), ptScheduledVisits, 100)
      : null;

  const ptRescheduleRate =
    (ptCancellations || 0) + (ptNoShows || 0) > 0
      ? safeDivide(ptReschedules, (ptCancellations || 0) + (ptNoShows || 0), 100)
      : null;

  return {
    ptVisitsSeen,
    unitsPerVisit,
    ptVisitsPerDay,
    ptCxnsRate,
    ptRescheduleRate
  };
}

/* ------------------------------
   SHARED — CXNS workbook formulas
   ------------------------------ */

function getCxnsCalculatedValues(inputs = {}) {
  const scheduledAppts = toNumber(inputs.scheduledAppts);
  const cancellations = toNumber(inputs.cancellations);
  const noShows = toNumber(inputs.noShows);
  const reschedules = toNumber(inputs.reschedules);

  const cxnsRate =
    scheduledAppts !== null
      ? safeDivide((cancellations || 0) + (noShows || 0), scheduledAppts, 100)
      : null;

  const rescheduleRate =
    (cancellations || 0) + (noShows || 0) > 0
      ? safeDivide(reschedules, (cancellations || 0) + (noShows || 0), 100)
      : null;

  return {
    cxnsRate,
    rescheduleRate
  };
}

/* --------------------------------
   SHARED — Capacity workbook logic
   -------------------------------- */

function providerAvailableDays(providerType, providerCount, ptoDays, workingDaysInMonth, mdClinicDays, mdSurgeryDays, paClinicDays) {
  const type = String(providerType || "").trim().toUpperCase();
  const count = toNumber(providerCount) || 0;
  const pto = toNumber(ptoDays) || 0;

  let baseDays = toNumber(workingDaysInMonth) || 0;

  if (type === "MD CLINIC") baseDays = toNumber(mdClinicDays) || 0;
  else if (type === "MD SURGERY") baseDays = toNumber(mdSurgeryDays) || 0;
  else if (type === "PA CLINIC") baseDays = toNumber(paClinicDays) || 0;

  return Math.max(0, (count * baseDays) - pto);
}

function getCapacityCalculatedValues(inputs = {}) {
  const workingDaysInMonth = toNumber(inputs.workingDaysInMonth);
  const mdClinicDays = toNumber(inputs.mdClinicDays);
  const mdSurgeryDays = toNumber(inputs.mdSurgeryDays);
  const paClinicDays = toNumber(inputs.paClinicDays);

  let totalProviders = 0;
  let capacityVisits = 0;

  for (let i = 1; i <= 7; i += 1) {
    const providerType = inputs[`providerType${i}`];
    const providerCount = toNumber(inputs[`providerCount${i}`]) || 0;
    const ptoDays = toNumber(inputs[`ptoDays${i}`]) || 0;
    const visitsPerDay = toNumber(inputs[`visitsPerDay${i}`]) || 0;

    totalProviders += providerCount;

    const availableDays = providerAvailableDays(
      providerType,
      providerCount,
      ptoDays,
      workingDaysInMonth,
      mdClinicDays,
      mdSurgeryDays,
      paClinicDays
    );

    capacityVisits += availableDays * visitsPerDay;
  }

  const capacityVisitsPerDay = safeDivide(capacityVisits, workingDaysInMonth);

  return {
    totalProviders,
    capacityVisits,
    capacityVisitsPerDay
  };
}

/* -------------------------------------------
   SHARED — Productivity Builder workbook logic
   ------------------------------------------- */

function getProductivityCalculatedValues(inputs = {}) {
  const workingDaysInPeriod = toNumber(inputs.workingDaysInPeriod);
  const newPatients = toNumber(inputs.newPatients);
  const npToEstablishedMultiplier = toNumber(inputs.npToEstablishedMultiplier);
  const npToSurgeryConversion = toNumber(inputs.npToSurgeryConversion);
  const callsVolume = toNumber(inputs.callsVolume);
  const maVisitsPerDayPerFte = toNumber(inputs.maVisitsPerDayPerFte);
  const psaVisitsPerDayPerFte = toNumber(inputs.psaVisitsPerDayPerFte);
  const callsPerDayPerFte = toNumber(inputs.callsPerDayPerFte);
  const surgeriesPerDayPerSchedulerFte = toNumber(inputs.surgeriesPerDayPerSchedulerFte);
  const staffingBufferPct = toNumber(inputs.staffingBufferPct) || 0;

  const forecastEstablishedVisits =
    newPatients !== null && npToEstablishedMultiplier !== null
      ? newPatients * npToEstablishedMultiplier
      : null;

  const forecastSurgeries =
    newPatients !== null && npToSurgeryConversion !== null
      ? newPatients * npToSurgeryConversion
      : null;

  const totalVisits =
    newPatients !== null && forecastEstablishedVisits !== null
      ? newPatients + forecastEstablishedVisits
      : null;

  const npPerDay = safeDivide(newPatients, workingDaysInPeriod);
  const estPerDay = safeDivide(forecastEstablishedVisits, workingDaysInPeriod);
  const surgPerDay = safeDivide(forecastSurgeries, workingDaysInPeriod);
  const visitsPerDay = safeDivide(totalVisits, workingDaysInPeriod);
  const callsPerDay = safeDivide(callsVolume, workingDaysInPeriod);

  const maFte =
    visitsPerDay !== null && maVisitsPerDayPerFte
      ? ceilQuarter((visitsPerDay / maVisitsPerDayPerFte) * (1 + staffingBufferPct))
      : null;

  const psaFte =
    visitsPerDay !== null && psaVisitsPerDayPerFte
      ? ceilQuarter((visitsPerDay / psaVisitsPerDayPerFte) * (1 + staffingBufferPct))
      : null;

  const callFte =
    callsPerDay !== null && callsPerDayPerFte
      ? ceilQuarter((callsPerDay / callsPerDayPerFte) * (1 + staffingBufferPct))
      : null;

  const surgerySchedulerFte =
    surgPerDay !== null && surgeriesPerDayPerSchedulerFte
      ? ceilQuarter((surgPerDay / surgeriesPerDayPerSchedulerFte) * (1 + staffingBufferPct))
      : null;

  const totalOpsFte =
    (maFte || 0) +
    (psaFte || 0) +
    (callFte || 0) +
    (surgerySchedulerFte || 0);

  return {
    forecastEstablishedVisits,
    forecastSurgeries,
    totalVisits,
    npPerDay,
    estPerDay,
    surgPerDay,
    maFte,
    psaFte,
    callFte,
    surgerySchedulerFte,
    totalOpsFte
  };
}

/* --------------------------------
   Shared page routing for app usage
   -------------------------------- */

export function getSharedCalculatedValues(pageName, inputs = {}) {
  switch (pageName) {
    case "PT":
      return getPtCalculatedValues(inputs);
    case "CXNS":
      return getCxnsCalculatedValues(inputs);
    case "Capacity":
      return getCapacityCalculatedValues(inputs);
    case "Productivity Builder":
      return getProductivityCalculatedValues(inputs);
    default:
      return {};
  }
}

export function calculateSharedSummaries(pageName, inputs = {}) {
  switch (pageName) {
    case "PT": {
      const calc = getPtCalculatedValues(inputs);
      return [
        {
          label: "PT Visits Seen",
          value: formatWhole(calc.ptVisitsSeen),
          meta: "Scheduled - Cancellations - No Shows"
        },
        {
          label: "Units / Visit",
          value: formatDecimal(calc.unitsPerVisit, 2),
          meta: "Total Units Billed / PT Visits Seen"
        },
        {
          label: "Visits / Day",
          value: formatDecimal(calc.ptVisitsPerDay, 2),
          meta: "PT Visits Seen / Working Days in Week"
        },
        {
          label: "PT CX/NS %",
          value: formatPercent(calc.ptCxnsRate),
          meta: "(Cancellations + No Shows) / Scheduled"
        },
        {
          label: "PT Reschedule %",
          value: formatPercent(calc.ptRescheduleRate),
          meta: "Reschedules / (Cancellations + No Shows)"
        }
      ];
    }

    case "CXNS": {
      const calc = getCxnsCalculatedValues(inputs);
      return [
        {
          label: "CX/NS Rate",
          value: formatPercent(calc.cxnsRate),
          meta: "(Cancellations + No Shows) / Scheduled"
        },
        {
          label: "Reschedule Rate",
          value: formatPercent(calc.rescheduleRate),
          meta: "Reschedules / (Cancellations + No Shows)"
        }
      ];
    }

    case "Capacity": {
      const calc = getCapacityCalculatedValues(inputs);
      return [
        {
          label: "Total Providers",
          value: formatWhole(calc.totalProviders),
          meta: "Sum of provider counts"
        },
        {
          label: "Capacity Visits",
          value: formatWhole(calc.capacityVisits),
          meta: "Sum of available days × visits/day"
        },
        {
          label: "Capacity Visits / Day",
          value: formatDecimal(calc.capacityVisitsPerDay, 2),
          meta: "Capacity Visits / Working Days"
        }
      ];
    }

    case "Productivity Builder": {
      const calc = getProductivityCalculatedValues(inputs);
      return [
        {
          label: "Forecast Established Visits",
          value: formatWhole(calc.forecastEstablishedVisits),
          meta: "NP × NP→Est Multiplier"
        },
        {
          label: "Forecast Surgeries",
          value: formatWhole(calc.forecastSurgeries),
          meta: "NP × NP→Surgery Conversion"
        },
        {
          label: "Total Visits",
          value: formatWhole(calc.totalVisits),
          meta: "NP + Forecast Established Visits"
        },
        {
          label: "MA FTE",
          value: formatDecimal(calc.maFte, 2),
          meta: "Quarter-step CEILING with buffer"
        },
        {
          label: "Total Ops FTE",
          value: formatDecimal(calc.totalOpsFte, 2),
          meta: "MA + PSA + Call + Surgery Scheduler"
        }
      ];
    }

    default:
      return [];
  }
}
