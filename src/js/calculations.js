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

export function calculateRegionSummaries(inputs = {}) {
  const visitVolume = toNumber(inputs.visitVolume);
  const callVolume = toNumber(inputs.callVolume);
  const newPatients = toNumber(inputs.newPatients);
  const noShowRate = toNumber(inputs.noShowRate);
  const cancellationRate = toNumber(inputs.cancellationRate);
  const abandonedCallRate = toNumber(inputs.abandonedCallRate);
  const capacityUtilization = toNumber(inputs.capacityUtilization);
  const ptUnits = toNumber(inputs.ptUnits);

  const callToVisitRatio =
    visitVolume && callVolume
      ? callVolume / Math.max(visitVolume, 1)
      : null;

  const accessPressureScore =
    [noShowRate, cancellationRate, abandonedCallRate]
      .filter((v) => Number.isFinite(v))
      .reduce((sum, v) => sum + v, 0);

  const visitToNewPatientMix =
    visitVolume && newPatients
      ? (newPatients / Math.max(visitVolume, 1)) * 100
      : null;

  return [
    { label: "Call-to-Visit Ratio", value: callToVisitRatio !== null ? callToVisitRatio.toFixed(2) : "—", meta: "Calls divided by visits" },
    { label: "Access Pressure Score", value: Number.isFinite(accessPressureScore) ? accessPressureScore.toFixed(1) : "—", meta: "No show + cancellation + abandoned call" },
    { label: "New Patient Mix", value: visitToNewPatientMix !== null ? `${visitToNewPatientMix.toFixed(1)}%` : "—", meta: "New patients as a share of visits" },
    { label: "Capacity Utilization", value: formatPercent(capacityUtilization), meta: "Region-level utilization" },
    { label: "PT Units", value: formatWhole(ptUnits), meta: "Weekly PT output" }
  ];
}

export function calculateSharedSummaries(pageName, inputs = {}) {
  switch (pageName) {
    case "PT": {
      const ptVisits = toNumber(inputs.ptVisits);
      const ptNewEvaluations = toNumber(inputs.ptNewEvaluations);
      const ptUnits = toNumber(inputs.ptUnits);
      const ptVisitsPerProvider = toNumber(inputs.ptVisitsPerProvider);
      const ptUnitsPerVisit = toNumber(inputs.ptUnitsPerVisit);

      return [
        { label: "PT Visits", value: formatWhole(ptVisits), meta: "Weekly PT visits" },
        { label: "New Evaluations", value: formatWhole(ptNewEvaluations), meta: "Weekly PT eval volume" },
        { label: "PT Units", value: formatWhole(ptUnits), meta: "Weekly PT units" },
        { label: "Visits per Provider", value: ptVisitsPerProvider !== null ? ptVisitsPerProvider.toFixed(2) : "—", meta: "PT productivity" },
        { label: "Units per Visit", value: ptUnitsPerVisit !== null ? ptUnitsPerVisit.toFixed(2) : "—", meta: "Efficiency signal" }
      ];
    }

    case "CXNS": {
      const callVolume = toNumber(inputs.callVolume);
      const scheduledVisits = toNumber(inputs.scheduledVisits);
      const noShowRate = toNumber(inputs.noShowRate);
      const cancellationRate = toNumber(inputs.cancellationRate);
      const abandonedCallRate = toNumber(inputs.abandonedCallRate);

      const conversionRate =
        callVolume && scheduledVisits
          ? (scheduledVisits / Math.max(callVolume, 1)) * 100
          : null;

      return [
        { label: "Call Volume", value: formatWhole(callVolume), meta: "Weekly call volume" },
        { label: "Scheduled Visits", value: formatWhole(scheduledVisits), meta: "Weekly scheduled visits" },
        { label: "Call Conversion", value: conversionRate !== null ? `${conversionRate.toFixed(1)}%` : "—", meta: "Scheduled visits / calls" },
        { label: "No Show Rate", value: formatPercent(noShowRate), meta: "Access reliability" },
        { label: "Abandoned Call Rate", value: formatPercent(abandonedCallRate), meta: "Call center health" }
      ];
    }

    case "Capacity": {
      const availableVisitSlots = toNumber(inputs.availableVisitSlots);
      const bookedVisitSlots = toNumber(inputs.bookedVisitSlots);
      const capacityUtilization = toNumber(inputs.capacityUtilization);
      const slotFillRate = toNumber(inputs.slotFillRate);
      const providerClinicDays = toNumber(inputs.providerClinicDays);

      const openSlots =
        Number.isFinite(availableVisitSlots) && Number.isFinite(bookedVisitSlots)
          ? availableVisitSlots - bookedVisitSlots
          : null;

      return [
        { label: "Clinic Days", value: formatWhole(providerClinicDays), meta: "Provider clinic days" },
        { label: "Available Slots", value: formatWhole(availableVisitSlots), meta: "Total visit capacity" },
        { label: "Open Slots", value: formatWhole(openSlots), meta: "Unbooked opportunity" },
        { label: "Capacity Utilization", value: formatPercent(capacityUtilization), meta: "Utilization signal" },
        { label: "Slot Fill Rate", value: formatPercent(slotFillRate), meta: "Booked / available" }
      ];
    }

    case "Productivity Builder": {
      const providerCount = toNumber(inputs.providerCount);
      const clinicSupportFte = toNumber(inputs.clinicSupportFte);
      const visitVolume = toNumber(inputs.visitVolume);
      const visitsPerProvider = toNumber(inputs.visitsPerProvider);
      const visitsPerSupportFte = toNumber(inputs.visitsPerSupportFte);

      return [
        { label: "Provider Count", value: formatWhole(providerCount), meta: "Active provider base" },
        { label: "Support FTE", value: clinicSupportFte !== null ? clinicSupportFte.toFixed(2) : "—", meta: "Clinic support staffing" },
        { label: "Visit Volume", value: formatWhole(visitVolume), meta: "Weekly visit volume" },
        { label: "Visits per Provider", value: visitsPerProvider !== null ? visitsPerProvider.toFixed(2) : "—", meta: "Provider productivity" },
        { label: "Visits per Support FTE", value: visitsPerSupportFte !== null ? visitsPerSupportFte.toFixed(2) : "—", meta: "Support leverage" }
      ];
    }

    default:
      return [];
  }
}
