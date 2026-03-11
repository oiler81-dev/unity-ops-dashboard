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
    {
      label: "Call-to-Visit Ratio",
      value: callToVisitRatio !== null ? callToVisitRatio.toFixed(2) : "—",
      meta: "Calls divided by visits"
    },
    {
      label: "Access Pressure Score",
      value: Number.isFinite(accessPressureScore) ? accessPressureScore.toFixed(1) : "—",
      meta: "No show + cancellation + abandoned call"
    },
    {
      label: "New Patient Mix",
      value: visitToNewPatientMix !== null ? `${visitToNewPatientMix.toFixed(1)}%` : "—",
      meta: "New patients as a share of visits"
    },
    {
      label: "Capacity Utilization",
      value: formatPercent(capacityUtilization),
      meta: "Region-level utilization"
    },
    {
      label: "PT Units",
      value: formatWhole(ptUnits),
      meta: "Weekly PT output"
    }
  ];
}
