const ENTITIES = ["LAOSS", "NES", "SpineOne", "MRO"];

function toNumber(value) {
  const n = Number(value);
  return Number.isFinite(n) ? n : null;
}

function formatPercent(value) {
  return value === null || value === undefined ? "—" : `${Number(value).toFixed(1)}%`;
}

function formatWhole(value) {
  return value === null || value === undefined ? "—" : Number(value).toLocaleString();
}

function formatValue(metricKey, value) {
  if (value === null || value === undefined || value === "") return "—";

  if (["noShowRate", "cancellationRate", "abandonedCallRate", "capacityUtilization", "slotFillRate"].includes(metricKey)) {
    return `${Number(value).toFixed(1)}%`;
  }

  if (typeof value === "number") {
    return Number(value).toLocaleString();
  }

  return String(value);
}

function computeStatusColor(metricKey, value) {
  if (value === null || value === undefined || value === "") return "yellow";

  if (["noShowRate", "cancellationRate", "abandonedCallRate"].includes(metricKey)) {
    if (Number(value) <= 5) return "green";
    if (Number(value) <= 10) return "yellow";
    return "red";
  }

  if (["visitVolume", "callVolume", "newPatients", "ptUnits"].includes(metricKey)) {
    if (Number(value) >= 100) return "green";
    if (Number(value) >= 50) return "yellow";
    return "red";
  }

  if (metricKey === "capacityUtilization") {
    if (Number(value) >= 90) return "green";
    if (Number(value) >= 75) return "yellow";
    return "red";
  }

  return "yellow";
}

function computeThresholdStatus(metricKey, value, threshold) {
  if (value === null || value === undefined || !threshold) {
    return computeStatusColor(metricKey, value);
  }

  const numericValue = Number(value);
  if (!Number.isFinite(numericValue)) return "yellow";

  const comparisonType = String(threshold.comparisonType || "").toLowerCase();

  if (comparisonType === "higher_better") {
    const greenMin = toNumber(threshold.greenMin);
    const yellowMin = toNumber(threshold.yellowMin);

    if (greenMin !== null && numericValue >= greenMin) return "green";
    if (yellowMin !== null && numericValue >= yellowMin) return "yellow";
    return "red";
  }

  if (comparisonType === "lower_better") {
    const greenMax = toNumber(threshold.greenMax);
    const yellowMax = toNumber(threshold.yellowMax);

    if (greenMax !== null && numericValue <= greenMax) return "green";
    if (yellowMax !== null && numericValue <= yellowMax) return "yellow";
    return "red";
  }

  return computeStatusColor(metricKey, numericValue);
}

function computeVariance(metricKey, value, targetValue) {
  const numericValue = toNumber(value);
  const numericTarget = toNumber(targetValue);

  if (numericValue === null || numericTarget === null) return null;

  if (["noShowRate", "cancellationRate", "abandonedCallRate"].includes(metricKey)) {
    return numericValue - numericTarget;
  }

  return numericValue - numericTarget;
}

function buildMetricCard({ metricKey, label, value, targetValue, threshold, fallbackMeta }) {
  const statusColor = computeThresholdStatus(metricKey, value, threshold);
  const variance = computeVariance(metricKey, value, targetValue);

  let meta = fallbackMeta || "";
  if (targetValue !== null && targetValue !== undefined && targetValue !== "") {
    const varianceText =
      variance === null
        ? "—"
        : ["noShowRate", "cancellationRate", "abandonedCallRate"].includes(metricKey)
          ? `${variance >= 0 ? "+" : ""}${variance.toFixed(1)} pts vs target`
          : `${variance >= 0 ? "+" : ""}${variance.toLocaleString()} vs target`;

    meta = `Target: ${formatValue(metricKey, targetValue)} • ${varianceText}`;
  }

  return {
    label,
    value: formatValue(metricKey, value),
    rawValue: value,
    targetValue,
    variance,
    meta,
    status: "Tracking",
    statusColor
  };
}

function buildRegionKpis(inputs = {}, referenceMap = {}) {
  const metrics = [
    { metricKey: "visitVolume", label: "Visit Volume", fallbackMeta: "Weekly total visits" },
    { metricKey: "callVolume", label: "Call Volume", fallbackMeta: "Weekly total calls" },
    { metricKey: "noShowRate", label: "No Show Rate", fallbackMeta: "Weekly no show percentage" },
    { metricKey: "abandonedCallRate", label: "Abandoned Call Rate", fallbackMeta: "Weekly abandoned call percentage" }
  ];

  return metrics.map((metric) => {
    const ref = referenceMap[metric.metricKey] || {};
    return buildMetricCard({
      metricKey: metric.metricKey,
      label: metric.label,
      value: inputs[metric.metricKey],
      targetValue: ref.targetValue,
      threshold: ref.threshold,
      fallbackMeta: metric.fallbackMeta
    });
  });
}

function normalizeEntityRow(entity, inputMap = {}, status = "Draft") {
  const visitVolume = toNumber(inputMap.visitVolume);
  const callVolume = toNumber(inputMap.callVolume);
  const noShowRate = toNumber(inputMap.noShowRate);
  const cancellationRate = toNumber(inputMap.cancellationRate);
  const abandonedCallRate = toNumber(inputMap.abandonedCallRate);

  return {
    entity,
    visitVolume: visitVolume !== null ? visitVolume.toLocaleString() : "—",
    callVolume: callVolume !== null ? callVolume.toLocaleString() : "—",
    noShowRate: formatPercent(noShowRate),
    cancellationRate: formatPercent(cancellationRate),
    abandonedCallRate: formatPercent(abandonedCallRate),
    status
  };
}

function buildExecutiveFromRows(weekEnding, entityRows, entityReferenceMaps = {}) {
  const totals = entityRows.reduce(
    (acc, row) => {
      const visitVolume = toNumber(row.raw?.visitVolume);
      const callVolume = toNumber(row.raw?.callVolume);
      const noShowRate = toNumber(row.raw?.noShowRate);
      const abandonedCallRate = toNumber(row.raw?.abandonedCallRate);

      if (visitVolume !== null) acc.visitVolume += visitVolume;
      if (callVolume !== null) acc.callVolume += callVolume;
      if (noShowRate !== null) {
        acc.noShowRateSum += noShowRate;
        acc.noShowRateCount += 1;
      }
      if (abandonedCallRate !== null) {
        acc.abandonedCallRateSum += abandonedCallRate;
        acc.abandonedCallRateCount += 1;
      }
      return acc;
    },
    {
      visitVolume: 0,
      callVolume: 0,
      noShowRateSum: 0,
      noShowRateCount: 0,
      abandonedCallRateSum: 0,
      abandonedCallRateCount: 0
    }
  );

  const avgNoShowRate =
    totals.noShowRateCount > 0 ? totals.noShowRateSum / totals.noShowRateCount : null;

  const avgAbandonedCallRate =
    totals.abandonedCallRateCount > 0
      ? totals.abandonedCallRateSum / totals.abandonedCallRateCount
      : null;

  const firstEntityRef = entityReferenceMaps[ENTITIES[0]] || {};

  return {
    weekEnding,
    kpis: [
      buildMetricCard({
        metricKey: "visitVolume",
        label: "Visit Volume",
        value: totals.visitVolume || null,
        targetValue: null,
        threshold: null,
        fallbackMeta: "All entities combined"
      }),
      buildMetricCard({
        metricKey: "callVolume",
        label: "Call Volume",
        value: totals.callVolume || null,
        targetValue: null,
        threshold: null,
        fallbackMeta: "All entities combined"
      }),
      buildMetricCard({
        metricKey: "noShowRate",
        label: "No Show Rate",
        value: avgNoShowRate,
        targetValue: firstEntityRef.noShowRate?.targetValue ?? null,
        threshold: firstEntityRef.noShowRate?.threshold ?? null,
        fallbackMeta: "Average across entities"
      }),
      buildMetricCard({
        metricKey: "abandonedCallRate",
        label: "Abandoned Call Rate",
        value: avgAbandonedCallRate,
        targetValue: firstEntityRef.abandonedCallRate?.targetValue ?? null,
        threshold: firstEntityRef.abandonedCallRate?.threshold ?? null,
        fallbackMeta: "Average across entities"
      })
    ],
    entities: entityRows.map((row) => normalizeEntityRow(row.entity, row.raw, row.status))
  };
}

module.exports = {
  ENTITIES,
  toNumber,
  formatValue,
  computeThresholdStatus,
  computeVariance,
  buildMetricCard,
  buildRegionKpis,
  buildExecutiveFromRows
};
