const ENTITIES = ["LAOSS", "NES", "SpineOne", "MRO"];

function toNumber(value) {
  const n = Number(value);
  return Number.isFinite(n) ? n : null;
}

function formatPercent(value) {
  return value === null || value === undefined ? "—" : `${Number(value).toFixed(1)}%`;
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

function buildRegionKpis(inputs = {}) {
  return [
    {
      label: "Visit Volume",
      value: inputs.visitVolume ?? "—",
      meta: "Weekly total visits",
      status: "Tracking",
      statusColor: computeStatusColor("visitVolume", inputs.visitVolume)
    },
    {
      label: "Call Volume",
      value: inputs.callVolume ?? "—",
      meta: "Weekly total calls",
      status: "Tracking",
      statusColor: computeStatusColor("callVolume", inputs.callVolume)
    },
    {
      label: "No Show Rate",
      value: inputs.noShowRate !== null && inputs.noShowRate !== undefined ? `${inputs.noShowRate}%` : "—",
      meta: "Weekly no show percentage",
      status: "Tracking",
      statusColor: computeStatusColor("noShowRate", inputs.noShowRate)
    },
    {
      label: "Abandoned Call Rate",
      value: inputs.abandonedCallRate !== null && inputs.abandonedCallRate !== undefined ? `${inputs.abandonedCallRate}%` : "—",
      meta: "Weekly abandoned call percentage",
      status: "Tracking",
      statusColor: computeStatusColor("abandonedCallRate", inputs.abandonedCallRate)
    }
  ];
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

function buildExecutiveFromRows(weekEnding, entityRows) {
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

  return {
    weekEnding,
    kpis: [
      {
        label: "Visit Volume",
        value: totals.visitVolume ? totals.visitVolume.toLocaleString() : "—",
        meta: "All entities combined",
        status: "Live",
        statusColor: computeStatusColor("visitVolume", totals.visitVolume)
      },
      {
        label: "Call Volume",
        value: totals.callVolume ? totals.callVolume.toLocaleString() : "—",
        meta: "All entities combined",
        status: "Live",
        statusColor: computeStatusColor("callVolume", totals.callVolume)
      },
      {
        label: "No Show Rate",
        value: avgNoShowRate !== null ? `${avgNoShowRate.toFixed(1)}%` : "—",
        meta: "Average across entities",
        status: "Live",
        statusColor: computeStatusColor("noShowRate", avgNoShowRate)
      },
      {
        label: "Abandoned Call Rate",
        value: avgAbandonedCallRate !== null ? `${avgAbandonedCallRate.toFixed(1)}%` : "—",
        meta: "Average across entities",
        status: "Live",
        statusColor: computeStatusColor("abandonedCallRate", avgAbandonedCallRate)
      }
    ],
    entities: entityRows.map((row) => normalizeEntityRow(row.entity, row.raw, row.status))
  };
}

module.exports = {
  ENTITIES,
  buildRegionKpis,
  buildExecutiveFromRows
};
