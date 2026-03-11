function computeStatusColor(metricKey, value) {
  if (value === null || value === undefined || value === "") return "yellow";

  if (["noShowRate", "cancellationRate", "abandonedCallRate"].includes(metricKey)) {
    if (Number(value) <= 5) return "green";
    if (Number(value) <= 10) return "yellow";
    return "red";
  }

  if (["visitVolume", "callVolume", "newPatients"].includes(metricKey)) {
    if (Number(value) >= 100) return "green";
    if (Number(value) >= 50) return "yellow";
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

function buildExecutiveSample(weekEnding) {
  return {
    weekEnding,
    kpis: [
      { label: "Visit Volume", value: "1,482", meta: "All entities combined", status: "Stable", statusColor: "green" },
      { label: "Call Volume", value: "3,906", meta: "All entities combined", status: "Stable", statusColor: "green" },
      { label: "No Show Rate", value: "6.2%", meta: "Companywide", status: "Watch", statusColor: "yellow" },
      { label: "Abandoned Call Rate", value: "4.7%", meta: "Companywide", status: "Good", statusColor: "green" }
    ],
    entities: [
      { entity: "LAOSS", visitVolume: 412, callVolume: 1084, noShowRate: "5.1%", cancellationRate: "8.4%", abandonedCallRate: "3.8%", status: "Draft" },
      { entity: "NES", visitVolume: 318, callVolume: 752, noShowRate: "6.0%", cancellationRate: "7.8%", abandonedCallRate: "4.2%", status: "Submitted" },
      { entity: "SpineOne", visitVolume: 409, callVolume: 1010, noShowRate: "7.2%", cancellationRate: "9.2%", abandonedCallRate: "5.1%", status: "Draft" },
      { entity: "MRO", visitVolume: 343, callVolume: 1060, noShowRate: "6.4%", cancellationRate: "8.9%", abandonedCallRate: "5.6%", status: "Approved" }
    ]
  };
}

module.exports = {
  buildRegionKpis,
  buildExecutiveSample
};
