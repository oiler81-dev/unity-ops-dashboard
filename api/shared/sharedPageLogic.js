function buildSharedKpis(page, inputs = {}) {
  switch (page) {
    case "PT":
      return [
        { label: "PT Visits", value: inputs.ptVisits ?? "—", meta: "Weekly PT visits", status: "Tracking", statusColor: "yellow" },
        { label: "New Evaluations", value: inputs.ptNewEvaluations ?? "—", meta: "Weekly PT evals", status: "Tracking", statusColor: "yellow" },
        { label: "PT Units", value: inputs.ptUnits ?? "—", meta: "Weekly PT units", status: "Tracking", statusColor: "yellow" },
        { label: "Visits per Provider", value: inputs.ptVisitsPerProvider ?? "—", meta: "Productivity", status: "Tracking", statusColor: "yellow" }
      ];

    case "CXNS":
      return [
        { label: "Call Volume", value: inputs.callVolume ?? "—", meta: "Weekly calls", status: "Tracking", statusColor: "yellow" },
        { label: "Scheduled Visits", value: inputs.scheduledVisits ?? "—", meta: "Scheduled volume", status: "Tracking", statusColor: "yellow" },
        { label: "No Show Rate", value: inputs.noShowRate !== undefined && inputs.noShowRate !== null ? `${inputs.noShowRate}%` : "—", meta: "Access", status: "Tracking", statusColor: "yellow" },
        { label: "Abandoned Call Rate", value: inputs.abandonedCallRate !== undefined && inputs.abandonedCallRate !== null ? `${inputs.abandonedCallRate}%` : "—", meta: "Call center health", status: "Tracking", statusColor: "yellow" }
      ];

    case "Capacity":
      return [
        { label: "Available Slots", value: inputs.availableVisitSlots ?? "—", meta: "Capacity", status: "Tracking", statusColor: "yellow" },
        { label: "Booked Slots", value: inputs.bookedVisitSlots ?? "—", meta: "Utilized", status: "Tracking", statusColor: "yellow" },
        { label: "Capacity Utilization", value: inputs.capacityUtilization !== undefined && inputs.capacityUtilization !== null ? `${inputs.capacityUtilization}%` : "—", meta: "Utilization", status: "Tracking", statusColor: "yellow" },
        { label: "Slot Fill Rate", value: inputs.slotFillRate !== undefined && inputs.slotFillRate !== null ? `${inputs.slotFillRate}%` : "—", meta: "Fill rate", status: "Tracking", statusColor: "yellow" }
      ];

    case "Productivity Builder":
      return [
        { label: "Provider Count", value: inputs.providerCount ?? "—", meta: "Base staffing", status: "Tracking", statusColor: "yellow" },
        { label: "Support FTE", value: inputs.clinicSupportFte ?? "—", meta: "Support model", status: "Tracking", statusColor: "yellow" },
        { label: "Visits per Provider", value: inputs.visitsPerProvider ?? "—", meta: "Productivity", status: "Tracking", statusColor: "yellow" },
        { label: "Visits per Support FTE", value: inputs.visitsPerSupportFte ?? "—", meta: "Leverage", status: "Tracking", statusColor: "yellow" }
      ];

    default:
      return [];
  }
}

module.exports = { buildSharedKpis };
