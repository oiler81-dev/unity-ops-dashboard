export const ENTITIES = ["LAOSS", "NES", "SpineOne", "MRO"];

export const REGION_SECTIONS = [
  {
    key: "coreVolume",
    title: "Core Volume Metrics",
    description: "Weekly visit and patient volume inputs used in executive reporting.",
    entities: ENTITIES,
    fields: [
      {
        key: "visitVolume",
        label: "Visit Volume",
        type: "number",
        step: "1",
        placeholder: "Enter total weekly visits"
      },
      {
        key: "callVolume",
        label: "Call Volume",
        type: "number",
        step: "1",
        placeholder: "Enter total weekly calls"
      },
      {
        key: "newPatients",
        label: "New Patients",
        type: "number",
        step: "1",
        placeholder: "Enter total new patients"
      }
    ]
  },
  {
    key: "accessMetrics",
    title: "Access Metrics",
    description: "Scheduling and access indicators used for weekly operations review.",
    entities: ENTITIES,
    fields: [
      {
        key: "noShowRate",
        label: "No Show Rate (%)",
        type: "number",
        step: "0.01",
        placeholder: "Example: 5.4"
      },
      {
        key: "cancellationRate",
        label: "Cancellation Rate (%)",
        type: "number",
        step: "0.01",
        placeholder: "Example: 7.1"
      },
      {
        key: "abandonedCallRate",
        label: "Abandoned Call Rate (%)",
        type: "number",
        step: "0.01",
        placeholder: "Example: 3.2"
      }
    ]
  },
  {
    key: "operationalHealth",
    title: "Operational Health",
    description: "Supporting metrics for staffing, throughput, and service level review.",
    entities: ENTITIES,
    fields: [
      {
        key: "capacityUtilization",
        label: "Capacity Utilization (%)",
        type: "number",
        step: "0.01",
        placeholder: "Example: 91.5"
      },
      {
        key: "ptUnits",
        label: "PT Units",
        type: "number",
        step: "1",
        placeholder: "Enter PT units"
      },
      {
        key: "staffingNotesCount",
        label: "Staffing Variance Count",
        type: "number",
        step: "1",
        placeholder: "Enter staffing variance count"
      }
    ]
  }
];

export function getRegionSections(entity) {
  return REGION_SECTIONS.filter((section) => section.entities.includes(entity));
}

export function getAllMetricKeysForEntity(entity) {
  return getRegionSections(entity).flatMap((section) => section.fields.map((field) => field.key));
}
