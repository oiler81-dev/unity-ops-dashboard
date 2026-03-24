const ENTITIES = ["LAOSS", "NES", "SpineOne", "MRO"];

const ENTITY_LABELS = {
  LAOSS: "Los Angeles Orthopedic Surgery Specialists",
  NES: "Northwest Extremity Specialists",
  SpineOne: "SpineOne",
  MRO: "Midland Riverside Orthopedics"
};

const KPI_FIELDS = [
  { key: "visitVolume", label: "Visit Volume", type: "number" },
  { key: "callVolume", label: "Call Volume", type: "number" },
  { key: "newPatients", label: "New Patients", type: "number" },
  { key: "noShowRate", label: "No Show Rate", type: "number" },
  { key: "cancellationRate", label: "Cancellation Rate", type: "number" },
  { key: "abandonedCallRate", label: "Abandoned Call Rate", type: "number" }
];

const REGION_USER_MAP = {
  "tony.guerrero@laorthos.com": { entity: "LAOSS", role: "regional" },

  "annette@nespecialists.com": { entity: "NES", role: "regional" },
  "marketa.stuck@nespecialists.com": { entity: "NES", role: "regional" },

  "chris.zamucen@spineone.com": { entity: "SpineOne", role: "regional" },
  "lauren.bradley@spineone.com": { entity: "SpineOne", role: "regional" },

  "greg.lundgren@mrorthopedics.com": { entity: "MRO", role: "regional" },

  "nperez@unitymsk.com": { entity: "admin", role: "admin" },
  "tessa.kelley@spineone.com": { entity: "admin", role: "admin" }
};

const WEEKLY_TABLE = "WeeklyRegionData";

module.exports = {
  ENTITIES,
  ENTITY_LABELS,
  KPI_FIELDS,
  REGION_USER_MAP,
  WEEKLY_TABLE
};
