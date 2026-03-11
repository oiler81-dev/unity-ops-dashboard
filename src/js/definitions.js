export const ENTITIES = ["LAOSS", "NES", "SpineOne", "MRO"];

export const REGION_SECTIONS = [
  {
    key: "coreVolume",
    title: "Core Volume Metrics",
    description: "Weekly visit and patient volume inputs used in executive reporting.",
    entities: ENTITIES,
    fields: [
      { key: "visitVolume", label: "Visit Volume", type: "number", step: "1", placeholder: "Enter total weekly visits" },
      { key: "callVolume", label: "Call Volume", type: "number", step: "1", placeholder: "Enter total weekly calls" },
      { key: "newPatients", label: "New Patients", type: "number", step: "1", placeholder: "Enter total new patients" }
    ],
    calculatedFields: [
      { key: "callToVisitRatio", label: "Call-to-Visit Ratio", format: "decimal2" },
      { key: "newPatientMix", label: "New Patient Mix (%)", format: "percent1" }
    ]
  },
  {
    key: "accessMetrics",
    title: "Access Metrics",
    description: "Scheduling and access indicators used for weekly operations review.",
    entities: ENTITIES,
    fields: [
      { key: "noShowRate", label: "No Show Rate (%)", type: "number", step: "0.01", placeholder: "Example: 5.4" },
      { key: "cancellationRate", label: "Cancellation Rate (%)", type: "number", step: "0.01", placeholder: "Example: 7.1" },
      { key: "abandonedCallRate", label: "Abandoned Call Rate (%)", type: "number", step: "0.01", placeholder: "Example: 3.2" }
    ],
    calculatedFields: [
      { key: "keptAppointmentRate", label: "Kept Appointment Rate (%)", format: "percent1" },
      { key: "accessPressureScore", label: "Access Pressure Score", format: "decimal1" }
    ]
  },
  {
    key: "operationalHealth",
    title: "Operational Health",
    description: "Supporting metrics for staffing, throughput, and service level review.",
    entities: ENTITIES,
    fields: [
      { key: "capacityUtilization", label: "Capacity Utilization (%)", type: "number", step: "0.01", placeholder: "Example: 91.5" },
      { key: "ptUnits", label: "PT Units", type: "number", step: "1", placeholder: "Enter PT units" },
      { key: "staffingNotesCount", label: "Staffing Variance Count", type: "number", step: "1", placeholder: "Enter staffing variance count" }
    ],
    calculatedFields: [
      { key: "unusedCapacity", label: "Unused Capacity (%)", format: "percent1" },
      { key: "ptUnitsPerVisit", label: "PT Units per Visit", format: "decimal2" }
    ]
  },
  {
    key: "throughputPlanning",
    title: "Throughput Planning",
    description: "Additional planning inputs aligned to visit throughput and scheduling health.",
    entities: ENTITIES,
    fields: [
      { key: "scheduledVisits", label: "Scheduled Visits", type: "number", step: "1", placeholder: "Enter scheduled visits" },
      { key: "rescheduledVisits", label: "Rescheduled Visits", type: "number", step: "1", placeholder: "Enter rescheduled visits" },
      { key: "providerClinicDays", label: "Provider Clinic Days", type: "number", step: "0.01", placeholder: "Example: 18.5" }
    ],
    calculatedFields: [
      { key: "visitsPerClinicDay", label: "Visits per Clinic Day", format: "decimal2" },
      { key: "scheduleConversionRate", label: "Schedule Conversion Rate (%)", format: "percent1" }
    ]
  }
];

export const SHARED_PAGE_DEFINITIONS = {
  PT: {
    key: "PT",
    title: "PT",
    description: "Physical therapy productivity and throughput measures.",
    sections: [
      {
        key: "ptVolume",
        title: "PT Volume",
        description: "Weekly PT output and new evaluation tracking.",
        fields: [
          { key: "ptVisits", label: "PT Visits", type: "number", step: "1", placeholder: "Enter PT visits" },
          { key: "ptNewEvaluations", label: "PT New Evaluations", type: "number", step: "1", placeholder: "Enter new evals" },
          { key: "ptUnits", label: "PT Units", type: "number", step: "1", placeholder: "Enter PT units" }
        ],
        calculatedFields: [
          { key: "ptEvalMix", label: "PT Eval Mix (%)", format: "percent1" },
          { key: "ptUnitsPerVisit", label: "PT Units per Visit", format: "decimal2" }
        ]
      },
      {
        key: "ptProductivity",
        title: "PT Productivity",
        description: "Productivity metrics supporting utilization review.",
        fields: [
          { key: "ptVisitsPerProvider", label: "PT Visits per Provider", type: "number", step: "0.01", placeholder: "Example: 16.5" },
          { key: "ptUnitsPerVisit", label: "PT Units per Visit", type: "number", step: "0.01", placeholder: "Example: 4.1" }
        ],
        calculatedFields: [
          { key: "ptProductivityScore", label: "PT Productivity Score", format: "decimal2" }
        ]
      }
    ]
  },

  CXNS: {
    key: "CXNS",
    title: "CXNS",
    description: "Call center, cancellations, no-shows, and access-related performance.",
    sections: [
      {
        key: "cxnsVolume",
        title: "Call and Scheduling Volume",
        description: "Weekly volume for access and scheduling.",
        fields: [
          { key: "callVolume", label: "Call Volume", type: "number", step: "1", placeholder: "Enter call volume" },
          { key: "scheduledVisits", label: "Scheduled Visits", type: "number", step: "1", placeholder: "Enter scheduled visits" },
          { key: "newPatients", label: "New Patients", type: "number", step: "1", placeholder: "Enter new patients" }
        ],
        calculatedFields: [
          { key: "callConversionRate", label: "Call Conversion Rate (%)", format: "percent1" },
          { key: "newPatientRate", label: "New Patient Rate (%)", format: "percent1" }
        ]
      },
      {
        key: "cxnsAccess",
        title: "Access Quality",
        description: "Core access health metrics for leadership review.",
        fields: [
          { key: "noShowRate", label: "No Show Rate (%)", type: "number", step: "0.01", placeholder: "Example: 5.4" },
          { key: "cancellationRate", label: "Cancellation Rate (%)", type: "number", step: "0.01", placeholder: "Example: 7.1" },
          { key: "abandonedCallRate", label: "Abandoned Call Rate (%)", type: "number", step: "0.01", placeholder: "Example: 3.2" }
        ],
        calculatedFields: [
          { key: "keptAppointmentRate", label: "Kept Appointment Rate (%)", format: "percent1" },
          { key: "accessPressureScore", label: "Access Pressure Score", format: "decimal1" }
        ]
      }
    ]
  },

  Capacity: {
    key: "Capacity",
    title: "Capacity",
    description: "Capacity assumptions and throughput support metrics.",
    sections: [
      {
        key: "capacityInputs",
        title: "Capacity Inputs",
        description: "Weekly inputs used for capacity modeling.",
        fields: [
          { key: "providerClinicDays", label: "Provider Clinic Days", type: "number", step: "0.01", placeholder: "Example: 18.5" },
          { key: "availableVisitSlots", label: "Available Visit Slots", type: "number", step: "1", placeholder: "Enter available slots" },
          { key: "bookedVisitSlots", label: "Booked Visit Slots", type: "number", step: "1", placeholder: "Enter booked slots" }
        ],
        calculatedFields: [
          { key: "openSlots", label: "Open Slots", format: "whole" },
          { key: "visitsPerClinicDay", label: "Visits per Clinic Day", format: "decimal2" }
        ]
      },
      {
        key: "capacityUtilizationSection",
        title: "Capacity Utilization",
        description: "Utilization and fill-rate related values.",
        fields: [
          { key: "capacityUtilization", label: "Capacity Utilization (%)", type: "number", step: "0.01", placeholder: "Example: 91.5" },
          { key: "slotFillRate", label: "Slot Fill Rate (%)", type: "number", step: "0.01", placeholder: "Example: 88.2" }
        ],
        calculatedFields: [
          { key: "unusedCapacity", label: "Unused Capacity (%)", format: "percent1" }
        ]
      }
    ]
  },

  "Productivity Builder": {
    key: "Productivity Builder",
    title: "Productivity Builder",
    description: "Operational assumptions and supporting productivity calculations.",
    sections: [
      {
        key: "productivityInputs",
        title: "Productivity Inputs",
        description: "Inputs used to drive productivity assumptions and outputs.",
        fields: [
          { key: "providerCount", label: "Provider Count", type: "number", step: "1", placeholder: "Enter provider count" },
          { key: "clinicSupportFte", label: "Clinic Support FTE", type: "number", step: "0.01", placeholder: "Example: 12.5" },
          { key: "visitVolume", label: "Visit Volume", type: "number", step: "1", placeholder: "Enter visit volume" }
        ],
        calculatedFields: [
          { key: "visitsPerProviderAuto", label: "Visits per Provider (Auto)", format: "decimal2" },
          { key: "visitsPerSupportFteAuto", label: "Visits per Support FTE (Auto)", format: "decimal2" }
        ]
      },
      {
        key: "productivityOutputs",
        title: "Productivity Outputs",
        description: "Output-oriented operational productivity values.",
        fields: [
          { key: "visitsPerProvider", label: "Visits per Provider", type: "number", step: "0.01", placeholder: "Example: 84.5" },
          { key: "visitsPerSupportFte", label: "Visits per Support FTE", type: "number", step: "0.01", placeholder: "Example: 42.1" }
        ],
        calculatedFields: [
          { key: "productivityAlignmentScore", label: "Productivity Alignment Score", format: "decimal2" }
        ]
      }
    ]
  }
};

export function getRegionSections(entity) {
  return REGION_SECTIONS.filter((section) => section.entities.includes(entity));
}

export function getAllMetricKeysForEntity(entity) {
  return getRegionSections(entity).flatMap((section) => section.fields.map((field) => field.key));
}

export function getSharedPageDefinition(pageName) {
  return SHARED_PAGE_DEFINITIONS[pageName] || null;
}

export function getAllMetricKeysForSharedPage(pageName) {
  const def = getSharedPageDefinition(pageName);
  if (!def) return [];
  return def.sections.flatMap((section) => section.fields.map((field) => field.key));
}
