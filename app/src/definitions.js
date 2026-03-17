export const ENTITIES = ["LAOSS", "NES", "SpineOne", "MRO"];

export const ENTITY_LABELS = {
  LAOSS: "Los Angeles Orthopedic Surgery Specialists",
  NES: "Northwest Extremity Specialists",
  SpineOne: "SpineOne",
  MRO: "Midland Riverside Orthopedics"
};

export const REGION_TO_WORKBOOK_LABEL = {
  LAOSS: "Chicago",
  NES: "Portland",
  SpineOne: "SpineOne",
  MRO: "LAOSS"
};

export const KPI_METRICS = [
  { key: "totalVisits", label: "Total Visits", format: "whole" },
  { key: "visitsPerDay", label: "Visits / Day", format: "decimal2" },
  { key: "npPerDay", label: "NP / Day", format: "decimal2" },
  { key: "abandonmentRate", label: "Abandonment Rate", format: "percent1" },
  { key: "npToEstablishedConversion", label: "NP → Established", format: "percent1" },
  { key: "npToSurgeryConversion", label: "NP → Surgery", format: "percent1" }
];

export const REGION_SECTIONS = [
  {
    key: "periodContext",
    title: "Period Context",
    description: "Workbook-aligned weekly period inputs.",
    entities: ENTITIES,
    fields: [
      {
        key: "daysInPeriod",
        label: "Days in Period",
        type: "number",
        step: "1",
        placeholder: "Example: 5"
      },
      {
        key: "monthTag",
        label: "Month Tag",
        type: "text",
        placeholder: "Example: Mar"
      },
      {
        key: "workingDaysInMonth",
        label: "Working Days (Month)",
        type: "number",
        step: "1",
        placeholder: "Example: 22"
      },
      {
        key: "ptVisitsSeen",
        label: "PT Visits Seen (linked/imported)",
        type: "number",
        step: "1",
        placeholder: "PT weekly visits seen for this region"
      }
    ],
    calculatedFields: []
  },
  {
    key: "volumeInputs",
    title: "Weekly Volume Inputs",
    description: "Workbook weekly entry fields used by region formulas.",
    entities: ENTITIES,
    fields: [
      {
        key: "npActual",
        label: "NP Actual",
        type: "number",
        step: "1",
        placeholder: "Enter new patients"
      },
      {
        key: "establishedActual",
        label: "Established Actual",
        type: "number",
        step: "1",
        placeholder: "Enter established visits"
      },
      {
        key: "surgeryActual",
        label: "Surgery Actual",
        type: "number",
        step: "1",
        placeholder: "Enter surgery count"
      },
      {
        key: "cashActual",
        label: "Cash",
        type: "number",
        step: "1",
        placeholder: "Enter weekly cash"
      }
    ],
    calculatedFields: [
      {
        key: "totalVisits",
        label: "Total Visits",
        format: "whole"
      },
      {
        key: "visitsPerDay",
        label: "Visits / Day",
        format: "decimal2"
      },
      {
        key: "npPerDay",
        label: "NP / Day",
        format: "decimal2"
      }
    ]
  },
  {
    key: "accessAndConversions",
    title: "Calls, Access, and Conversion",
    description: "Workbook-aligned call, abandonment, and conversion formulas.",
    entities: ENTITIES,
    fields: [
      {
        key: "totalCalls",
        label: "Total Calls",
        type: "number",
        step: "1",
        placeholder: "Enter total calls"
      },
      {
        key: "abandonedCalls",
        label: "Abandoned Calls",
        type: "number",
        step: "1",
        placeholder: "Enter abandoned calls"
      }
    ],
    calculatedFields: [
      {
        key: "abandonmentRate",
        label: "Abandonment Rate",
        format: "percent1"
      },
      {
        key: "answeredCallRate",
        label: "Answered Call Rate",
        format: "percent1"
      },
      {
        key: "npToEstablishedConversion",
        label: "NP → Established Conversion",
        format: "percent1"
      },
      {
        key: "npToSurgeryConversion",
        label: "NP → Surgery Conversion",
        format: "percent1"
      }
    ]
  },
  {
    key: "varianceAndTargets",
    title: "Budget / Variance Support",
    description: "Workbook-aligned variance drivers.",
    entities: ENTITIES,
    fields: [
      {
        key: "monthlyBudgetNp",
        label: "Monthly Budget NP",
        type: "number",
        step: "1",
        placeholder: "Reference/imported"
      },
      {
        key: "monthlyBudgetEstablished",
        label: "Monthly Budget Established",
        type: "number",
        step: "1",
        placeholder: "Reference/imported"
      },
      {
        key: "monthlyBudgetSurgery",
        label: "Monthly Budget Surgery",
        type: "number",
        step: "1",
        placeholder: "Reference/imported"
      }
    ],
    calculatedFields: [
      {
        key: "npVariance",
        label: "NP Variance",
        format: "whole"
      },
      {
        key: "establishedVariance",
        label: "Established Variance",
        format: "whole"
      },
      {
        key: "surgeryVariance",
        label: "Surgery Variance",
        format: "whole"
      }
    ]
  }
];

export const SHARED_PAGE_DEFINITIONS = {
  PT: {
    key: "PT",
    title: "PT",
    description: "Workbook-aligned PT scheduling, visits seen, units, and access measures.",
    sections: [
      {
        key: "ptInputs",
        title: "PT Weekly Entry",
        description: "Direct PT workbook input fields.",
        fields: [
          {
            key: "monthTag",
            label: "Month Tag",
            type: "text",
            placeholder: "Example: Mar"
          },
          {
            key: "workingDaysInWeek",
            label: "Working Days in Week",
            type: "number",
            step: "1",
            placeholder: "Example: 5"
          },
          {
            key: "ptScheduledVisits",
            label: "PT Scheduled Visits",
            type: "number",
            step: "1",
            placeholder: "Enter scheduled visits"
          },
          {
            key: "ptCancellations",
            label: "Cancellations",
            type: "number",
            step: "1",
            placeholder: "Enter cancellations"
          },
          {
            key: "ptNoShows",
            label: "No Shows",
            type: "number",
            step: "1",
            placeholder: "Enter no shows"
          },
          {
            key: "ptReschedules",
            label: "Reschedules",
            type: "number",
            step: "1",
            placeholder: "Enter reschedules"
          },
          {
            key: "totalUnitsBilled",
            label: "Total Units Billed",
            type: "number",
            step: "1",
            placeholder: "Enter units billed"
          }
        ],
        calculatedFields: [
          {
            key: "ptVisitsSeen",
            label: "PT Visits Seen",
            format: "whole"
          },
          {
            key: "unitsPerVisit",
            label: "Units / Visit",
            format: "decimal2"
          },
          {
            key: "ptVisitsPerDay",
            label: "Visits / Day",
            format: "decimal2"
          },
          {
            key: "ptCxnsRate",
            label: "PT CX/NS %",
            format: "percent1"
          },
          {
            key: "ptRescheduleRate",
            label: "PT Reschedule %",
            format: "percent1"
          }
        ]
      }
    ]
  },

  CXNS: {
    key: "CXNS",
    title: "CXNS",
    description: "Workbook-aligned cancellations, no-shows, and reschedules.",
    sections: [
      {
        key: "cxnsInputs",
        title: "CXNS Weekly Entry",
        description: "Direct CXNS workbook input fields.",
        fields: [
          {
            key: "monthTag",
            label: "Month Tag",
            type: "text",
            placeholder: "Example: Mar"
          },
          {
            key: "scheduledAppts",
            label: "Scheduled Appts",
            type: "number",
            step: "1",
            placeholder: "Enter scheduled appointments"
          },
          {
            key: "cancellations",
            label: "Cancellations",
            type: "number",
            step: "1",
            placeholder: "Enter cancellations"
          },
          {
            key: "noShows",
            label: "No Shows",
            type: "number",
            step: "1",
            placeholder: "Enter no shows"
          },
          {
            key: "reschedules",
            label: "Reschedules",
            type: "number",
            step: "1",
            placeholder: "Enter reschedules"
          }
        ],
        calculatedFields: [
          {
            key: "cxnsRate",
            label: "CX/NS Rate",
            format: "percent1"
          },
          {
            key: "rescheduleRate",
            label: "Reschedule Rate",
            format: "percent1"
          }
        ]
      }
    ]
  },

  Capacity: {
    key: "Capacity",
    title: "Capacity",
    description: "Workbook-aligned capacity builder with seven provider rows.",
    sections: [
      {
        key: "capacityPeriod",
        title: "Capacity Period Inputs",
        description: "Monthly context inputs.",
        fields: [
          {
            key: "monthTag",
            label: "Month",
            type: "text",
            placeholder: "Example: Jan"
          },
          {
            key: "workingDaysInMonth",
            label: "Working Days",
            type: "number",
            step: "1",
            placeholder: "Example: 22"
          },
          {
            key: "mdClinicDays",
            label: "MD Clinic Days",
            type: "number",
            step: "0.01",
            placeholder: "Enter MD clinic days"
          },
          {
            key: "mdSurgeryDays",
            label: "MD Surgery Days",
            type: "number",
            step: "0.01",
            placeholder: "Enter MD surgery days"
          },
          {
            key: "paClinicDays",
            label: "PA Clinic Days",
            type: "number",
            step: "0.01",
            placeholder: "Enter PA clinic days"
          }
        ],
        calculatedFields: []
      },
      {
        key: "capacityProviderRows",
        title: "Provider Capacity Inputs",
        description: "Seven workbook-style provider rows.",
        fields: [
          { key: "providerType1", label: "Provider Type 1", type: "text", placeholder: "MD Clinic / MD Surgery / PA Clinic / Other" },
          { key: "providerCount1", label: "Provider Count 1", type: "number", step: "1", placeholder: "0" },
          { key: "ptoDays1", label: "PTO Days 1", type: "number", step: "0.01", placeholder: "0" },
          { key: "visitsPerDay1", label: "Visits / Day 1", type: "number", step: "0.01", placeholder: "0" },

          { key: "providerType2", label: "Provider Type 2", type: "text", placeholder: "Provider type" },
          { key: "providerCount2", label: "Provider Count 2", type: "number", step: "1", placeholder: "0" },
          { key: "ptoDays2", label: "PTO Days 2", type: "number", step: "0.01", placeholder: "0" },
          { key: "visitsPerDay2", label: "Visits / Day 2", type: "number", step: "0.01", placeholder: "0" },

          { key: "providerType3", label: "Provider Type 3", type: "text", placeholder: "Provider type" },
          { key: "providerCount3", label: "Provider Count 3", type: "number", step: "1", placeholder: "0" },
          { key: "ptoDays3", label: "PTO Days 3", type: "number", step: "0.01", placeholder: "0" },
          { key: "visitsPerDay3", label: "Visits / Day 3", type: "number", step: "0.01", placeholder: "0" },

          { key: "providerType4", label: "Provider Type 4", type: "text", placeholder: "Provider type" },
          { key: "providerCount4", label: "Provider Count 4", type: "number", step: "1", placeholder: "0" },
          { key: "ptoDays4", label: "PTO Days 4", type: "number", step: "0.01", placeholder: "0" },
          { key: "visitsPerDay4", label: "Visits / Day 4", type: "number", step: "0.01", placeholder: "0" },

          { key: "providerType5", label: "Provider Type 5", type: "text", placeholder: "Provider type" },
          { key: "providerCount5", label: "Provider Count 5", type: "number", step: "1", placeholder: "0" },
          { key: "ptoDays5", label: "PTO Days 5", type: "number", step: "0.01", placeholder: "0" },
          { key: "visitsPerDay5", label: "Visits / Day 5", type: "number", step: "0.01", placeholder: "0" },

          { key: "providerType6", label: "Provider Type 6", type: "text", placeholder: "Provider type" },
          { key: "providerCount6", label: "Provider Count 6", type: "number", step: "1", placeholder: "0" },
          { key: "ptoDays6", label: "PTO Days 6", type: "number", step: "0.01", placeholder: "0" },
          { key: "visitsPerDay6", label: "Visits / Day 6", type: "number", step: "0.01", placeholder: "0" },

          { key: "providerType7", label: "Provider Type 7", type: "text", placeholder: "Provider type" },
          { key: "providerCount7", label: "Provider Count 7", type: "number", step: "1", placeholder: "0" },
          { key: "ptoDays7", label: "PTO Days 7", type: "number", step: "0.01", placeholder: "0" },
          { key: "visitsPerDay7", label: "Visits / Day 7", type: "number", step: "0.01", placeholder: "0" }
        ],
        calculatedFields: [
          { key: "totalProviders", label: "Total Providers", format: "whole" },
          { key: "capacityVisits", label: "Capacity Visits", format: "whole" },
          { key: "capacityVisitsPerDay", label: "Capacity Visits / Day", format: "decimal2" }
        ]
      }
    ]
  },

  "Productivity Builder": {
    key: "Productivity Builder",
    title: "Productivity Builder",
    description: "Workbook-aligned staffing builder with throughput and buffer logic.",
    sections: [
      {
        key: "productivityInputs",
        title: "Productivity Inputs",
        description: "Direct workbook staffing inputs.",
        fields: [
          { key: "monthTag", label: "Month", type: "text", placeholder: "Example: Jan" },
          { key: "workingDaysInPeriod", label: "Working Days", type: "number", step: "1", placeholder: "Example: 22" },
          { key: "newPatients", label: "New Patients (NP)", type: "number", step: "1", placeholder: "Enter NP forecast" },
          { key: "npToEstablishedMultiplier", label: "NP → Est Multiplier", type: "number", step: "0.01", placeholder: "Example: 3.5" },
          { key: "npToSurgeryConversion", label: "NP → Surgery Conversion", type: "number", step: "0.01", placeholder: "Example: 0.20" },
          { key: "callsVolume", label: "Calls Volume", type: "number", step: "1", placeholder: "Enter calls volume" },
          { key: "maVisitsPerDayPerFte", label: "MA Visits / Day / FTE", type: "number", step: "0.01", placeholder: "Example: 20" },
          { key: "psaVisitsPerDayPerFte", label: "PSA Visits / Day / FTE", type: "number", step: "0.01", placeholder: "Example: 25" },
          { key: "callsPerDayPerFte", label: "Calls / Day / FTE", type: "number", step: "0.01", placeholder: "Example: 60" },
          { key: "surgeriesPerDayPerSchedulerFte", label: "Surgeries / Day / Scheduler FTE", type: "number", step: "0.01", placeholder: "Example: 8" },
          { key: "staffingBufferPct", label: "Staffing Buffer %", type: "number", step: "0.01", placeholder: "Example: 0.10" }
        ],
        calculatedFields: [
          { key: "forecastEstablishedVisits", label: "Forecast Established Visits", format: "whole" },
          { key: "forecastSurgeries", label: "Forecast Surgeries", format: "whole" },
          { key: "totalVisits", label: "Total Visits (NP + Est)", format: "whole" },
          { key: "npPerDay", label: "NP / Day", format: "decimal2" },
          { key: "estPerDay", label: "Est / Day", format: "decimal2" },
          { key: "surgPerDay", label: "Surg / Day", format: "decimal2" },
          { key: "maFte", label: "MA FTE", format: "decimal2" },
          { key: "psaFte", label: "PSA FTE", format: "decimal2" },
          { key: "callFte", label: "Call FTE", format: "decimal2" },
          { key: "surgerySchedulerFte", label: "Surgery Scheduler FTE", format: "decimal2" },
          { key: "totalOpsFte", label: "Total Ops FTE", format: "decimal2" }
        ]
      }
    ]
  }
};

export function getRegionSections(entity) {
  return REGION_SECTIONS.filter((section) => section.entities.includes(entity));
}

export function getAllMetricKeysForEntity(entity) {
  const sections = getRegionSections(entity);
  return sections.flatMap((section) => [
    ...section.fields.map((field) => field.key),
    ...((section.calculatedFields || []).map((field) => field.key))
  ]);
}

export function getSharedPageDefinition(pageName) {
  return SHARED_PAGE_DEFINITIONS[pageName] || null;
}

export function getAllMetricKeysForSharedPage(pageName) {
  const def = getSharedPageDefinition(pageName);
  if (!def) return [];
  return def.sections.flatMap((section) => [
    ...section.fields.map((field) => field.key),
    ...((section.calculatedFields || []).map((field) => field.key))
  ]);
}
