// GET /api/amd-weekly-preview?weekEnding=YYYY-MM-DD
//
// Read-only preview of LAOSS weekly volume + cash, derived live from the
// AMD OAK staging tables (staging.AdvMD137388_*). Operators see these as
// pre-fills on the weekly-entry form; overtyping a field overrides the
// AMD-derived number.
//
// LAOSS-only for now (office key 137388). Other regions use ECW/ModMed
// exports via the excel-upload flow (phase 4).

const { queryOak } = require("../shared/oak-sql");
const { getUserFromRequest } = require("../shared/auth");
const {
  resolveAccess,
  requireAccess,
  canAccessEntity,
  safeErrorResponse
} = require("../shared/permissions");

const ENTITY = "LAOSS";

// Authoritative mapping rules (from UnityMSK Appointment Mapping.xlsx).
// If this list changes, update the rules meta returned to the client too.
const SURGERY_TYPES = [
  "MARCUS WC SURGERY",
  "PRIVATE - SURGERY",
  "PRIVATE - SURGERY, PRIVATE - ORTHOVISC / PRP",
  "PRIVATE - SURGERY, PRIVATE - PRE-OP",
  "WC - SURGERY",
  "WC - SURGERY, MARCUS PROCEDURE"
];

const NEW_PATIENT_TYPES = [
  "PRIVATE - NEW",
  "WC - NEW",
  "NEW WC",
  "WC-NEW-CONSULT"
];

const APPT_STATUS_NO_SHOW = 12;
const APPT_STATUS_CANCELLED = 10;

function weekRangeFromEnding(weekEnding) {
  // weekEnding is a Friday date; range is Mon-Fri (Mon = weekEnding - 4)
  const end = new Date(`${weekEnding}T12:00:00Z`);
  const start = new Date(end);
  start.setUTCDate(end.getUTCDate() - 4);
  const fmt = (d) => d.toISOString().slice(0, 10);
  return { startDate: fmt(start), endDate: fmt(end) };
}

function buildQuery() {
  // Parameterize all literals. In-clause has to be interpolated
  // since mssql doesn't support array params; use a TVP alternative:
  // build named params @st0, @st1, ... for each element.
  const surgParams = SURGERY_TYPES.map((_, i) => `@surg${i}`).join(",");
  const newParams = NEW_PATIENT_TYPES.map((_, i) => `@newp${i}`).join(",");

  return `
    DECLARE @start date = @startDate;
    DECLARE @end date = @endDate;

    ;WITH seen AS (
      SELECT Appointment_UID, ApptTypes
      FROM staging.AdvMD137388_appts_Appointments
      WHERE SeenTime IS NOT NULL
        AND CAST(SeenTime AS date) BETWEEN @start AND @end
    ),
    sched AS (
      SELECT Appointment_UID, ApptStatus
      FROM staging.AdvMD137388_appts_Appointments
      WHERE CAST(StartDateTime AS date) BETWEEN @start AND @end
    )
    SELECT
      -- Seen-based counts (authoritative visits)
      (SELECT COUNT(*) FROM seen WHERE ApptTypes IN (${newParams})) AS newPatients,
      (SELECT COUNT(*) FROM seen WHERE ApptTypes IN (${surgParams})) AS surgeries,
      (SELECT COUNT(*) FROM seen WHERE
        ApptTypes NOT IN (${newParams})
        AND ApptTypes NOT IN (${surgParams})
      ) AS established,
      (SELECT COUNT(*) FROM seen) AS visitVolume,

      -- Scheduled-but-not-seen status counts
      (SELECT COUNT(*) FROM sched WHERE ApptStatus = @noShowCode) AS noShows,
      (SELECT COUNT(*) FROM sched WHERE ApptStatus = @cancelCode) AS cancelled,

      -- Cash collected
      (
        SELECT ISNULL(SUM(PaymentAmount), 0)
        FROM staging.AdvMD137388_actv_Payments
        WHERE CAST(EntryDate AS date) BETWEEN @start AND @end
      ) AS cashCollected,

      -- Freshness signal for the UI
      (
        SELECT MAX(DateStaged)
        FROM staging.AdvMD137388_appts_Appointments
      ) AS dataStagedAt
  `;
}

async function fetchAmdWeekly(startDate, endDate) {
  const params = { startDate, endDate, noShowCode: APPT_STATUS_NO_SHOW, cancelCode: APPT_STATUS_CANCELLED };
  SURGERY_TYPES.forEach((t, i) => { params[`surg${i}`] = t; });
  NEW_PATIENT_TYPES.forEach((t, i) => { params[`newp${i}`] = t; });

  const result = await queryOak(buildQuery(), params);
  const row = result.recordset[0] || {};
  const num = (v) => Number(v || 0);
  return {
    newPatients: num(row.newPatients),
    surgeries: num(row.surgeries),
    established: num(row.established),
    visitVolume: num(row.visitVolume),
    noShows: num(row.noShows),
    cancelled: num(row.cancelled),
    cashCollected: num(row.cashCollected),
    dataStagedAt: row.dataStagedAt || null
  };
}

module.exports = async function (context, req) {
  try {
    const user = getUserFromRequest(req);
    const access = resolveAccess(user);
    const authError = requireAccess(access);
    if (authError) {
      context.res = { status: authError.status, headers: { "Content-Type": "application/json" }, body: authError.body };
      return;
    }

    if (!canAccessEntity(access, ENTITY)) {
      context.res = { status: 404, headers: { "Content-Type": "application/json" }, body: { ok: false, error: "Not found" } };
      return;
    }

    const weekEnding = String(req.query?.weekEnding || "").trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(weekEnding)) {
      context.res = { status: 400, headers: { "Content-Type": "application/json" }, body: { ok: false, error: "Provide weekEnding=YYYY-MM-DD" } };
      return;
    }

    const { startDate, endDate } = weekRangeFromEnding(weekEnding);
    const fields = await fetchAmdWeekly(startDate, endDate);

    context.res = {
      status: 200,
      headers: { "Content-Type": "application/json" },
      body: {
        ok: true,
        entity: ENTITY,
        weekEnding,
        weekStart: startDate,
        source: "amd-137388",
        fetchedAt: new Date().toISOString(),
        dataStagedAt: fields.dataStagedAt,
        fields: {
          newPatients: fields.newPatients,
          surgeries: fields.surgeries,
          established: fields.established,
          visitVolume: fields.visitVolume,
          noShows: fields.noShows,
          cancelled: fields.cancelled,
          cashCollected: fields.cashCollected
        },
        rules: {
          visits: "Counts appointments where SeenTime is within the Mon-Fri week.",
          newPatients: `ApptTypes in ${JSON.stringify(NEW_PATIENT_TYPES)}`,
          surgeries: `ApptTypes in ${JSON.stringify(SURGERY_TYPES)}`,
          established: "Seen visits that are neither new-patient nor surgery types",
          noShows: `ApptStatus = ${APPT_STATUS_NO_SHOW}, scheduled within the week`,
          cancelled: `ApptStatus = ${APPT_STATUS_CANCELLED}, scheduled within the week`,
          cashCollected: "Sum of Payments.PaymentAmount with EntryDate in the week"
        }
      }
    };
  } catch (error) {
    return safeErrorResponse(context, error, "Failed to fetch AMD weekly preview");
  }
};
