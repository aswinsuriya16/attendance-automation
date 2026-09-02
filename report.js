/**
 * Daily Attendance Report — Exact Working Hours
 * =============================================
 * Emails each reporting manager the exact working hours of their reportees for
 * the previous day, as Excel + PDF attachments.
 *
 *   node report.js                 # live run
 *   DRY_RUN=true node report.js    # build files, send no email
 *
 * ------------------------------------------------------------------
 * WHAT "EXACT WORKING HOURS" MEANS HERE
 * ------------------------------------------------------------------
 * The sum of every paired check-in/check-out session in the day. Time between
 * sessions — lunch, a commute home, any gap — is excluded.
 *
 * Worked example (a 10-to-7 shift plus an evening login from home):
 *
 *     10:00 -> 13:00   3h 00m   morning
 *     13:00 -> 14:00      —     lunch, excluded
 *     14:00 -> 19:00   5h 00m   afternoon
 *     19:00 -> 20:00      —     commute, excluded
 *     20:00 -> 21:00   1h 00m   evening login from home
 *     ------------------------------------
 *     TOTAL            9h 00m
 *
 * Zoho's own figures would report this as 11h00m, because `totalHrs` is simply
 * lastOut minus firstIn and counts lunch and the commute as work. This script
 * deliberately ignores `totalHrs`, `firstIn` and `lastOut` for that reason.
 *
 * ------------------------------------------------------------------
 * WHY THIS RUNS IN THE MORNING, NOT AT END OF DAY
 * ------------------------------------------------------------------
 * Exact hours require everyone to have clocked out. In the example above the
 * employee is still working at 21:00, and staff on US-aligned shifts work until
 * ~01:00 IST. A report sent at 18:00 or 20:00 would truncate those sessions and
 * understate hours.
 *
 * So the report covers YESTERDAY and runs the next morning:
 *
 *     CRON_TZ=Asia/Kolkata
 *     0 9 * * * /usr/bin/node /path/to/report.js >> /var/log/attendance.log 2>&1
 *
 * The script warns at startup if anyone is still clocked in, which is the
 * signal that it is running too early.
 *
 * ------------------------------------------------------------------
 * THE ZOHO TIMEZONE ISSUE — HANDLED AUTOMATICALLY, NO SETTINGS CHANGE
 * ------------------------------------------------------------------
 * Zoho resolves the `date` parameter in the timezone of the OAuth user, which
 * on this account is America/Los_Angeles. Asking for "20-Aug-2026" returns only
 * punches falling in 20-Aug 12:30 PM -> 21-Aug 12:29 PM IST — the back half of
 * one working day plus the front half of the next.
 *
 * Consequences if unhandled: a 09:30 IST check-in is before that window opens,
 * so it is absent from the response, its check-out arrives orphaned, and hours
 * are understated. Meanwhile punches from two different days get summed.
 *
 * This script reads `responseTimezone` at startup, works out which Zoho dates
 * are needed to cover the full IST day, fetches them all, merges and
 * de-duplicates the punches, and converts each to a true UTC instant before
 * doing any arithmetic. Durations are therefore exact regardless of which
 * timezone Zoho answers in, and it keeps working if that setting ever changes.
 *
 * ------------------------------------------------------------------
 * SETUP
 * ------------------------------------------------------------------
 *   1. Two OAuth apps at https://api-console.zoho.in:
 *        ZOHOPEOPLE.forms.READ       -> employee + reporting manager data
 *        ZOHOPEOPLE.attendance.READ  -> attendance entries
 *      Each needs its own Client ID, Secret and refresh token.
 *   2. npm install axios exceljs pdfkit nodemailer dotenv
 *   3. Create .env (never commit it):
 *
 *        ZOHO_FORMS_CLIENT_ID=...
 *        ZOHO_FORMS_CLIENT_SECRET=...
 *        ZOHO_FORMS_REFRESH_TOKEN=...
 *        ZOHO_ATTENDANCE_CLIENT_ID=...
 *        ZOHO_ATTENDANCE_CLIENT_SECRET=...
 *        ZOHO_ATTENDANCE_REFRESH_TOKEN=...
 *        EMAIL_PROVIDER=smtp
 *        MAIL_FROM=no-reply@clouddestinations.com
 *        SMTP_HOST=email-smtp.us-east-2.amazonaws.com
 *        SMTP_USER=AKIA6LP3BDTBYCHYRLKQ
 *        SMTP_PASSWORD=...              (the SES SMTP password, not an IAM secret key)
 *        REPORTS_DIR=C:\Reports
 *
 *      Two SES-specific things that fail silently if missed — see section
 *      "AMAZON SES — TWO THINGS THAT WILL BITE YOU" below.
 *
 *      For Microsoft Graph instead (no password held anywhere):
 *        EMAIL_PROVIDER=graph
 *        MS_TENANT_ID=... / MS_CLIENT_ID=... / MS_CLIENT_SECRET=...
 */

require("dotenv").config();

const axios = require("axios");
const fs = require("fs");
const path = require("path");
const ExcelJS = require("exceljs");
const PDFDocument = require("pdfkit");
const nodemailer = require("nodemailer");

// ==================================================================
// CONFIG
// ==================================================================

const ZOHO_ACCOUNTS_BASE = "https://accounts.zoho.in";
const ZOHO_PEOPLE_BASE = "https://people.zoho.in/people/api";

const FORMS_CLIENT_ID = process.env.ZOHO_FORMS_CLIENT_ID || "";
const FORMS_CLIENT_SECRET = process.env.ZOHO_FORMS_CLIENT_SECRET || "";
const FORMS_REFRESH_TOKEN = process.env.ZOHO_FORMS_REFRESH_TOKEN || "";

const ATTENDANCE_CLIENT_ID = process.env.ZOHO_ATTENDANCE_CLIENT_ID || "";
const ATTENDANCE_CLIENT_SECRET = process.env.ZOHO_ATTENDANCE_CLIENT_SECRET || "";
const ATTENDANCE_REFRESH_TOKEN = process.env.ZOHO_ATTENDANCE_REFRESH_TOKEN || "";

const GMAIL_APP_PASSWORD = process.env.GMAIL_APP_PASSWORD || "";

const REPORTS_BASE_DIR = process.env.REPORTS_DIR || path.join(__dirname, "reports_out");

// All staff are in India, so the business day is a fixed IST calendar day.
// IST is +05:30 year-round with no daylight saving.
const BUSINESS_TZ = "Asia/Kolkata";
const IST_OFFSET_MINUTES = 330;

/**
 * How far past midnight a session that STARTED in the reported day may run.
 * A session is attributed to the day its check-in falls in and is never cut off
 * at midnight, so this only controls how much extra data is fetched to find the
 * matching check-out. Staff on US-aligned shifts finish around 01:00 IST; 12
 * hours gives comfortable headroom.
 */
const LATE_SESSION_HOURS = parseFloat(process.env.LATE_SESSION_HOURS || "12");

/**
 * GATE FILTERING
 *
 * The office has multiple badge readers: an entry gate that records both
 * check-in and check-out, plus washroom and gym gates that record ONLY a
 * check-in. Those scan-only punches are not work events and must not enter the
 * pairing, or they corrupt every figure (see computeHours).
 *
 * Rather than hardcoding gate names, the script infers them: it censuses every
 * location seen across the whole run and treats any location that produced
 * check-ins but never a single check-out as a scan-only gate. A real entry gate
 * always produces check-outs, so it can never be misclassified.
 *
 * Override the inference if needed:
 *   IGNORE_LOCATIONS=COIMBATORE G2,CHENNAI GYM   explicit deny list
 *   ENTRY_LOCATIONS=COIMBATORE,CHENNAI           explicit allow list (wins)
 *   AUTO_DETECT_GATES=false                      disable inference entirely
 */
const AUTO_DETECT_GATES = String(process.env.AUTO_DETECT_GATES ?? "true").toLowerCase() === "true";
const IGNORE_LOCATIONS = (process.env.IGNORE_LOCATIONS || "").split(",").map((x) => x.trim()).filter(Boolean);
const ENTRY_LOCATIONS = (process.env.ENTRY_LOCATIONS || "").split(",").map((x) => x.trim()).filter(Boolean);

/** A location needs at least this many check-ins before it can be judged scan-only. */
const GATE_MIN_SAMPLE = parseInt(process.env.GATE_MIN_SAMPLE || "5", 10);

/**
 * Two punches in the same direction, at the same reader, within this many
 * seconds are one physical event badged twice. Collapsed to a single punch.
 */
const DOUBLE_SWIPE_SECONDS = parseInt(process.env.DOUBLE_SWIPE_SECONDS || "90", 10);

/** Sessions longer than this are implausible and get flagged for review. */
const MAX_SESSION_HOURS = parseFloat(process.env.MAX_SESSION_HOURS || "16");

/** Report a specific day instead of yesterday: REPORT_DATE=20-Aug-2026 */
const REPORT_DATE_OVERRIDE = process.env.REPORT_DATE || "";

const DRY_RUN = String(process.env.DRY_RUN || "").toLowerCase() === "true";

/**
 * Filter to a single reporting manager's email for testing (e.g. TEST_MANAGER_EMAIL=kandhakt@clouddestinations.com)
 */
const TEST_MANAGER_EMAIL = (process.env.TEST_MANAGER_EMAIL || "").trim().toLowerCase();

/**
 * RATE LIMITING — adaptive, not pre-emptive.
 *
 * Zoho documents getAttendanceEntries as "Threshold Limit: 100 requests,
 * Lock period: 10 minutes", where the threshold is the number of calls allowed
 * WITHIN A MINUTE and the lock period is the penalty served after tripping it.
 *
 * So the correct strategy is to pace just under 100/min and only back off if
 * Zoho actually returns 429 — not to sleep 10 minutes every 90 calls, which
 * turned a 5-minute job into a 40-minute one.
 *
 * If the limit does turn out to be 100 per 10 minutes on your plan, this still
 * behaves correctly: the first 429 triggers a full lock-period wait and the run
 * continues. Adaptive is never slower than pre-emptive pausing.
 */
const REQUEST_DELAY_MS = parseInt(process.env.REQUEST_DELAY_MS || "700", 10); // ~85 req/min
const LOCK_PERIOD_MS = parseInt(process.env.LOCK_PERIOD_MS || "600000", 10); // 10 min penalty
const PROGRESS_EVERY = parseInt(process.env.PROGRESS_EVERY || "25", 10);

/**
 * Pulls the manager's employee ID out of a "Reporting_To" string.
 *
 * The org uses more than one ID format — "CD-CJB01-00025" (two hyphens) and
 * "CD-A0118" (one). A single regex cannot cover both without either missing
 * one or matching too greedily, and a miss is silent: managerId comes back
 * null, the entire "Reporting_To" string is mistaken for the manager's name,
 * and the employee is dropped or attached to the wrong group.
 *
 * So this matches against the REAL employee IDs instead of a guessed pattern.
 * Longest first, so "CD-A0118" is never shadowed by a shorter "CD-A011".
 * The regex stays only as a fallback for a manager missing from the roster.
 */
let _idMatcher = null;
function extractManagerId(reportingTo, knownIds) {
  if (!reportingTo) return null;

  if (!_idMatcher || _idMatcher.size !== knownIds.length) {
    _idMatcher = { size: knownIds.length, sorted: [...knownIds].sort((a, b) => b.length - a.length) };
  }
  for (const id of _idMatcher.sorted) {
    if (containsIdToken(reportingTo, id)) return id;
  }

  const m = reportingTo.match(EMP_ID_PATTERN_FALLBACK);
  return m ? m[0] : null;
}

/**
 * Substring match anchored to token boundaries.
 *
 * A plain includes() is unsafe here because the roster holds very short IDs —
 * this org has one that is literally "9". Bare includes() would match that
 * inside "CD-MAA01-00925" and hand the employee to the wrong manager, while
 * still counting as a clean resolution. Requiring a non-alphanumeric character
 * either side means "9" only matches a standalone 9.
 */
function containsIdToken(haystack, id) {
  if (!id) return false;
  let i = haystack.indexOf(id);
  while (i !== -1) {
    const before = i === 0 ? "" : haystack[i - 1];
    const after = i + id.length >= haystack.length ? "" : haystack[i + id.length];
    if (!/[A-Za-z0-9]/.test(before) && !/[A-Za-z0-9]/.test(after)) return true;
    i = haystack.indexOf(id, i + 1);
  }
  return false;
}

/** Fallback only — covers both "CD-CJB01-00025" and "CD-A0118" shapes. */
const EMP_ID_PATTERN_FALLBACK = /[A-Z]{2,}-[A-Z0-9]+(?:-\d+)?/;

/** getRecords caps at 200 per call (error 7021 above that) and pages via 1-based sIndex. */
const EMPLOYEE_PAGE_SIZE = 200;
const EMPLOYEE_PAGE_CAP = parseInt(process.env.EMPLOYEE_PAGE_CAP || "50", 10); // 10,000 staff

/**
 * Employment statuses to drop, e.g. EXCLUDE_EMPLOYEE_STATUS=Terminated,Resigned
 * Empty by default: the roster's actual status values are printed at startup so
 * you can see what exists before deciding, rather than guessing at labels.
 */
const EXCLUDE_EMPLOYEE_STATUS = (
  process.env.EXCLUDE_EMPLOYEE_STATUS ?? "Resigned,Terminated,Abscond"
).split(",").map((x) => x.trim()).filter(Boolean);

const EMP_ID_PATTERN = /[A-Z]{2}-[A-Z0-9]+-\d+/;
const MONTH_NAMES = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
const MONTHS = MONTH_NAMES.reduce((a, m, i) => ((a[m.toLowerCase()] = i), a), {});

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

// ==================================================================
// TIME HANDLING
//
// Zoho sends wall-clock strings in whatever zone it chooses. Those are
// converted to true UTC instants on the way in, so all arithmetic is exact and
// unaffected by the host server's timezone or by DST in Zoho's zone. Output is
// rendered back into IST.
// ==================================================================

/** Offset of `timeZone` in minutes at a given instant. Positive = east of UTC. */
function zoneOffsetMinutes(instant, timeZone) {
  const p = new Intl.DateTimeFormat("en-US", {
    timeZone,
    hour12: false,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
  })
    .formatToParts(instant)
    .reduce((a, x) => ((a[x.type] = x.value), a), {});
  let hour = parseInt(p.hour, 10);
  if (hour === 24) hour = 0;
  return (Date.UTC(+p.year, +p.month - 1, +p.day, hour, +p.minute, +p.second) - instant.getTime()) / 60000;
}

/** Converts a wall-clock reading in `timeZone` to a true UTC instant. */
function wallClockToInstant(y, mo, d, h, mi, s, timeZone) {
  const naive = Date.UTC(y, mo, d, h, mi, s || 0);
  let off = zoneOffsetMinutes(new Date(naive), timeZone);
  off = zoneOffsetMinutes(new Date(naive - off * 60000), timeZone);
  return new Date(naive - off * 60000);
}

/** Renders an instant as IST wall-clock, for logs and diagnostics. */
function formatIst(instant) {
  if (!instant) return "-";
  const p = new Intl.DateTimeFormat("en-GB", {
    timeZone: BUSINESS_TZ,
    hour12: true,
    day: "2-digit",
    month: "short",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
  })
    .formatToParts(instant)
    .reduce((a, x) => ((a[x.type] = x.value), a), {});
  return `${p.day}-${p.month.slice(0, 3)}-${p.year} ${p.hour}:${p.minute} ${(p.dayPeriod || "").toUpperCase()}`;
}

/** The Zoho calendar date (dd-MMM-yyyy) an instant falls on, in Zoho's zone. */
function zohoDateFor(instant, zohoTz) {
  const p = new Intl.DateTimeFormat("en-GB", { timeZone: zohoTz, day: "2-digit", month: "short", year: "numeric" })
    .formatToParts(instant)
    .reduce((a, x) => ((a[x.type] = x.value), a), {});
  return `${p.day}-${p.month.slice(0, 3)}-${p.year}`;
}

function parseDateStr(str) {
  const m = String(str).match(/^(\d{1,2})-([A-Za-z]{3})-(\d{4})$/);
  if (!m) throw new Error(`Unparseable date: ${str}. Expected dd-MMM-yyyy, e.g. 20-Aug-2026`);
  const mo = MONTHS[m[2].toLowerCase()];
  if (mo === undefined) throw new Error(`Unknown month in: ${str}`);
  return { y: +m[3], mo, d: +m[1] };
}

/** Midnight IST on a given business date, as a UTC instant. */
function istMidnight(y, mo, d) {
  return new Date(Date.UTC(y, mo, d, 0, 0, 0) - IST_OFFSET_MINUTES * 60000);
}

/** Yesterday on the IST calendar, independent of the host server's clock. */
function istYesterday() {
  const ist = new Date(Date.now() + IST_OFFSET_MINUTES * 60000);
  ist.setUTCDate(ist.getUTCDate() - 1);
  return `${String(ist.getUTCDate()).padStart(2, "0")}-${MONTH_NAMES[ist.getUTCMonth()]}-${ist.getUTCFullYear()}`;
}

/**
 * Which Zoho calendar dates must be fetched to fully cover the IST day.
 *
 * The IST day plus LATE_SESSION_HOURS is projected into Zoho's own timezone,
 * and every Zoho date it touches is returned. With Zoho on Pacific this yields
 * two dates; if the account were on IST it would yield one or two depending on
 * the overrun.
 */
function requiredZohoDates(dayStart, dayEnd, zohoTz) {
  const end = new Date(dayEnd.getTime() + LATE_SESSION_HOURS * 3600000);
  const dates = [];
  const seen = new Set();
  for (let t = dayStart.getTime(); t <= end.getTime(); t += 6 * 3600000) {
    const ds = zohoDateFor(new Date(t), zohoTz);
    if (!seen.has(ds)) (seen.add(ds), dates.push(ds));
  }
  const last = zohoDateFor(end, zohoTz);
  if (!seen.has(last)) dates.push(last);
  return dates;
}

// ==================================================================
// ZOHO API
// ==================================================================

async function getAccessToken(clientId, clientSecret, refreshToken) {
  const resp = await axios.post(`${ZOHO_ACCOUNTS_BASE}/oauth/v2/token`, null, {
    params: {
      refresh_token: refreshToken,
      client_id: clientId,
      client_secret: clientSecret,
      grant_type: "refresh_token",
    },
    timeout: 30000,
  });
  const { access_token, error } = resp.data;
  if (error || !access_token) {
    throw new Error(
      `OAuth token exchange failed for client_id=${clientId}: ${error || "no access_token returned"}. ` +
        `The refresh token may have expired — regenerate it at https://api-console.zoho.in`
    );
  }
  return access_token;
}

const authHeaders = (t) => ({ Authorization: `Zoho-oauthtoken ${t}` });

/** Retries rate-limit and transient server errors; surfaces everything else. */
async function withRetry(fn, label, attempts = 5) {
  let lastErr;
  for (let i = 1; i <= attempts; i++) {
    try {
      return await fn();
    } catch (err) {
      lastErr = err;
      const status = err?.response?.status;
      const retryable =
        status === 429 || (status >= 500 && status < 600) || err.code === "ECONNRESET" || err.code === "ETIMEDOUT";
      if (!retryable || i === attempts) break;
      // A 429 means the per-minute threshold was tripped, so the full lock
      // period must be served. Shorter backoff just burns another attempt.
      const wait = status === 429 ? LOCK_PERIOD_MS : 3000 * i;
      if (status === 429) {
        console.warn(`  [rate limit] ${label} — serving ${Math.round(wait / 60000)} min lock period`);
      } else {
        console.warn(`  [retry ${i}/${attempts}] ${label} (${status || err.code}) — waiting ${Math.round(wait / 1000)}s`);
      }
      await sleep(wait);
    }
  }
  throw lastErr;
}

/** Unwraps Zoho's result envelope. Three shapes are seen across org versions. */
function unwrapRecords(raw) {
  if (Array.isArray(raw)) return raw;
  if (Array.isArray(raw?.response?.result)) {
    return raw.response.result.flatMap((i) => Object.values(i).flatMap((v) => (Array.isArray(v) ? v : [v])));
  }
  if (raw?.response?.result) {
    return Object.values(raw.response.result).flatMap((v) => (Array.isArray(v) ? v : [v]));
  }
  return [];
}

/**
 * Fetches EVERY employee, paging until the records run out.
 *
 * getRecords returns at most 200 records per call and pages with a 1-based
 * sIndex. A single unpaginated call therefore returns only the first 200
 * employees and gives no indication that anything is missing — managers appear
 * to have two reportees instead of eleven, and the employees who fell outside
 * the window are silently absent from every report. Passing limit above 200
 * does not help; Zoho rejects it with error 7021.
 */
async function fetchEmployees(token) {
  const all = [];
  let sIndex = 1;
  let page = 0;

  while (page < EMPLOYEE_PAGE_CAP) {
    const resp = await withRetry(
      () =>
        axios.get(`${ZOHO_PEOPLE_BASE}/forms/employee/getRecords`, {
          headers: authHeaders(token),
          params: { sIndex, limit: EMPLOYEE_PAGE_SIZE },
          timeout: 30000,
        }),
      `fetchEmployees(sIndex=${sIndex})`
    );

    const batch = unwrapRecords(resp.data);
    page++;

    // Stop only on a genuinely EMPTY page, never on a short one. A short page
    // usually means the end, but Zoho can apply filters after slicing, so a
    // partial page can still be followed by more records. Treating "short" as
    // "last" would silently truncate the roster — the same class of bug as not
    // paginating at all. One extra call is a cheap price for certainty.
    if (batch.length === 0) {
      console.log(`  page ${page}: empty — end of records`);
      break;
    }

    all.push(...batch);
    console.log(`  page ${page}: ${batch.length} records (sIndex ${sIndex})`);

    sIndex += EMPLOYEE_PAGE_SIZE;
    await sleep(REQUEST_DELAY_MS);
  }

  if (page >= EMPLOYEE_PAGE_CAP) {
    console.warn(`  [WARN] Stopped at the ${EMPLOYEE_PAGE_CAP}-page cap. Raise EMPLOYEE_PAGE_CAP if you have more staff.`);
  }

  // Zoho can repeat a record across page boundaries if the underlying list
  // shifts mid-fetch; de-duplicate on EmployeeID.
  // Normalise IDs before anything else. The roster contains at least one
  // EmployeeID with a leading space; untrimmed, it fails every lookup because
  // other records reference it without the space.
  let trimmed = 0;
  for (const f of all) {
    if (typeof f["EmployeeID"] === "string" && f["EmployeeID"] !== f["EmployeeID"].trim()) {
      f["EmployeeID"] = f["EmployeeID"].trim();
      trimmed++;
    }
  }
  if (trimmed) console.log(`  normalised ${trimmed} EmployeeID(s) with surrounding whitespace`);

  const seen = new Set();
  const records = all.filter((f) => {
    const id = f["EmployeeID"];
    if (!id || seen.has(id)) return false;
    seen.add(id);
    return true;
  });

  const shortIds = records.map((f) => f["EmployeeID"]).filter((id) => id.length < 4);
  if (shortIds.length) {
    console.warn(
      `  [WARN] ${shortIds.length} very short EmployeeID(s): ${shortIds.join(", ")}. ` +
        `Matched on token boundaries so they cannot collide, but verify they are real.`
    );
  }

  reportStatusFields(records);

  // Build the set of real employee IDs FIRST, so manager IDs can be matched
  // against actual data instead of guessed at with a pattern. See extractManagerId.
  const knownIds = records.map((f) => f["EmployeeID"]).filter(Boolean);

  return records.map((f) => {
    // "Reporting_To" is a combined string: "<Manager Name> <Manager EmpID>",
    // e.g. "Sona Sujanan CD-MAA01-00540". Not an email — resolved in groupByManager.
    const mgrRaw = f["Reporting_To"] || "";
    const managerId = extractManagerId(mgrRaw, knownIds);
    return {
      empId: f["EmployeeID"],
      name: [f["FirstName"], f["LastName"]].filter(Boolean).join(" "),
      email: f["EmailID"],
      managerId,
      managerName: (managerId ? mgrRaw.replace(managerId, "") : mgrRaw).trim(),
      managerEmail: f["Reporting_To.MailID"] || null,
      status: statusOf(f),
      // Flagged, NOT removed. Former staff are excluded from attendance and
      // from reports, but must remain resolvable: if a resigned manager were
      // dropped from the lookup, their reportees would fall through to
      // Reporting_To.MailID and the report would be emailed to someone who
      // has left the company.
      isActive: isActiveStatus(statusOf(f)),
    };
  });
}

function isActiveStatus(status) {
  if (!EXCLUDE_EMPLOYEE_STATUS.length) return true;
  const st = String(status || "").toLowerCase();
  return !EXCLUDE_EMPLOYEE_STATUS.some((x) => st === x.toLowerCase());
}

/** Common field names Zoho uses for employment status; orgs vary. */
const STATUS_FIELDS = ["Employeestatus", "Employee_status", "Employee_Status", "EmployeeStatus", "Employment_Status"];

function statusOf(fields) {
  for (const k of STATUS_FIELDS) if (fields[k]) return String(fields[k]);
  // Fall back to any key that looks like a status field.
  const key = Object.keys(fields).find((k) => /status/i.test(k) && !/task|approval/i.test(k));
  return key ? String(fields[key]) : null;
}

/**
 * Prints the spread of employment statuses in the roster.
 *
 * getRecords returns whatever the employee form holds, which in most orgs
 * includes former staff. Without this you cannot tell whether a roster of 700
 * is 700 current employees or 450 current plus 250 who have left — and the
 * latter would double the attendance calls and put departed people on their old
 * manager's report.
 */
function reportStatusFields(records) {
  const counts = new Map();
  for (const f of records) {
    const st = statusOf(f) || "(no status field)";
    counts.set(st, (counts.get(st) || 0) + 1);
  }
  if (counts.size <= 1 && counts.has("(no status field)")) {
    console.log("  No employment-status field found on these records — all are included.");
    return;
  }
  console.log("  Employment status spread:");
  for (const [st, n] of [...counts.entries()].sort((a, b) => b[1] - a[1])) {
    const excluded = EXCLUDE_EMPLOYEE_STATUS.some((x) => x.toLowerCase() === st.toLowerCase());
    console.log(`    ${st.padEnd(24)} ${String(n).padStart(5)}${excluded ? "   EXCLUDED" : ""}`);
  }
  if (!EXCLUDE_EMPLOYEE_STATUS.length) {
    console.log("    (all included — set EXCLUDE_EMPLOYEE_STATUS in .env to drop former staff)");
  }
}

async function fetchAttendance(token, empId, zohoDate) {
  const resp = await withRetry(
    () =>
      axios.get(`${ZOHO_PEOPLE_BASE}/attendance/getAttendanceEntries`, {
        headers: authHeaders(token),
        params: { date: zohoDate, empId },
        timeout: 30000,
      }),
    `attendance(${empId} @ ${zohoDate})`
  );
  // Zoho signals "no records" with a 200 and status:1, not an HTTP error.
  return resp.data?.response?.status === 1 ? null : resp.data;
}

/** Probes a few employees to learn which timezone Zoho is answering in. */
async function detectZohoTimezone(token, employees, dateStr) {
  for (const emp of employees.slice(0, 8)) {
    if (!emp.empId) continue;
    try {
      const data = await fetchAttendance(token, emp.empId, dateStr);
      if (data?.responseTimezone) return data.responseTimezone;
    } catch (err) {
      console.warn(`  [probe] ${emp.empId}: ${err.message}`);
    }
  }
  throw new Error(
    "Could not read responseTimezone from any probe employee. " +
      "Set ZOHO_TIMEZONE_OVERRIDE in .env (e.g. America/Los_Angeles) to proceed."
  );
}

// ==================================================================
// PUNCH PARSING AND HOURS CALCULATION
// ==================================================================

// Entry timestamps: "20-Aug-2026 - 09:52 PM". Single-digit hours are accepted;
// requiring two digits silently dropped punches like "9:36 AM".
const ENTRY_TS = /^(\d{1,2})-([A-Za-z]{3})-(\d{4})\s*-\s*(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)$/i;
// Top-level timestamps use a different format: "2026-08-20 21:52:00".
const ISO_TS = /^(\d{4})-(\d{2})-(\d{2})[ T](\d{1,2}):(\d{2})(?::(\d{2}))?$/;

function parsePunch(str, zohoTz) {
  if (!str || str === "-") return null;
  const s = String(str).trim();

  let m = s.match(ENTRY_TS);
  if (m) {
    const mo = MONTHS[m[2].toLowerCase()];
    if (mo === undefined) return null;
    let hour = parseInt(m[4], 10) % 12;
    if (/PM/i.test(m[7])) hour += 12;
    return wallClockToInstant(+m[3], mo, +m[1], hour, +m[5], +(m[6] || 0), zohoTz);
  }

  m = s.match(ISO_TS);
  if (m) return wallClockToInstant(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +(m[6] || 0), zohoTz);

  return null;
}

function hhmm(mins) {
  const safe = Math.max(0, Math.round(mins));
  return `${String(Math.floor(safe / 60)).padStart(2, "0")}:${String(safe % 60).padStart(2, "0")}`;
}

/**
 * Flattens every entry from every fetched Zoho date into one de-duplicated,
 * chronologically sorted punch stream.
 *
 * De-duplication matters because adjacent Zoho date windows overlap at the
 * seam, so the same physical punch can arrive in two responses. Without this,
 * a 6h30m session would be counted twice as 13h.
 */
function buildPunchStream(responses, zohoTz) {
  const punches = [];
  const seen = new Set();

  for (const raw of responses) {
    for (const e of raw?.entries || []) {
      for (const p of [
        { val: e.checkIn, dir: "IN", src: e.sourceOfPunchIn, loc: e.checkIn_Location },
        { val: e.checkOut, dir: "OUT", src: e.sourceOfPunchOut, loc: e.checkOut_Location },
      ]) {
        if (!p.val || p.val === "-") continue;
        const t = parsePunch(p.val, zohoTz);
        if (!t) {
          console.warn(`  [warn] unparseable ${p.dir} timestamp: "${p.val}"`);
          continue;
        }
        const key = `${p.dir}|${t.getTime()}`;
        if (seen.has(key)) continue;
        seen.add(key);
        punches.push({
          t,
          dir: p.dir,
          src: p.src && p.src !== "-" ? p.src : "-",
          loc: p.loc && p.loc !== "-" ? p.loc : "-",
        });
      }
    }
  }

  punches.sort((a, b) => a.t - b.t);
  return collapseDoubleSwipes(punches);
}

/**
 * Collapses a badge swiped twice at the same reader within seconds. Two
 * check-outs in a row would otherwise leave the second orphaned and flag a
 * clean record as suspect.
 */
function collapseDoubleSwipes(punches) {
  const out = [];
  for (const p of punches) {
    const prev = out[out.length - 1];
    if (prev && prev.dir === p.dir && prev.loc === p.loc && (p.t - prev.t) / 1000 <= DOUBLE_SWIPE_SECONDS) {
      continue; // same event, badged twice
    }
    out.push(p);
  }
  return out;
}

/**
 * Counts check-ins and check-outs per location across every employee.
 *
 * This is what identifies the gates. A door people badge through in both
 * directions produces both. A washroom or gym reader produces check-ins only.
 */
function censusLocations(allPunches) {
  const census = new Map();
  for (const p of allPunches) {
    const key = p.loc;
    if (!census.has(key)) census.set(key, { in: 0, out: 0, sources: new Set() });
    const c = census.get(key);
    c[p.dir === "IN" ? "in" : "out"]++;
    if (p.src !== "-") c.sources.add(p.src);
  }
  return census;
}

/**
 * Decides which locations to exclude, from the census plus any manual override.
 * An explicit ENTRY_LOCATIONS allow list wins outright; otherwise the deny list
 * and the inferred scan-only gates are combined.
 */
function resolveIgnoredLocations(census) {
  if (ENTRY_LOCATIONS.length) {
    const ignored = [...census.keys()].filter((l) => !ENTRY_LOCATIONS.includes(l));
    return { ignored: new Set(ignored), reason: "ENTRY_LOCATIONS allow list" };
  }

  const ignored = new Set(IGNORE_LOCATIONS);
  if (AUTO_DETECT_GATES) {
    for (const [loc, c] of census.entries()) {
      // "-" is Zoho web check-in, which legitimately carries no location.
      if (loc === "-") continue;
      if (c.in >= GATE_MIN_SAMPLE && c.out === 0) ignored.add(loc);
    }
  }
  return { ignored, reason: IGNORE_LOCATIONS.length ? "deny list + inference" : "inference" };
}

/**
 * Removes punches recorded at scan-only readers (restroom, gym).
 *
 * Filtering happens at PUNCH level, never at entry level. Zoho pairs across
 * readers — a Coimbatore entry can hold a check-in at "COIMBATORE G2" and its
 * check-out at "COIMBATORE". Discarding the whole entry would throw away that
 * real main-gate check-out, orphan the earlier check-in and report 00:00.
 * Dropping only the G2 check-in leaves the check-out to pair correctly.
 */
function filterPunches(punches, ignoredLocations) {
  if (!ignoredLocations.size) return punches;
  return punches.filter((p) => !ignoredLocations.has(p.loc));
}

/**
 * Pairs the punch stream into sessions and sums the ones belonging to the day.
 *
 * A session is attributed to the day its CHECK-IN falls in, and is never
 * truncated at midnight — a US-aligned shift running 16:00 to 01:00 counts in
 * full against the day it started.
 *
 * Because only paired sessions are summed, every gap between a check-out and
 * the next check-in is excluded automatically: lunch, a commute home, or any
 * other break. This is what makes the total differ from Zoho's own `totalHrs`,
 * which is lastOut minus firstIn and counts those gaps as work.
 *
 * Unmatched punches are counted and reported rather than silently dropped, so
 * an unreliable figure is visibly unreliable instead of quietly wrong.
 */
function computeHours(punches, dayStart, dayEnd) {
  const sessions = [];
  let orphanIn = 0;
  let orphanOut = 0;
  let stillOpen = false;
  let open = null;

  const inDay = (t) => t >= dayStart && t < dayEnd;

  let strayIns = 0;

  for (const p of punches) {
    if (p.dir === "IN") {
      // FIRST CHECK-IN WINS. Overwriting `open` here was a serious bug: a
      // washroom or gym scan mid-session would replace the real check-in with a
      // later timestamp, so the eventual check-out paired against the stray
      // scan instead. A 9-hour day with a gate scan a minute before leaving
      // reported as 00:01. Keeping the earliest open check-in makes scan-only
      // punches harmless even when they slip past the location filter.
      if (open) {
        strayIns++;
        continue;
      }
      open = p;
    } else {
      if (!open) {
        if (inDay(p.t)) orphanOut++; // check-out with no preceding check-in
        continue;
      }
      const mins = (p.t - open.t) / 60000;
      if (mins > 0) sessions.push({ start: open.t, end: p.t, mins });
      open = null;
    }
  }
  if (open && inDay(open.t)) {
    orphanIn++;
    stillOpen = true;
  }

  const daySessions = sessions.filter((s) => inDay(s.start));
  const workedMinutes = daySessions.reduce((sum, s) => sum + s.mins, 0);
  const longSessions = daySessions.filter((x) => x.mins > MAX_SESSION_HOURS * 60).length;

  // Invariant: worked time can never exceed the span from first to last punch.
  // A violation means a punch was counted twice — a real bug, not bad data.
  const dayPunches = punches.filter((p) => inDay(p.t));
  const spanMinutes =
    dayPunches.length >= 2 ? (dayPunches[dayPunches.length - 1].t - dayPunches[0].t) / 60000 : 0;

  let breakMinutes = 0;
  for (let i = 1; i < daySessions.length; i++) {
    breakMinutes += (daySessions[i].start - daySessions[i - 1].end) / 60000;
  }

  return {
    workedMinutes,
    breakMinutes,
    sessions: daySessions,
    orphanIn,
    orphanOut,
    stillOpen,
    strayIns,
    longSessions,
    spanMinutes,
    exceedsSpan: daySessions.length > 0 && workedMinutes > spanMinutes + 1,
    punchCount: dayPunches.length,
  };
}

function buildRow(emp, punches, dayStart, dayEnd, status, site) {
  const calc = computeHours(punches, dayStart, dayEnd);

  // Notes are terse so they fit the PDF column without truncating.
  //
  // Deliberately NOT flagged: extra check-ins mid-session, and gaps bridged by
  // the first-in-wins rule. Staff eat in the on-site cafeteria without badging
  // out, so an unbroken stretch across lunch is the normal case, not a fault.
  // Working hours are defined as time on premises; a badged check-out means the
  // employee actually left the building.
  let note = "";
  let unreliable = false;
  if (calc.punchCount === 0) {
    note = status && status !== "-" ? "" : "No record";
  } else if (calc.exceedsSpan) {
    note = "Total exceeds punch span — bug";
    unreliable = true;
  } else if (calc.stillOpen) {
    note = "Still clocked in";
    unreliable = true;
  } else if (calc.orphanOut) {
    // After gate filtering this means a check-out with no entry-gate check-in.
    note = `${calc.orphanOut} check-out, no entry`;
    unreliable = true;
  } else if (calc.longSessions) {
    note = `Session over ${MAX_SESSION_HOURS}h — verify`;
    unreliable = true;
  }

  return {
    empId: emp.empId,
    name: emp.name,
    workingHours: hhmm(calc.workedMinutes),
    workedMinutes: calc.workedMinutes,
    breakTime: hhmm(calc.breakMinutes),
    sessions: calc.sessions.length,
    strayIns: calc.strayIns,
    site: site || "-",
    status,
    note,
    unreliable,
    stillOpen: calc.stillOpen,
    detail: calc.sessions.map((s) => `${formatIst(s.start)} -> ${formatIst(s.end)} = ${hhmm(s.mins)}`),
  };
}

// ==================================================================
// GROUP BY MANAGER
// ==================================================================

function groupByManager(employees, rows) {
  const lookup = new Map(employees.map((e) => [e.empId, e]));
  const grouped = new Map();
  let noManager = 0;
  let noEmail = 0;
  const departedManagers = new Map();

  for (const row of rows) {
    const emp = lookup.get(row.empId);
    if (!emp) continue;

    // Managers are employees too, so the manager's EmployeeID resolves against
    // the same list to get their email.
    let mgrEmail = null;
    let mgrName = emp.managerName;

    let mgrActive = true;
    if (emp.managerId) {
      const mgr = lookup.get(emp.managerId);
      if (mgr?.email) {
        mgrEmail = mgr.email;
        mgrName = mgr.name || emp.managerName;
        mgrActive = mgr.isActive;
        if (!mgr.isActive) {
          // An active employee reporting to someone who has left is an
          // orphaned reporting line in Zoho — HR data to fix, not a code bug.
          const k = `${mgr.name} (${mgr.empId}, ${mgr.status})`;
          departedManagers.set(k, (departedManagers.get(k) || 0) + 1);
        }
      }
    }
    if (!mgrEmail && emp.managerEmail) mgrEmail = emp.managerEmail; // fallback

    if (!mgrEmail) {
      if (!emp.managerId) {
        noManager++;
        if (noManager <= 5) console.warn(`  [skip] ${row.empId} (${row.name}) — no Reporting_To set`);
      } else {
        noEmail++;
        console.warn(`  [skip] ${row.empId} — manager "${emp.managerId}" not found and no fallback email`);
      }
      continue;
    }

    if (!grouped.has(mgrEmail)) grouped.set(mgrEmail, { managerName: mgrName, rows: [], managerActive: mgrActive });
    grouped.get(mgrEmail).rows.push(row);
  }

  if (departedManagers.size) {
    console.warn(`\n  [WARN] ${departedManagers.size} manager(s) have left but still have active reportees:`);
    for (const [k, n] of [...departedManagers].sort((a, b) => b[1] - a[1]).slice(0, 10)) {
      console.warn(`    ${k} — ${n} active reportee(s)`);
    }
    console.warn(`    These reporting lines need reassigning in Zoho. Reports for them are held.`);
  }

  if (noManager) console.warn(`  -> skipped, no manager set: ${noManager} of ${rows.length}`);
  if (noEmail) console.warn(`  -> skipped, manager email not found: ${noEmail}`);

  for (const g of grouped.values()) g.rows.sort((a, b) => String(a.name).localeCompare(String(b.name)));
  return grouped;
}

// ==================================================================
// EXCEL
// ==================================================================

async function buildExcel(managerName, rows, dateStr, outDir) {
  const wb = new ExcelJS.Workbook();
  const sheet = wb.addWorksheet("Working Hours");

  sheet.columns = [
    { header: "Employee ID", key: "empId", width: 20 },
    { header: "Name", key: "name", width: 30 },
    { header: "Working Hours", key: "workingHours", width: 16 },
    { header: "Break Time", key: "breakTime", width: 13 },
    { header: "Sessions", key: "sessions", width: 10 },
    { header: "Site", key: "site", width: 18 },
    { header: "Note", key: "note", width: 36 },
  ];

  sheet.getRow(1).eachCell((c) => {
    c.font = { bold: true, color: { argb: "FFFFFFFF" } };
    c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF2E7D32" } };
    c.alignment = { horizontal: "center" };
  });

  rows.forEach((row) => {
    const r = sheet.addRow(row);
    r.getCell("workingHours").alignment = { horizontal: "center" };
    r.getCell("workingHours").font = { bold: true };
    if (row.unreliable) {
      r.eachCell((c) => {
        c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFF3CD" } };
      });
      r.getCell("note").font = { color: { argb: "FF9C6500" }, bold: true };
    }
  });

  sheet.views = [{ state: "frozen", ySplit: 1 }];
  sheet.autoFilter = { from: "A1", to: "G1" };

  const total = rows.reduce((s, r) => s + r.workedMinutes, 0);
  sheet.addRow([]);
  const tr = sheet.addRow({ name: "TEAM TOTAL", workingHours: hhmm(total) });
  tr.font = { bold: true };
  tr.getCell("workingHours").alignment = { horizontal: "center" };

  sheet.addRow([]);
  const note = sheet.addRow([
    "Working Hours = the sum of all paired check-in/check-out sessions, i.e. time on premises. " +
      "Restroom and gym gate scans are excluded. Meals taken in the on-site cafeteria are not badged " +
      "and so count as working time; Break Time reflects only periods where the employee badged out of " +
      "the building. Times are IST. Amber rows need verification.",
  ]);
  note.font = { italic: true, size: 9, color: { argb: "FF666666" } };

  const safe = String(managerName).replace(/[^\w\s-]/g, "").replace(/\s+/g, "_");
  const fp = path.join(outDir, `working_hours_${safe}_${dateStr}.xlsx`);
  await wb.xlsx.writeFile(fp);
  return fp;
}

// ==================================================================
// PDF
// ==================================================================

const COLS = [
  { header: "Emp ID", key: "empId", width: 90 },
  { header: "Name", key: "name", width: 165 },
  { header: "Working Hours", key: "workingHours", width: 75, align: "center" },
  { header: "Breaks", key: "breakTime", width: 55, align: "center" },
  { header: "Note", key: "note", width: 130 },
];
const TABLE_W = COLS.reduce((a, c) => a + c.width, 0); // 515pt = A4 minus 40pt margins

function pdfHeader(doc, y) {
  doc.fontSize(9);
  doc.rect(40, y, TABLE_W, 22).fill("#2E7D32");
  let x = 40;
  for (const c of COLS) {
    doc.fillColor("#fff").text(c.header, x + 5, y + 7, {
      width: c.width - 10,
      align: c.align || "left",
      lineBreak: false,
      ellipsis: true,
    });
    x += c.width;
  }
  return y + 22;
}

function buildPdf(managerName, rows, dateStr, outDir) {
  return new Promise((resolve, reject) => {
    const safe = String(managerName).replace(/[^\w\s-]/g, "").replace(/\s+/g, "_");
    const fp = path.join(outDir, `working_hours_${safe}_${dateStr}.pdf`);
    const doc = new PDFDocument({ margin: 40, size: "A4" });
    const stream = fs.createWriteStream(fp);
    doc.pipe(stream);

    doc.fontSize(17).fillColor("#000").text(`Daily Working Hours — ${dateStr}`);
    doc.fontSize(11).fillColor("#444").text(`Reporting Manager: ${managerName}`);
    // doc
    //   .fontSize(8)
    //   .fillColor("#666")
    //   .text(
    //     "Working Hours = total time on premises, summed across check-in/check-out sessions. Restroom and gym " +
    //       "scans excluded. Breaks shows time badged out of the building; cafeteria meals are not badged and " +
    //       "count as working time. Times IST.",
    //     { width: TABLE_W }
    //   );
    doc.moveDown(0.8);

    let y = pdfHeader(doc, doc.y);
    doc.fontSize(8.5);

    rows.forEach((row, i) => {
      if (y > 750) {
        doc.addPage();
        y = pdfHeader(doc, 40);
        doc.fontSize(8.5);
      }
      // Two-line row height so long names wrap rather than truncate.
      const nameLines = String(row.name).length > 26 ? 2 : 1;
      const h = nameLines === 2 ? 26 : 19;

      doc.rect(40, y, TABLE_W, h).fill(row.unreliable ? "#FFF3CD" : i % 2 === 0 ? "#FFFFFF" : "#F5F5F5");

      let x = 40;
      for (const c of COLS) {
        const isHours = c.key === "workingHours";
        doc
          .fillColor(c.key === "note" && row.unreliable ? "#9C6500" : "#000")
          .font(isHours ? "Helvetica-Bold" : "Helvetica")
          .fontSize(isHours ? 9.5 : 8.5)
          .text(String(row[c.key] ?? "-"), x + 5, y + 5, {
            width: c.width - 10,
            align: c.align || "left",
            lineBreak: c.key === "name",
            ellipsis: true,
            height: h - 6,
          });
        x += c.width;
      }
      doc.font("Helvetica").fontSize(8.5);
      y += h;
    });

    const total = rows.reduce((s, r) => s + r.workedMinutes, 0);
    doc.rect(40, y, TABLE_W, 20).fill("#E8F5E9");
    doc.fillColor("#000").font("Helvetica-Bold").fontSize(9);
    doc.text("TEAM TOTAL", 45, y + 6, { width: 250, lineBreak: false });
    doc.text(hhmm(total), 295, y + 6, { width: 75, align: "center", lineBreak: false });
    doc.font("Helvetica");
    y += 20;

    const bad = rows.filter((r) => r.unreliable).length;
    if (bad) {
      doc
        .fontSize(8)
        .fillColor("#9C6500")
        .text(
          `${bad} highlighted row(s) need verification — see the Note column.`,
          40,
          y + 12,
          { width: TABLE_W }
        );
    }

    doc.end();
    stream.on("finish", () => resolve(fp));
    stream.on("error", reject);
  });
}

// ==================================================================
// EMAIL
//
// Three providers, selected with EMAIL_PROVIDER in .env:
//
//   smtp   (default) — SMTP AUTH. Currently configured for Amazon SES:
//          host email-smtp.us-east-2.amazonaws.com, port 587, STARTTLS.
//          Also works unmodified against Microsoft 365's smtp.office365.com
//          by changing SMTP_HOST — same protocol, different server.
//
//          AMAZON SES — TWO THINGS THAT WILL BITE YOU:
//
//          1. SANDBOX MODE. Every new SES account starts sandboxed: max 200
//             messages/24h, max 1/second, and — this is the one that actually
//             bites — you can ONLY SEND TO PRE-VERIFIED RECIPIENT ADDRESSES.
//             With 88 reporting managers this WILL be hit immediately if the
//             account has not been moved to production. AWS Console -> SES ->
//             Account dashboard shows the current status. Moving to
//             production is a support request, approved in ~24h, and does
//             not require code changes here.
//
//          2. VERIFIED SENDER IDENTITY. Regardless of sandbox/production
//             status, the FROM address — no-reply@clouddestinations.com —
//             must be a verified identity in SES (domain or single-address
//             verification) in this exact AWS region (us-east-2). An
//             unverified sender is rejected at submission time with
//             "554 Message rejected: Email address is not verified", not
//             delivered-then-bounced.
//
//          The startup check in verifyMailConnection() cannot detect either
//          condition — SMTP AUTH succeeding only proves the credentials are
//          valid, not that the account can reach these recipients. Do a
//          single live send to your own address before pointing this at the
//          full manager list.
//
//   graph  — Microsoft Graph API with OAuth2 client credentials. No password
//          held anywhere. Kept available; not the current default.
//
//   gmail  — Gmail App Password. Kept for local testing.
// ==================================================================

const EMAIL_PROVIDER = (process.env.EMAIL_PROVIDER || "smtp").toLowerCase();

// Shared: the address managers see in the From field.
const MAIL_FROM = process.env.MAIL_FROM || process.env.GMAIL_USER || "";
// Optional: replies go here instead of to the sending mailbox.
const MAIL_REPLY_TO = process.env.MAIL_REPLY_TO || "";

// --- Microsoft Graph (OAuth2 client credentials) ---
const MS_TENANT_ID = process.env.MS_TENANT_ID || "";
const MS_CLIENT_ID = process.env.MS_CLIENT_ID || "";
const MS_CLIENT_SECRET = process.env.MS_CLIENT_SECRET || "";

// --- SMTP AUTH — configured for Amazon SES by default ---
// SMTP_USER is an SES SMTP username (looks like an IAM access key, e.g.
// "AKIA..."), which is NOT the same value as MAIL_FROM and must not default
// to it — unlike Microsoft 365, where the SMTP username usually is the
// sending mailbox. Getting this wrong fails auth outright, so no fallback
// is provided; SMTP_USER must be set explicitly.
const SMTP_HOST = process.env.SMTP_HOST || "email-smtp.us-east-2.amazonaws.com";
const SMTP_PORT = parseInt(process.env.SMTP_PORT || "587", 10);
const SMTP_USER = process.env.SMTP_USER || "";
const SMTP_PASSWORD = process.env.SMTP_PASSWORD || "";

/** Graph caps a single sendMail request at 4 MB including base64 attachments. */
const GRAPH_ATTACHMENT_LIMIT_BYTES = 3 * 1024 * 1024;

let _graphToken = null;
let _graphTokenExpiry = 0;

/**
 * Client-credentials token for Graph. Cached until shortly before expiry so a
 * run with 88 manager emails does not fetch 88 tokens.
 */
async function getGraphToken() {
  if (_graphToken && Date.now() < _graphTokenExpiry - 60000) return _graphToken;

  const url = `https://login.microsoftonline.com/${MS_TENANT_ID}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: MS_CLIENT_ID,
    client_secret: MS_CLIENT_SECRET,
    scope: "https://graph.microsoft.com/.default",
    grant_type: "client_credentials",
  });

  try {
    const resp = await axios.post(url, body.toString(), {
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      timeout: 30000,
    });
    _graphToken = resp.data.access_token;
    _graphTokenExpiry = Date.now() + (resp.data.expires_in || 3600) * 1000;
    return _graphToken;
  } catch (err) {
    const d = err.response?.data;
    throw new Error(
      `Microsoft token request failed: ${d?.error || err.message}` +
        (d?.error_description ? `\n  ${String(d.error_description).split("\n")[0]}` : "") +
        `\n  Check MS_TENANT_ID, MS_CLIENT_ID and MS_CLIENT_SECRET. A secret that has expired ` +
        `is the most common cause — they have a maximum lifetime of 24 months.`
    );
  }
}

function fileToAttachment(filePath) {
  const buf = fs.readFileSync(filePath);
  return {
    "@odata.type": "#microsoft.graph.fileAttachment",
    name: path.basename(filePath),
    contentBytes: buf.toString("base64"),
    size: buf.length,
  };
}

async function sendViaGraph(toEmail, subject, text, html, files) {
  const token = await getGraphToken();
  const attachments = files.map(fileToAttachment);

  const totalBytes = attachments.reduce((a, x) => a + x.size, 0);
  if (totalBytes > GRAPH_ATTACHMENT_LIMIT_BYTES) {
    // Beyond this, Graph requires a chunked upload session rather than an
    // inline sendMail. Reports are tens of KB, so this should never fire.
    throw new Error(
      `Attachments total ${(totalBytes / 1048576).toFixed(1)} MB, over the ${GRAPH_ATTACHMENT_LIMIT_BYTES / 1048576} MB inline limit.`
    );
  }
  attachments.forEach((a) => delete a.size);

  const message = {
    subject,
    body: { contentType: "HTML", content: html },
    toRecipients: [{ emailAddress: { address: toEmail } }],
    attachments,
  };
  if (MAIL_REPLY_TO) message.replyTo = [{ emailAddress: { address: MAIL_REPLY_TO } }];

  // /users/{sender} sends AS that mailbox using application permissions, so no
  // interactive sign-in and no stored mailbox password.
  const url = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(MAIL_FROM)}/sendMail`;

  try {
    await axios.post(url, { message, saveToSentItems: true }, {
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
      timeout: 60000,
    });
    // Graph returns 202 Accepted with an empty body — there is no message ID.
    return "accepted-by-graph";
  } catch (err) {
    const d = err.response?.data?.error;
    if (err.response?.status === 403) {
      throw new Error(
        `Graph refused the send (403 ${d?.code || ""}). Usual causes: the Mail.Send ` +
          `APPLICATION permission was never granted admin consent, or an Application ` +
          `Access Policy excludes ${MAIL_FROM}.` + (d?.message ? `\n  ${d.message}` : "")
      );
    }
    throw new Error(`Graph sendMail failed (${err.response?.status || err.code}): ${d?.message || err.message}`);
  }
}

// --- Nodemailer path (Microsoft SMTP AUTH, or Gmail) ---

let _transporter = null;

function getTransporter() {
  if (_transporter) return _transporter;

  if (EMAIL_PROVIDER === "gmail") {
    _transporter = nodemailer.createTransport({
      service: "gmail",
      pool: true,
      maxConnections: 5,
      maxMessages: 100,
      auth: { user: MAIL_FROM, pass: process.env.GMAIL_APP_PASSWORD || "" },
    });
  } else {
    _transporter = nodemailer.createTransport({
      host: SMTP_HOST,
      port: SMTP_PORT,
      secure: false, // port 587 upgrades via STARTTLS
      requireTLS: true,
      pool: true,
      maxConnections: 3,
      maxMessages: 50,
      auth: { user: SMTP_USER, pass: SMTP_PASSWORD },
    });
  }
  return _transporter;
}

/**
 * SES rejects an unverified sender or a sandbox-blocked recipient at SMTP
 * submission time with a distinctive 554/postmaster response, not a bounce
 * email later. Recognising it here turns a cryptic nodemailer stack trace
 * into an actionable message.
 */
function isSesRejection(err) {
  return /email address is not verified|not authorized to send/i.test(String(err?.response || err?.message || ""));
}

async function sendViaSmtp(toEmail, subject, text, html, files) {
  try {
    const info = await getTransporter().sendMail({
      from: MAIL_FROM,
      to: toEmail,
      replyTo: MAIL_REPLY_TO || undefined,
      subject,
      text,
      html,
      attachments: files.map((f) => ({ filename: path.basename(f), path: f })),
    });
    return info.messageId;
  } catch (err) {
    if (isSesRejection(err)) {
      throw new Error(
        `SES rejected this send: ${err.response || err.message}\n` +
          `  Either "${MAIL_FROM}" is not a verified identity in SES (region ${SMTP_HOST.split(".")[1] || "?"}), ` +
          `or the account is still in sandbox mode and "${toEmail}" is not a pre-verified recipient. ` +
          `Check AWS Console -> SES -> Verified identities, and the Account dashboard for sandbox status.`
      );
    }
    throw err;
  }
}

/** Confirms credentials work BEFORE spending 10 minutes on Zoho API calls. */
async function verifyMailConnection() {
  if (EMAIL_PROVIDER === "graph") {
    await getGraphToken();
    console.log(`Microsoft Graph token acquired. Sending as ${MAIL_FROM}.`);
    return;
  }

  await getTransporter().verify();
  console.log(`${EMAIL_PROVIDER === "gmail" ? "Gmail" : SMTP_HOST} SMTP verified. Sending as ${MAIL_FROM}.`);

  if (EMAIL_PROVIDER === "smtp" && /office365|outlook/i.test(SMTP_HOST)) {
    console.warn(
      "[NOTE] SMTP AUTH Basic Authentication is disabled by default for existing\n" +
        "       Exchange Online tenants from the end of December 2026. Move to\n" +
        "       EMAIL_PROVIDER=graph before then."
    );
  }

  if (EMAIL_PROVIDER === "smtp" && /email-smtp\..*\.amazonaws\.com/i.test(SMTP_HOST)) {
    console.warn(
      "[NOTE] Sending via Amazon SES. This verify() step only confirms the SMTP\n" +
        "       credentials are valid — it does NOT confirm the account can reach\n" +
        "       real recipients. Two things to check in the AWS Console before a\n" +
        `       live run: (1) "${MAIL_FROM}" is a verified identity in SES, and\n` +
        "       (2) the account has PRODUCTION ACCESS, not sandbox — sandboxed\n" +
        "       accounts can only send to pre-verified recipient addresses and cap\n" +
        "       out at 200 messages/24h. A single test send to your own address\n" +
        "       is the fastest way to confirm both before emailing every manager."
    );
  }
}

async function sendEmail(toEmail, managerName, dateStr, pdf, rows) {
  const total = hhmm(rows.reduce((s, r) => s + r.workedMinutes, 0));
  const absent = rows.filter((r) => r.workedMinutes === 0).length;

  const subject = `Daily Working Hours — ${dateStr}`;
  const text =
    `Hi ${managerName},\n\n` +
    `Attached are the working hours for your team for ${dateStr}.\n\n` +
    `Team size: ${rows.length}  |  No hours recorded: ${absent}  |  Team total: ${total}`;
  const html =
    `<p>Hi <strong>${managerName}</strong>,</p>` +
    `<p>Attached are the working hours for your team for <strong>${dateStr}</strong>.</p>` +
    `<p>Team size: ${rows.length} &nbsp;|&nbsp; No hours recorded: ${absent} &nbsp;|&nbsp; Team total: <strong>${total}</strong></p>`;

  const files = [pdf];
  return EMAIL_PROVIDER === "graph"
    ? sendViaGraph(toEmail, subject, text, html, files)
    : sendViaSmtp(toEmail, subject, text, html, files);
}

// ==================================================================
// MAIN
// ==================================================================

function validateConfig() {
  const missing = Object.entries({
    ZOHO_FORMS_CLIENT_ID: FORMS_CLIENT_ID,
    ZOHO_FORMS_CLIENT_SECRET: FORMS_CLIENT_SECRET,
    ZOHO_FORMS_REFRESH_TOKEN: FORMS_REFRESH_TOKEN,
    ZOHO_ATTENDANCE_CLIENT_ID: ATTENDANCE_CLIENT_ID,
    ZOHO_ATTENDANCE_CLIENT_SECRET: ATTENDANCE_CLIENT_SECRET,
    ZOHO_ATTENDANCE_REFRESH_TOKEN: ATTENDANCE_REFRESH_TOKEN,
    MAIL_FROM,
    ...(EMAIL_PROVIDER === "graph"
      ? { MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET }
      : EMAIL_PROVIDER === "gmail"
      ? { GMAIL_APP_PASSWORD: process.env.GMAIL_APP_PASSWORD }
      : { SMTP_USER, SMTP_PASSWORD }),
  })
    .filter(([, v]) => !v)
    .map(([k]) => k);
  if (missing.length) throw new Error(`Missing in .env: ${missing.join(", ")}`);

  // Intl timezone support is required. Node 14+ ships full ICU by default, but
  // a small-icu build would silently ignore timeZone options and corrupt every
  // calculation, so this fails loudly instead.
  const probe = new Intl.DateTimeFormat("en-US", { timeZone: BUSINESS_TZ, hour: "2-digit" }).format(new Date());
  if (!probe) throw new Error("This Node build lacks full ICU timezone data, which this script requires.");
}

async function main() {
  validateConfig();

  if (!DRY_RUN) {
    await verifyMailConnection();
  } else {
    console.log("DRY_RUN — files will be built, no emails sent.");
  }

  const dateStr = REPORT_DATE_OVERRIDE || istYesterday();
  const { y, mo, d } = parseDateStr(dateStr);
  const dayStart = istMidnight(y, mo, d);
  const dayEnd = new Date(dayStart.getTime() + 24 * 3600000);

  console.log(`\nReporting on ${dateStr} (IST)`);

  const formsToken = await getAccessToken(FORMS_CLIENT_ID, FORMS_CLIENT_SECRET, FORMS_REFRESH_TOKEN);
  const attToken = await getAccessToken(ATTENDANCE_CLIENT_ID, ATTENDANCE_CLIENT_SECRET, ATTENDANCE_REFRESH_TOKEN);

  console.log("Fetching employees...");
  const employees = await fetchEmployees(formsToken);
  console.log(`Employees fetched: ${employees.length}`);

  // INTEGRITY CHECK. Every manager is also an employee, so a Reporting_To that
  // points at somebody absent from this list means the list is incomplete —
  // the exact symptom of an unpaginated fetch, which silently returned the
  // first 200 records and made managers look like they had two reportees.
  const ids = new Set(employees.map((e) => e.empId));
  const danglingManagers = new Set(
    employees.map((e) => e.managerId).filter((m) => m && !ids.has(m))
  );
  if (danglingManagers.size) {
    console.warn(
      `[WARN] ${danglingManagers.size} manager ID(s) referenced but not present in the employee list, e.g. ` +
        `${[...danglingManagers].slice(0, 3).join(", ")}.\n` +
        `       The employee fetch may be incomplete. Those reportees fall back to Reporting_To.MailID.`
    );
  }
  if (employees.length % EMPLOYEE_PAGE_SIZE === 0 && employees.length > 0) {
    console.warn(
      `[WARN] Employee count is an exact multiple of the ${EMPLOYEE_PAGE_SIZE}-record page size. ` +
        `Verify no page was missed.`
    );
  }

  const zohoTz = process.env.ZOHO_TIMEZONE_OVERRIDE || (await detectZohoTimezone(attToken, employees, dateStr));
  const zohoDates = requiredZohoDates(dayStart, dayEnd, zohoTz);
  const offsetH = (IST_OFFSET_MINUTES - zoneOffsetMinutes(dayStart, zohoTz)) / 60;

  console.log(`Zoho answers in ${zohoTz} (IST is ${offsetH >= 0 ? "+" : ""}${offsetH}h ahead)`);
  console.log(`Fetching Zoho dates: ${zohoDates.join(", ")} per employee`);

  const activeStaff = employees.filter((e) => e.empId && e.isActive);
  const inactiveCount = employees.filter((e) => e.empId).length - activeStaff.length;
  if (inactiveCount) {
    console.log(
      `Excluding ${inactiveCount} former employee(s) [${EXCLUDE_EMPLOYEE_STATUS.join(", ")}] — ` +
        `${activeStaff.length} active. Saves roughly ${inactiveCount * 2} API calls.`
    );
  }
  const withIds = activeStaff;
  const totalCalls = withIds.length * zohoDates.length;
  const estMin = Math.ceil((totalCalls * REQUEST_DELAY_MS) / 60000);
  console.log(`~${totalCalls} API calls at ${Math.round(60000 / REQUEST_DELAY_MS)}/min — roughly ${estMin} min\n`);

  // PHASE 1 — fetch everything. Punches are held unfiltered so the gate census
  // can see the whole picture before any of them are excluded.
  const perEmployee = [];
  let calls = 0;
  const startedAt = Date.now();

  for (let i = 0; i < withIds.length; i++) {
    const emp = withIds[i];
    const responses = [];

    for (const zd of zohoDates) {
      try {
        const r = await fetchAttendance(attToken, emp.empId, zd);
        if (r) responses.push(r);
      } catch (err) {
        console.error(`  [error] ${emp.empId} @ ${zd}: ${err.message}`);
      }
      calls++;
      // Pace requests instead of batching them, so the run stays under the
      // per-minute threshold without ever idling for a full lock period.
      if (calls < totalCalls) await sleep(REQUEST_DELAY_MS);
    }

    perEmployee.push({
      emp,
      punches: buildPunchStream(responses, zohoTz),
      status: responses.find((r) => r?.status)?.status || "-",
    });

    if ((i + 1) % PROGRESS_EVERY === 0 || i + 1 === withIds.length) {
      const elapsed = (Date.now() - startedAt) / 1000;
      const rate = calls / (elapsed / 60);
      const remain = Math.ceil(((withIds.length - i - 1) * zohoDates.length) / Math.max(rate, 1));
      console.log(
        `  ${i + 1}/${withIds.length} employees | ${calls} calls | ${rate.toFixed(0)}/min | ~${remain} min left`
      );
    }
  }

  // PHASE 2 — census every badge location, then decide which are scan-only.
  const census = censusLocations(perEmployee.flatMap((x) => x.punches));
  const { ignored, reason } = resolveIgnoredLocations(census);

  console.log("\nBadge locations seen:");
  const width = Math.max(12, ...[...census.keys()].map((k) => k.length));
  for (const [loc, c] of [...census.entries()].sort((a, b) => b[1].in + b[1].out - (a[1].in + a[1].out))) {
    const verdict = ignored.has(loc) ? "IGNORED — no check-outs, scan-only gate" : "counted";
    const src = [...c.sources].join("/") || "-";
    console.log(`  ${loc.padEnd(width)}  in:${String(c.in).padStart(5)}  out:${String(c.out).padStart(5)}  [${src}]  ${verdict}`);
  }
  if (ignored.size) console.log(`  -> ${ignored.size} location(s) excluded via ${reason}`);
  else console.log("  -> no scan-only gates detected; all locations counted");

  // PHASE 3 — filter and compute.
  let lostAll = 0;
  const rows = perEmployee.map(({ emp, punches, status }) => {
    const kept = filterPunches(punches, ignored);

    // A scan-only reader is never the way into the building, so nobody should
    // lose their entire punch stream to filtering. If someone does, the gate
    // classification is wrong for them — keep the raw stream and flag it rather
    // than silently reporting a zero.
    const usable = kept.length === 0 && punches.length > 0 ? punches : kept;
    if (kept.length === 0 && punches.length > 0) lostAll++;

    // Site = the reader the employee actually entered through.
    const site = usable.find((p) => p.dir === "IN" && p.loc !== "-")?.loc || usable[0]?.loc || "-";
    return buildRow(emp, usable, dayStart, dayEnd, status, site);
  });

  const strayTotal = rows.reduce((a, r) => a + r.strayIns, 0);
  if (strayTotal) {
    console.log(`\n${strayTotal} extra check-in(s) ignored mid-session (restroom/gym scans while already clocked in).`);
  }
  if (lostAll) {
    console.warn(
      `[WARN] ${lostAll} employee(s) had ONLY scan-only-gate punches. Raw stream kept for them.\n` +
        `       Check the census — a real entry gate may have been misclassified.`
    );
  }
  const bugRows = rows.filter((r) => r.note.includes("bug")).length;
  if (bugRows) console.error(`[BUG] ${bugRows} row(s) exceed their punch span — double counting. Investigate before sending.`);

  const stillIn = rows.filter((r) => r.stillOpen).length;
  const unreliable = rows.filter((r) => r.unreliable).length;
  const worked = rows.filter((r) => r.workedMinutes > 0).length;
  console.log(`\nRows: ${rows.length} | with hours: ${worked} | flagged: ${unreliable}`);

  if (stillIn > 0) {
    console.warn(
      `\n[WARN] ${stillIn} employee(s) are still clocked in for ${dateStr}.\n` +
        `       Their hours are partial. The report is running too early — schedule\n` +
        `       it later in the morning so every session has closed.\n`
    );
  }
  if (unreliable / Math.max(1, rows.length) > 0.3) {
    console.warn(
      `[WARN] Over 30% of records still look unreliable after gate filtering.\n` +
        `       Check the badge location census above — if the entry gate was\n` +
        `       misclassified, set ENTRY_LOCATIONS explicitly in .env.\n`
    );
  }

  const grouped = groupByManager(employees, rows);
  console.log(`Manager groups: ${grouped.size}`);

  // Cross-check each group against the manager's actual reportee count in the
  // employee master, so a short team shows up here rather than in someone's inbox.
  const expected = new Map();
  for (const e of employees) {
    if (e.managerId && e.isActive) expected.set(e.managerId, (expected.get(e.managerId) || 0) + 1);
  }
  const byEmail = new Map(employees.filter((e) => e.email).map((e) => [e.email, e.empId]));
  let mismatches = 0;
  for (const [mgrEmail, g] of grouped.entries()) {
    const mgrId = byEmail.get(mgrEmail);
    const exp = mgrId ? expected.get(mgrId) : undefined;
    if (exp !== undefined && exp !== g.rows.length) {
      mismatches++;
      if (mismatches <= 10) {
        console.warn(`  [WARN] ${g.managerName}: ${g.rows.length} row(s) but ${exp} reportee(s) in Zoho`);
      }
    }
  }
  if (mismatches) console.warn(`  -> ${mismatches} manager group(s) short of their Zoho reportee count`);
  else console.log("  All manager groups match their Zoho reportee counts.");

  const outDir = path.join(REPORTS_BASE_DIR, dateStr);
  fs.mkdirSync(outDir, { recursive: true });
  console.log(`Output: ${outDir}\n`);

  let sent = 0;
  let failed = 0;
  let consecutiveSesRejections = 0;
  let sesCircuitTripped = false;

  // SES caps SMTP submission at 1 message/second by default (sandbox and most
  // fresh production accounts). A brief pause between sends keeps the run
  // under that without needing to know the account's exact quota.
  const sesPaceMs = /email-smtp\..*\.amazonaws\.com/i.test(SMTP_HOST)
    ? parseInt(process.env.SES_SEND_DELAY_MS || "1100", 10)
    : 0;

  if (TEST_MANAGER_EMAIL) {
    console.log(`[TEST FILTER] TEST_MANAGER_EMAIL is set. Only processing manager: ${TEST_MANAGER_EMAIL}\n`);
  }

  for (const [mgrEmail, g] of grouped.entries()) {
    const mgrName = g.managerName || mgrEmail;
    if (!g.rows.length) continue;

    if (TEST_MANAGER_EMAIL && mgrEmail.toLowerCase() !== TEST_MANAGER_EMAIL) {
      continue;
    }

    // Never email someone who has left. Their address is usually deactivated,
    // and if it is not, a departed employee receives current staff data.
    if (g.managerActive === false) {
      console.warn(`  [HELD] ${mgrName} (${mgrEmail}) has left — not emailing. Reassign these reportees in Zoho.`);
      continue;
    }

    // SANDBOX CIRCUIT BREAKER. A sandboxed SES account rejects every
    // unverified recipient identically. Without this, a sandboxed run prints
    // the same "not verified" failure up to 88 times and still takes as long
    // as a real run. Three in a row is enough to conclude the account is
    // sandboxed rather than one manager having a stale address, so the rest
    // are held with a single explanation instead of individually retried.
    if (sesCircuitTripped) {
      console.warn(`  [HELD] ${mgrName} (${mgrEmail}) — skipped, SES sandbox mode detected above.`);
      continue;
    }

    try {

      const pdf = await buildPdf(mgrName, g.rows, dateStr, outDir);
      if (DRY_RUN) {
        console.log(`  [dry] ${mgrName} (${mgrEmail}) — ${g.rows.length} reportee(s), files built`);
        continue;
      }
      if (sesPaceMs && sent > 0) await sleep(sesPaceMs);
      const id = await sendEmail(mgrEmail, mgrName, dateStr, pdf, g.rows);
      sent++;
      consecutiveSesRejections = 0;
      console.log(`  sent -> ${mgrEmail} (${g.rows.length} reportees) [${id}]`);
    } catch (err) {
      // One manager failing must not abort the rest of the run.
      failed++;
      console.error(`  [FAILED] ${mgrName} (${mgrEmail}): ${err.message}`);

      if (isSesRejection(err)) {
        consecutiveSesRejections++;
        if (consecutiveSesRejections >= 3) {
          sesCircuitTripped = true;
          console.error(
            `\n[STOPPED] ${consecutiveSesRejections} consecutive SES rejections — this account is almost\n` +
              `          certainly in SANDBOX MODE, which only permits sending to pre-verified\n` +
              `          recipient addresses. Remaining managers will be listed as held rather\n` +
              `          than retried individually. Request production access in the AWS SES\n` +
              `          console (Account dashboard -> Request production access), or verify\n` +
              `          each recipient address individually for a small pilot group.\n`
          );
        }
      }
    }
  }

  console.log(`\n${sent} sent, ${failed} failed, across ${grouped.size} manager(s).`);
  console.log(`Files saved in ${outDir}`);
}

if (require.main === module) {
  main()
    .then(() => {
      if (_transporter) _transporter.close(); // null when EMAIL_PROVIDER=graph
      console.log("Done.");
    })
    .catch((err) => {
      console.error("\nFatal:", err.message || err);
      if (_transporter) _transporter.close(); // null when EMAIL_PROVIDER=graph
      process.exit(1);
    });
}

module.exports = {
  parsePunch,
  extractManagerId,
  isSesRejection,
  containsIdToken,
  buildPunchStream,
  collapseDoubleSwipes,
  censusLocations,
  resolveIgnoredLocations,
  filterPunches,
  computeHours,
  requiredZohoDates,
  istMidnight,
  formatIst,
  hhmm,
  zoneOffsetMinutes,
};

// ==================================================================
// SCHEDULING
// ==================================================================
// Run mid-morning so every session from the previous day has closed —
// including evening logins from home and US-aligned shifts ending ~01:00 IST.
//
//     CRON_TZ=Asia/Kolkata
//     0 9 * * 2-6  /usr/bin/node /path/to/report.js >> /var/log/attendance.log 2>&1
//
// (2-6 = Tue-Sat, so Monday's run does not report on a Sunday. Use * * * for
// all seven days if weekend attendance matters.)
//
// The target date comes from the IST calendar, not the host clock, so a UTC
// server still reports the correct day.
