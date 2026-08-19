/**
 * BRYC Senior Portal -- backend for bryc-senior-portal.html and bryc-admin.html.
 *
 * IMPORTANT: deploy this as a STANDALONE Apps Script project, NOT bound to the
 * tracker spreadsheet. The tracker already has its own bound script (the
 * "Request Router" menu and the _RoutingLog tab); pasting this into that
 * project would collide with it. This script reaches the tracker with
 * SpreadsheetApp.openById(TRACKER_SHEET_ID) instead, so the two never touch.
 *
 * Setup:
 * 1. script.google.com > New project. Paste this file in. Name it
 *    "BRYC Senior Portal Backend".
 * 2. Deploy > New deployment > Web app.
 *      - Execute as: Me
 *      - Who has access: Anyone
 * 3. Paste the Web App URL into APPS_SCRIPT_URL in BOTH bryc-senior-portal.html
 *    and bryc-admin.html, and set PORTAL_URL below.
 * 4. Run diagnoseColumns() once from the editor and read the Execution log --
 *    it confirms every response column was found. Then run diagnoseAttendance().
 *
 * The first run will ask for authorization (it reads two spreadsheets, sends
 * mail, and calls Google's userinfo endpoint to check who an admin is).
 * ---------------------------------------------------------------------------
 * WHERE THE DATA LIVES
 *   Tracker  "PY27_BRYC Senior Request Form Tracker" (TRACKER_SHEET_ID)
 *     "Form Responses"      -- raw Google Form submissions, one row per request
 *     "Fee Payment Balance" -- one row per fellow: Starting Balance, Additional
 *                              Funds, Total, Current Balance, then repeating
 *                              [Type of Fee | Amount | Date Paid] triplets
 *   Attendance "PY27 Fall Fellow Attendance" (ATTENDANCE_SHEET_ID), tab
 *     "Grade 12" -- weekly P / A / HP marks and a Total % column
 *
 * COLUMNS ARE RESOLVED BY HEADER TEXT, never by position. The response layout
 * has already changed once between program years; RESPONSE_FIELDS below claims
 * headers in form order, so repeated headers ("Certification", "What is your
 * Username and Password?") still land on the right keys, and re-ordering or
 * inserting a question does not silently shift everybody's data.
 *
 * TRACKING COLUMNS appended after the form's own columns, created automatically
 * by bryc-admin.html: Fulfilled | Notes | Stage | Message to Student.
 *   "Notes" is internal and is NEVER sent to a student.
 *   "Message to Student" is written expressly to be read by them.
 * ---------------------------------------------------------------------------
 */

const TRACKER_SHEET_ID = '1J56K1q9DrIUC4YiLx_5R9vC3e2uq2KExcl6t1-0VhKw';
const RESPONSE_SHEET = 'Form Responses';
const FEE_BALANCE_SHEET = 'Fee Payment Balance';

// Paste the deployed URL of bryc-senior-portal.html here; blank just omits the
// "check your portal" link from student emails.
const PORTAL_URL = '';

// Counselors allowed to use bryc-admin.html and to trigger a student email.
// Keep in sync with ALLOWED_ADMIN_EMAILS there. Addresses follow BRYC's
// firstname@thebryc.org pattern -- correct any that differ.
const ADMIN_EMAILS = [
  'jakia@thebryc.org',      // runs senior funds
  'catherine@thebryc.org', 'rachel@thebryc.org', 'tim@thebryc.org',
  'rhonda@thebryc.org', 'tavidee@thebryc.org', 'richard@thebryc.org',
  'kirsten@thebryc.org',
  'aareena@thebryc.org'     // UWorld / MasteryPrep
];

// $150 for the ACADEMIC YEAR, the same for every fellow -- not per semester,
// and no spring reset.
//
// This deliberately OVERRIDES the Fee Payment Balance tab, whose Starting
// Balance column still reads $100 from last year. Because that tab's
// "Current Balance" formula is derived from the stale $100, honouring it would
// show every fellow $50 short -- so the balance here is computed as
// DEFAULT_STARTING_FUNDS + Additional Funds - payments instead. Once the sheet's
// Starting Balance column is updated to 150, the two agree and this override
// stops mattering.
const DEFAULT_STARTING_FUNDS = 150;

// Values for the "Stage" tracking column, in order.
const PORTAL_STAGES = ['Received', 'In review', 'Approved', 'Complete'];

const STAGE_COLUMNS = ['Fulfilled', 'Notes', 'Stage', 'Message to Student'];

/* ---- Live attendance (separate spreadsheet, updated weekly) ----
 * "PY27 Fall Fellow Attendance". Layout confirmed 2026-08-19:
 *   header row 5, data from row 6
 *   B "First Name", C "Last Name", F "Grade"
 *   G "Startup Aug 10", then H..X = "Week 1" .. "Week 17"
 *   Y "Total %"  <- the percentage the portal shows seniors
 * Marks are P (present), A (absent), HP (half present).
 * Columns are located BY HEADER TEXT, so adding weeks won't break the lookup.
 * NOTE: this is the FALL sheet -- spring will need a new id/tab. */
const ATTENDANCE_SHEET_ID = '13j0yKuyPdOx_zdQgzyegD0yiAiuu3BQEVtKEWlJ20yU';
const ATTENDANCE_TAB = 'Grade 12';       // seniors; other grades have their own tabs
const ATTENDANCE_HEADER_ROW = 5;         // 1-based
const ATTENDANCE_MIN_PERCENT = 80;       // BRYC's 80% requirement
const ATTENDANCE_CACHE_SECONDS = 900;    // sheet changes weekly; 15 min is plenty

/* ---- Sign-in security ---- */
const PIN_TTL_SECONDS = 600;             // one-time PIN lifetime: 10 minutes
const SESSION_TTL_SECONDS = 1800;        // signed-in session: 30 minutes
const MAX_PIN_ATTEMPTS = 5;              // wrong PINs before lockout
const LOCKOUT_SECONDS = 900;             // lockout length: 15 minutes
const MAX_PIN_REQUESTS_PER_HOUR = 5;     // anti inbox-spam throttle

/* ---------------------------------------------------------------------------
 * RESPONSE_FIELDS -- the whole column map, in the order the questions appear.
 *
 * `m` matches a header that CONTAINS the string; `eq` matches exactly. Each
 * entry claims the leftmost header not already claimed, so the duplicated
 * headers resolve in form order.
 *
 * `safe: true` means the value may be shown back to the senior in the portal.
 * Anything without it is withheld -- deliberately an allowlist, so a question
 * added to the form stays hidden until somebody marks it safe here. Every
 * credential field and the date of birth are withheld on purpose.
 * ------------------------------------------------------------------------ */
const RESPONSE_FIELDS = [
  { k: 'name',            m: 'first and last name' },
  { k: 'email',           m: 'professional (non-school) email' },
  { k: 'counselor',       m: 'who is your counselor',                     safe: true, label: 'Counselor/Advisor' },
  { k: 'role',            m: 'choose which one best applies to you' },
  { k: 'fellowSupport',   m: 'which type of support are you requesting? (bryc fellows' },
  { k: 'adviseeSupport',  m: 'which type of support are you requesting? (bryc advisees' },
  { k: 'capWho',          m: 'choose which one best applies to you',      safe: true, label: 'College Aid Pro: requested by' },
  { k: 'guardianName',    m: 'guardian first and last name',              safe: true, label: 'Guardian name' },
  { k: 'capHighSchool',   m: 'student high school',                       safe: true, label: 'High school' },
  { k: 'capCounselor',    m: "school counselor's full name",              safe: true, label: 'School counselor' },
  { k: 'cert1',           eq: 'certification',                            safe: true, label: 'Certification' },
  { k: 'cert2',           eq: 'certification',                            safe: true, label: 'Certification' },
  { k: 'writingProject',  m: 'which writing project',                     safe: true, label: 'Writing project' },
  { k: 'writingDetail',   m: 'detail as possible about the writing',      safe: true, label: 'About the writing' },
  { k: 'writingStage',    m: 'where are you in your writing process',     safe: true, label: 'Writing stage' },
  { k: 'writingSupport',  m: 'what type of support do you need right now', safe: true, label: 'Support needed' },
  { k: 'writingDocLink',  m: 'link to your writing document',             safe: true, label: 'Writing document' },
  { k: 'writingExtra',    m: 'additional information that will help your writing coach', safe: true, label: 'Additional info' },
  { k: 'partnerApplied',  m: 'have you completed an application to a bryc partner school', safe: true, label: 'Applied to a partner school?' },
  { k: 'partnerSchool',   m: 'which bryc partner school have you applied to', safe: true, label: 'Partner school' },
  { k: 'partnerCreds',    eq: 'username and password' },                  // withheld
  { k: 'dob',             m: 'what is your date of birth' },              // withheld
  { k: 'highSchool',      m: 'which high school do you attend',           safe: true, label: 'High school' },
  { k: 'waiverAck',       m: 'fee waiver process for partner schools',    safe: true, label: 'Fee waiver process' },
  { k: 'feeType',         m: 'which type of fee are you requesting your senior funds for', safe: true, label: 'Fee type' },
  { k: 'feeOther',        m: 'explain what you are requesting your senior funds for', safe: true, label: 'What it is for' },
  { k: 'college',         eq: 'college',                                  safe: true, label: 'College' },
  { k: 'loginPageUrl',    m: 'login page: copy and paste the url' },      // withheld
  { k: 'appFeeCreds',     m: 'what is your username and password' },      // withheld
  { k: 'appFeeAmount',    m: 'how much is the application fee',           safe: true, label: 'Application fee' },
  { k: 'appFeeDeadline',  m: 'when is the deadline for the fee to be paid', safe: true, label: 'Fee deadline' },
  { k: 'actTestDate',     m: 'date of the test you are requesting to take', safe: true, label: 'ACT test date' },
  { k: 'actTestLocation', m: 'where would you like to take the act',      safe: true, label: 'ACT test location' },
  { k: 'actScoreFour',    m: 'which four colleges do you want your score sent to', safe: true, label: 'Colleges for scores' },
  { k: 'actCreds1',       m: 'what is your act username and password' },  // withheld
  { k: 'actTestsToSend',  m: 'which test(s) do you want to send scores from', safe: true, label: 'Tests to send' },
  { k: 'actScoreColleges', m: 'which college(s) do you want your score sent to', safe: true, label: 'Colleges for score report' },
  { k: 'actCreds2',       m: 'what is your act username and password' },  // withheld
  { k: 'depositType',     m: 'which type of deposit do you want paid',    safe: true, label: 'Deposit type' },
  { k: 'depositAmount',   m: 'how much is the deposit',                   safe: true, label: 'Deposit amount' },
  { k: 'depositDueDate',  m: 'what is the due date',                      safe: true, label: 'Deposit due' },
  { k: 'depositSite',     m: 'what is the login website' },               // withheld
  { k: 'depositCreds',    m: 'what is your username and password' },      // withheld
  { k: 'masteryActRetake', m: 'i plan to take at least one more act test', safe: true, label: 'Plans to retake the ACT' },
  { k: 'masteryRequest',  m: 'requesting a login for masteryprep',        safe: true, label: 'MasteryPrep login requested' },
  { k: 'uworldActRetake', m: 'i plan to take at least one more act test', safe: true, label: 'Plans to retake the ACT' },
  { k: 'uworldRequest',   m: 'requesting access to a uworld account',     safe: true, label: 'UWorld account requested' }
];

// Exact option strings the form writes, used for routing. If the form's
// wording changes these must change with it.
const ROLE_FELLOW = 'I am a Fellow at a BRYC campus';
const ROLE_ADVISEE = 'I am an Advisee';
const TYPE_SENIOR_FUNDS = 'I am requesting to use my senior funds';
const TYPE_WRITING = 'I am requesting writing coach support';
const TYPE_CAP = 'I am requesting College Aid Pro access';
const TYPE_UWORLD = 'I am requesting a UWorld account';
const TYPE_MASTERYPREP = 'I am requesting a MasteryPrep login for ACT support';
const TYPE_PARTNER_WAIVER = 'I am requesting a college application fee waiver for a BRYC partner school';

/* ---------------------------------------------------------------------------
 * Column resolution
 * ------------------------------------------------------------------------ */

function tracker_() {
  return SpreadsheetApp.openById(TRACKER_SHEET_ID);
}

function normHeader_(h) {
  return String(h == null ? '' : h).replace(/\s+/g, ' ').trim().toLowerCase();
}

/**
 * headers -> { key: columnIndex }. Claims left to right in RESPONSE_FIELDS
 * order so repeated headers map to successive keys. Missing keys are simply
 * absent, which every caller tolerates.
 */
function resolveColumns_(headers) {
  const norm = headers.map(normHeader_);
  const taken = {};
  const map = {};
  RESPONSE_FIELDS.forEach(function (f) {
    for (let i = 0; i < norm.length; i++) {
      if (taken[i] || !norm[i]) continue;
      const hit = f.eq ? (norm[i] === f.eq) : (norm[i].indexOf(f.m) > -1);
      if (!hit) continue;
      taken[i] = true;
      map[f.k] = i;
      return;
    }
  });
  // The tracking columns the admin dashboard appends.
  STAGE_COLUMNS.forEach(function (name) {
    const i = norm.indexOf(name.toLowerCase());
    if (i > -1) map['_' + name.replace(/\s+/g, '')] = i;
  });
  return map;
}

function cell_(row, map, key) {
  const i = map[key];
  if (i === undefined || i === null) return '';
  const v = row[i];
  return v === null || v === undefined ? '' : v;
}

function str_(row, map, key) {
  return String(cell_(row, map, key)).trim();
}

/** Which support type this row represents, whichever branch it came down. */
function requestTypeOf_(row, map) {
  const role = str_(row, map, 'role');
  const t = role === ROLE_ADVISEE ? str_(row, map, 'adviseeSupport') : str_(row, map, 'fellowSupport');
  return t || str_(row, map, 'fellowSupport') || str_(row, map, 'adviseeSupport');
}

/** True when BRYC actually pays money for this request. */
function isPaymentRequest_(row, map) {
  if (requestTypeOf_(row, map) !== TYPE_SENIOR_FUNDS) return false;
  return str_(row, map, 'feeType') !== '';
}

/**
 * Run from the editor after deploying. Confirms every response column was
 * located, without printing anybody's data.
 */
function diagnoseColumns() {
  const sheet = tracker_().getSheetByName(RESPONSE_SHEET);
  if (!sheet) { console.log('FAILED: no tab named "' + RESPONSE_SHEET + '".'); return; }
  const headers = sheet.getDataRange().getValues()[0] || [];
  const map = resolveColumns_(headers);
  const missing = RESPONSE_FIELDS.filter(function (f) { return map[f.k] === undefined; }).map(function (f) { return f.k; });
  console.log('Header columns found: ' + headers.filter(String).length);
  console.log('Mapped ' + (RESPONSE_FIELDS.length - missing.length) + ' of ' + RESPONSE_FIELDS.length + ' fields.');
  if (missing.length) {
    console.log('NOT FOUND (check the question wording in the form): ' + missing.join(', '));
  } else {
    console.log('All response fields resolved.');
  }
  STAGE_COLUMNS.forEach(function (n) {
    console.log('Tracking column "' + n + '": ' + (map['_' + n.replace(/\s+/g, '')] !== undefined ? 'present' : 'not created yet (open bryc-admin.html once)'));
  });
}

/* ---------------------------------------------------------------------------
 * FEE PAYMENT BALANCE
 * ---------------------------------------------------------------------------
 * One row per fellow. Fixed columns then repeating [Type of Fee | Amount |
 * Date Paid] triplets running rightwards. "Current Balance" is BRYC's own
 * formula -- this script reads it and never writes to it.
 * ------------------------------------------------------------------------ */

function feeBalanceLayout_(headers) {
  const norm = headers.map(normHeader_);
  const layout = {
    name: norm.indexOf('fellow name'),
    email: norm.indexOf('email'),
    starting: norm.indexOf('starting balance'),
    additional: norm.indexOf('additional funds'),
    total: norm.indexOf('total'),
    current: norm.indexOf('current balance'),
    triplets: []
  };
  // Every "Type of Fee" starts a triplet; Amount and Date Paid follow it.
  for (let i = 0; i < norm.length; i++) {
    if (norm[i] !== 'type of fee') continue;
    layout.triplets.push({
      type: i,
      amount: (norm[i + 1] === 'amount') ? i + 1 : -1,
      date: (norm[i + 2] === 'date paid') ? i + 2 : -1
    });
  }
  return layout;
}

function moneyNumber_(v) {
  if (v === '' || v === null || v === undefined) return null;
  if (typeof v === 'number') return v;
  const n = parseFloat(String(v).replace(/[^0-9.\-]/g, ''));
  return isNaN(n) ? null : n;
}

/** Reads one fellow's funds row into the shape the portal renders. */
function fundsFor_(email, displayName) {
  const sheet = tracker_().getSheetByName(FEE_BALANCE_SHEET);
  if (!sheet) return null;

  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return null;
  const layout = feeBalanceLayout_(rows[0]);
  if (layout.name === -1) return null;

  const wantEmail = normalizeEmail_(email);
  const wantName = nameKey_(displayName);

  let match = null;
  if (layout.email > -1 && wantEmail) {
    for (let i = 1; i < rows.length; i++) {
      if (normalizeEmail_(rows[i][layout.email]) === wantEmail) { match = rows[i]; break; }
    }
  }
  if (!match && wantName) {
    for (let i = 1; i < rows.length; i++) {
      if (nameKey_(rows[i][layout.name]) === wantName) { match = rows[i]; break; }
    }
  }
  if (!match) return null;

  // Everyone gets the same annual amount; only Additional Funds varies.
  const additional = layout.additional > -1 ? (moneyNumber_(match[layout.additional]) || 0) : 0;
  const startingFunds = DEFAULT_STARTING_FUNDS + additional;

  const spending = [];
  let totalSpent = 0;
  layout.triplets.forEach(function (t) {
    const type = String(match[t.type] || '').trim();
    const amt = t.amount > -1 ? moneyNumber_(match[t.amount]) : null;
    if (!type && amt === null) return;
    if (amt !== null) totalSpent += amt;
    const rawDate = t.date > -1 ? match[t.date] : '';
    spending.push({
      description: type || 'Senior funds payment',
      amount: amt === null ? 0 : amt,
      date: rawDate instanceof Date
        ? Utilities.formatDate(rawDate, Session.getScriptTimeZone(), 'MMM d, yyyy')
        : String(rawDate || '')
    });
  });

  // Computed here rather than read from the sheet's Current Balance, which is
  // still derived from the stale $100 starting figure (see DEFAULT_STARTING_FUNDS).
  const remaining = startingFunds - totalSpent;

  return {
    startingFunds: startingFunds,
    totalSpent: totalSpent,
    remainingFunds: remaining,
    additionalFunds: additional,
    spending: spending
  };
}

/* ---------------------------------------------------------------------------
 * PORTAL DATA
 * ------------------------------------------------------------------------ */

/** True if this email appears on any submitted request. */
function portalEmailIsKnown_(email) {
  const sheet = tracker_().getSheetByName(RESPONSE_SHEET);
  if (!sheet) return false;
  const rows = sheet.getDataRange().getValues();
  if (!rows.length) return false;
  const map = resolveColumns_(rows[0]);
  if (map.email === undefined) return false;
  for (let i = 1; i < rows.length; i++) {
    if (normalizeEmail_(rows[i][map.email]) === email) return true;
  }
  return false;
}

function buildPortalData_(email) {
  const tz = Session.getScriptTimeZone();
  const sheet = tracker_().getSheetByName(RESPONSE_SHEET);

  const requests = [];
  const roles = {};
  let displayName = '';

  if (sheet) {
    const rows = sheet.getDataRange().getValues();
    const map = rows.length ? resolveColumns_(rows[0]) : {};

    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (map.email === undefined || normalizeEmail_(row[map.email]) !== email) continue;

      if (!displayName) displayName = str_(row, map, 'name');
      const role = str_(row, map, 'role');
      if (role) roles[role] = true;

      const type = requestTypeOf_(row, map);

      // Only fields explicitly marked safe are ever sent to the browser.
      const details = [];
      RESPONSE_FIELDS.forEach(function (f) {
        if (!f.safe) return;
        const raw = cell_(row, map, f.k);
        if (raw === '' ) return;
        details.push({
          label: f.label || f.k,
          value: raw instanceof Date ? Utilities.formatDate(raw, tz, 'MMM d, yyyy') : String(raw)
        });
      });

      const fulfilled = String(cell_(row, map, '_Fulfilled') || '').toUpperCase() === 'TRUE';
      const rawStage = String(cell_(row, map, '_Stage') || '').trim();
      const stage = PORTAL_STAGES.indexOf(rawStage) > -1 ? rawStage : (fulfilled ? 'Complete' : 'Received');

      const ts = row[0];
      requests.push({
        submitted: ts instanceof Date ? Utilities.formatDate(ts, tz, 'MMM d, yyyy') : String(ts || ''),
        submittedLong: ts instanceof Date ? Utilities.formatDate(ts, tz, "MMM d, yyyy 'at' h:mm a") : String(ts || ''),
        sortKey: ts instanceof Date ? ts.getTime() : 0,
        type: type || 'BRYC support request',
        friendlyType: friendlyType_(type),
        role: role,
        fulfilled: fulfilled,
        stage: stage,
        // "Notes" is internal and is never included here.
        message: String(cell_(row, map, '_MessagetoStudent') || '').trim(),
        isPayment: isPaymentRequest_(row, map),
        details: details
      });
    }
  }

  requests.sort(function (a, b) { return b.sortKey - a.sortKey; });
  requests.forEach(function (r) { delete r.sortKey; });

  // Advisees are not eligible for senior funds and are not on the Fellow
  // Attendance sheet, so they get neither panel. Anyone with even one Fellow
  // request counts as a Fellow.
  const isFellow = !!roles[ROLE_FELLOW];
  const isAdviseeOnly = !isFellow && !!roles[ROLE_ADVISEE];

  const funds = isAdviseeOnly ? null : fundsFor_(email, displayName);
  const attendance = isAdviseeOnly ? null : attendanceFor_(displayName);

  return {
    status: 'ok',
    name: displayName,
    firstName: (displayName || '').split(/\s+/)[0] || '',
    role: isAdviseeOnly ? 'advisee' : 'fellow',
    funds: funds,
    attendance: attendance,
    requests: requests
  };
}

/* ---------------------------------------------------------------------------
 * ROUTER
 * ------------------------------------------------------------------------ */

function doGet(e) {
  const callback = e.parameter.callback;
  const action = String(e.parameter.action || '').trim();
  let result;

  try {
    if (action === 'requestPin') {
      result = handleRequestPin(e.parameter.email);
    } else if (action === 'verifyPin') {
      result = handleVerifyPin(e.parameter.email, e.parameter.pin);
    } else if (action === 'portal') {
      result = handlePortal(e.parameter.token);
    } else if (action === 'signOut') {
      result = handleSignOut(e.parameter.token);
    } else if (action === 'notifyStudent') {
      result = handleNotifyStudent(e);
    } else {
      result = { status: 'error', message: 'Unknown action.' };
    }
  } catch (err) {
    // Never leak a stack trace or sheet internals to an unauthenticated caller.
    console.error('doGet failed (action=' + action + '): ' + err);
    result = { status: 'error', message: 'Something went wrong. Please try again.' };
  }

  const json = JSON.stringify(result);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

/* ===========================================================================
 * STUDENT NOTIFICATIONS (triggered by bryc-admin.html)
 * ---------------------------------------------------------------------------
 * The admin dashboard writes to the sheet with the Sheets API, which this
 * script never sees, so an onEdit trigger would not fire -- the dashboard
 * calls this endpoint after a successful save instead. Because the Web App is
 * public, the caller is verified here: their live Google OAuth token is checked
 * against Google's own userinfo endpoint and the resulting email must be in
 * ADMIN_EMAILS. The message body is composed from sheet data, never from
 * anything the caller sends, so this cannot be used to mail arbitrary text.
 * ======================================================================== */

function friendlyType_(type) {
  const map = {};
  map[TYPE_SENIOR_FUNDS]    = 'senior funds request';
  map[TYPE_WRITING]         = 'writing coach request';
  map[TYPE_CAP]             = 'College Aid Pro request';
  map[TYPE_UWORLD]          = 'UWorld account request';
  map[TYPE_MASTERYPREP]     = 'MasteryPrep login request';
  map[TYPE_PARTNER_WAIVER]  = 'partner school fee waiver request';
  return map[String(type || '').trim()] || 'BRYC support request';
}
function verifyAdmin_(accessToken) {
  const token = String(accessToken || '').trim();
  if (!token) return null;
  try {
    const res = UrlFetchApp.fetch('https://www.googleapis.com/oauth2/v3/userinfo', {
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    });
    if (res.getResponseCode() !== 200) return null;
    const info = JSON.parse(res.getContentText());
    if (!info.email || String(info.email_verified) !== 'true') return null;
    const email = normalizeEmail_(info.email);
    return ADMIN_EMAILS.indexOf(email) > -1 ? email : null;
  } catch (err) {
    console.error('Admin verification failed: ' + err);
    return null;
  }
}

function portalLinkLine_() {
  return PORTAL_URL ? '\nYou can see the details anytime in your BRYC Senior Portal:\n' + PORTAL_URL + '\n' : '';
}


function handleNotifyStudent(e) {
  const admin = verifyAdmin_(e.parameter.adminToken);
  if (!admin) return { status: 'forbidden' };

  const sheetRow = Number(e.parameter.row || 0);
  const kind = String(e.parameter.kind || '').trim();
  if (!sheetRow || sheetRow < 2) return { status: 'error', message: 'Bad row number.' };

  const sheet = tracker_().getSheetByName(RESPONSE_SHEET);
  if (!sheet) return { status: 'error', message: 'Response tab not found.' };

  const rows = sheet.getDataRange().getValues();
  if (sheetRow > rows.length) return { status: 'error', message: 'Row not found.' };
  const map = resolveColumns_(rows[0]);
  const row = rows[sheetRow - 1];

  const to = normalizeEmail_(cell_(row, map, 'email'));
  if (!isEmailShaped_(to)) return { status: 'error', message: 'That row has no usable student email.' };

  const first = str_(row, map, 'name').split(/\s+/)[0] || 'there';
  const kindOf = friendlyType_(requestTypeOf_(row, map));

  let subject, body;

  if (kind === 'fulfilled') {
    subject = 'Your BRYC ' + kindOf + ' is complete';
    body = 'Hi ' + first + ',\n\n'
      + (isPaymentRequest_(row, map)
          ? 'Good news -- BRYC has paid the fee for your ' + kindOf + '. It has come out of your senior funds, and your updated balance is in your portal.\n'
          : 'Good news -- your ' + kindOf + ' has been taken care of.\n')
      + portalLinkLine_()
      + '\nIf something does not look right, reply to your BRYC counselor or advisor.\n\nThe BRYC Team';

  } else if (kind === 'message') {
    const message = String(cell_(row, map, '_MessagetoStudent') || '').trim();
    if (!message) return { status: 'error', message: 'There is no message on that row to send.' };
    subject = 'An update on your BRYC ' + kindOf;
    body = 'Hi ' + first + ',\n\n'
      + 'Your BRYC counselor left you a note about your ' + kindOf + ':\n\n'
      + '    ' + message + '\n'
      + portalLinkLine_()
      + '\nThe BRYC Team';

  } else {
    return { status: 'error', message: 'Unknown notification type.' };
  }

  try {
    MailApp.sendEmail({ to: to, subject: subject, body: body });
  } catch (err) {
    console.error('Student notification failed: ' + err);
    return { status: 'error', message: 'The email could not be sent.' };
  }

  return { status: 'ok', sentTo: to };
}

function portalSecret_() {
  const props = PropertiesService.getScriptProperties();
  let secret = props.getProperty('PORTAL_SECRET');
  if (!secret) {
    secret = Utilities.getUuid() + Utilities.getUuid();
    props.setProperty('PORTAL_SECRET', secret);
  }
  return secret;
}

function normalizeEmail_(raw) {
  return String(raw || '').trim().toLowerCase();
}

function isEmailShaped_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email);
}

function sha256_(text) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8);
  return bytes.map(function (b) {
    return ('0' + (b & 0xff).toString(16)).slice(-2);
  }).join('');
}

// Cache keys are hashed so a raw email address is never sitting in the cache.
function emailKey_(email) {
  return sha256_(normalizeEmail_(email) + '|' + portalSecret_()).slice(0, 32);
}

function hashPin_(pin, email) {
  return sha256_(String(pin) + '|' + normalizeEmail_(email) + '|' + portalSecret_());
}

// Length-independent comparison so response timing doesn't leak the PIN.
function safeEquals_(a, b) {
  a = String(a);
  b = String(b);
  if (a.length !== b.length) return false;
  let diff = 0;
  for (let i = 0; i < a.length; i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return diff === 0;
}

function randomToken_() {
  // 256 bits from Apps Script's CSPRNG-backed UUIDs, hashed into hex.
  return sha256_(Utilities.getUuid() + Utilities.getUuid() + portalSecret_());
}

function randomPin_() {
  const n = Math.floor(Math.random() * 900000) + 100000; // 100000-999999
  return String(n);
}

/* --------------------------- sign-in: step 1 --------------------------- */

function handleRequestPin(rawEmail) {
  const email = normalizeEmail_(rawEmail);

  // Same response no matter what, so nobody can probe which emails exist.
  const genericOk = {
    status: 'ok',
    message: 'If that email is on file with BRYC, a 6-digit code is on its way. It expires in 10 minutes.'
  };

  if (!isEmailShaped_(email)) {
    return { status: 'error', message: 'That does not look like an email address.' };
  }

  const cache = CacheService.getScriptCache();
  const key = emailKey_(email);

  if (cache.get('lock:' + key)) {
    return {
      status: 'locked',
      message: 'Too many attempts. For your security this email is locked for 15 minutes. Try again later or contact your counselor.'
    };
  }

  const reqCount = Number(cache.get('req:' + key) || 0);
  if (reqCount >= MAX_PIN_REQUESTS_PER_HOUR) {
    return {
      status: 'throttled',
      message: 'You have requested several codes recently. Please wait an hour before requesting another, or contact your counselor.'
    };
  }
  cache.put('req:' + key, String(reqCount + 1), 3600);

  if (!portalEmailIsKnown_(email)) return genericOk; // silent no-op

  const pin = randomPin_();
  cache.put('pin:' + key, hashPin_(pin, email), PIN_TTL_SECONDS);
  cache.remove('att:' + key);

  try {
    MailApp.sendEmail({
      to: email,
      subject: 'Your BRYC sign-in code: ' + pin,
      body:
        'Your one-time BRYC Senior Portal sign-in code is:\n\n' +
        '    ' + pin + '\n\n' +
        'It expires in 10 minutes and can only be used once.\n\n' +
        'If you did not try to sign in, you can ignore this email -- nobody can see your\n' +
        'information without this code. Let your BRYC counselor or advisor know if you\n' +
        'keep getting codes you did not ask for.\n\n' +
        'The BRYC Team'
    });
  } catch (err) {
    console.error('PIN email failed: ' + err);
    return { status: 'error', message: 'We could not send your code right now. Please try again in a minute.' };
  }

  return genericOk;
}

/* --------------------------- sign-in: step 2 --------------------------- */

function handleVerifyPin(rawEmail, rawPin) {
  const email = normalizeEmail_(rawEmail);
  const pin = String(rawPin || '').trim();
  const cache = CacheService.getScriptCache();
  const key = emailKey_(email);

  if (cache.get('lock:' + key)) {
    return { status: 'locked', message: 'Too many incorrect codes. This email is locked for 15 minutes.' };
  }

  const stored = cache.get('pin:' + key);
  const bad = { status: 'bad_pin', message: 'That code is not right, or it has expired. Request a new one.' };

  if (!stored || !/^\d{6}$/.test(pin)) {
    registerFailedPin_(cache, key);
    return bad;
  }

  if (!safeEquals_(stored, hashPin_(pin, email))) {
    registerFailedPin_(cache, key);
    return bad;
  }

  // Correct: burn the PIN and its attempt counter, then issue a session.
  cache.remove('pin:' + key);
  cache.remove('att:' + key);

  const token = randomToken_();
  cache.put('sess:' + token, email, SESSION_TTL_SECONDS);

  return { status: 'ok', token: token, expiresInSeconds: SESSION_TTL_SECONDS };
}

function registerFailedPin_(cache, key) {
  const attempts = Number(cache.get('att:' + key) || 0) + 1;
  if (attempts >= MAX_PIN_ATTEMPTS) {
    cache.put('lock:' + key, '1', LOCKOUT_SECONDS);
    cache.remove('att:' + key);
    cache.remove('pin:' + key);
  } else {
    cache.put('att:' + key, String(attempts), LOCKOUT_SECONDS);
  }
}

function handleSignOut(token) {
  if (token) CacheService.getScriptCache().remove('sess:' + String(token).trim());
  return { status: 'ok' };
}

/* ------------------------------ the data ------------------------------ */

function handlePortal(rawToken) {
  const token = String(rawToken || '').trim();
  if (!token) return { status: 'unauthorized' };

  const cache = CacheService.getScriptCache();
  const email = cache.get('sess:' + token);
  if (!email) return { status: 'unauthorized' };

  // Sliding expiry: staying active keeps you signed in, idling signs you out.
  cache.put('sess:' + token, email, SESSION_TTL_SECONDS);

  return buildPortalData_(email);
}

/**
 * True if this email appears anywhere we'd let someone sign in from:
 * a submitted request row, or a Fellow Status row.
 */
/* ===========================================================================
 * LIVE ATTENDANCE LOOKUP
 * ---------------------------------------------------------------------------
 * Reads the weekly "PY27 Fall Fellow Attendance" spreadsheet directly, so
 * nobody has to retype absences into the Fellow Status tab. The whole tab is
 * parsed once and cached; a senior's own row is then picked out of it.
 *
 * The attendance sheet splits names into First Name / Last Name, while the
 * request form collects a single "Name" field -- matchAttendanceName_ bridges
 * that, exact-match first, then a first-token + last-token fallback so
 * middle names and suffixes ("Lanieu, Jr") still line up.
 * ======================================================================== */

/** Strips punctuation/extra spaces and lowercases, for name comparison. */
function nameKey_(s) {
  return String(s || '')
    .toLowerCase()
    .replace(/[.,'`\u2019-]/g, ' ')
    .replace(/\b(jr|sr|ii|iii|iv)\b/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

/** "92.5%" / 0.925 / 92.5 -> 92.5. Returns null when there's nothing usable. */
function toPercent_(raw) {
  if (raw === '' || raw === null || raw === undefined) return null;
  if (typeof raw === 'number') {
    if (isNaN(raw)) return null;
    // Percent-formatted cells come back as a fraction (0.925), plain numbers don't.
    return raw <= 1 ? raw * 100 : raw;
  }
  const txt = String(raw).trim();
  if (!txt) return null;
  const n = parseFloat(txt.replace('%', ''));
  if (isNaN(n)) return null;
  return (txt.indexOf('%') === -1 && n <= 1) ? n * 100 : n;
}

/**
 * Parses the attendance tab into { key: {...} } keyed by normalized name.
 * Cached, because every portal sign-in would otherwise re-read the sheet.
 */
function attendanceIndex_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('attendance_index');
  if (cached) {
    try { return JSON.parse(cached); } catch (err) { /* fall through and rebuild */ }
  }

  let sheet;
  try {
    const ss = SpreadsheetApp.openById(ATTENDANCE_SHEET_ID);
    sheet = ss.getSheetByName(ATTENDANCE_TAB);
  } catch (err) {
    console.error('Could not open the attendance spreadsheet: ' + err);
    return null;
  }
  if (!sheet) {
    console.error('Attendance tab "' + ATTENDANCE_TAB + '" not found.');
    return null;
  }

  const rows = sheet.getDataRange().getValues();
  if (rows.length < ATTENDANCE_HEADER_ROW) return null;

  const headers = rows[ATTENDANCE_HEADER_ROW - 1].map(function (h) { return String(h).trim(); });
  const lower = headers.map(function (h) { return h.toLowerCase(); });

  const firstIdx = lower.indexOf('first name');
  const lastIdx = lower.indexOf('last name');
  const totalIdx = lower.indexOf('total %');
  if (firstIdx === -1 || lastIdx === -1) {
    console.error('Attendance tab is missing "First Name" / "Last Name" headers.');
    return null;
  }

  // Session columns: anything headed "Startup ..." or "Week N ...".
  const sessionCols = [];
  for (let c = 0; c < headers.length; c++) {
    if (/^(startup|week)\b/i.test(headers[c])) sessionCols.push(c);
  }

  const index = {};
  for (let r = ATTENDANCE_HEADER_ROW; r < rows.length; r++) {
    const row = rows[r];
    const full = String(row[firstIdx] || '').trim() + ' ' + String(row[lastIdx] || '').trim();
    const key = nameKey_(full);
    if (!key) continue;

    let present = 0, absent = 0, half = 0, asOf = '';
    const absentSessions = [], halfSessions = [];
    for (let i = 0; i < sessionCols.length; i++) {
      const c = sessionCols[i];
      const mark = String(row[c] || '').trim().toUpperCase();
      if (!mark) continue;
      if (mark === 'P') present++;
      else if (mark === 'A') { absent++; absentSessions.push(headers[c]); }
      else if (mark === 'HP') { half++; halfSessions.push(headers[c]); }
      else continue;
      asOf = headers[c];
    }

    const recorded = present + absent + half;
    let percent = totalIdx > -1 ? toPercent_(row[totalIdx]) : null;
    // Fall back to computing it ourselves if the Total % cell is empty,
    // counting a half-present session as half a session attended.
    if (percent === null && recorded > 0) percent = ((present + half * 0.5) / recorded) * 100;

    index[key] = {
      percent: percent,
      present: present,
      absent: absent,
      halfPresent: half,
      sessionsRecorded: recorded,
      asOf: asOf,
      absentSessions: absentSessions,
      halfSessions: halfSessions
    };
  }

  try { cache.put('attendance_index', JSON.stringify(index), ATTENDANCE_CACHE_SECONDS); }
  catch (err) { /* index too big to cache; just don't cache it */ }

  return index;
}

/** Finds one senior's attendance row by their single-field name. */
function matchAttendanceName_(index, fellowName) {
  if (!index) return null;
  const key = nameKey_(fellowName);
  if (!key) return null;
  if (index[key]) return index[key];

  // Fallback: match on first token + last token, so a middle name on one side
  // and not the other doesn't cause a miss.
  const parts = key.split(' ');
  if (parts.length < 2) return null;
  const wanted = parts[0] + ' ' + parts[parts.length - 1];
  const keys = Object.keys(index);
  for (let i = 0; i < keys.length; i++) {
    const p = keys[i].split(' ');
    if (p.length >= 2 && (p[0] + ' ' + p[p.length - 1]) === wanted) return index[keys[i]];
  }
  return null;
}

/** The attendance object handed to the portal, or null if nothing is known. */
function attendanceFor_(fellowName) {
  let rec = null;
  try { rec = matchAttendanceName_(attendanceIndex_(), fellowName); }
  catch (err) { console.error('Attendance lookup failed: ' + err); }
  if (!rec || rec.sessionsRecorded === 0 || rec.percent === null) return null;

  const percent = Math.round(rec.percent * 10) / 10;
  return {
    percent: percent,
    present: rec.present,
    absent: rec.absent,
    halfPresent: rec.halfPresent,
    sessionsRecorded: rec.sessionsRecorded,
    minPercent: ATTENDANCE_MIN_PERCENT,
    meetsRequirement: percent >= ATTENDANCE_MIN_PERCENT,
    asOf: rec.asOf,
    // The exact sessions missed, so a senior below the line knows which days
    // to actually ask their counselor about.
    absentSessions: rec.absentSessions || [],
    halfSessions: rec.halfSessions || [],
    source: 'sheet'
  };
}

/**
 * Run this once from the Apps Script editor (Run > diagnoseAttendance) after
 * deploying, and read the Execution log. It confirms the tab was found, which
 * columns were detected, and how many fellows were indexed -- without printing
 * anybody's attendance record.
 */
function diagnoseAttendance() {
  const index = attendanceIndex_();
  if (!index) {
    console.log('FAILED: could not read "' + ATTENDANCE_TAB + '" in the attendance spreadsheet. '
      + 'Check ATTENDANCE_SHEET_ID / ATTENDANCE_TAB, and that this script is authorized.');
    return;
  }
  const keys = Object.keys(index);
  console.log('OK: indexed ' + keys.length + ' fellows from tab "' + ATTENDANCE_TAB + '".');
  if (keys.length) {
    const sample = index[keys[0]];
    console.log('Session columns detected per fellow: ' + sample.sessionsRecorded
      + ' recorded so far, most recent column header "' + sample.asOf + '".');
  }
  console.log('If the count looks wrong, check that the header row is row '
    + ATTENDANCE_HEADER_ROW + ' and that name columns are "First Name" / "Last Name".');
}
