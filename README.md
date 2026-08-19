# BRYC Senior Portal

Two pages plus the backend that serves them.

| File | Who it's for | What it does |
|---|---|---|
| `index.html` | Seniors | Sign in with an emailed 6-digit code; see senior funds, spending, attendance, and the status of every request |
| `admin.html` | BRYC counselors | Google sign-in; work the request queue, record payments, message students |
| `apps-script-code.gs` | — | Standalone Apps Script backend. **Not** served by this site |

## Where the data lives

- **Tracker** `1J56K1q9DrIUC4YiLx_5R9vC3e2uq2KExcl6t1-0VhKw` — *PY27_BRYC Senior Request Form Tracker*
  - `Form Responses` — raw Google Form submissions
  - `Fee Payment Balance` — one row per fellow; funds are read *and written* here
- **Attendance** `13j0yKuyPdOx_zdQgzyegD0yiAiuu3BQEVtKEWlJ20yU` — *PY27 Fall Fellow Attendance*, tab `Grade 12`. Read-only, and doubles as the senior roster.
- **Form** — [Class of 2027: Request for BRYC Support](https://docs.google.com/forms/d/e/1FAIpQLSfXlBXAMZ1SXLVGY6o3Y5YfYwhrFZM76leJtioZnNWuFEvZhQ/viewform)

Columns are matched **by header text, not position**, so re-ordering or adding a form question won't silently shift anyone's data.

## Setup

The Apps Script is **standalone** — do not paste it into the tracker's own script project, which already runs the *Request Router*.

1. script.google.com → New project → paste `apps-script-code.gs`
2. Deploy → New deployment → Web app → *Execute as: Me*, *Who has access: Anyone*
3. Put the Web App URL into `APPS_SCRIPT_URL` in **both** `index.html` and `admin.html`, and into `PORTAL_URL` in `Code.gs`
4. Run `diagnoseColumns()` then `diagnoseAttendance()` from the editor and read the log — both report what they found without printing student data

Firebase (project `senior-team-site`) is already configured: Google sign-in on, Sheets API enabled, OAuth user type Internal, `auth/spreadsheets` scope registered. Once this site has a real domain, add it under Firebase → Authentication → Settings → Authorized domains.

Counselors need **Editor** on the tracker and **Viewer** on the attendance workbook. Signing in does not grant sheet access on its own.

## Things worth knowing

- Senior funds are **$150 for the academic year**, the same for everyone, with no per-semester reset. The `Fee Payment Balance` tab still reads `$100` in its Starting Balance column, so the code overrides it; once that column is updated the override is a no-op.
- The admin dashboard asks for sign-in on every page load, deliberately — the Sheets token is never persisted.
- Sign-in is gated on the BRYC email domain in `admin.html`; the per-person allowlist lives in `ADMIN_EMAILS` in `Code.gs`, which is never served to a browser. Two of those addresses were inferred from the organisation's usual `firstname@` pattern and should be confirmed against the real accounts.
