---
name: "ctsppum-patient-list"
description: "Generate a formatted Excel patient surgery list from the CTSPPUM Google Calendar (ctsppum@gmail.com). Use this skill whenever the user wants to extract a patient list, surgery schedule, operation list, or patient register from the calendar — even if they say \"get my patients\", \"who's booked for surgery\", \"make the op list\", or just \"export to Excel\". Also use it when patients need to be notified of a POSTPONED or RESCHEDULED operation, since the output feeds the CTSPPUM Patient Notifier. Columns produced: Name, RN, Operation Planned, Date of Surgery, Contact Numbers, Old Date. Missing phone numbers are filled from iPesakit Patient Info."
---

# CTSPPUM Patient List Extractor

Pulls events from the CTSPPUM Google Calendar and parses them into a structured Excel file
with columns: **Name · RN · Operation Planned · Date of Surgery · Contact Numbers · Old Date**.

When a calendar event has no phone number, the patient's registered contact number is
scraped from **iPesakit → Patient Info** (step 6) so no row goes out blank.

The output is drop-in input for the **CTSPPUM Patient Notifier** (v3.3, a claude.ai Artifact —
see the `ctsppum-patient-notifier` skill in this plugin).

---

## How events are structured

Each patient event in the calendar has its summary in one of these patterns:

| Pattern | Example |
|---------|---------|
| `RN - NAME OPERATION` | `40001244 - LEE MING FATT CABG` |
| `RN NAME _ OPERATION` | `40002203 MOHAMAD FAIZ BIN OTHMAN _ VATS Sympathetectomy` |
| `RN NAME OPERATION` (no separator) | `40001381 GOH SIEW LIN Right Upper Lobectomy` |
| `OPERATION RN - NAME` (op prefix) | `MVR 40000696 - NURUL AIN BINTI IBRAHIM` |
| `NAME RN _ OPERATION` | `KAMALA A/P RAMASAMY 40002340 _ CABG` |
| `NAME_RN_ OPERATION` (no spaces) | `SAMAD BIN HASSAN_40000148_ mitral valve` |

**RN is always an 8-digit number.** Events without one (holidays, workshops, MDTs, etc.) are skipped.
Phone numbers appear in the event **description** in Malaysian format (`01x-xxxxxxx`).

Status prefixes the team adds to titles are stripped automatically before parsing:
`<!>`, `(To reschedule)`, `(To be postponed to Sept, KIV 9/9)`, a leading `Mr`/`Mrs`,
and a trailing `confirmed`.

---

## The Old Date column

`Old Date` is **blank by default**. Fill it only for a patient whose operation date has
actually moved. `Date of Surgery` always holds the ACTUAL / current date.

| Situation | Date of Surgery | Old Date | What the notifier sends |
|---|---|---|---|
| Normal booking, never moved | 30/09/2026 | *(blank)* | Standard admission script |
| Postponed | 22/09/2026 | 08/09/2026 | Change-of-date script |
| Brought forward | 27/08/2026 | 08/09/2026 | Change-of-date script |

The notifier decides purely from the spreadsheet — no toggle to set. Admission
(`Date-2days`) and medication-cessation (`Date-5days`) are always computed from
`Date of Surgery`, never from the old date.

---

## Step-by-step workflow

### 1 · Ask for the date range

Ask the user which date range to extract (e.g. "next 30 days", "this month", or specific
start/end dates). Convert their answer to ISO 8601 strings (e.g. `2026-06-01`, `2026-06-30`).
"Next week" = the coming Monday to Sunday.

### 2 · Ask whether any dates have changed

Ask if any patient in this batch has had their operation date moved. If yes, collect each
one's **RN** and **previous date**. Sources for the old date, in order of reliability:

1. The user tells you directly (handwritten shift note, verbal handover).
2. The calendar event description — postponements are often logged there
   (e.g. "Planned for AVR + CABG on 17/8/26", "Postponed x 1 (patient request)",
   "informed pt date change on 22/8 (old date 09/09/2026)").

Cross-check the two when both exist and flag any disagreement rather than guessing.
If no dates changed, skip to step 3 and omit `old_dates.json` entirely.

### 3 · Fetch all events from the CTSPPUM calendar

Use the Google Calendar MCP tool `list_events` with:
- `calendarId`: `ctsppum@gmail.com`
- `startTime` / `endTime`: the date range from step 1
- `pageSize`: `250`

If the response includes `nextPageToken`, call `list_events` again with that token to retrieve
all pages. Collect every event.

### 4 · Save events (and any old dates) to JSON

```python
import json
events = [...]   # all events collected above
with open('events.json', 'w') as f:
    json.dump(events, f)

# Only if some dates changed — RN -> previous date, DD/MM/YYYY
old_dates = {"40000277": "08/09/2026", "40000148": "23/09/2026"}
with open('old_dates.json', 'w') as f:
    json.dump(old_dates, f)
```

### 5 · Install openpyxl (if needed) and run the parser script

```bash
pip install openpyxl --break-system-packages -q
python scripts/build_patient_list.py events.json patient_list.xlsx old_dates.json
```

Drop the third argument when nothing was rescheduled.

The script will:
- Strip status prefixes, then skip any event without an 8-digit RN
- Parse Name, RN, and Operation from the summary
- Extract all Malaysian phone numbers from the description
- Attach the previous date to any RN listed in `old_dates.json`
- Sort patients by surgery date
- Output a colour-formatted Excel file with changed dates highlighted in amber

It prints `{"status": "ok", "patients": N, "date_changed": M, "output": "patient_list.xlsx"}`

Then print every row and fix any misparsed Name / Operation cells by hand with openpyxl
(e.g. lower-case titles like `tan ah kow- 40000389- left uniportal vats ...`, or a
title with only RN + name — take the operation from the event description).

### 6 · Fill missing phone numbers from iPesakit

**Only for rows whose `Contact Numbers` cell is blank** (no number in the calendar
description). Rows that already have a calendar number are left as they are.

Where the number lives (user-confirmed path): iPesakit → **Inpatient** → key in RN → Search
→ **Patient Info** → *Contact No (Mobile)* / *Contact No (Home)*. Do it by fetch, not by clicking:

1. **Get a logged-in iPesakit tab** (Claude in Chrome). Call `tabs_context_mcp`; reuse a tab
   already on `ipesakit.ummc.edu.my` with a `keyid=` in its URL. If none, open a tab to
   `https://ipesakit.ummc.edu.my`. If it lands on the login page, **do not sign in** — ask the
   user to click Sign In themselves, then continue.
2. **Read the live keyid from the tab URL** — never navigate to the root `Default.asp` or
   reload the login page (that kills the SSO session):
   `location.href.match(/keyid=([A-Z0-9]+)/)[1]`
3. **One `javascript_tool` call for all missing RNs.** POST the Registration search form; the
   response IS the Patient Info page (no result-list click needed):

```js
const keyid = location.href.match(/keyid=([A-Z0-9]+)/)[1];
const out = [];
for (const rn of ['RN1','RN2']) {
  const body = new URLSearchParams({rn, alternateid:'', namapesakit:'', tarikhlahir:'', keyid, btnsearch:'Search'});
  const h = await fetch('/Registration/ResultCariPesakit.asp', {method:'POST',
    headers:{'Content-Type':'application/x-www-form-urlencoded'}, body}).then(r=>r.text());
  const t = new DOMParser().parseFromString(h,'text/html').body.textContent.replace(/\s+/g,' ');
  const g = re => { const m = t.match(re); return m ? m[1].trim() : ''; };
  out.push({rn,
    name:   g(/NAME\s*:\s*([A-Z@'\/ .]+?)\s+(RELIGION|MYKAD)/i),
    home:   g(/Contact No \(Home\)\s*:?\s*([0-9\- ]*?)\s*(Prefered|Contact|Address)/i),
    mobile: g(/Contact No \(Mobile\)\s*:?\s*([0-9\- ]*?)\s*(EMAIL|Address|PDPA|Contact)/i)});
}
JSON.stringify(out)
```

4. **Check the returned `name` matches the calendar name** for that RN before using the number.
   If every row comes back with an empty name, the keyid is stale — re-read it from
   `location.href` and retry once; if the page looks like a login page, ask the user to sign in.
5. **Write the numbers into the Excel** `Contact Numbers` cell (mobile first, then home,
   joined with ` / `), preserving the existing formatting:

```python
import openpyxl
wb = openpyxl.load_workbook('patient_list.xlsx'); ws = wb.active
fill = {'40000512': '0110000000'}   # RN -> number(s) from iPesakit
for r in ws.iter_rows(min_row=2):
    if r[1].value in fill and not r[4].value:
        r[4].value = fill[r[1].value]
wb.save('patient_list.xlsx')
```

6. If iPesakit also has no number, leave the cell blank and flag that RN to the user.

Notes:
- `EMR/search_patient.asp` (the EMR dashboard search) does **not** return Patient Info —
  use `/Registration/ResultCariPesakit.asp`.
- Keep the JS return compact (just rn/name/home/mobile). Returning page snippets that contain
  `?`, `&` or `=` can be swallowed by the output filter.
- Read-only: never click UPDATE / Update Record on the Patient Info page.

### 7 · Verify before handing over

Check that `patients` matches the number of real cases in the range, and that
`date_changed` equals the number of postponements the user described. A count of 0
patients almost always means the RN pattern failed — see Troubleshooting. Confirm every
row now has a contact number or is explicitly flagged as missing.

### 8 · Ask where to save the file

Ask the user where they'd like the file saved. If they have a folder connected, move it there.
Otherwise place it in the outputs directory for download.

### 9 · Present the file to the user

Share `patient_list.xlsx`, and mention it can be dropped straight into the CTSPPUM
Patient Notifier. Then offer to open the notifier (follow the `ctsppum-patient-notifier`
skill: open https://claude.ai/artifact/PN4YWSqYjDKnqUJcgBekuz with the Artifact tool). In the summary, say which numbers came from iPesakit rather than the
calendar, note any number that belongs to a relative (e.g. "pt son" in the description),
and list any RN still without a number.

---

## Parsing notes

- The `_` and `—` characters are treated as hard separators between name and operation.
- Underscores are padded before RN extraction, so `NAME_40000148_` still yields the RN.
- Without a separator, the script accumulates consecutive ALL-CAPS words as the patient name
  and uses the first transition to mixed-case as the start of the operation.
- Known surgical abbreviations (CABG, MVR, AVR, VATS, etc.) are never treated as name words.
- Multiple phone numbers are joined with ` / `. The notifier splits on `/` and normalises
  each to international format, so keep that separator.
- If a summary generates an unexpected name or operation, review the raw `summary` field in
  the event and consider whether the calendar entry uses a new format not yet handled.

---

## Troubleshooting

| Symptom | Likely cause | Fix |
|---------|-------------|-----|
| `patients: 0` | No events have 8-digit RNs in the date range | Widen the date range or confirm calendar ID |
| One patient missing | RN hugged by underscores, or a `(...)` prefix on the title | Both are handled now; if it persists, check the RN really is 8 digits |
| Name contains operation abbreviation | Abbreviation not in `KNOWN_OP_WORDS` | Add it to the set in `build_patient_list.py` |
| Phone number missing | Number not in the calendar description | Run step 6 — pull it from iPesakit Patient Info |
| iPesakit lookup returns empty name for every RN | Stale keyid, or session logged out | Re-read keyid from the tab URL and retry once; if on the login page, ask the user to sign in |
| iPesakit returns a number but name doesn't match | Wrong RN in the calendar title | Don't use the number; flag the RN to the user |
| Operation blank | Title has only RN + name (e.g. "Mr X (40001107) confirmed") | Take it from the event description, or fill manually — the notifier prints "Operation TBC" otherwise |
| `date_changed: 0` but dates did move | RN key in `old_dates.json` doesn't match, or old date equals the new one | Keys must be the 8-digit RN as a string, dates as `DD/MM/YYYY` |
| Operation shows disease context (e.g. "MVR SLE") | Summary uses `—` to append clinical notes | Expected; user can edit the cell manually |