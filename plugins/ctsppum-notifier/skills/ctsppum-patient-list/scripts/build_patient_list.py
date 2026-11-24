#!/usr/bin/env python3
"""
Build patient surgery list Excel from CTSPPUM Google Calendar events.

Usage:
    python build_patient_list.py events.json output.xlsx [old_dates.json]

Input:  events.json     — either a list of events or {"events": [...]}
        old_dates.json  — OPTIONAL {"<RN>": "DD/MM/YYYY", ...} mapping an RN to the
                          patient's PREVIOUS operation date, for cases that were
                          postponed or brought forward. Omit a patient (or the whole
                          file) when the date never changed.

Output: Formatted Excel file with columns: Name, RN, Operation Planned,
        Date of Surgery, Contact Numbers, Old Date

The output is drop-in input for the CTSPPUM Patient Notifier (v2.1+):
  - "Date of Surgery" is always the ACTUAL / current operation date.
  - "Old Date" is blank unless the date changed. When it is filled in and differs
    from Date of Surgery, the notifier automatically switches that patient to the
    change-of-date script; admission and medication-cessation dates are still
    calculated from Date of Surgery.
"""

import json
import re
import sys
from datetime import datetime

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


# ─── Parsing helpers ──────────────────────────────────────────────────────────

def extract_rn(text):
    """Return the first 8-digit registration number found in text, or None.

    Underscores are word characters, so a summary like "SAMAD BIN HASSAN_40000148_"
    defeats a plain \b...\b match. Pad separators before searching.
    """
    padded = re.sub(r'[_\u2014]', ' ', text)
    m = re.search(r'(?<!\d)(\d{8})(?!\d)', padded)
    return m.group(1) if m else None


def clean_summary(summary):
    """Strip status prefixes the team adds to event titles.

    Handles: "<!>", "(To reschedule)", "(To be postponed to Sept, KIV 9/9)",
             a leading "Mr"/"Mrs", and a trailing "confirmed".
    """
    s = re.sub(r'^\s*<!>\s*', '', summary)
    s = re.sub(r'^\s*\([^)]*\)\s*', '', s)
    s = re.sub(r'\bconfirmed\b', '', s, flags=re.IGNORECASE)
    s = s.replace('(', ' ').replace(')', ' ')
    s = re.sub(r'^\s*(Mr|Mrs|Ms|Encik|Puan)\.?\s+', '', s, flags=re.IGNORECASE)
    return re.sub(r'\s+', ' ', s).strip()


def extract_contacts(description):
    """
    Find phone numbers in the event description.
    Matches any number of the form: 3 digits + optional separator (dash/space) + 7 or 8 digits.
    Handles formats: 012-3456789 / 012 3456789 / 0123456789 / 0121234567
    Also catches numbers with labels like "Dr X - 012 0000000" or "016 88776655 - son"
    Returns all found numbers joined with ' / ', or empty string.
    """
    if not description:
        return ''

    # Strip basic HTML tags (calendar sometimes sends <br>)
    text = re.sub(r'<br\s*/?>', '\n', description, flags=re.IGNORECASE)
    text = re.sub(r'<[^>]+>', '', text)

    # Match: 3 digits, optional dash or space separator, then 7 or 8 digits
    # This catches: 012-3456789, 012 3456789, 0123456789, 016 88776655, etc.
    # Negative lookbehind/lookahead prevent matching middle of longer digit strings (e.g. RNs)
    raw = re.findall(r'(?<!\d)(\d{3})[\s\-]?(\d{7,8})(?!\d)', text)

    phones = set()
    for prefix, suffix in raw:
        digits = prefix + suffix
        if 10 <= len(digits) <= 11:
            phones.add(digits)

    return ' / '.join(sorted(phones))


# Common CTS operation abbreviations — these are NOT patient name words
KNOWN_OP_WORDS = {
    'CABG', 'MVR', 'AVR', 'TVR', 'DVR', 'TVT', 'TAVI', 'TAVR',
    'VATS', 'UVATS', 'MIDCAB', 'MIS', 'MAZE', 'LAA', 'REDO',
    'EVAR', 'TEVAR', 'PEA', 'LVAD', 'ECMO', 'MV', 'AV', 'TV',
    'PV', 'TAP', 'SOE', 'IBN',
}


def _is_name_word(word):
    """True if a word looks like part of a Malaysian patient name."""
    alpha = re.sub(r'[^A-Za-z/]', '', word)
    if not alpha:
        return False
    # Never treat known surgical abbreviations as name words
    if alpha.upper() in KNOWN_OP_WORDS:
        return False
    # Name connectors (common in Malay/Indian names)
    if alpha.upper() in ('BIN', 'BINTI', 'BINTE', 'AP', 'AL'):
        return True
    if alpha in ('A/P', 'A/L', 'A/C'):
        return True
    # General rule: all-uppercase and at least 2 characters
    return alpha == alpha.upper() and len(alpha) >= 2


def parse_summary(summary, rn):
    """
    Extract (patient_name, operation) from the event summary.

    Common patterns observed in CTSPPUM calendar:
      1. "40001244 - LEE MING FATT CABG"
      2. "40002203  MOHAMAD FAIZ BIN OTHMAN  _ VATS Sympathetectomy"
      3. "40000011 CHEONG SOOK MEI  MV repair KIV replace + TAP"
      4. "MVR 40000696 - NURUL AIN BINTI IBRAHIM— SLE"
      5. "40001381 GOH SIEW LIN Right Upper Lobectomy"
      6. "KAMALA A/P RAMASAMY 40002340 _ CABG"
      7. "MOHAMED RAFIQ BIN JALIL_ 40001655 _ avr"

    Strategy:
      - Remove the RN.
      - If a hard separator ( _ or — ) exists, split on it and identify which
        part is the name (longest run of all-caps words).
      - Otherwise, scan words left-to-right: accumulate all-caps words as the
        name until the first clearly mixed-case word (= start of operation).
    """
    if not rn:
        return '', ''

    text = summary.replace(rn, '', 1)

    # Normalise separators: replace — and _ with a pipe so we can split later
    text = re.sub(r'\s*[—_]\s*', ' | ', text)
    # Clean up dashes that are separators (surrounded by spaces or at edges)
    text = re.sub(r'(?<=[A-Z\s])\s*-\s*(?=[A-Z])', ' | ', text)
    text = re.sub(r'\s+', ' ', text).strip(' |-')

    parts = [p.strip() for p in text.split('|') if p.strip()]

    if len(parts) >= 2:
        # Identify the name part: the segment whose words are mostly ALL CAPS
        def name_score(s):
            words = s.split()
            if not words:
                return 0
            caps = sum(1 for w in words if _is_name_word(w))
            return caps / len(words)

        # Pick the part with highest name-score that has at least 2 words
        name_part = max(
            (p for p in parts if len(p.split()) >= 2),
            key=name_score,
            default=parts[0],
        )
        op_parts = [p for p in parts if p != name_part]
        operation = ' '.join(op_parts).strip()
        return name_part.strip(), operation.strip()

    # No separator — walk word by word
    words = text.split()
    name_words = []
    op_words = []
    name_done = False

    for i, word in enumerate(words):
        if name_done:
            op_words.append(word)
            continue

        if _is_name_word(word):
            name_words.append(word)
        else:
            # First non-name word
            if len(name_words) >= 2:
                # We have enough name words already — everything else is op
                name_done = True
                op_words.append(word)
            else:
                # Haven't accumulated a name yet (e.g. "MVR" prefix before name)
                # Treat this as part of the operation prefix; keep scanning
                op_words.append(word)

    # Edge case: nothing classified as op yet but last word is a short all-caps abbreviation
    if name_words and not op_words:
        last = name_words[-1]
        alpha = re.sub(r'[^A-Za-z]', '', last)
        KNOWN_OPS = {'CABG', 'MVR', 'AVR', 'TVR', 'VATS', 'MIDCAB', 'TAVI',
                     'TAVR', 'DVR', 'TVT', 'MIS', 'AVR'}
        if alpha in KNOWN_OPS:
            op_words = [name_words.pop()]

    name = ' '.join(name_words).strip()
    operation = ' '.join(op_words).strip()

    # If op_words came before the name (pattern 4 / prefix pattern),
    # the "operation" string starts with op prefix + name — tease them apart.
    # Heuristic: if operation contains a long run of caps, that's the name.
    if not name:
        # Fallback: whole text is the name
        name = text
        operation = ''

    return name, operation


def get_date(event):
    """Return event start date as YYYY-MM-DD string."""
    start = event.get('start', {})
    raw = start.get('date') or start.get('dateTime', '')
    return raw[:10]


def fmt_date(date_str):
    """Convert YYYY-MM-DD to DD/MM/YYYY for display."""
    try:
        return datetime.strptime(date_str, '%Y-%m-%d').strftime('%d/%m/%Y')
    except ValueError:
        return date_str


# ─── Main processing ──────────────────────────────────────────────────────────

def process_events(events, old_dates=None):
    """Convert a list of calendar event dicts into patient records.

    old_dates: optional {RN: 'DD/MM/YYYY'} of previous operation dates.
    """
    old_dates = old_dates or {}
    records = []
    for ev in events:
        summary = clean_summary(ev.get('summary', ''))
        description = ev.get('description', '') or ''

        rn = extract_rn(summary)
        if not rn:
            continue  # Skip non-patient events (holidays, workshops, etc.)

        name, operation = parse_summary(summary, rn)
        date_iso = get_date(ev)
        contacts = extract_contacts(description)

        old = str(old_dates.get(rn, '') or '').strip()
        # An old date equal to the current date is not a change
        if old == fmt_date(date_iso):
            old = ''

        records.append({
            'Name': name,
            'RN': rn,
            'Operation': operation,
            'Date': date_iso,
            'Date Display': fmt_date(date_iso),
            'Contacts': contacts,
            'Old Date': old,
        })

    # Sort chronologically
    records.sort(key=lambda r: r['Date'])
    return records


# ─── Excel output ─────────────────────────────────────────────────────────────

ALT_ROW_BG  = 'EEF2F7'  # light blue-grey
OLD_DATE_BG = 'FFEDD5'  # amber — marks a changed operation date

def _border():
    thin = Side(style='thin', color='C0C8D4')
    return Border(left=thin, right=thin, top=thin, bottom=thin)


def create_excel(records, output_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Patient List'


    # ── Column headers (row 1) ──
    col_defs = [
        ('Name',              32),
        ('RN',                13),
        ('Operation Planned', 38),
        ('Date of Surgery',   16),
        ('Contact Numbers',   28),
        ('Old Date',          14),
    ]
    h_font  = Font(name='Arial', bold=True, size=10, color='FFFFFF')
    h_fill  = PatternFill('solid', start_color='4A6FA5')
    h_align = Alignment(horizontal='center', vertical='center', wrap_text=True)

    for col, (header, width) in enumerate(col_defs, 1):
        c = ws.cell(row=1, column=col, value=header)
        c.font    = h_font
        c.fill    = h_fill
        c.alignment = h_align
        c.border  = _border()
        ws.column_dimensions[get_column_letter(col)].width = width

    ws.row_dimensions[1].height = 22

    # ── Data rows ──
    d_font  = Font(name='Arial', size=10)
    d_align = Alignment(vertical='center', wrap_text=True)
    alt_fill = PatternFill('solid', start_color=ALT_ROW_BG)

    for i, rec in enumerate(records):
        row = i + 2
        row_data = [
            rec['Name'],
            rec['RN'],
            rec['Operation'],
            rec['Date Display'],
            rec['Contacts'],
            rec.get('Old Date', ''),
        ]
        fill = alt_fill if i % 2 == 0 else None
        for col, val in enumerate(row_data, 1):
            c = ws.cell(row=row, column=col, value=val)
            c.font      = d_font
            c.alignment = d_align
            c.border    = _border()
            if fill:
                c.fill = fill
            # Flag a changed date so it is obvious at a glance
            if col == 6 and val:
                c.fill = PatternFill('solid', start_color=OLD_DATE_BG)
                c.font = Font(name='Arial', size=10, bold=True, color='9A3412')
        ws.row_dimensions[row].height = 18

    # Freeze header rows
    ws.freeze_panes = 'A2'

    wb.save(output_path)
    return len(records)


# ─── Entry point ──────────────────────────────────────────────────────────────

if __name__ == '__main__':
    if len(sys.argv) < 3:
        print('Usage: python build_patient_list.py events.json output.xlsx '
              '[old_dates.json]', file=sys.stderr)
        sys.exit(1)

    events_path = sys.argv[1]
    output_path = sys.argv[2]
    old_path    = sys.argv[3] if len(sys.argv) > 3 else None

    with open(events_path) as f:
        data = json.load(f)

    old_dates = {}
    if old_path:
        with open(old_path) as f:
            old_dates = {str(k): v for k, v in json.load(f).items()}

    events = data if isinstance(data, list) else data.get('events', [])
    records = process_events(events, old_dates)

    n = create_excel(records, output_path)
    n_changed = sum(1 for r in records if r.get('Old Date'))
    print(json.dumps({'status': 'ok', 'patients': n,
                      'date_changed': n_changed, 'output': output_path}))
