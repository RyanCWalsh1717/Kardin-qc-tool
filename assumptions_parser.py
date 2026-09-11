"""
Parsers for GRP's own planning/assumption workbooks - independent inputs GRP
provides to build the budget, not read from a PM's Kardin export. Each check
here answers: does Kardin's actual exported budget reflect what GRP told the
PM to assume, not just whether Kardin's own reports are internally
consistent (which is what kardin_parser.py and the other bucket parsers
check).

These workbooks are portfolio-wide (one sheet per property) and NOT
Kardin's own format, so nothing here reuses kardin_parser's PDF-text
helpers - it's straight openpyxl cell reading.
"""
import datetime


def parse_grp_budget_assumptions(xlsx_file, sheet_name):
    """
    "2027 GRP Budget Assumptions.xlsx" (or equivalent for another budget
    cycle) - one sheet per property, GL-code-keyed rows (column B = GL code,
    column C = label) with a monthly $ column per dated header. Some GL
    codes are section headers whose actual numbers live on indented
    sub-item rows below (blank column B) - e.g. "637370 Software:" rolls up
    " - Yardi / Nexus", " - MRI", " - Kardin", " - VTS" underneath it; those
    are summed together under that one GL. A GL row with a text note
    instead of a number (e.g. "Use FY 2026 actuals increased by 3%") has no
    assumption $ to compare for that month - has_note=True flags this so
    callers skip it rather than silently comparing against $0.

    Returns [{'gl': str, 'label': str, 'has_note': bool,
               'monthly': {date: float or None}}] - one entry per GL section
    (sub-items already summed in), in sheet order. 'gl' may be a non-numeric
    placeholder like 'n/a' or 'MRI only' - those simply won't match any real
    Kardin GL code, which is correct (nothing to tie out).
    """
    import openpyxl
    wb = openpyxl.load_workbook(xlsx_file, data_only=True)
    ws = wb[sheet_name]

    header_row = None
    month_cols = []
    for r in range(1, min(ws.max_row, 10) + 1):
        date_cols = [c for c in range(1, ws.max_column + 1)
                     if isinstance(ws.cell(row=r, column=c).value, datetime.datetime)]
        if len(date_cols) >= 2:
            header_row, month_cols = r, date_cols
            break
    if header_row is None:
        return []
    month_dates = [ws.cell(row=header_row, column=c).value.date() for c in month_cols]

    sections = []
    current = None
    for r in range(header_row + 1, ws.max_row + 1):
        gl_val = ws.cell(row=r, column=2).value
        label_val = ws.cell(row=r, column=3).value
        if gl_val is not None:
            current = {'gl': str(gl_val).strip(), 'label': (label_val or '').strip(),
                       'has_note': False, 'monthly_lists': {d: [] for d in month_dates}}
            sections.append(current)
        if current is None:
            continue
        for c, d in zip(month_cols, month_dates):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, (int, float)):
                current['monthly_lists'][d].append(v)
            elif isinstance(v, str) and v.strip():
                current['has_note'] = True

    return [{
        'gl': sec['gl'], 'label': sec['label'], 'has_note': sec['has_note'],
        'monthly': {d: (sum(vals) if vals else None) for d, vals in sec['monthly_lists'].items()},
    } for sec in sections]


def check_assumptions_vs_bucket1(assumption_rows, bucket1_monthly_rows, source_label, budget_year, tolerance=1):
    """
    Compares GRP's own monthly $ assumption per GL against Kardin's actual
    Monthly Budget Detail (bucket 1) for the same GL, month by month. Only
    compares months present in BOTH sources (assumption workbooks often
    cover more months - e.g. a reforecast tail - than a single Kardin
    export), and skips any GL where the assumption is a text note rather
    than a number (has_note) rather than treating that as a $0 assumption.

    budget_year: REQUIRED - the assumption workbook typically spans both a
    reforecast tail (e.g. Jun-Dec 2026) and the next full budget year (e.g.
    Jan-Dec 2027) in one sheet. bucket1_monthly_rows' 'months' is a plain
    12-value Jan-Dec list for ONE specific year (whatever Kardin's Monthly
    Detail PDF covers) - aligning purely by month NAME without pinning the
    year would silently compare e.g. June 2026's assumption against
    Kardin's June 2027 figure. Only assumption columns matching this year
    are used.
    """
    by_gl = {r['gl']: r for r in bucket1_monthly_rows if r.get('gl') and not r.get('is_total')}
    findings = []
    for a in assumption_rows:
        if a['gl'] not in by_gl:
            continue
        b1 = by_gl[a['gl']]
        mismatches = []
        for d, assumed in a['monthly'].items():
            if assumed is None or d.year != budget_year:
                continue
            month_idx = d.month - 1
            if month_idx >= len(b1['months']):
                continue
            actual = b1['months'][month_idx]
            if abs(actual - assumed) > tolerance:
                mismatches.append(f"{d.strftime('%b-%y')}: assumed ${assumed:,.0f} vs Kardin ${actual:,.0f}")
        if mismatches:
            findings.append({
                'Report Section': '1. General', 'GL Acct': a['gl'], 'Line Item': a['label'],
                'Budget Year': 'Next Year Budget', 'Priority': 'For Discussion',
                'Comment': (
                    f"GRP's own budget assumption for GL {a['gl']} ({a['label']}) doesn't match what "
                    f"Kardin's Monthly Budget Detail actually shows, in {len(mismatches)} month(s): "
                    + '; '.join(mismatches[:6]) + ('...' if len(mismatches) > 6 else '') +
                    f". Per {source_label}. Confirm the PM's Kardin entry reflects the latest assumption, "
                    "or that the assumption itself is current - either could be stale."
                ),
                'Status': 'Open', 'Source Check': 'GRP assumption vs Kardin Monthly Detail',
            })
    return findings
