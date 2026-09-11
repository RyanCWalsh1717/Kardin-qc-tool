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
import re


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


BLOCK_TITLE_RE = re.compile(r'^Leasing Assumptions\s*-\s*.*?(\d{4})\s*Budget\s*$', re.IGNORECASE)


def parse_leasing_assumptions(xlsx_file, sheet_name, budget_year):
    """
    "2027 Leasing Assumptions_Final.xlsx" (or equivalent) - one sheet per
    property, with REPEATED stacked blocks: one per budget cycle this same
    file has been reused for (e.g. "Leasing Assumptions - 2024 Budget", then
    "... - 2024 Reforecast/2025 Budget", etc., appended below each other
    year over year - each is that cycle's own snapshot, not a correction of
    the one before). Only the block whose title ends in "{budget_year}
    Budget" is parsed - earlier cycles are history, not this run's
    assumptions.

    Returns [{'building': str, 'suite_num': str, 'rsf': float or None,
              'new_lease_cd': date or str or None, 'term': str or None,
              'free_rent': str or None, 'starting_base_rent': float or None}]
    - one row per vacant/expiring suite assumption in that block. Header
    columns are read dynamically (case-insensitive) since column order/count
    drifts slightly between cycles (a 'New Tenant' column was added later).
    """
    import openpyxl
    wb = openpyxl.load_workbook(xlsx_file, data_only=True)
    ws = wb[sheet_name]

    block_start = None
    for r in range(1, ws.max_row + 1):
        v = ws.cell(row=r, column=1).value
        if isinstance(v, str) and v.strip().lower().startswith('leasing assumptions'):
            m = BLOCK_TITLE_RE.match(v.strip())
            if m and int(m.group(1)) == budget_year:
                block_start = r
                break
    if block_start is None:
        return []

    header_row = block_start + 2  # title, then "The following suites...", then the header
    headers = [str(ws.cell(row=header_row, column=c).value or '').strip().lower()
               for c in range(1, ws.max_column + 1)]

    def col_idx(*names):
        for name in names:
            for i, h in enumerate(headers):
                if h == name:
                    return i + 1
        return None

    c = {
        'building': col_idx('building'), 'suite': col_idx('suite #', 'suite#'),
        'rsf': col_idx('rsf'), 'cd': col_idx('new lease cd'), 'term': col_idx('term'),
        'free': col_idx('free rent'), 'rent': col_idx('starting base rent'),
    }

    rows = []
    for r in range(header_row + 1, ws.max_row + 1):
        building = ws.cell(row=r, column=c['building']).value if c['building'] else None
        if building is None:
            if all(ws.cell(row=r, column=col).value is None for col in range(1, ws.max_column + 1)):
                break
            continue
        rows.append({
            'building': str(building).strip(),
            'suite_num': str(ws.cell(row=r, column=c['suite']).value or '').strip() if c['suite'] else '',
            'rsf': ws.cell(row=r, column=c['rsf']).value if c['rsf'] else None,
            'new_lease_cd': ws.cell(row=r, column=c['cd']).value if c['cd'] else None,
            'term': ws.cell(row=r, column=c['term']).value if c['term'] else None,
            'free_rent': ws.cell(row=r, column=c['free']).value if c['free'] else None,
            'starting_base_rent': ws.cell(row=r, column=c['rent']).value if c['rent'] else None,
        })
    return rows


def _match_kardin_suite(suite_num, rsf, cost_center_code, kardin_suites):
    """Best-effort match of one assumptions row to a Kardin suite code.
    Confirmed against real Riverside Labs data: Kardin's suite code is
    usually just the assumption's suite number zero-padded by one leading
    digit under the cost center (e.g. '300' -> 'west20-0300'). BUT a suite
    can get subdivided in Kardin after the assumption was made - several
    assumption rows can then share the same suite # (e.g. three separate
    '300' rows), each really corresponding to a DIFFERENT real Kardin suite
    (0300/0301/0302), distinguishable only by RSF. Trying the padded-suite-
    number match first would wrongly match all three to the same one, so
    RSF (combined with the padded number when both agree, otherwise alone)
    is checked before falling back to suite-number-only. Returns a suite
    dict or None if no confident match is found (ambiguous or absent)."""
    candidates = [s for s in kardin_suites if s['suite'].lower().startswith(f'{cost_center_code.lower()}-')]
    padded = f'{cost_center_code.lower()}-0{suite_num}' if suite_num.isdigit() else None

    if padded and rsf:
        combined = [s for s in candidates
                    if s['suite'].lower() == padded and s.get('rsf') and abs(s['rsf'] - rsf) <= 5]
        if len(combined) == 1:
            return combined[0]
    if rsf:
        rsf_matches = [s for s in candidates if s.get('rsf') and abs(s['rsf'] - rsf) <= 5]
        if len(rsf_matches) == 1:
            return rsf_matches[0]
    if padded:
        exact = [s for s in candidates if s['suite'].lower() == padded]
        if len(exact) == 1:
            return exact[0]
    return None


def check_leasing_assumptions_vs_bucket2(assumption_rows, occupancy_rows, building_to_cost_center,
                                          budget_year, source_label):
    """
    For each vacant/expiring suite GRP assumed a new lease for, confirms:
      1. The suite can be matched to a real Kardin suite at all.
      2. RSF ties out.
      3. If the assumed commencement date falls within this budget year (or
         the tail end of the prior year, still active during it), Kardin's
         Occupancy Summary should show a real lease-up (status Contract/New
         with an actual commence date) for that suite - not 'Unknown' (no
         leasing assumption modeled). A commencement date in a LATER year is
         correctly not yet modeled - noted, not flagged.

    Deliberately does NOT compare Starting Base Rent or Free Rent months
    against Kardin's own monthly schedule - deriving a single "rate" from a
    ramping/escalating monthly schedule isn't reliable enough yet to compare
    confidently; scoped out rather than guessed at.

    building_to_cost_center: {building name as it appears in the assumptions
    sheet (e.g. '20 Riverside'): Kardin cost center code (e.g. 'west20')} -
    from a property's config.yaml.
    """
    findings = []
    unmatched = []
    for a in assumption_rows:
        cost_center = building_to_cost_center.get(a['building'])
        if not cost_center:
            unmatched.append(f"{a['building']} suite {a['suite_num']} (unknown building - no cost center mapped)")
            continue
        match = _match_kardin_suite(a['suite_num'], a['rsf'], cost_center, occupancy_rows)
        if match is None:
            unmatched.append(f"{a['building']} suite {a['suite_num']} ({a['rsf']} RSF)")
            continue

        if a['rsf'] and match.get('rsf') and abs(match['rsf'] - a['rsf']) > 5:
            findings.append({
                'Report Section': '9. Leasing & Rent', 'GL Acct': '', 'Line Item': match['suite'],
                'Budget Year': 'Next Year Budget', 'Priority': 'For Discussion',
                'Comment': (
                    f"Leasing Assumptions lists {match['suite']} at {a['rsf']:,} RSF, but Kardin's "
                    f"Occupancy Summary shows {match['rsf']:,} RSF. Per {source_label}."
                ),
                'Status': 'Open', 'Source Check': 'Leasing assumption RSF mismatch',
            })

        cd = a['new_lease_cd']
        cd_date = cd if isinstance(cd, (datetime.date, datetime.datetime)) else None
        if cd_date and cd_date.year <= budget_year:
            if match['status'] == 'Unknown':
                findings.append({
                    'Report Section': '9. Leasing & Rent', 'GL Acct': '', 'Line Item': match['suite'],
                    'Budget Year': 'Next Year Budget', 'Priority': 'Must Fix',
                    'Comment': (
                        f"Leasing Assumptions has {match['suite']} leasing up starting "
                        f"{cd_date.strftime('%b %Y') if hasattr(cd_date, 'strftime') else cd_date} "
                        f"({a['term']}, {a['free_rent']} free), which is within or before {budget_year} - "
                        f"but Kardin's Occupancy Summary shows no leasing assumption at all for this suite "
                        f"('Unknown' status). Per {source_label}. Confirm Kardin was updated to reflect this."
                    ),
                    'Status': 'Open', 'Source Check': 'Leasing assumption not modeled in Kardin',
                })
    if unmatched:
        findings.append({
            'Report Section': '9. Leasing & Rent', 'GL Acct': '', 'Line Item': 'GENERAL',
            'Budget Year': 'N/A', 'Priority': 'For Discussion',
            'Comment': (
                f"{len(unmatched)} suite(s) from Leasing Assumptions couldn't be confidently matched to a "
                f"Kardin suite (by suite number or RSF): {'; '.join(unmatched)}. Per {source_label}. "
                "May just need a manual look, not necessarily an error."
            ),
            'Status': 'Open', 'Source Check': 'Leasing assumption suite not matched',
        })
    return findings


# Kardin's Capex.pdf also includes Leasing Commissions (181200) and Tenant
# Improvements (181400) - those are leasing-driven capital costs, already
# tracked separately in bucket 5 via its own TIs.pdf/LCs.pdf, and are NOT
# part of the 5-Year CapEx Plan's scope (building capital projects - HVAC,
# roof, etc.). Confirmed against real Riverside Labs data: restricting to
# just these two GLs made both buildings' totals tie out exactly to the
# 5-Year Plan; including 181200/181400 made west20 off by ~$1.45M.
CAPEX_BUILDING_IMPROVEMENT_GLS = {'154500', '171300'}


def parse_capex_plan_portfolio_summary(xlsx_file, target_year):
    """
    GRP's "5-Year CapEx Plan" workbook - reads the "Portfolio Summary"
    sheet's own pre-aggregated "PORTFOLIO TOTAL BY YEAR" section (already
    summed across every project category - HVAC, Roof, Chiller, etc. - so
    this doesn't need to re-derive it from the many per-category detail
    blocks above it). Building names come from that section's own column
    header row (e.g. '20 Riverside', '1 Riverside', 'Building 3'...).

    Returns {building_name: annual_total} for the target year, or {} if
    that year isn't found in the sheet (a 5-year plan only covers a fixed
    window - e.g. 2026-2030 plus a "2030+" catch-all - so a year outside
    that range legitimately has nothing to compare).
    """
    import openpyxl
    wb = openpyxl.load_workbook(xlsx_file, data_only=True)
    ws = wb['Portfolio Summary']

    header_row = None
    for r in range(1, ws.max_row + 1):
        v = ws.cell(row=r, column=1).value
        if isinstance(v, str) and v.strip().upper().startswith('PORTFOLIO TOTAL BY YEAR'):
            header_row = r
            break
    if header_row is None:
        return {}

    # The building-name column layout is set by the sheet's main column
    # header row (shared with the per-category blocks above) - find it by
    # locating the 'Year' column header rather than assuming a fixed row.
    col_header_row = None
    for r in range(1, header_row):
        vals = [str(ws.cell(row=r, column=c).value or '').strip().lower() for c in range(1, ws.max_column + 1)]
        if 'year' in vals:
            col_header_row = r
            break
    if col_header_row is None:
        return {}
    col_headers = [str(ws.cell(row=col_header_row, column=c).value or '').strip()
                   for c in range(1, ws.max_column + 1)]
    building_cols = {h: i + 1 for i, h in enumerate(col_headers)
                     if h and h.lower() not in ('category', 'year', 'portfolio total', 'notes')}

    for r in range(header_row + 1, ws.max_row + 1):
        yr_val = ws.cell(row=r, column=1).value
        if yr_val == target_year:
            return {b: (ws.cell(row=r, column=c).value or 0) for b, c in building_cols.items()}
    return {}


def check_capex_plan_vs_bucket5(plan_totals, capex_line_rows, building_to_cost_center, target_year,
                                 source_label, tolerance=1):
    """
    Compares GRP's 5-Year CapEx Plan's per-building annual total (for
    target_year) against Kardin's actual Capex.pdf (bucket 5) - restricted
    to CAPEX_BUILDING_IMPROVEMENT_GLS only, since Capex.pdf also carries
    Leasing Commissions/Tenant Improvements which the 5-Year Plan doesn't
    track (those are compared separately via bucket 5's own TIs.pdf/LCs.pdf
    checks already).

    plan_totals: parse_capex_plan_portfolio_summary() output for target_year.
    capex_line_rows: expense_parser.parse_expense_detail() output for
    Capex.pdf (bucket 5 already parses this for its own GL tie-out check).
    building_to_cost_center: {building name as it appears in the CapEx Plan
    (e.g. '20 Riverside'): Kardin cost center code (e.g. 'west20')}.
    """
    findings = []
    for building, plan_total in plan_totals.items():
        cost_center = building_to_cost_center.get(building)
        if not cost_center:
            continue
        kardin_total = sum(
            r['total'] for r in capex_line_rows
            if (r.get('cost_center') or '').lower() == cost_center.lower() and r.get('gl') in CAPEX_BUILDING_IMPROVEMENT_GLS
        )
        if abs(kardin_total - plan_total) > tolerance:
            findings.append({
                'Report Section': '8. CapEx', 'GL Acct': '', 'Line Item': building,
                'Budget Year': 'Next Year Budget', 'Priority': 'For Discussion',
                'Comment': (
                    f"GRP's 5-Year CapEx Plan budgets ${plan_total:,.0f} in building-improvement capital "
                    f"({target_year}) for {building}, but Kardin's Capex.pdf (GL 154500 + 171300 only - "
                    f"excludes Leasing Commissions/Tenant Improvements, tracked separately) totals "
                    f"${kardin_total:,.0f} (diff ${kardin_total-plan_total:+,.0f}). Per {source_label}."
                ),
                'Status': 'Open', 'Source Check': 'CapEx Plan vs Kardin Capex.pdf',
            })
    return findings
