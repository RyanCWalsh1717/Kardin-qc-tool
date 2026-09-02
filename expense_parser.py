"""
Kardin '4. Expense Back-up' bucket parser (v1).

Two inputs:
  - Expense Detail    (PDF: one page per GL code, vendor/allocation-level line
                        items each tagged by cost center - "West01 (West01)"
                        or "west20 (west20)" - with a monthly $ x12 + total)
  - Mgmt Fee Calc x2   (PDF: Kardin's User Defined Calculation for management
                        fee, one per cost center, income basis -> 3% factor ->
                        floor/minimum -> final fee)

Checks:
  1. Per-GL internal consistency: do a page's cost-center-tagged line totals
     sum to its own "Totals:" row?
  2. Management Fee tie-out: does Expense Detail's GL 637130 Admin-Management
     Fees page (split by cost center) match the two Mgmt Fee Calc PDFs'
     final (post-minimum-floor) fee amounts?
  3. (optional, needs bucket-1 data) west20 GL-level tie-out: does each GL
     page's west20-tagged line total match bucket 1's already-parsed
     "20 Riverside" Detail report for that GL code?
"""
import re

from kardin_parser import MONEY_RE, DECIMAL_RE, is_boilerplate, to_money

GL_HEADER_RE = re.compile(r'^(\d{6})\s+(\S.*)$')
COST_CENTER_TAG_RE = re.compile(r'^(\S+)\s*\(\1\)$', re.IGNORECASE)


def all_lines(pdf_file):
    if pdf_file is None:
        return []
    import pdfplumber
    lines = []
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ''
            lines.extend(text.split('\n'))
    return lines


def _trailing_numeric_14(toks):
    """Return (label_end_idx, months, total) if the line ends in 12 money + total + $/RSF."""
    if len(toks) < 14:
        return None
    tail = toks[-14:]
    if all(MONEY_RE.match(t) for t in tail[:13]) and DECIMAL_RE.match(tail[13]):
        return len(toks) - 14, [to_money(t) for t in tail[:12]], to_money(tail[12])
    return None


def _trailing_numeric_13(toks):
    """Return (label_end_idx, months, total) if the line ends in 12 money + total (no $/RSF) -
    the Mgmt Fee Calc report's format, unlike Expense Detail's."""
    if len(toks) < 13:
        return None
    tail = toks[-13:]
    if all(MONEY_RE.match(t) for t in tail):
        return len(toks) - 13, [to_money(t) for t in tail[:12]], to_money(tail[12])
    return None


def parse_expense_detail(pdf_file):
    """
    Returns (line_rows, totals_rows).
    line_rows: [{gl, gl_label, cost_center, description, total}]
    totals_rows: [{gl, gl_label, total}]  (the page's own "Totals:" row)
    """
    line_rows = []
    totals_rows = []
    current_gl = None
    current_gl_label = None
    current_cost_center = None
    for raw in all_lines(pdf_file):
        line = raw.rstrip()
        if is_boilerplate(line) or not line.strip():
            continue
        toks = line.split()
        if not toks:
            continue

        m_tag = COST_CENTER_TAG_RE.match(line.strip())
        if m_tag:
            current_cost_center = m_tag.group(1)
            continue

        numeric = _trailing_numeric_14(toks)

        if toks[0] == 'Totals:':
            if numeric:
                _, months, total = numeric
                totals_rows.append({'gl': current_gl, 'gl_label': current_gl_label, 'total': total})
            continue

        m_gl = GL_HEADER_RE.match(line.strip())
        if m_gl and numeric is None:
            current_gl = m_gl.group(1)
            current_gl_label = m_gl.group(2).strip()
            current_cost_center = None
            continue

        if numeric and current_gl:
            label_end_idx, months, total = numeric
            description = ' '.join(toks[:label_end_idx]).strip()
            line_rows.append({
                'gl': current_gl, 'gl_label': current_gl_label,
                'cost_center': current_cost_center, 'description': description,
                'total': total,
            })
    return line_rows, totals_rows


# ---------------------------------------------------------------- Mgmt Fee Calc

MGMT_FEE_ROW_RE = re.compile(r'^(Management Fee\S*)\s+(.*)$')


def parse_mgmt_fee_calc(pdf_file):
    """One entry per page: {report_name, cost_center, total} using the FINAL
    (post-minimum-floor) 'Management Fee' row - the one following
    'Minimum Amount :' and preceding 'Allocation Name:'."""
    results = []
    report_name = None
    cost_center = None
    seen_minimum = False
    for raw in all_lines(pdf_file):
        line = raw.rstrip()
        stripped = line.strip()
        if stripped.startswith('User Defined Calculations -'):
            report_name = stripped.replace('User Defined Calculations -', '').strip()
            seen_minimum = False
            continue
        if stripped.startswith('Minimum Amount'):
            seen_minimum = True
            continue
        if stripped.startswith('Allocation Name:'):
            cost_center = stripped.replace('Allocation Name:', '').strip()
            continue
        m = MGMT_FEE_ROW_RE.match(stripped)
        if m and seen_minimum:
            toks = stripped.split()
            numeric = _trailing_numeric_13(toks)
            if numeric:
                _, months, total = numeric
                results.append({'report_name': report_name, 'cost_center': None, 'total': total})
            seen_minimum = False
    # cost_center lines come AFTER the row they describe, so backfill in order
    cc_names = []
    for raw in all_lines(pdf_file):
        s = raw.strip()
        if s.startswith('Allocation Name:'):
            cc_names.append(s.replace('Allocation Name:', '').strip())
    for r, cc in zip(results, cc_names):
        r['cost_center'] = cc
    return results


# --------------------------------------------------------------------- Checks

def check_gl_totals_internal_consistency(line_rows, totals_rows, tolerance=1):
    findings = []
    sums = {}
    for r in line_rows:
        sums.setdefault(r['gl'], []).append(r)
    for t in totals_rows:
        rows = sums.get(t['gl'], [])
        computed = sum(r['total'] for r in rows)
        if abs(computed - t['total']) > tolerance:
            findings.append({
                'Report Section': '5. Expense Detail',
                'GL Acct': t['gl'], 'Line Item': t['gl_label'], 'Budget Year': 'Next Year Budget',
                'Priority': 'Must Fix',
                'Comment': (
                    f"Expense Detail's line items for GL {t['gl']} sum to ${computed:,} but the page's "
                    f"own 'Totals:' row shows ${t['total']:,} (diff ${t['total']-computed:+,}). "
                    "Likely a stale export or a parsing edge case worth spot-checking."
                ),
                'Status': 'Open', 'Source Check': 'GL page internal sum mismatch',
            })
    return findings


def check_mgmt_fee_files_are_duplicates(mgmt_fee_20r, mgmt_fee_1r, mgmt_fee_20r_name, mgmt_fee_1r_name):
    """
    Both Mgmt Fee Calc PDFs are supposed to be named per-building, but Kardin's
    export apparently always dumps the FULL portfolio calc (all cost centers)
    regardless of which one is requested - so the two files end up byte-for-
    byte equivalent in content. Not necessarily an error, but worth surfacing
    so it isn't mistaken for two independent confirmations of the same number.
    """
    key = lambda rows: sorted((r['cost_center'], r['total']) for r in rows)
    if key(mgmt_fee_20r) != key(mgmt_fee_1r) or not mgmt_fee_20r:
        return []
    return [{
        'Report Section': '5. Expense Detail',
        'GL Acct': '', 'Line Item': 'Mgmt Fee Calc', 'Budget Year': 'Next Year Budget',
        'Priority': 'For Discussion',
        'Comment': (
            f"'{mgmt_fee_20r_name}' and '{mgmt_fee_1r_name}' contain identical data (both cost centers "
            "in both files), despite being named per-building. Likely just how Kardin exports this report "
            "(not a data error), but treat them as one source, not two independent confirmations."
        ),
        'Status': 'Open', 'Source Check': 'Duplicate Mgmt Fee Calc files',
    }]


def check_mgmt_fee_tie_out(line_rows, mgmt_fee_rows, tolerance=1):
    """mgmt_fee_rows: rows from ONE Mgmt Fee Calc file - since both files
    contain the same full set of cost centers (see check_mgmt_fee_files_are_
    duplicates), summing both would double-count."""
    findings = []
    mgmt_637130 = [r for r in line_rows if r['gl'] == '637130']
    by_cc = {}
    for r in mgmt_637130:
        cc = (r['cost_center'] or '').lower()
        by_cc.setdefault(cc, 0)
        by_cc[cc] += r['total']

    expected = {}
    for r in mgmt_fee_rows:
        cc = (r['cost_center'] or '').lower()
        expected[cc] = expected.get(cc, 0) + r['total']

    for cc, exp_total in expected.items():
        actual = by_cc.get(cc)
        if actual is None:
            continue
        if abs(actual - exp_total) > tolerance:
            findings.append({
                'Report Section': '5. Expense Detail',
                'GL Acct': '637130', 'Line Item': 'Admin-Management Fees', 'Budget Year': 'Next Year Budget',
                'Priority': 'Must Fix',
                'Comment': (
                    f"Expense Detail's GL 637130 lines for cost center '{cc}' total ${actual:,}, but the "
                    f"Mgmt Fee Calc backup shows the final management fee for '{cc}' should be ${exp_total:,} "
                    f"(diff ${actual-exp_total:+,}). Confirm the Kardin GL entry was updated after the last "
                    "management fee recalculation."
                ),
                'Status': 'Open', 'Source Check': 'Management fee tie-out',
            })
    return findings


def check_west20_ties_to_bucket1_detail(line_rows, bucket1_detail_rows, tolerance=1):
    """Optional cross-bucket check: requires bucket 1's already-parsed
    '20 Riverside' Detail rows (kardin_parser.parse_analysis_report output)."""
    findings = []
    bucket1_by_gl = {r['gl']: r for r in bucket1_detail_rows if r.get('gl')}
    west20_sums = {}
    for r in line_rows:
        cc = (r['cost_center'] or '').lower()
        if cc != 'west20':
            continue
        west20_sums.setdefault(r['gl'], 0)
        west20_sums[r['gl']] += r['total']

    for gl, total in west20_sums.items():
        b1 = bucket1_by_gl.get(gl)
        if b1 is None:
            continue
        expected = b1['v3_2027B']
        if abs(total - expected) > tolerance:
            findings.append({
                'Report Section': '5. Expense Detail',
                'GL Acct': gl, 'Line Item': b1['label'], 'Budget Year': 'Next Year Budget',
                'Priority': 'Must Fix',
                'Comment': (
                    f"Expense Detail's west20-tagged lines for GL {gl} total ${total:,}, but bucket 1's "
                    f"Budget Analysis Detail shows 2027 Budget = ${expected:,} for this GL (diff "
                    f"${total-expected:+,}). Reports may be out of sync."
                ),
                'Status': 'Open', 'Source Check': 'Expense Detail vs Budget Analysis Detail mismatch',
            })
    return findings


def run(expense_detail_pdf, mgmt_fee_20r_pdf, mgmt_fee_1r_pdf=None, bucket1_west20_detail_rows=None,
        mgmt_fee_20r_name='Mgmt Fee Calc #1', mgmt_fee_1r_name='Mgmt Fee Calc #2'):
    """mgmt_fee_1r_pdf is optional - some PMs only send one Mgmt Fee Calc export
    since Kardin dumps the same full-portfolio calc into every one regardless of
    which building was requested (see check_mgmt_fee_files_are_duplicates), so a
    second copy is redundant. The duplicate check and its stats are skipped when
    it's not provided; the tie-out still runs against the one file we have."""
    line_rows, totals_rows = parse_expense_detail(expense_detail_pdf)
    mgmt_fee_20r = parse_mgmt_fee_calc(mgmt_fee_20r_pdf)
    mgmt_fee_1r = parse_mgmt_fee_calc(mgmt_fee_1r_pdf) if mgmt_fee_1r_pdf is not None else []

    from kardin_parser import missing_file_finding
    findings = []
    if expense_detail_pdf is None:
        findings.append(missing_file_finding('Expense Detail'))
    if mgmt_fee_20r_pdf is None:
        findings.append(missing_file_finding('Mgmt Fee Calc - File 1'))
    findings += check_gl_totals_internal_consistency(line_rows, totals_rows)
    if mgmt_fee_1r_pdf is not None:
        findings += check_mgmt_fee_files_are_duplicates(
            mgmt_fee_20r, mgmt_fee_1r, mgmt_fee_20r_name, mgmt_fee_1r_name)
    findings += check_mgmt_fee_tie_out(line_rows, mgmt_fee_20r)
    if bucket1_west20_detail_rows is not None:
        findings += check_west20_ties_to_bucket1_detail(line_rows, bucket1_west20_detail_rows)

    stats = {
        'gl_pages': len(totals_rows),
        'line_rows': len(line_rows),
        'mgmt_fee_20r_rows': len(mgmt_fee_20r),
        'mgmt_fee_1r_rows': len(mgmt_fee_1r),
    }
    return {
        'line_rows': line_rows, 'totals_rows': totals_rows,
        'mgmt_fee_20r': mgmt_fee_20r, 'mgmt_fee_1r': mgmt_fee_1r,
        'findings': findings, 'stats': stats,
    }
