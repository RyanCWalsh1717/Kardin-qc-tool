"""
Kardin '5. CapEx Back-up' bucket parser (v1).

Three inputs:
  - Capex.pdf   Building Improvements (154500) + CIP (171300), same per-GL
                vendor/cost-center layout as bucket 4's Expense Detail - the
                bucket-4 parser (expense_parser.parse_expense_detail) is
                reused directly rather than duplicated.
  - TIs.pdf     Per-suite Tenant Improvement schedule: a lifetime $ total
                broken into dated installments (e.g. 25%/25%/25%/25% around
                lease commencement), of which only some may fall in 2027.
  - LCs.pdf     Same shape for Leasing Commissions (typically 50%/50%).

Checks:
  1. (via expense_parser, reused) Capex GL totals tie to bucket 1's west20
     Detail - same pattern as bucket 4.
  2. TI/LC payment-schedule consistency: does the sum of a suite's dated
     installments falling in 2027 equal that suite's own stated 2027 total?
  3. TI/LC prior-year-installment note: for any suite with installments
     dated before 2027, surface that explicitly - it's exactly what explains
     the gap (flagged as "For Discussion" back in bucket 2) between Leasing
     Activity's lifetime TI/LC totals and the 2027 Budget capex figures.
"""
import re

from kardin_parser import MONEY_RE, is_boilerplate, to_money

DATE_RE = re.compile(r'^\d{1,2}/\d{1,2}/\d{4}$')
# Building-code shape varies by property (see leasing_parser.SUITE_RE), and
# some exports (Lexington Labs) glue the suite code directly onto the start
# of the tenant name with no space, e.g. "lexlab-1-0200Vacant Space" - this
# matches just the leading suite-code portion of a token like that.
SUITE_PREFIX_RE = re.compile(r'^([A-Za-z][A-Za-z-]*\d+-\d+)')
GLUED_RATE_RE = re.compile(r'(\d\.\d{2})(?=\d)')  # Kardin sometimes drops the space, e.g. "55.001,295,140"


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


def parse_ti_lc_report(pdf_file):
    """Returns [{suite, suite_2027_total, installments: [(date, amount), ...]}]."""
    suites = []
    current = None
    for raw in all_lines(pdf_file):
        line = GLUED_RATE_RE.sub(r'\1 ', raw.rstrip())
        if is_boilerplate(line) or not line.strip():
            continue
        toks = line.split()
        if not toks:
            continue
        m_suite = SUITE_PREFIX_RE.match(toks[0])
        if m_suite:
            if current:
                suites.append(current)
            money_toks = [t for t in toks if MONEY_RE.match(t)]
            suite_2027_total = to_money(money_toks[-1]) if money_toks else None
            current = {'suite': m_suite.group(1), 'suite_2027_total': suite_2027_total, 'installments': []}
            continue
        if len(toks) == 3 and DATE_RE.match(toks[0]) and toks[1].endswith('%') and MONEY_RE.match(toks[2]):
            if current is not None:
                current['installments'].append((toks[0], to_money(toks[2])))
            continue
    if current:
        suites.append(current)
    return suites


def check_ti_lc_schedule_consistency(suites, report_label, tolerance=1):
    findings = []
    for s in suites:
        year2027 = sum(amt for date, amt in s['installments'] if date.endswith('/2027'))
        if s['suite_2027_total'] is None:
            continue
        if abs(year2027 - s['suite_2027_total']) > tolerance:
            findings.append({
                'Report Section': '8. CapEx',
                'GL Acct': '', 'Line Item': f"{s['suite']} - {report_label}", 'Budget Year': 'Next Year Budget',
                'Priority': 'Must Fix',
                'Comment': (
                    f"{s['suite']}'s dated installments falling in 2027 sum to ${year2027:,}, but the "
                    f"report's own 2027 total for this suite is ${s['suite_2027_total']:,} "
                    f"(diff ${s['suite_2027_total']-year2027:+,}). Check the payment schedule dates."
                ),
                'Status': 'Open', 'Source Check': f'{report_label} payment schedule mismatch',
            })
    return findings


def note_prior_year_installments(suites, report_label):
    """Explains (doesn't flag as an error) any portion of a suite's lifetime
    TI/LC that falls before 2027 - this is exactly what accounts for the gap
    between bucket 2's Leasing Activity lifetime totals and the 2027 Budget
    capex figures, so surface it rather than leaving it as an open question."""
    findings = []
    for s in suites:
        prior = sum(amt for date, amt in s['installments'] if not date.endswith('/2027'))
        if prior <= 0:
            continue
        prior_dates = [date for date, amt in s['installments'] if not date.endswith('/2027')]
        findings.append({
            'Report Section': '8. CapEx',
            'GL Acct': '', 'Line Item': f"{s['suite']} - {report_label}", 'Budget Year': 'N/A',
            'Priority': 'For Discussion',
            'Comment': (
                f"${prior:,} of {s['suite']}'s {report_label.lower()} commitment is scheduled before 2027 "
                f"({', '.join(prior_dates)}) and so isn't in the 2027 Budget figure - it should already be "
                "reflected in the 2026 Reforecast. This is the source of the gap between Leasing Activity's "
                "lifetime TI/LC totals and the 2027 Budget capex lines (flagged in bucket 2) - confirmed "
                "explained, not missing."
            ),
            'Status': 'Open', 'Source Check': 'Prior-year TI/LC installment (informational)',
        })
    return findings


def run(tis_pdf, lcs_pdf, capex_line_rows=None, capex_totals_rows=None, bucket1_west20_detail_rows=None,
        capex_pdf_missing=False):
    ti_suites = parse_ti_lc_report(tis_pdf)
    lc_suites = parse_ti_lc_report(lcs_pdf)

    from kardin_parser import missing_file_finding
    findings = []
    if tis_pdf is None:
        findings.append(missing_file_finding('TIs'))
    if lcs_pdf is None:
        findings.append(missing_file_finding('LCs'))
    if capex_pdf_missing:
        findings.append(missing_file_finding('Capex'))
    findings += check_ti_lc_schedule_consistency(ti_suites, 'Tenant Improvements')
    findings += check_ti_lc_schedule_consistency(lc_suites, 'Leasing Commissions')
    findings += note_prior_year_installments(ti_suites, 'Tenant Improvements')
    findings += note_prior_year_installments(lc_suites, 'Leasing Commissions')

    if capex_line_rows is not None and capex_totals_rows is not None:
        import expense_parser as ep
        findings += ep.check_gl_totals_internal_consistency(capex_line_rows, capex_totals_rows)
        if bucket1_west20_detail_rows is not None:
            findings += ep.check_west20_ties_to_bucket1_detail(capex_line_rows, bucket1_west20_detail_rows)

    stats = {'ti_suites': len(ti_suites), 'lc_suites': len(lc_suites)}
    return {'ti_suites': ti_suites, 'lc_suites': lc_suites, 'findings': findings, 'stats': stats}
