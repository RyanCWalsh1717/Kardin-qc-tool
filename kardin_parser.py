"""
Kardin 'Bgt and Fcst Summaries' bucket parser (v1).

Parses the three Kardin PDF exports for one cost center / building:
  - Budget Analysis Summary   (category-level: 2026B / 2026F / 2027B + variance + explanation)
  - Budget Analysis Detail    (GL-line-level: same columns)
  - Monthly Budget Detail     (GL-line-level: 12 months + Total + $/RSF)

Extracts structured rows and runs the QC checks derived from the GRP Budget SOP:
  1. Missing variance explanation where BOTH $2,500+ AND 5%+ thresholds are met
     (checked independently for Reforecast-vs-Budget and NextBudget-vs-Reforecast)
  2. Tenant Electric Reimbursement (613115) vs Recovery - Electricity (440500)
     monthly tie-out (SOP: must match dollar-for-dollar by month)
  3. Detail annual total vs Monthly Detail Total, per GL code
  4. Revision-number mismatch across the three files (possible stale export)

Every function accepts either a file path (str) or a file-like object
(e.g. a Streamlit UploadedFile), since pdfplumber handles both.
"""
import re

import pdfplumber

MONEY_RE = re.compile(r'^-?[\d,]+$')
PCT_RE = re.compile(r'^-?\d+\.\d+%$')
NA_RE = re.compile(r'^N/A$')
GLCODE_RE = re.compile(r'^\d{6}$')
DECIMAL_RE = re.compile(r'^-?\d+\.\d+$')

BOILERPLATE_PREFIXES = (
    'Prepared For:', 'Prepared By:', 'Property ID:', 'Property RSF:',
    'Cost Center(s) RSF:', 'Selected Cost Centers', 'Software:', 'File:',
    'Revision:', 'Date (EDT):', 'Page:', 'Explanation of Variance',
    'Budget Analysis', 'Monthly Budget Detail', 'vs.',
)

MATERIALITY_DOLLAR = 2500
MATERIALITY_PCT = 0.05

BASIS_TO_BUDGET_YEAR = {
    '2026 Reforecast': 'Current Year Reforecast',
    '2027 Budget': 'Next Year Budget',
}


def to_money(tok):
    return int(tok.replace(',', ''))


def to_pct(tok):
    if tok == 'N/A':
        return None
    return float(tok.rstrip('%')) / 100.0


def is_boilerplate(line):
    s = line.strip()
    if not s:
        return True
    for p in BOILERPLATE_PREFIXES:
        if s.startswith(p):
            return True
    if re.match(r'^Jan-\d{2}', s) or re.match(r'^(Budget\s*)+$', s):
        return True
    return False


def is_pct_or_na(tok):
    return bool(PCT_RE.match(tok)) or bool(NA_RE.match(tok))


def extract_pages_text(pdf_file):
    pages = []
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ''
            pages.append(text.split('\n'))
    return pages


def parse_analysis_report(pdf_file):
    """Summary or Detail report: label [+ optional GL code] + 3 values + 2 (delta,pct) pairs + explanation."""
    rows = []
    for page_lines in extract_pages_text(pdf_file):
        current = None
        for raw in page_lines:
            line = raw.rstrip()
            if is_boilerplate(line):
                current = None
                continue
            toks = line.split()
            if not toks:
                current = None
                continue
            idx = 0
            gl = None
            if GLCODE_RE.match(toks[0]):
                gl = toks[0]
                idx = 1
            found = None
            for j in range(idx, len(toks)):
                window = toks[j:j + 7]
                if len(window) < 7:
                    break
                if (MONEY_RE.match(window[0]) and MONEY_RE.match(window[1]) and
                        MONEY_RE.match(window[2]) and MONEY_RE.match(window[3]) and
                        is_pct_or_na(window[4]) and MONEY_RE.match(window[5]) and
                        is_pct_or_na(window[6])):
                    found = j
                    break
            if found is None:
                if current is not None:
                    current['explanation'] = (current['explanation'] + ' ' + line.strip()).strip()
                continue
            label = ' '.join(toks[idx:found]).strip()
            v1, v2, v3, d1, p1, d2, p2 = toks[found:found + 7]
            explanation = ' '.join(toks[found + 7:]).strip()
            current = {
                'gl': gl,
                'label': label,
                'v1_2026B': to_money(v1),
                'v2_2026F': to_money(v2),
                'v3_2027B': to_money(v3),
                'd1_F_vs_B': to_money(d1),
                'p1_F_vs_B': to_pct(p1),
                'd2_27B_vs_F': to_money(d2),
                'p2_27B_vs_F': to_pct(p2),
                'explanation': explanation,
                'is_total': label.startswith('Total'),
            }
            rows.append(current)
    return rows


def parse_monthly_detail(pdf_file):
    rows = []
    for page_lines in extract_pages_text(pdf_file):
        for raw in page_lines:
            line = raw.rstrip()
            if is_boilerplate(line):
                continue
            toks = line.split()
            if not toks:
                continue
            idx = 0
            gl = None
            if GLCODE_RE.match(toks[0]):
                gl = toks[0]
                idx = 1
            found = None
            for j in range(idx, len(toks)):
                window = toks[j:j + 14]
                if len(window) < 14:
                    break
                if all(MONEY_RE.match(t) for t in window[:13]) and DECIMAL_RE.match(window[13]):
                    found = j
                    break
            if found is None:
                continue
            label = ' '.join(toks[idx:found]).strip()
            months = [to_money(t) for t in toks[found:found + 12]]
            total = to_money(toks[found + 12])
            per_rsf = float(toks[found + 13])
            rows.append({
                'gl': gl,
                'label': label,
                'months': months,
                'total': total,
                'per_rsf': per_rsf,
                'is_total': label.startswith('Total'),
            })
    return rows


def get_revision(pdf_file):
    with pdfplumber.open(pdf_file) as pdf:
        text = pdf.pages[0].extract_text() or ''
    m = re.search(r'Revision:\s*(\d+)', text)
    return m.group(1) if m else None


def get_selected_cost_centers(pdf_file):
    """['west20'] for a single-building export, ['West01', 'west20'] for a
    combined one. This report type (unlike Expense Detail/Capex) has never
    been observed in combined form, so parse_analysis_report/parse_monthly_
    detail are NOT validated against it - see check_multiple_cost_centers."""
    with pdfplumber.open(pdf_file) as pdf:
        text = pdf.pages[0].extract_text() or ''
    m = re.search(r'Selected Cost Centers\s*:\s*(.+)', text)
    if not m:
        return []
    return [c.strip() for c in m.group(1).split(',') if c.strip()]


def parse_cost_center_roster(pdf_file):
    """
    Kardin's dedicated 'Selected Cost Centers' report (rptSelectedCostCenters) -
    NOT the same as the "Selected Cost Centers :" line every other report
    carries (see get_selected_cost_centers above, which reads ONE report's own
    scope). This one is a standalone two-column table (Cost Center ID / Cost
    Center Name) listing EVERY cost center Kardin has defined for the
    property - the authoritative source for a property's config.yaml
    cost_centers list (see config_loader.py), rather than inferring it from
    which per-building files happen to have been uploaded.

    Returns [{'code': ..., 'name': ...}], in the report's own row order.
    """
    rows = []
    header_seen = False
    for page_lines in extract_pages_text(pdf_file):
        for raw in page_lines:
            line = raw.rstrip()
            stripped = line.strip()
            if not header_seen:
                if stripped.startswith('Cost Center ID'):
                    header_seen = True
                continue
            if is_boilerplate(line) or not stripped:
                continue
            toks = stripped.split()
            if len(toks) < 2:
                continue
            rows.append({'code': toks[0], 'name': ' '.join(toks[1:])})
    return rows


def meets_materiality(dollar, pct):
    if dollar is None or pct is None:
        return False
    return abs(dollar) >= MATERIALITY_DOLLAR and abs(pct) >= MATERIALITY_PCT


def check_missing_explanations(detail_rows):
    findings = []
    for r in detail_rows:
        if r['is_total'] or r['gl'] is None:
            continue
        gaps = []
        if meets_materiality(r['d1_F_vs_B'], r['p1_F_vs_B']):
            gaps.append(('2026 Reforecast', r['d1_F_vs_B'], r['p1_F_vs_B']))
        if meets_materiality(r['d2_27B_vs_F'], r['p2_27B_vs_F']):
            gaps.append(('2027 Budget', r['d2_27B_vs_F'], r['p2_27B_vs_F']))
        if gaps and not r['explanation']:
            for basis, dollar, pct in gaps:
                findings.append({
                    'Report Section': '2. Budget Analysis',
                    'GL Acct': r['gl'],
                    'Line Item': r['label'],
                    'Budget Year': BASIS_TO_BUDGET_YEAR[basis],
                    'Priority': 'Must Fix',
                    'Comment': (
                        f"Missing variance comment: {basis} variance is "
                        f"${dollar:,.0f} ({pct:.1%}) - both SOP thresholds "
                        f"($2,500 & 5%) are met and no explanation is provided. "
                        f"Per SOP Phase 8, a comment is required."
                    ),
                    'Status': 'Open',
                    'Source Check': 'Missing variance explanation',
                })
    return findings


def check_electric_tie_out(monthly_rows):
    findings = []
    by_gl = {r['gl']: r for r in monthly_rows if r['gl']}
    recovery = by_gl.get('440500')
    reimb = by_gl.get('613115')
    if not recovery or not reimb:
        return findings
    month_names = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    mismatches = []
    for name, rec_v, reimb_v in zip(month_names, recovery['months'], reimb['months']):
        if rec_v != reimb_v:
            mismatches.append(f"{name}: Recovery ${rec_v:,} vs Reimb ${reimb_v:,} (diff ${reimb_v - rec_v:+,})")
    if mismatches:
        findings.append({
            'Report Section': '6. Monthly Detail',
            'GL Acct': '440500 / 613115',
            'Line Item': 'Recovery - Electricity vs Tenant Electric Reimbursement',
            'Budget Year': 'Next Year Budget',
            'Priority': 'Must Fix',
            'Comment': (
                "SOP requires tenant electric recovery to match tenant electric "
                "expense dollar-for-dollar by month. Mismatch found in "
                f"{len(mismatches)} month(s): " + '; '.join(mismatches) +
                f". Annual totals {'match' if recovery['total'] == reimb['total'] else 'DO NOT match'} "
                f"(${recovery['total']:,} vs ${reimb['total']:,})."
            ),
            'Status': 'Open',
            'Source Check': 'Electric recovery tie-out',
        })
    return findings


def check_detail_vs_monthly_totals(detail_rows, monthly_rows):
    findings = []
    monthly_by_gl = {r['gl']: r for r in monthly_rows if r['gl']}
    for r in detail_rows:
        if r['is_total'] or not r['gl']:
            continue
        m = monthly_by_gl.get(r['gl'])
        if m is None:
            continue
        if r['v3_2027B'] != m['total']:
            findings.append({
                'Report Section': '2. Budget Analysis',
                'GL Acct': r['gl'],
                'Line Item': r['label'],
                'Budget Year': 'Next Year Budget',
                'Priority': 'For Discussion',
                'Comment': (
                    f"Detail report shows 2027 Budget = ${r['v3_2027B']:,} but Monthly "
                    f"Detail annual total = ${m['total']:,} (diff ${m['total'] - r['v3_2027B']:+,}). "
                    "Reports may be out of sync - confirm both were exported from the same Kardin revision."
                ),
                'Status': 'Open',
                'Source Check': 'Detail vs Monthly total mismatch',
            })
    return findings


def check_revision_mismatch(summary_file, detail_file, monthly_file):
    revs = {
        'Summary': get_revision(summary_file),
        'Detail': get_revision(detail_file),
        'Monthly Detail': get_revision(monthly_file),
    }
    findings = []
    if len(set(revs.values())) > 1:
        findings.append({
            'Report Section': '1. General',
            'GL Acct': '',
            'Line Item': 'GENERAL',
            'Budget Year': 'N/A',
            'Priority': 'For Discussion',
            'Comment': (
                "Report revision numbers do not match across the three files: "
                + ', '.join(f"{k}=Rev{v}" for k, v in revs.items()) +
                ". Confirm all reports were regenerated after the PM's final Kardin edits "
                "before relying on them for review."
            ),
            'Status': 'Open',
            'Source Check': 'Revision mismatch',
        })
    return findings


def check_multiple_cost_centers(summary_file, detail_file, monthly_file, expected_cost_center=None):
    """
    parse_analysis_report/parse_monthly_detail have only ever been validated
    against single-cost-center exports of this report type. If Kardin breaks
    out a combined export the same way it does for Expense Detail/Capex (GL
    code once, then "West01 (West01)"/"west20 (west20)" tag lines with no
    repeated GL code on the numbers themselves), those rows come back with
    gl=None - which every check in this module silently skips. That means a
    combined-scope export could produce a clean-looking zero-finding result
    that's actually just not checking anything. Flag it instead of guessing.

    expected_cost_center: optional - from a property's config.yaml (see
    config_loader.py). When given, ALSO flags a single-cost-center file that's
    scoped to the WRONG cost center (e.g. B3's Detail file accidentally
    carries B4's export) - not just files scoped to more than one.
    """
    findings = []
    files = {'Summary': summary_file, 'Detail': detail_file, 'Monthly Detail': monthly_file}
    for label, f in files.items():
        centers = get_selected_cost_centers(f)
        if len(centers) > 1:
            findings.append({
                'Report Section': '1. General',
                'GL Acct': '',
                'Line Item': 'GENERAL',
                'Budget Year': 'N/A',
                'Priority': 'Must Fix',
                'Comment': (
                    f"{label} was exported with multiple cost centers selected ({', '.join(centers)}), "
                    "not one building. This parser has only been validated against single-building "
                    "exports of this report type - findings from this file may be incomplete or "
                    "silently wrong. Re-export with a single cost center selected, or treat any "
                    "'no findings' result from this file with suspicion until combined-format support "
                    "is built and validated against a real sample."
                ),
                'Status': 'Open',
                'Source Check': 'Multiple cost centers in single-building report',
            })
        elif expected_cost_center and centers and centers[0].lower() != expected_cost_center.lower():
            findings.append({
                'Report Section': '1. General',
                'GL Acct': '',
                'Line Item': 'GENERAL',
                'Budget Year': 'N/A',
                'Priority': 'Must Fix',
                'Comment': (
                    f"{label} is scoped to cost center '{centers[0]}', but this building is configured "
                    f"as '{expected_cost_center}'. Likely the wrong file was picked for this building - "
                    "every finding below may actually belong to a different building."
                ),
                'Status': 'Open',
                'Source Check': 'Cost center mismatch',
            })
    return findings


def run(summary_file, detail_file, monthly_file, expected_cost_center=None):
    """Parse all three files and run all bucket-1 checks. Files can be paths or
    Streamlit UploadedFile objects (each is read multiple times, so callers
    passing UploadedFile should not need to seek - pdfplumber reads fully
    each time it's opened, and Streamlit's UploadedFile supports re-reading).

    expected_cost_center: optional, see check_multiple_cost_centers."""
    summary_rows = parse_analysis_report(summary_file)
    detail_rows = parse_analysis_report(detail_file)
    monthly_rows = parse_monthly_detail(monthly_file)

    findings = []
    findings += check_multiple_cost_centers(summary_file, detail_file, monthly_file, expected_cost_center)
    findings += check_missing_explanations(detail_rows)
    findings += check_electric_tie_out(monthly_rows)
    findings += check_detail_vs_monthly_totals(detail_rows, monthly_rows)
    findings += check_revision_mismatch(summary_file, detail_file, monthly_file)

    stats = {
        'summary_rows': len(summary_rows),
        'detail_rows': len(detail_rows),
        'detail_gl_rows': len([r for r in detail_rows if r['gl']]),
        'monthly_rows': len(monthly_rows),
        'monthly_gl_rows': len([r for r in monthly_rows if r['gl']]),
    }
    return {'summary_rows': summary_rows, 'detail_rows': detail_rows,
            'monthly_rows': monthly_rows, 'findings': findings, 'stats': stats}
