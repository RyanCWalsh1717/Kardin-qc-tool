"""
Kardin '6. Forecast Back-up' bucket parser (v1).

Two report types per building (1 Riverside, 20 Riverside):
  - {Building} Detail          2-way "Budget Analysis Detail": 2026 Budget vs
                                2026 Reforecast only (4 numeric columns: v1,
                                v2, delta, %change - NOT bucket 1's 7-column
                                3-way layout), plus an Explanation column.
  - {Building} Monthly Detail  "2026 Monthly Reforecast Detail": Jan-Jul (or
                                Jan-May for Barings) are Actuals, the rest are
                                Reforecast - but the column shape (12 months +
                                total + $/RSF) matches bucket 1's monthly
                                format exactly, so kardin_parser.parse_monthly_
                                detail is reused directly rather than rewritten.

Checks:
  1. Missing variance explanation (SOP Phase 8, same $2,500-and-5% rule as
     bucket 1) on the 2026 Refore-vs-2026 Budget comparison.
  2. Tenant electric tie-out on the 2026 reforecast year (reuses kardin_
     parser.check_electric_tie_out on this bucket's Monthly Detail).
  3. Detail vs Monthly Detail total tie-out, within this bucket.
  4. Reforecast drift: this file's 2026 Reforecast figures vs bucket 1's
     already-parsed 2026 Reforecast column for the same GL - a real check
     since the two reports are frequently exported at different Kardin
     revisions (this bucket's files run days-to-weeks behind bucket 1's).
"""
import re

from kardin_parser import (
    MONEY_RE, PCT_RE, is_boilerplate, to_money, to_pct, meets_materiality,
    get_revision, check_electric_tie_out, parse_monthly_detail,
)

BOILERPLATE_EXTRA = ('Budget Analysis Detail', 'Explanation of Variance')


def _is_boilerplate(line):
    if is_boilerplate(line):
        return True
    return any(line.strip().startswith(p) for p in BOILERPLATE_EXTRA)


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


def is_pct_or_na(tok):
    return bool(PCT_RE.match(tok)) or tok == 'N/A'


def parse_2way_detail(pdf_file):
    """2026 Budget vs 2026 Reforecast Detail: label [+ GL] + v1 + v2 + delta + pct + explanation."""
    rows = []
    current = None
    for raw in all_lines(pdf_file):
        line = raw.rstrip()
        if _is_boilerplate(line):
            current = None
            continue
        toks = line.split()
        if not toks:
            current = None
            continue
        idx = 0
        gl = None
        if re.match(r'^\d{6}$', toks[0]):
            gl = toks[0]
            idx = 1
        found = None
        for j in range(idx, len(toks)):
            window = toks[j:j + 4]
            if len(window) < 4:
                break
            if (MONEY_RE.match(window[0]) and MONEY_RE.match(window[1]) and
                    MONEY_RE.match(window[2]) and is_pct_or_na(window[3])):
                found = j
                break
        if found is None:
            if current is not None:
                current['explanation'] = (current['explanation'] + ' ' + line.strip()).strip()
            continue
        label = ' '.join(toks[idx:found]).strip()
        v1, v2, d, p = toks[found:found + 4]
        explanation = ' '.join(toks[found + 4:]).strip()
        current = {
            'gl': gl, 'label': label,
            'v1_2026B': to_money(v1), 'v2_2026F': to_money(v2),
            'd_F_vs_B': to_money(d), 'p_F_vs_B': to_pct(p),
            'explanation': explanation, 'is_total': label.startswith('Total'),
        }
        rows.append(current)
    return rows


# --------------------------------------------------------------------- Checks

def check_missing_explanations(detail_rows):
    findings = []
    for r in detail_rows:
        if r['is_total'] or r['gl'] is None:
            continue
        if not meets_materiality(r['d_F_vs_B'], r['p_F_vs_B']):
            continue
        if r['explanation']:
            continue
        findings.append({
            'Report Section': '10. Forecast Back-up',
            'GL Acct': r['gl'], 'Line Item': r['label'], 'Budget Year': 'Current Year Reforecast',
            'Priority': 'Must Fix',
            'Comment': (
                f"Missing variance comment: 2026 Reforecast vs 2026 Budget variance is "
                f"${r['d_F_vs_B']:,} ({r['p_F_vs_B']:.1%}) - both SOP thresholds ($2,500 & 5%) are met "
                "and no explanation is provided. Per SOP Phase 8, a comment is required."
            ),
            'Status': 'Open', 'Source Check': 'Missing variance explanation',
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
        if r['v2_2026F'] != m['total']:
            findings.append({
                'Report Section': '10. Forecast Back-up',
                'GL Acct': r['gl'], 'Line Item': r['label'], 'Budget Year': 'Current Year Reforecast',
                'Priority': 'For Discussion',
                'Comment': (
                    f"Detail report shows 2026 Reforecast = ${r['v2_2026F']:,} but Monthly Reforecast "
                    f"Detail annual total = ${m['total']:,} (diff ${m['total']-r['v2_2026F']:+,}). "
                    "Reports may be out of sync - confirm both were exported from the same Kardin revision."
                ),
                'Status': 'Open', 'Source Check': 'Detail vs Monthly total mismatch',
            })
    return findings


def check_refore_drift_vs_bucket1(bucket6_detail_rows, bucket1_detail_rows, rev6, rev1, tolerance=1):
    """
    Bucket 1's original 3-way Detail (2026B/2026F/2027B) already contains a
    2026 Reforecast column. This bucket's dedicated 2-way file is usually
    exported later in the cycle, so its 2026 Reforecast figures can drift
    from what bucket 1 originally showed - a real, actionable staleness
    signal, not just a formality.
    """
    if rev6 == rev1:
        return []
    b1_by_gl = {r['gl']: r for r in bucket1_detail_rows if r.get('gl')}
    findings = []
    for r in bucket6_detail_rows:
        if r['is_total'] or not r['gl']:
            continue
        b1 = b1_by_gl.get(r['gl'])
        if b1 is None:
            continue
        if abs(r['v2_2026F'] - b1['v2_2026F']) > tolerance:
            findings.append({
                'Report Section': '10. Forecast Back-up',
                'GL Acct': r['gl'], 'Line Item': r['label'], 'Budget Year': 'Current Year Reforecast',
                'Priority': 'Must Fix',
                'Comment': (
                    f"2026 Reforecast for this GL has drifted between exports: bucket 1's original Detail "
                    f"(Rev {rev1}) showed ${b1['v2_2026F']:,}, but this Forecast Back-up Detail (Rev {rev6}) "
                    f"shows ${r['v2_2026F']:,} (diff ${r['v2_2026F']-b1['v2_2026F']:+,}). Any 2027-budget "
                    "variance comment written against the older figure may now be stale - re-check it "
                    "against the current reforecast."
                ),
                'Status': 'Open', 'Source Check': 'Reforecast drift vs bucket 1',
            })
    return findings


def run(detail_pdf, monthly_pdf, bucket1_detail_rows=None, bucket1_detail_pdf=None):
    detail_rows = parse_2way_detail(detail_pdf)
    monthly_rows = parse_monthly_detail(monthly_pdf)

    from kardin_parser import missing_file_finding
    findings = []
    if detail_pdf is None:
        findings.append(missing_file_finding('2026B v 2026F Detail'))
    if monthly_pdf is None:
        findings.append(missing_file_finding('2026F Monthly Detail'))
    findings += check_missing_explanations(detail_rows)
    findings += check_electric_tie_out(monthly_rows)
    findings += check_detail_vs_monthly_totals(detail_rows, monthly_rows)

    if bucket1_detail_rows is not None:
        rev6 = get_revision(detail_pdf)
        rev1 = get_revision(bucket1_detail_pdf) if bucket1_detail_pdf else None
        findings += check_refore_drift_vs_bucket1(detail_rows, bucket1_detail_rows, rev6, rev1 or '?')

    stats = {
        'detail_rows': len(detail_rows),
        'detail_gl_rows': len([r for r in detail_rows if r['gl']]),
        'monthly_rows': len(monthly_rows),
    }
    return {'detail_rows': detail_rows, 'monthly_rows': monthly_rows, 'findings': findings, 'stats': stats}
