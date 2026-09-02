"""
Kardin '7. Xtra rpts' bucket parser (v1).

A catch-all bucket for supplementary, non-standard reports outside Kardin's
core package - e.g. a Lease Expiration Schedule pulled ad hoc from Downloads
rather than the standard Kardin export folder. One file type handled here:

  - Lease Expiration Schedule: per leased/assumed suite, RSF expiring in each
    of the next 10 years, plus a Total row and a %-of-building-RSF row. A
    suite whose lease runs past the 10-year window (e.g. 2037) is listed
    with no year columns at all.

Checks:
  1. Total/% row internal consistency: does each year's % = that year's
     total RSF / building RSF?
  2. Cross-tie to bucket 2: does this file's suite roster, expire date, and
     expiring RSF match bucket 2's Occupancy Summary (Contract + New status
     only - Unknown-status suites have no lease to expire, so are correctly
     absent here)?
"""
import re

from kardin_parser import MONEY_RE, PCT_RE, is_boilerplate, to_money, to_pct

DATE_RE = re.compile(r'^\d{1,2}/\d{1,2}/\d{4}$')
# Building-code shape varies by property - see leasing_parser.SUITE_RE.
SUITE_RE = re.compile(r'^[A-Za-z][A-Za-z-]*\d+-\d+$')


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


def get_building_rsf(pdf_file):
    for line in all_lines(pdf_file):
        m = re.search(r'Cost Center\(s\) RSF:\s*([\d,]+)', line)
        if m:
            return to_money(m.group(1))
    return None


def parse_lease_expiration_schedule(pdf_file):
    """Returns {suites: [{suite, tenant, expire_date, by_year}],
    years: [...], total_by_year: {year: rsf}, pct_by_year: {year: pct},
    building_rsf}."""
    lines = all_lines(pdf_file)
    years = None
    suites = []
    total_by_year = {}
    pct_by_year = {}
    building_rsf = get_building_rsf(pdf_file)

    for raw in lines:
        line = raw.rstrip()
        if is_boilerplate(line) or not line.strip():
            continue
        toks = line.split()
        if not toks:
            continue

        if toks[0] == 'Suite' and 'Expire' in toks:
            years = [int(t) for t in toks if re.match(r'^20\d{2}$', t)]
            continue
        if years is None:
            continue

        if line.strip().startswith('Total Lease Expirations:'):
            values = [t for t in toks if MONEY_RE.match(t)]
            for y, v in zip(years, values):
                total_by_year[y] = to_money(v)
            continue
        if line.strip().startswith('% of Building Total:'):
            values = [t for t in toks if PCT_RE.match(t)]
            for y, v in zip(years, values):
                pct_by_year[y] = to_pct(v)
            continue

        suite_idx = next((i for i, t in enumerate(toks[:2]) if SUITE_RE.match(t)), None)
        if suite_idx is None:
            continue
        date_idx = next((i for i, t in enumerate(toks) if DATE_RE.match(t)), None)
        if date_idx is None:
            continue
        tenant = ' '.join(toks[suite_idx + 1:date_idx]).strip()
        expire_date = toks[date_idx]
        year_values = [t for t in toks[date_idx + 1:] if MONEY_RE.match(t)]
        by_year = {y: to_money(v) for y, v in zip(years, year_values)}
        suites.append({'suite': toks[suite_idx], 'tenant': tenant,
                       'expire_date': expire_date, 'by_year': by_year})

    return {'suites': suites, 'years': years or [], 'total_by_year': total_by_year,
            'pct_by_year': pct_by_year, 'building_rsf': building_rsf}


# --------------------------------------------------------------------- Checks

def check_percent_consistency(schedule, tolerance=0.001):
    findings = []
    if not schedule['building_rsf']:
        return findings
    for year in schedule['years']:
        total = schedule['total_by_year'].get(year)
        pct = schedule['pct_by_year'].get(year)
        if total is None or pct is None:
            continue
        expected_pct = total / schedule['building_rsf']
        if abs(expected_pct - pct) > tolerance:
            findings.append({
                'Report Section': '1. General',
                'GL Acct': '', 'Line Item': f'Lease Expiration Schedule - {year}', 'Budget Year': 'N/A',
                'Priority': 'For Discussion',
                'Comment': (
                    f"{year} shows {total:,} RSF expiring ({pct:.2%} of building), but "
                    f"{total:,}/{schedule['building_rsf']:,} = {expected_pct:.2%} - the printed percentage "
                    "doesn't match. Possibly a stale export."
                ),
                'Status': 'Open', 'Source Check': 'Lease Expiration % mismatch',
            })
    return findings


def check_cross_tie_vs_occupancy(schedule, occupancy_rows):
    """occupancy_rows: leasing_parser.parse_occupancy_summary() output for
    the same building. Only Contract/New status suites have a lease to
    expire - Unknown-status suites are correctly absent from this report."""
    findings = []
    occ_by_suite = {r['suite']: r for r in occupancy_rows if r['status'] in ('Contract', 'New')}
    schedule_suites = {s['suite'] for s in schedule['suites']}

    for s in schedule['suites']:
        occ = occ_by_suite.get(s['suite'])
        if occ is None:
            findings.append({
                'Report Section': '1. General',
                'GL Acct': '', 'Line Item': s['suite'], 'Budget Year': 'N/A',
                'Priority': 'For Discussion',
                'Comment': (
                    f"{s['suite']} appears in the Lease Expiration Schedule but not as a Contract/New "
                    "suite in the Occupancy Summary (bucket 2). Confirm both were exported from the same "
                    "Kardin revision."
                ),
                'Status': 'Open', 'Source Check': 'Lease Expiration vs Occupancy Summary roster mismatch',
            })
            continue
        if s['expire_date'] != occ['exp']:
            findings.append({
                'Report Section': '1. General',
                'GL Acct': '', 'Line Item': s['suite'], 'Budget Year': 'N/A',
                'Priority': 'Must Fix',
                'Comment': (
                    f"{s['suite']}'s expiration date differs between reports: Lease Expiration Schedule "
                    f"shows {s['expire_date']}, Occupancy Summary shows {occ['exp']}."
                ),
                'Status': 'Open', 'Source Check': 'Lease Expiration date mismatch',
            })
        expiring_rsf = sum(s['by_year'].values())
        if expiring_rsf and expiring_rsf != occ['rsf']:
            findings.append({
                'Report Section': '1. General',
                'GL Acct': '', 'Line Item': s['suite'], 'Budget Year': 'N/A',
                'Priority': 'Must Fix',
                'Comment': (
                    f"{s['suite']}'s expiring RSF ({expiring_rsf:,}) doesn't match its Occupancy Summary "
                    f"RSF ({occ['rsf']:,})."
                ),
                'Status': 'Open', 'Source Check': 'Lease Expiration RSF mismatch',
            })

    for suite, occ in occ_by_suite.items():
        if suite not in schedule_suites:
            findings.append({
                'Report Section': '1. General',
                'GL Acct': '', 'Line Item': suite, 'Budget Year': 'N/A',
                'Priority': 'For Discussion',
                'Comment': (
                    f"{suite} is a Contract/New suite in the Occupancy Summary but doesn't appear in the "
                    "Lease Expiration Schedule at all - confirm it was included in the export."
                ),
                'Status': 'Open', 'Source Check': 'Lease Expiration vs Occupancy Summary roster mismatch',
            })
    return findings


def run(lease_expiration_pdf, occupancy_rows=None):
    schedule = parse_lease_expiration_schedule(lease_expiration_pdf)

    from kardin_parser import missing_file_finding
    findings = []
    if lease_expiration_pdf is None:
        findings.append(missing_file_finding('Lease Expiration Schedule'))
    findings += check_percent_consistency(schedule)
    if occupancy_rows is not None:
        findings += check_cross_tie_vs_occupancy(schedule, occupancy_rows)

    stats = {'suites': len(schedule['suites']), 'years': len(schedule['years'])}
    return {'schedule': schedule, 'findings': findings, 'stats': stats}
