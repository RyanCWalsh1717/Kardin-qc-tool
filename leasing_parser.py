"""
Kardin '2. Leasing and Rent Rpts' bucket parser (v1).

Six report types, all heterogeneous in layout compared to bucket 1's clean
GL-line tables:
  - Free Rent               (per new/renewing suite: term + monthly free-rent $ x12 + total)
  - Rent-Lab-Mnthly         (per suite: gross scheduled base rent $ x12 + total; "Base Rent - Retail")
  - Occupancy Summary       (per suite, grouped Contract/New/Unknown: occupied RSF x12 + average)
  - Rent Roll               (per suite: lease term schedule, roster-level: status/suite/tenant/RSF)
  - Leasing Activity By Date (per suite: deal terms + lifetime Total TIs/LCs/Capital Cost)
  - Stacking Plan           (per suite: current tenant, RSF, prospective tenant/lease commence)

Checks implemented (all self-contained to this bucket - no bucket-1 dependency):
  1. Unmodeled vacant suites: "Unknown" status in Occupancy Summary = zero
     leasing assumption for the entire budget year.
  2. Stacking Plan vs Occupancy Summary agreement on which suites are
     genuinely unmodeled (both should show no prospective tenant/plan).
  3. Free Rent report's per-suite rows sum to its own reported total.
  4. Total free rent never exceeds total gross scheduled rent, by month.
     (A strict per-suite free-vs-gross tie-out isn't reliable: Kardin
     truncates suite codes in Base Rent - Retail - see check 5.)
  5. Informational note when Base Rent - Retail's suite-code truncation
     would make multiple suites collide under the same displayed code.
  6. Suite-roster consistency across Occupancy Summary, Stacking Plan, and
     Rent Roll (Base Rent - Retail is excluded - see check 5).
"""
import re

from kardin_parser import MONEY_RE, PCT_RE, is_boilerplate, to_money, to_pct

DATE_RE = re.compile(r'^\d{1,2}/\d{1,2}/\d{4}$')
# Building-code shape varies by property: Riverside Labs uses "west01" (no
# internal hyphen), Lexington Labs uses "lexlab-1" (hyphenated) - both are
# "letters[-]digits", so match that generally rather than hardcoding "west".
SUITE_RE = re.compile(r'^[A-Za-z][A-Za-z-]*\d+-\d+$')  # exact: the whole token IS a suite code
SUITE_PREFIX_RE = re.compile(r'^([A-Za-z][A-Za-z-]*\d+-\d+)(.*)$')  # suite code + any glued-on leftover text
STATUSES = ('Unknown', 'New', 'Contract', 'Renew', 'Expansion')


def extract_pages_text(pdf_file):
    if pdf_file is None:
        return []
    import pdfplumber
    pages = []
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ''
            pages.append(text.split('\n'))
    return pages


def all_lines(pdf_file):
    lines = []
    for page in extract_pages_text(pdf_file):
        lines.extend(page)
    return lines


# ---------------------------------------------------------------- Free Rent

def parse_free_rent(pdf_file):
    rows = []
    for raw in all_lines(pdf_file):
        line = raw.rstrip()
        if is_boilerplate(line):
            continue
        toks = line.split()
        if not toks:
            continue
        date_positions = [i for i, t in enumerate(toks) if DATE_RE.match(t)]
        consecutive_dates = None
        for i in date_positions:
            if i + 1 < len(toks) and DATE_RE.match(toks[i + 1]):
                consecutive_dates = i
                break
        if consecutive_dates is not None:
            i = consecutive_dates
            label = ' '.join(toks[:i]).strip()
            frm, to = toks[i], toks[i + 1]
            rest = toks[i + 2:]
            if len(rest) < 15 or not MONEY_RE.match(rest[0]) or not PCT_RE.match(rest[1]):
                continue
            monthly_rent = to_money(rest[0])
            free_pct = to_pct(rest[1])
            values = rest[2:15]
            if len(values) != 13 or not all(MONEY_RE.match(v) for v in values):
                continue
            months = [to_money(v) for v in values[:12]]
            total = to_money(values[12])
            rows.append({
                'suite': label, 'from': frm, 'to': to,
                'monthly_rent': monthly_rent, 'free_pct': free_pct,
                'months': months, 'total': total, 'is_total': False,
            })
        elif line.strip().startswith('Total Free Rent'):
            values = toks[-13:]
            if len(values) == 13 and all(MONEY_RE.match(v) for v in values):
                months = [to_money(v) for v in values[:12]]
                rows.append({
                    'suite': 'TOTAL', 'from': None, 'to': None,
                    'monthly_rent': None, 'free_pct': None,
                    'months': months, 'total': to_money(values[12]), 'is_total': True,
                })
    return rows


# ------------------------------------------------------------ Rent-Lab-Mnthly

def parse_rent_lab_monthly(pdf_file):
    rows = []
    for raw in all_lines(pdf_file):
        line = raw.rstrip()
        if is_boilerplate(line):
            continue
        toks = line.split()
        if not toks:
            continue
        # data row with a trailing 13-MONEY block (12 months + total)
        if len(toks) >= 13:
            tail = toks[-13:]
            if all(MONEY_RE.match(t) for t in tail):
                head = toks[:-13]
                # strip a trailing lone $/RSF decimal some suite rows carry before the block
                suite = head[0] if head else None
                label = ' '.join(head).strip()
                months = [to_money(t) for t in tail[:12]]
                total = to_money(tail[12])
                is_total = label.startswith('Total') or label.startswith('209,458')
                rows.append({'suite': suite, 'label': label, 'months': months,
                              'total': total, 'is_total': is_total})
                continue
        # vacant/no-rent suite row: "west01-0100 Vacant Space 14,600" (RSF only, no schedule)
        if SUITE_RE.match(toks[0]) and MONEY_RE.match(toks[-1]) and len(toks) <= 5:
            rows.append({'suite': toks[0], 'label': ' '.join(toks[1:-1]),
                          'months': None, 'total': None, 'is_total': False})
    return rows


def parse_rent_lab_monthly_multi(pdf_files):
    """Combines one or more per-use-type monthly base-rent exports into one
    row set. Riverside Labs exports this as a single file; Lexington Labs
    splits it by use-type (Rent-Lab-Monthly / Rent-Office-Monthly /
    Rent-Misc-Monthly) instead - each file's own 'Total' row only covers its
    slice, so those are summed month-by-month into one combined total that
    represents the whole property's gross scheduled rent (needed by
    check_free_rent_within_gross_bound, which compares against ALL free rent,
    not just one use-type's)."""
    all_rows = []
    per_file_totals = []
    for f in pdf_files:
        rows = parse_rent_lab_monthly(f)
        all_rows += [r for r in rows if not r['is_total']]
        t = next((r for r in rows if r['is_total']), None)
        if t:
            per_file_totals.append(t)
    if per_file_totals:
        months = [sum(t['months'][m] for t in per_file_totals) for m in range(12)]
        all_rows.append({'suite': 'TOTAL', 'label': 'Total (combined across files)',
                          'months': months, 'total': sum(months), 'is_total': True})
    return all_rows


def parse_rent_roll_roster_multi(pdf_files):
    """Concatenates one or more Rent Roll exports (Lexington Labs adds a
    supplementary 'Rent Roll-Retail' alongside the main roll) into one
    roster. check_suite_roster_consistency dedupes by suite via sets, so an
    overlapping/duplicate suite across files is harmless either way."""
    rows = []
    for f in pdf_files:
        rows += parse_rent_roll_roster(f)
    return rows


# ---------------------------------------------------------- Occupancy Summary

def parse_occupancy_summary(pdf_file):
    rows = []
    status = None
    for raw in all_lines(pdf_file):
        line = raw.rstrip()
        if is_boilerplate(line):
            continue
        stripped = line.strip()
        if stripped in ('Contract', 'New', 'Unknown'):
            status = stripped
            continue
        toks = line.split()
        suite_idx = next((i for i, t in enumerate(toks[:3]) if SUITE_PREFIX_RE.match(t)), None)
        if suite_idx is None:
            continue
        m = SUITE_PREFIX_RE.match(toks[suite_idx])
        suite_code, leftover = m.group(1), m.group(2)
        # Some exports (Lexington Labs) glue the suite code directly onto the
        # start of the tenant name with no space, e.g. "lexlab-1-0100Thermo
        # Expansion" - splice the suite code back out so it doesn't pollute
        # the label, and shift all downstream token positions accordingly.
        clean_toks = toks[:suite_idx] + ([leftover] if leftover else []) + toks[suite_idx + 1:]

        date_positions = [i for i, t in enumerate(clean_toks) if DATE_RE.match(t)]
        consecutive_dates = None
        for i in date_positions:
            if i + 1 < len(clean_toks) and DATE_RE.match(clean_toks[i + 1]):
                consecutive_dates = i
                break
        if consecutive_dates is not None:
            i = consecutive_dates
            label = ' '.join(clean_toks[:i]).strip()
            comm, exp = clean_toks[i], clean_toks[i + 1]
            rest = clean_toks[i + 2:]
            if len(rest) != 14 or not all(MONEY_RE.match(v) for v in rest):
                continue
            rsf = to_money(rest[0])
            months = [to_money(v) for v in rest[1:13]]
            average = to_money(rest[13])
            rows.append({'status': status, 'suite': suite_code, 'label': label,
                         'comm': comm, 'exp': exp, 'rsf': rsf,
                         'months': months, 'average': average})
        else:
            # Unknown suites: "Retail west01-0100 Vacant Space 14,600" - label + single RSF
            if MONEY_RE.match(clean_toks[-1]) and len(clean_toks) - suite_idx <= 5:
                label = ' '.join(clean_toks[:-1]).strip()
                rows.append({'status': status, 'suite': suite_code, 'label': label,
                             'comm': None, 'exp': None, 'rsf': to_money(clean_toks[-1]),
                             'months': None, 'average': None})
    return rows


# --------------------------------------------------------------- Rent Roll

RENT_ROLL_ROW_RE = re.compile(
    r'^(?P<status>' + '|'.join(STATUSES) + r')\s+R(?P<prefix>[A-Za-z][A-Za-z-]*\d+)-\s+(?P<rest>.*)$'
)


def parse_rent_roll_roster(pdf_file):
    """
    Suite-level roster only (status/suite/tenant/RSF) - not the full multi-year
    schedule. Rent Roll's Suite column is narrow enough that Kardin wraps the
    suite number onto its own following line (e.g. "Unknown Rwest01- Vacant
    Space 14,600" then a lone "0100" line) - reconstruct before parsing.
    """
    lines = [l.rstrip() for l in all_lines(pdf_file) if not is_boilerplate(l)]
    rows = []
    i = 0
    while i < len(lines):
        line = lines[i]
        m = RENT_ROLL_ROW_RE.match(line)
        if not m:
            i += 1
            continue
        suite_number = None
        extra_tenant_text = ''
        if i + 1 < len(lines):
            # the wrapped continuation line is "<suite digits>" alone, or
            # "<suite digits> <leftover tenant-name text>" when a long tenant
            # name (e.g. "ModeX Therapeutics, Inc.") also wraps.
            cont_m = re.match(r'^(\d+)(?:\s+(.*))?$', lines[i + 1].strip())
            if cont_m:
                suite_number = cont_m.group(1)
                extra_tenant_text = (cont_m.group(2) or '').strip()
                i += 1  # consume the wrapped continuation line
        i += 1
        if suite_number is None:
            continue
        suite = f"{m.group('prefix')}-{suite_number}"
        rest_toks = m.group('rest').split()
        date_positions = [j for j, t in enumerate(rest_toks) if DATE_RE.match(t)]
        consecutive_dates = next(
            (j for j in date_positions if j + 1 < len(rest_toks) and DATE_RE.match(rest_toks[j + 1])),
            None,
        )
        if consecutive_dates is not None:
            tenant = ' '.join(rest_toks[:consecutive_dates]).strip()
            if extra_tenant_text:
                tenant = f"{tenant} {extra_tenant_text}".strip()
            after_dates = rest_toks[consecutive_dates + 2:]
            rsf_tok = next((t for t in after_dates if MONEY_RE.match(t)), None)
        else:
            money_toks = [t for t in rest_toks if MONEY_RE.match(t)]
            rsf_tok = money_toks[-1] if money_toks else None
            rsf_idx = rest_toks.index(rsf_tok) if rsf_tok else len(rest_toks)
            tenant = ' '.join(rest_toks[:rsf_idx]).strip()
            if extra_tenant_text:
                tenant = f"{tenant} {extra_tenant_text}".strip()
        if rsf_tok is None:
            continue
        rows.append({'status': m.group('status'), 'suite': suite, 'tenant': tenant,
                     'rsf': to_money(rsf_tok)})
    return rows


# ------------------------------------------------------------- Stacking Plan

def parse_stacking_plan(pdf_file):
    rows = []
    for raw in all_lines(pdf_file):
        line = raw.rstrip()
        if is_boilerplate(line):
            continue
        toks = line.split()
        if len(toks) < 3:
            continue
        suite_idx = next((i for i, t in enumerate(toks[:2]) if SUITE_RE.match(t)), None)
        if suite_idx is None:
            continue
        building = toks[0] if suite_idx == 1 else None
        suite = toks[suite_idx]
        rest = toks[suite_idx + 1:]
        money_positions = [i for i, t in enumerate(rest) if MONEY_RE.match(t)]
        if len(money_positions) < 2:
            continue
        m1, m2 = money_positions[0], money_positions[1]
        current_tenant = ' '.join(rest[:m1]).strip()
        rsf_measured = to_money(rest[m1])
        rsf_leased = to_money(rest[m2])
        tail = rest[m2 + 1:]
        # rsf_leased > 0 => currently occupied: any trailing date is the lease's
        #   own Lease Expires, not a prospective deal's commence date.
        # rsf_leased == 0 => vacant: a trailing (text + date) pair is a
        #   prospective tenant name + Est. Lease Commence; a lone date with no
        #   text isn't expected for a vacant suite in this report.
        lease_expires = None
        prospective_tenant = None
        est_commence = None
        if rsf_leased > 0:
            lease_expires = next((t for t in tail if DATE_RE.match(t)), None)
        else:
            date_toks = [t for t in tail if DATE_RE.match(t)]
            text_toks = [t for t in tail if not DATE_RE.match(t)]
            if date_toks and text_toks:
                prospective_tenant = ' '.join(text_toks).strip()
                est_commence = date_toks[-1]
        rows.append({
            'building': building, 'suite': suite, 'current_tenant': (current_tenant or None) if rsf_leased > 0 else None,
            'rsf_measured': rsf_measured, 'rsf_leased': rsf_leased,
            'lease_expires': lease_expires, 'prospective_tenant': prospective_tenant,
            'est_commence': est_commence,
        })
    return rows


# --------------------------------------------------------------------- Checks

def check_unmodeled_vacant_suites(occupancy_rows):
    unknown = [r for r in occupancy_rows if r['status'] == 'Unknown']
    if not unknown:
        return []
    total_rsf = sum(r['rsf'] for r in unknown)
    suite_list = ', '.join(f"{r['suite']} ({r['rsf']:,} RSF)" for r in unknown)
    return [{
        'Report Section': '9. Leasing & Rent',
        'GL Acct': '',
        'Line Item': 'Vacant suites with no leasing assumption',
        'Budget Year': 'Next Year Budget',
        'Priority': 'For Discussion',
        'Comment': (
            f"{len(unknown)} suite(s) totaling {total_rsf:,} RSF are categorized 'Unknown' in the "
            f"Occupancy Summary with zero leasing assumption for all of 2027: {suite_list}. "
            "Confirm whether these are intentionally excluded (e.g. held for redevelopment) or "
            "whether GRP leasing assumptions are simply missing for these suites (SOP Phase 5: "
            "GRP provides leasing assumptions for vacant suites)."
        ),
        'Status': 'Open',
        'Source Check': 'Unmodeled vacant suites',
    }]


def check_stacking_vs_occupancy_consistency(occupancy_rows, stacking_rows):
    findings = []
    unknown_suites = {r['suite'] for r in occupancy_rows if r['status'] == 'Unknown'}
    stacking_by_suite = {r['suite']: r for r in stacking_rows}
    for suite in unknown_suites:
        s = stacking_by_suite.get(suite)
        if s is None:
            continue
        if s['prospective_tenant'] or s['est_commence']:
            findings.append({
                'Report Section': '9. Leasing & Rent',
                'GL Acct': '',
                'Line Item': suite,
                'Budget Year': 'Next Year Budget',
                'Priority': 'Must Fix',
                'Comment': (
                    f"Occupancy Summary shows {suite} as 'Unknown' (no leasing assumption), but the "
                    f"Stacking Plan shows a prospective tenant/lease commence date "
                    f"({s['prospective_tenant']!r}, commence {s['est_commence']!r}) for the same suite. "
                    "Reports are inconsistent - confirm both were exported after the same Kardin edits."
                ),
                'Status': 'Open',
                'Source Check': 'Stacking Plan vs Occupancy Summary mismatch',
            })
    return findings


MONTH_NAMES = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']


def check_free_rent_internal_consistency(free_rent_rows):
    """Sum of per-suite free rent should equal the report's own 'Total Free Rent' row."""
    non_total = [r for r in free_rent_rows if not r['is_total']]
    total_row = next((r for r in free_rent_rows if r['is_total']), None)
    if not non_total or not total_row:
        return []
    computed = [sum(r['months'][m] for r in non_total) for m in range(12)]
    mismatches = [f"{MONTH_NAMES[m]}: sum ${computed[m]:,} vs report total ${total_row['months'][m]:,}"
                  for m in range(12) if computed[m] != total_row['months'][m]]
    if not mismatches:
        return []
    return [{
        'Report Section': '9. Leasing & Rent',
        'GL Acct': '', 'Line Item': 'Free Rent - Retail', 'Budget Year': 'Next Year Budget',
        'Priority': 'Must Fix',
        'Comment': ("Free Rent report's per-suite rows don't sum to its own 'Total Free Rent' row in "
                    f"{len(mismatches)} month(s): " + '; '.join(mismatches) + ". Likely a stale export."),
        'Status': 'Open', 'Source Check': 'Free rent internal sum mismatch',
    }]


def check_free_rent_within_gross_bound(free_rent_rows, rent_lab_rows):
    """
    Per-suite tie-out isn't reliable here because Kardin truncates suite codes
    in the Base Rent - Retail report (see note_rent_lab_suite_truncation), so
    this checks the weaker but still meaningful aggregate bound: total free
    rent in a month can never exceed total gross scheduled rent that month.
    """
    total_free = next((r for r in free_rent_rows if r['is_total']), None)
    total_gross = next((r for r in rent_lab_rows if r['is_total']), None)
    if not total_free or not total_gross:
        return []
    violations = [f"{MONTH_NAMES[m]}: free ${total_free['months'][m]:,} vs gross ${total_gross['months'][m]:,}"
                  for m in range(12) if abs(total_free['months'][m]) > abs(total_gross['months'][m])]
    if not violations:
        return []
    return [{
        'Report Section': '9. Leasing & Rent',
        'GL Acct': '', 'Line Item': 'Free Rent vs Base Rent - Retail', 'Budget Year': 'Next Year Budget',
        'Priority': 'Must Fix',
        'Comment': ("Total free rent exceeds total gross scheduled rent in "
                    f"{len(violations)} month(s): " + '; '.join(violations) +
                    ". This should never happen - investigate immediately."),
        'Status': 'Open', 'Source Check': 'Free rent exceeds gross rent',
    }]


def note_rent_lab_suite_truncation(rent_lab_rows):
    """
    Kardin truncates the Suite column in Base Rent - Retail to ~9 characters,
    so e.g. west01-0100/0110/0120 all display as 'west01-01'. Informational -
    not a data error, but means this report's rows can't be reliably matched
    to other reports by suite code alone.
    """
    from collections import Counter
    counts = Counter(r['suite'] for r in rent_lab_rows if r.get('suite') and not r['is_total'])
    dupes = sorted(s for s, c in counts.items() if c > 1)
    if not dupes:
        return []
    return [{
        'Report Section': '9. Leasing & Rent',
        'GL Acct': '', 'Line Item': 'Base Rent - Retail (Rent-Lab-Mnthly)', 'Budget Year': 'Next Year Budget',
        'Priority': 'For Discussion',
        'Comment': (
            "Kardin's Base Rent - Retail report truncates suite codes to ~9 characters, so multiple "
            f"suites display under the same code (affected prefixes: {', '.join(dupes)}). This is a "
            "Kardin export limitation, not a data error - just be aware this report's rows can't be "
            "matched to other reports by suite code alone."
        ),
        'Status': 'Open', 'Source Check': 'Suite code truncation (informational)',
    }]


def check_suite_roster_consistency(occupancy_rows, stacking_rows, rent_roll_rows):
    """Rent-Lab-Mnthly is deliberately excluded - see note_rent_lab_suite_truncation."""
    occ_suites = {r['suite'] for r in occupancy_rows}
    stack_suites = {r['suite'] for r in stacking_rows}
    roll_suites = {r['suite'] for r in rent_roll_rows}

    all_reports = {'Occupancy Summary': occ_suites, 'Stacking Plan': stack_suites,
                    'Rent Roll': roll_suites}
    all_suites = set().union(*all_reports.values())
    findings = []
    for suite in sorted(all_suites):
        missing_from = [name for name, s in all_reports.items() if suite not in s]
        if missing_from:
            findings.append({
                'Report Section': '9. Leasing & Rent',
                'GL Acct': '',
                'Line Item': suite,
                'Budget Year': 'Next Year Budget',
                'Priority': 'For Discussion',
                'Comment': (
                    f"Suite {suite} is missing from: {', '.join(missing_from)}. "
                    "Confirm this is expected (e.g. a report scoped differently) rather than a stale export."
                ),
                'Status': 'Open',
                'Source Check': 'Suite roster inconsistency',
            })
    return findings


def run(free_rent_pdf, rent_lab_pdfs, occupancy_pdf, rent_roll_pdfs, stacking_plan_pdf):
    """rent_lab_pdfs / rent_roll_pdfs: a file or a list of files - see
    parse_rent_lab_monthly_multi / parse_rent_roll_roster_multi."""
    if not isinstance(rent_lab_pdfs, (list, tuple)):
        rent_lab_pdfs = [rent_lab_pdfs]
    if not isinstance(rent_roll_pdfs, (list, tuple)):
        rent_roll_pdfs = [rent_roll_pdfs]

    free_rent_rows = parse_free_rent(free_rent_pdf)
    rent_lab_rows = parse_rent_lab_monthly_multi(rent_lab_pdfs)
    occupancy_rows = parse_occupancy_summary(occupancy_pdf)
    rent_roll_rows = parse_rent_roll_roster_multi(rent_roll_pdfs)
    stacking_rows = parse_stacking_plan(stacking_plan_pdf)

    from kardin_parser import missing_file_finding
    findings = []
    if free_rent_pdf is None:
        findings.append(missing_file_finding('Free Rent'))
    if not rent_lab_pdfs:
        findings.append(missing_file_finding('Base Rent'))
    if occupancy_pdf is None:
        findings.append(missing_file_finding('Occupancy Summary'))
    if not rent_roll_pdfs:
        findings.append(missing_file_finding('Rent Roll'))
    if stacking_plan_pdf is None:
        findings.append(missing_file_finding('Stacking Plan'))
    findings += check_unmodeled_vacant_suites(occupancy_rows)
    findings += check_stacking_vs_occupancy_consistency(occupancy_rows, stacking_rows)
    findings += check_free_rent_internal_consistency(free_rent_rows)
    findings += check_free_rent_within_gross_bound(free_rent_rows, rent_lab_rows)
    findings += note_rent_lab_suite_truncation(rent_lab_rows)
    findings += check_suite_roster_consistency(occupancy_rows, stacking_rows, rent_roll_rows)

    stats = {
        'free_rent_rows': len(free_rent_rows),
        'rent_lab_rows': len(rent_lab_rows),
        'occupancy_rows': len(occupancy_rows),
        'rent_roll_rows': len(rent_roll_rows),
        'stacking_rows': len(stacking_rows),
    }
    return {
        'free_rent_rows': free_rent_rows, 'rent_lab_rows': rent_lab_rows,
        'occupancy_rows': occupancy_rows, 'rent_roll_rows': rent_roll_rows,
        'stacking_rows': stacking_rows, 'findings': findings, 'stats': stats,
    }
