import re
import traceback
from datetime import date

import openpyxl
import pandas as pd
import streamlit as st

import kardin_parser
import leasing_parser
import recoveries_parser
import expense_parser
import capex_parser
import forecast_parser
import xtra_parser
import tracker_writer

st.set_page_config(page_title="Kardin Budget QC", layout="wide")

st.title("Kardin Budget QC")
st.caption("All 7 file buckets. Cross-bucket checks (4/5/6 → 1, 7 → 2) run automatically "
           "when the earlier bucket has already been analyzed this session for the same building.")

if 'building' not in st.session_state:
    st.session_state.building = ''
if 'bucket_results' not in st.session_state:
    st.session_state.bucket_results = {}  # (bucket_num, building) -> dict


def show_stats(stats):
    st.json(stats, expanded=False)


def show_findings(findings):
    if not findings:
        st.info("No issues found by the current checks.")
        return
    df = pd.DataFrame(findings)[['Priority', 'Report Section', 'GL Acct', 'Line Item',
                                  'Budget Year', 'Comment', 'Source Check']]
    must_fix = df[df['Priority'] == 'Must Fix']
    discuss = df[df['Priority'] != 'Must Fix']
    st.subheader(f"Must Fix ({len(must_fix)})")
    st.caption("Straightforward, rule-based - SOP thresholds, tie-outs, cross-report mismatches.")
    st.dataframe(must_fix, use_container_width=True, hide_index=True)
    st.subheader(f"For Discussion ({len(discuss)})")
    st.caption("Flagged by the parser but needs a human read before it's a real comment.")
    st.dataframe(discuss, use_container_width=True, hide_index=True)


def run_button(key, requirements):
    """requirements: {label: is_satisfied_bool}. Shows exactly what's missing
    instead of just leaving the button greyed out with no explanation."""
    missing = [label for label, ok in requirements.items() if not ok]
    if missing:
        st.caption(f"⚠️ Run Analysis is disabled - still need: {', '.join(missing)}.")
    return st.button("Run Analysis", type="primary", key=key, disabled=bool(missing))


def report_error():
    st.error("Parsing failed. Full traceback below - this is exactly the kind of bug we want to surface fast.")
    st.code(traceback.format_exc())


# Riverside Labs' own bucket-2 file was "...Rent-Lab-Mnthly.pdf" - not
# spelled out. Every rule that keys off "monthly" (to find/exclude a Monthly
# Detail file) checks all of these instead of just the one literal spelling.
MONTHLY_TOKENS = ['monthly', 'mnthly', 'mthly']


def has_bucket_tag(fname_lower, n):
    """True if the filename carries an explicit 'B1'/'B2'/... bucket-number
    tag (as a whole token, so 'b1' doesn't match inside 'b10' etc). Some
    properties' files are pre-labeled this way (discovered on Lexington
    Labs), which is a much stronger signal than keyword-guessing alone."""
    return re.search(rf'(?<![a-z0-9])b0*{n}(?![a-z0-9])', fname_lower) is not None


def classify_and_pick(files, slot_rules, key_prefix, bucket_number=None):
    """
    All files for every bucket are dropped into ONE global uploader; this
    guesses which uploaded file goes in which slot for THIS bucket, then
    renders an editable dropdown per slot (defaulted to the guess, but
    listing every uploaded file) so a naming-convention mismatch never
    blocks you - worst case you just pick manually.

    Matching is two-pass when bucket_number is given: first among files
    carrying an explicit 'B{bucket_number}' tag, then (if nothing tagged
    matches) across the full file list by keyword alone - so this works
    whether or not a property's files use the B-number convention.

    files: list of UploadedFile (or None/empty) - the FULL shared pool.
    slot_rules: {slot_label: [(must_contain_all, must_not_contain_any), ...]}
                rules tried in order, first match wins; substrings matched
                case-insensitively against the filename.
    Returns {slot_label: UploadedFile or None}.
    """
    files = files or []
    by_name = {f.name: f for f in files}
    names = list(by_name.keys())
    assignment = {}
    used = set()

    tagged = [f for f in files if bucket_number is not None and has_bucket_tag(f.name.lower(), bucket_number)]

    for slot, rule_groups in slot_rules.items():
        guess = None
        for pool in (tagged, files):
            for f in pool:
                if f.name in used:
                    continue
                fname_lower = f.name.lower()
                for must, must_not in rule_groups:
                    if (all(s.lower() in fname_lower for s in must)
                            and not any(s.lower() in fname_lower for s in must_not)):
                        guess = f.name
                        break
                if guess:
                    break
            if guess:
                break
        options = ['(none)'] + names
        default_idx = options.index(guess) if guess in options else 0
        sel = st.selectbox(slot, options, index=default_idx, key=f'{key_prefix}_{slot}')
        if sel != '(none)':
            assignment[slot] = by_name[sel]
            used.add(sel)
        else:
            assignment[slot] = None
    return assignment


st.header("Building")
building = st.text_input("Building name (must match the tracker's tab name)", value=st.session_state.building)
st.session_state.building = building
st.caption("e.g. \"20 Riverside\" - exactly as it appears as a sheet tab in your Comment Tracker. "
           "Re-run a bucket if you change this and want its cross-checks against other buckets to line up.")

st.header("Files")
all_files = st.file_uploader(
    "Drop every file for this budget draft here - all buckets, both buildings if applicable, all at once. "
    "Each tab below auto-picks its own files from this pool (using a 'B1'/'B2'/... tag in the filename if "
    "present, otherwise by keyword) - override any guess with its dropdown.",
    type=["pdf", "xlsx"], accept_multiple_files=True, key="all_files",
)
if all_files:
    st.caption(f"{len(all_files)} file(s) uploaded.")
    with st.expander("Show uploaded filenames"):
        for f in all_files:
            st.write(f.name)

tabs = st.tabs([
    "1. Bgt & Fcst Summaries", "2. Leasing & Rent", "3. Recoveries", "4. Expense Back-up",
    "5. CapEx Back-up", "6. Forecast Back-up", "7. Xtra rpts", "Merge to Tracker",
])

# ------------------------------------------------------------- 1. Bgt & Fcst Summaries
with tabs[0]:
    st.caption("Budget Analysis Summary / Detail / Monthly Detail")
    picked = classify_and_pick(all_files, {
        'Budget Analysis Summary': [(['summary'], [])],
        'Budget Analysis Detail': [(['detail'], MONTHLY_TOKENS)],
        'Monthly Budget Detail': [([t], []) for t in MONTHLY_TOKENS],
    }, 'b1', bucket_number=1)
    summary_pdf = picked['Budget Analysis Summary']
    detail_pdf = picked['Budget Analysis Detail']
    monthly_pdf = picked['Monthly Budget Detail']

    if run_button("b1_run", {
        'Budget Analysis Summary': bool(summary_pdf), 'Budget Analysis Detail': bool(detail_pdf),
        'Monthly Budget Detail': bool(monthly_pdf), 'Building name': bool(building),
    }):
        try:
            results = kardin_parser.run(summary_pdf, detail_pdf, monthly_pdf)
            st.session_state.bucket_results[(1, building)] = {'results': results, 'detail_pdf': detail_pdf}
            st.success(f"Parsed - {len(results['findings'])} finding(s).")
        except Exception:
            report_error()

    entry = st.session_state.bucket_results.get((1, building))
    if entry:
        show_stats(entry['results']['stats'])
        show_findings(entry['results']['findings'])

# ------------------------------------------------------------------- 2. Leasing & Rent
with tabs[1]:
    st.caption("Free Rent / Base Rent - Retail (Rent-Lab-Mnthly) / Occupancy Summary / Rent Roll / Stacking Plan")
    picked = classify_and_pick(all_files, {
        'Free Rent': [(['free', 'rent'], [])],
        'Base Rent - Retail (Rent-Lab-Mnthly)': [(['rent-lab'], []), (['base rent'], [])],
        'Occupancy Summary': [(['occupancy'], [])],
        'Rent Roll': [(['rent roll'], [])],
        'Stacking Plan': [(['stacking'], [])],
    }, 'b2', bucket_number=2)
    free_rent_pdf = picked['Free Rent']
    rent_lab_pdf = picked['Base Rent - Retail (Rent-Lab-Mnthly)']
    occupancy_pdf = picked['Occupancy Summary']
    rent_roll_pdf = picked['Rent Roll']
    stacking_pdf = picked['Stacking Plan']

    if run_button("b2_run", {
        'Free Rent': bool(free_rent_pdf), 'Base Rent - Retail': bool(rent_lab_pdf),
        'Occupancy Summary': bool(occupancy_pdf), 'Rent Roll': bool(rent_roll_pdf),
        'Stacking Plan': bool(stacking_pdf), 'Building name': bool(building),
    }):
        try:
            results = leasing_parser.run(free_rent_pdf, rent_lab_pdf, occupancy_pdf, rent_roll_pdf, stacking_pdf)
            st.session_state.bucket_results[(2, building)] = {'results': results}
            st.success(f"Parsed - {len(results['findings'])} finding(s).")
        except Exception:
            report_error()

    entry = st.session_state.bucket_results.get((2, building))
    if entry:
        show_stats(entry['results']['stats'])
        show_findings(entry['results']['findings'])

# --------------------------------------------------------------------------- 3. Recoveries
with tabs[2]:
    st.caption("Recovery Calc Est / Recovery Monthly / Gross Up Schedule / Fixed Factor Calcs (.xlsx)")
    picked = classify_and_pick(all_files, {
        'Recovery Calc Est': [(['calc est'], [])],
        'Recovery Monthly': [(['recovery', t], []) for t in MONTHLY_TOKENS],
        'Gross Up Schedule': [(['gross up'], [])],
    }, 'b3', bucket_number=3)
    calc_est_pdf = picked['Recovery Calc Est']
    recovery_monthly_pdf = picked['Recovery Monthly']
    gross_up_pdf = picked['Gross Up Schedule']

    xlsx_files = [f for f in (all_files or []) if f.name.lower().endswith('.xlsx')]
    picked_xlsx = classify_and_pick(xlsx_files, {'Fixed Factor Calcs (.xlsx)': [([], [])]}, 'b3x', bucket_number=3)
    fixed_factor_xlsx = picked_xlsx['Fixed Factor Calcs (.xlsx)']

    if run_button("b3_run", {
        'Recovery Calc Est': bool(calc_est_pdf), 'Recovery Monthly': bool(recovery_monthly_pdf),
        'Gross Up Schedule': bool(gross_up_pdf), 'Fixed Factor Calcs (.xlsx)': bool(fixed_factor_xlsx),
        'Building name': bool(building),
    }):
        try:
            results = recoveries_parser.run(calc_est_pdf, recovery_monthly_pdf, gross_up_pdf, fixed_factor_xlsx)
            st.session_state.bucket_results[(3, building)] = {'results': results}
            st.success(f"Parsed - {len(results['findings'])} finding(s).")
        except Exception:
            report_error()

    entry = st.session_state.bucket_results.get((3, building))
    if entry:
        show_stats(entry['results']['stats'])
        show_findings(entry['results']['findings'])

# ----------------------------------------------------------------------- 4. Expense Back-up
with tabs[3]:
    st.caption("Expense Detail / Mgmt Fee Calc x2. Cross-checks against bucket 1's Detail when available.")
    picked = classify_and_pick(all_files, {
        'Expense Detail': [(['expense detail'], [])],
        'Mgmt Fee Calc - File 1': [(['mgmt fee'], [])],
        'Mgmt Fee Calc - File 2': [(['mgmt fee'], [])],
    }, 'b4', bucket_number=4)
    expense_detail_pdf = picked['Expense Detail']
    mgmt_fee_a_pdf = picked['Mgmt Fee Calc - File 1']
    mgmt_fee_b_pdf = picked['Mgmt Fee Calc - File 2']

    b1_entry = st.session_state.bucket_results.get((1, building))
    if b1_entry:
        st.caption("✓ Bucket 1 data found for this building - GL tie-out will run.")
    else:
        st.caption("Bucket 1 not yet run for this building - GL tie-out will be skipped.")

    if run_button("b4_run", {
        'Expense Detail': bool(expense_detail_pdf), 'Mgmt Fee Calc - File 1': bool(mgmt_fee_a_pdf),
        'Mgmt Fee Calc - File 2': bool(mgmt_fee_b_pdf), 'Building name': bool(building),
    }):
        try:
            bucket1_rows = b1_entry['results']['detail_rows'] if b1_entry else None
            results = expense_parser.run(
                expense_detail_pdf, mgmt_fee_a_pdf, mgmt_fee_b_pdf,
                bucket1_west20_detail_rows=bucket1_rows,
                mgmt_fee_20r_name=mgmt_fee_a_pdf.name, mgmt_fee_1r_name=mgmt_fee_b_pdf.name,
            )
            st.session_state.bucket_results[(4, building)] = {'results': results}
            st.success(f"Parsed - {len(results['findings'])} finding(s).")
        except Exception:
            report_error()

    entry = st.session_state.bucket_results.get((4, building))
    if entry:
        show_stats(entry['results']['stats'])
        show_findings(entry['results']['findings'])

# ------------------------------------------------------------------------ 5. CapEx Back-up
with tabs[4]:
    st.caption("Capex / TIs / LCs. Cross-checks against bucket 1's Detail when available.")
    picked = classify_and_pick(all_files, {
        'Capex': [(['capex'], [])],
        'TIs': [(['tis'], []), (['tenant improvement'], [])],
        'LCs': [(['lcs'], []), (['leasing commission'], [])],
    }, 'b5', bucket_number=5)
    capex_pdf = picked['Capex']
    tis_pdf = picked['TIs']
    lcs_pdf = picked['LCs']

    b1_entry = st.session_state.bucket_results.get((1, building))
    if b1_entry:
        st.caption("✓ Bucket 1 data found for this building - GL tie-out will run.")
    else:
        st.caption("Bucket 1 not yet run for this building - GL tie-out will be skipped.")

    if run_button("b5_run", {
        'Capex': bool(capex_pdf), 'TIs': bool(tis_pdf), 'LCs': bool(lcs_pdf), 'Building name': bool(building),
    }):
        try:
            capex_line_rows, capex_totals_rows = expense_parser.parse_expense_detail(capex_pdf)
            bucket1_rows = b1_entry['results']['detail_rows'] if b1_entry else None
            results = capex_parser.run(tis_pdf, lcs_pdf, capex_line_rows=capex_line_rows,
                                       capex_totals_rows=capex_totals_rows,
                                       bucket1_west20_detail_rows=bucket1_rows)
            st.session_state.bucket_results[(5, building)] = {'results': results}
            st.success(f"Parsed - {len(results['findings'])} finding(s).")
        except Exception:
            report_error()

    entry = st.session_state.bucket_results.get((5, building))
    if entry:
        show_stats(entry['results']['stats'])
        show_findings(entry['results']['findings'])

# -------------------------------------------------------------------- 6. Forecast Back-up
with tabs[5]:
    st.caption("2026B v 2026F Detail / 2026F Monthly Detail. Cross-checks against bucket 1's Detail when available.")
    picked = classify_and_pick(all_files, {
        '2026B v 2026F Detail': [(['detail'], MONTHLY_TOKENS)],
        '2026F Monthly Detail': [([t], []) for t in MONTHLY_TOKENS],
    }, 'b6', bucket_number=6)
    fc_detail_pdf = picked['2026B v 2026F Detail']
    fc_monthly_pdf = picked['2026F Monthly Detail']

    b1_entry = st.session_state.bucket_results.get((1, building))
    if b1_entry:
        st.caption("✓ Bucket 1 data found for this building - reforecast drift check will run.")
    else:
        st.caption("Bucket 1 not yet run for this building - reforecast drift check will be skipped.")

    if run_button("b6_run", {
        '2026B v 2026F Detail': bool(fc_detail_pdf), '2026F Monthly Detail': bool(fc_monthly_pdf),
        'Building name': bool(building),
    }):
        try:
            bucket1_rows = b1_entry['results']['detail_rows'] if b1_entry else None
            bucket1_pdf = b1_entry['detail_pdf'] if b1_entry else None
            results = forecast_parser.run(fc_detail_pdf, fc_monthly_pdf,
                                          bucket1_detail_rows=bucket1_rows, bucket1_detail_pdf=bucket1_pdf)
            st.session_state.bucket_results[(6, building)] = {'results': results}
            st.success(f"Parsed - {len(results['findings'])} finding(s).")
        except Exception:
            report_error()

    entry = st.session_state.bucket_results.get((6, building))
    if entry:
        show_stats(entry['results']['stats'])
        show_findings(entry['results']['findings'])

# --------------------------------------------------------------------------- 7. Xtra rpts
with tabs[6]:
    st.caption("Lease Expiration Schedule. Cross-checks against bucket 2's Occupancy Summary when available.")
    picked = classify_and_pick(all_files, {
        'Lease Expiration Schedule': [(['lease expiration'], []), (['lease', 'exp'], [])],
    }, 'b7', bucket_number=7)
    lease_exp_pdf = picked['Lease Expiration Schedule']

    b2_entry = st.session_state.bucket_results.get((2, building))
    if b2_entry:
        st.caption("✓ Bucket 2 data found for this building - Occupancy Summary tie-out will run.")
    else:
        st.caption("Bucket 2 not yet run for this building - Occupancy Summary tie-out will be skipped.")

    if run_button("b7_run", {'Lease Expiration Schedule': bool(lease_exp_pdf), 'Building name': bool(building)}):
        try:
            occupancy_rows = b2_entry['results']['occupancy_rows'] if b2_entry else None
            results = xtra_parser.run(lease_exp_pdf, occupancy_rows=occupancy_rows)
            st.session_state.bucket_results[(7, building)] = {'results': results}
            st.success(f"Parsed - {len(results['findings'])} finding(s).")
        except Exception:
            report_error()

    entry = st.session_state.bucket_results.get((7, building))
    if entry:
        show_stats(entry['results']['stats'])
        show_findings(entry['results']['findings'])

# --------------------------------------------------------------------------- Merge to Tracker
with tabs[7]:
    st.caption("Combines findings from every bucket you've analyzed this session for this building.")

    combined_findings = []
    bucket_labels = {
        1: "1. Bgt & Fcst Summaries", 2: "2. Leasing & Rent", 3: "3. Recoveries",
        4: "4. Expense Back-up", 5: "5. CapEx Back-up", 6: "6. Forecast Back-up", 7: "7. Xtra rpts",
    }
    present = []
    for n in range(1, 8):
        e = st.session_state.bucket_results.get((n, building))
        if e:
            combined_findings += e['results']['findings']
            present.append(bucket_labels[n])
    missing = [v for k, v in bucket_labels.items() if bucket_labels[k] not in present]

    if present:
        st.write(f"**Included:** {', '.join(present)}")
    if missing:
        st.write(f"**Not yet run for '{building}':** {', '.join(missing)}")

    st.metric("Total findings to merge", len(combined_findings))

    tracker_file = st.file_uploader("Upload the current GRP Budget Comment Tracker (.xlsx)", type="xlsx",
                                     key="tracker_upload")

    if tracker_file and combined_findings:
        try:
            wb_probe = openpyxl.load_workbook(tracker_file, data_only=True, read_only=True)
            sheet_names = wb_probe.sheetnames
            tracker_file.seek(0)
        except Exception:
            sheet_names = []
            st.error("Couldn't read that file as an Excel workbook.")
            st.code(traceback.format_exc())

        default_idx = sheet_names.index(building) if building in sheet_names else 0
        if sheet_names:
            sheet_name = st.selectbox("Which tab is this building?", sheet_names, index=default_idx)
            comment_by = st.selectbox(
                "Comment By (attributed reviewer)",
                ["Ryan Walsh", "Lauren Sullivan", "Natasha Parker", "Justin Bentayou"],
                index=0,
            )

            if st.button("Merge & prepare download"):
                try:
                    tracker_file.seek(0)
                    buf, n_inserted, warnings = tracker_writer.merge_findings_into_tracker(
                        tracker_file, sheet_name, combined_findings, comment_by=comment_by
                    )
                    for w in warnings:
                        st.warning(w)
                    st.success(f"Inserted {n_inserted} row(s) into '{sheet_name}'.")
                    out_name = f"GRP Budget Comment Tracker (with QC - {date.today().isoformat()}).xlsx"
                    st.download_button(
                        "Download updated tracker",
                        data=buf,
                        file_name=out_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                except tracker_writer.TrackerStructureError as e:
                    st.error(f"Tracker structure issue: {e}")
                except Exception:
                    st.error("Merge failed. Full traceback below.")
                    st.code(traceback.format_exc())
    elif not combined_findings and tracker_file:
        st.info("No findings to merge yet - run at least one bucket above first.")
