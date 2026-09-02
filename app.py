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


def run_button(key, ready):
    return st.button("Run Analysis", type="primary", key=key, disabled=not ready)


def report_error():
    st.error("Parsing failed. Full traceback below - this is exactly the kind of bug we want to surface fast.")
    st.code(traceback.format_exc())


st.header("Building")
building = st.text_input("Building name (must match the tracker's tab name)", value=st.session_state.building)
st.session_state.building = building
st.caption("e.g. \"20 Riverside\" - exactly as it appears as a sheet tab in your Comment Tracker. "
           "Re-run a bucket if you change this and want its cross-checks against other buckets to line up.")

tabs = st.tabs([
    "1. Bgt & Fcst Summaries", "2. Leasing & Rent", "3. Recoveries", "4. Expense Back-up",
    "5. CapEx Back-up", "6. Forecast Back-up", "7. Xtra rpts", "Merge to Tracker",
])

# ------------------------------------------------------------- 1. Bgt & Fcst Summaries
with tabs[0]:
    st.caption("Budget Analysis Summary / Detail / Monthly Detail")
    c1, c2, c3 = st.columns(3)
    summary_pdf = c1.file_uploader("Budget Analysis Summary", type="pdf", key="b1_summary")
    detail_pdf = c2.file_uploader("Budget Analysis Detail", type="pdf", key="b1_detail")
    monthly_pdf = c3.file_uploader("Monthly Budget Detail", type="pdf", key="b1_monthly")

    if run_button("b1_run", summary_pdf and detail_pdf and monthly_pdf and building):
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
    c1, c2, c3, c4, c5 = st.columns(5)
    free_rent_pdf = c1.file_uploader("Free Rent", type="pdf", key="b2_freerent")
    rent_lab_pdf = c2.file_uploader("Base Rent - Retail", type="pdf", key="b2_rentlab")
    occupancy_pdf = c3.file_uploader("Occupancy Summary", type="pdf", key="b2_occupancy")
    rent_roll_pdf = c4.file_uploader("Rent Roll", type="pdf", key="b2_rentroll")
    stacking_pdf = c5.file_uploader("Stacking Plan", type="pdf", key="b2_stacking")

    ready2 = all([free_rent_pdf, rent_lab_pdf, occupancy_pdf, rent_roll_pdf, stacking_pdf, building])
    if run_button("b2_run", ready2):
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
    c1, c2, c3, c4 = st.columns(4)
    calc_est_pdf = c1.file_uploader("Recovery Calc Est", type="pdf", key="b3_calcest")
    recovery_monthly_pdf = c2.file_uploader("Recovery Monthly", type="pdf", key="b3_recmonthly")
    gross_up_pdf = c3.file_uploader("Gross Up Schedule", type="pdf", key="b3_grossup")
    fixed_factor_xlsx = c4.file_uploader("Fixed Factor Calcs", type="xlsx", key="b3_fixedfactor")

    ready3 = all([calc_est_pdf, recovery_monthly_pdf, gross_up_pdf, fixed_factor_xlsx, building])
    if run_button("b3_run", ready3):
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
    st.caption("Expense Detail / Mgmt Fee Calc x2 - cross-checks against bucket 1's west20 Detail when available")
    c1, c2, c3 = st.columns(3)
    expense_detail_pdf = c1.file_uploader("Expense Detail", type="pdf", key="b4_expdetail")
    mgmt_fee_20r_pdf = c2.file_uploader("Mgmt Fee Calc (20 Riverside)", type="pdf", key="b4_mgmt20r")
    mgmt_fee_1r_pdf = c3.file_uploader("Mgmt Fee Calc (1 Riverside)", type="pdf", key="b4_mgmt1r")

    b1_entry = st.session_state.bucket_results.get((1, building))
    if b1_entry:
        st.caption("✓ Bucket 1 data found for this building - west20 GL tie-out will run.")
    else:
        st.caption("Bucket 1 not yet run for this building - west20 GL tie-out will be skipped.")

    if run_button("b4_run", expense_detail_pdf and mgmt_fee_20r_pdf and mgmt_fee_1r_pdf and building):
        try:
            bucket1_rows = b1_entry['results']['detail_rows'] if b1_entry else None
            results = expense_parser.run(expense_detail_pdf, mgmt_fee_20r_pdf, mgmt_fee_1r_pdf,
                                          bucket1_west20_detail_rows=bucket1_rows)
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
    st.caption("Capex / TIs / LCs - cross-checks against bucket 1's west20 Detail when available")
    c1, c2, c3 = st.columns(3)
    capex_pdf = c1.file_uploader("Capex", type="pdf", key="b5_capex")
    tis_pdf = c2.file_uploader("TIs", type="pdf", key="b5_tis")
    lcs_pdf = c3.file_uploader("LCs", type="pdf", key="b5_lcs")

    b1_entry = st.session_state.bucket_results.get((1, building))
    if b1_entry:
        st.caption("✓ Bucket 1 data found for this building - west20 GL tie-out will run.")
    else:
        st.caption("Bucket 1 not yet run for this building - west20 GL tie-out will be skipped.")

    if run_button("b5_run", capex_pdf and tis_pdf and lcs_pdf and building):
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
    st.caption("2026B v 2026F Detail / 2026F Monthly Detail - cross-checks against bucket 1's Detail when available")
    c1, c2 = st.columns(2)
    fc_detail_pdf = c1.file_uploader("2026B v 2026F Detail", type="pdf", key="b6_detail")
    fc_monthly_pdf = c2.file_uploader("2026F Monthly Detail", type="pdf", key="b6_monthly")

    b1_entry = st.session_state.bucket_results.get((1, building))
    if b1_entry:
        st.caption("✓ Bucket 1 data found for this building - reforecast drift check will run.")
    else:
        st.caption("Bucket 1 not yet run for this building - reforecast drift check will be skipped.")

    if run_button("b6_run", fc_detail_pdf and fc_monthly_pdf and building):
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
    st.caption("Lease Expiration Schedule - cross-checks against bucket 2's Occupancy Summary when available")
    lease_exp_pdf = st.file_uploader("Lease Expiration Schedule", type="pdf", key="b7_leaseexp")

    b2_entry = st.session_state.bucket_results.get((2, building))
    if b2_entry:
        st.caption("✓ Bucket 2 data found for this building - Occupancy Summary tie-out will run.")
    else:
        st.caption("Bucket 2 not yet run for this building - Occupancy Summary tie-out will be skipped.")

    if run_button("b7_run", lease_exp_pdf and building):
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
