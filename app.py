from __future__ import annotations

from datetime import datetime
import pandas as pd
import streamlit as st

from report_builder import (
    build_report,
    apply_reporting_period,
    build_reports_by_major_zip,
    clean_placement_data,
    detect_available_majors,
    excel_files_from_zip,
    read_excel,
    read_excel_sheets,
    read_zip_placements,
)

st.set_page_config(
    page_title="Major Company Recruiting Report Builder",
    page_icon="📊",
    layout="wide",
)

st.title("Major Company Recruiting Report Builder")
st.caption("Version: Class Year filter — 2026-06-30")
st.caption(
    "Upload one placement Excel file or a ZIP of major-specific Excel exports. "
    "The app generates the same four-sheet Excel format as the IS/MISM 5-year company report for every major."
)

with st.sidebar:
    st.header("1. Report Period")
    current_class_year = int(st.number_input(
        "Current class year",
        min_value=2000,
        max_value=2100,
        value=2026,
        step=1,
        help="The 1-year report uses this class year. The 5-year report uses this class year plus the previous four class years.",
        key="current_class_year",
    ))
    five_year_count = int(st.number_input(
        "Number of class years in 5-year report",
        min_value=1,
        max_value=10,
        value=5,
        step=1,
        help="Default is 5: Class of 2026, 2025, 2024, 2023, and 2022.",
        key="five_year_count",
    ))
    five_years = [current_class_year - i for i in range(five_year_count)]
    five_year_range_label = f"Class of {five_years[0]}–{five_years[-1]}" if len(five_years) > 1 else f"Class of {five_years[0]}"
    report_period_choice = st.radio(
        "Choose the class-year window",
        [f"5-Year report ({five_year_range_label})", f"Past year report (Class of {current_class_year})"],
        index=0,
        help=(
            "This app filters by Class Year / Class Of, not Start Date. "
            "Past year = current class year only. 5-year = current class year plus previous class years."
        ),
        key="report_period_choice_sidebar",
    )
    report_window = "past_year" if report_period_choice.startswith("Past year") else "all"
    default_scope_label = (
        f"Past Year View — Class of {current_class_year}"
        if report_window == "past_year"
        else f"5-Year View — {five_year_range_label}"
    )
    st.success(f"Current selection: {default_scope_label}")

    st.header("2. Dashboard Options")
    include_major_distribution_chart = st.checkbox(
        "Include major distribution graph",
        value=True,
        help=(
            "Turn this off when the file is already filtered to one major. "
            "The workbook will keep the same template layout but will not add the Major Mix chart."
        ),
        key="include_major_distribution_chart",
    )

    st.header("3. Report Settings")
    report_title = st.text_input("Combined report title", "Employer Recruiting Report")
    scope_label = st.text_input("Scope label shown in Excel", default_scope_label)
    default_output_name = (
        f"Employer_Recruiting_Report_Class_of_{current_class_year}.xlsx"
        if report_window == "past_year"
        else f"Employer_Recruiting_Report_Class_of_{five_years[0]}_{five_years[-1]}.xlsx"
    )
    output_name = st.text_input("Combined report file name", default_output_name)
    st.divider()
    st.markdown("**Template behavior**")
    st.write("Each generated major report uses the exact same structure as the uploaded IS/MISM 5-year report, with the label updated for 5-year or past-year class-year scope:")
    st.write("Executive Dashboard, Company Targets, Summary Tables, Placement Detail")
    st.divider()
    st.write("Optional: upload a CRM/Handshake/contact export to append employer contact fields to Company Targets.")

active_class_years = [current_class_year] if report_window == "past_year" else five_years
active_class_years_text = ", ".join(str(y) for y in active_class_years)
st.info(
    f"Active report period: **{scope_label}**. "
    f"The generated Excel report will filter by Class Year / Class Of: **{active_class_years_text}**."
)

placement_file = st.file_uploader("Upload placement Excel file or ZIP folder export", type=["xlsx", "xls", "zip"])
contact_file = st.file_uploader("Optional contact / Handshake / CRM file", type=["xlsx", "xls", "zip"], key="contact_file")

if not placement_file:
    st.info(
        "Upload a placement workbook or a ZIP like the Master Report folder. Required fields: Company/Employer and Major/Program. "
        "Recommended fields: Job Title, Job Role, Functional Area, Industry, State, Class Year / Class Of, Start Date, Company ID, Job Offer ID, Record ID."
    )
    st.stop()

placement_bytes = placement_file.getvalue()
placement_name = placement_file.name.lower()
manifest = None

try:
    if placement_name.endswith(".zip"):
        excel_manifest = excel_files_from_zip(placement_bytes)
        if not excel_manifest:
            st.error("No Excel files were found inside the ZIP.")
            st.stop()
        placement_df, manifest = read_zip_placements(placement_bytes)
        st.success(f"Imported {len(excel_manifest)} Excel files from the ZIP.")
        with st.expander("Imported files", expanded=False):
            st.dataframe(manifest, width="stretch")
    else:
        sheets = read_excel_sheets(placement_bytes)
        lowered = [s.lower() for s in sheets]
        default_idx = lowered.index("sheet0") if "sheet0" in lowered else lowered.index("placement detail") if "placement detail" in lowered else 0
        sheet_name = st.selectbox("Placement sheet", sheets, index=default_idx)
        placement_df = read_excel(placement_bytes, sheet_name)
        placement_df["Source File"] = placement_file.name
except Exception as exc:
    st.error(f"Could not read the placement source: {exc}")
    st.stop()

contact_df = None
if contact_file:
    contact_bytes = contact_file.getvalue()
    try:
        if contact_file.name.lower().endswith(".zip"):
            contact_df, contact_manifest = read_zip_placements(contact_bytes)
            with st.expander("Imported contact files", expanded=False):
                st.dataframe(contact_manifest, width="stretch")
        else:
            contact_sheets = read_excel_sheets(contact_bytes)
            contact_sheet = st.selectbox("Contact sheet", contact_sheets, index=0)
            contact_df = read_excel(contact_bytes, contact_sheet)
    except Exception as exc:
        st.warning(f"Contact workbook was uploaded but could not be read: {exc}")
        contact_df = None

try:
    detected_majors = detect_available_majors(placement_df)
except Exception:
    detected_majors = []

st.subheader("Major Selection")
if detected_majors:
    selected_majors = st.multiselect(
        "Majors/programs to include",
        detected_majors,
        default=detected_majors,
        help="Select all for a combined report. The separate reports ZIP creates one template-formatted workbook per selected major.",
    )
else:
    default_major = st.text_input("Major/program label", "UNSPECIFIED")
    selected_majors = []

make_combined_report = st.checkbox("Generate combined cross-major report", value=True)
make_per_major_zip = bool(detected_majors and selected_majors)
if make_per_major_zip:
    make_per_major_zip = st.checkbox("Generate one separate template-formatted workbook per selected major", value=True)

st.subheader("Source Preview")
st.dataframe(placement_df.head(25), width="stretch")

try:
    preview_clean, _, _ = clean_placement_data(
        placement_df,
        selected_majors=selected_majors if selected_majors else None,
        default_major=None if detected_majors else default_major,
    )
    preview_clean, period_metadata = apply_reporting_period(
        preview_clean,
        report_window=report_window,
        anchor_class_year=current_class_year,
        five_year_count=five_year_count,
    )
    p1, p2, p3, p4 = st.columns(4)
    p1.metric("Rows after filtering", f"{len(preview_clean):,}")
    p2.metric("Detected companies", f"{preview_clean['Company'].nunique():,}")
    p3.metric("Majors/programs", f"{preview_clean['Major'].nunique():,}")
    period_text = period_metadata.get("Class Years Included", active_class_years_text)
    p4.metric("Class years", period_text)
    st.caption(f"Class-year filter method: {period_metadata.get('Filter Method', 'Unknown')}.")
    if not include_major_distribution_chart:
        st.caption("Major distribution graph will be skipped in the Excel dashboard.")
except Exception as exc:
    st.warning(f"Preview could not be cleaned yet: {exc}")
    preview_clean = None

if st.button("Generate report", type="primary"):
    if not make_combined_report and not make_per_major_zip:
        st.warning("Choose at least one output: combined report or separate reports by major.")
        st.stop()

    if make_combined_report:
        try:
            report_bytes, metrics, company_targets, cleaned = build_report(
                placement_df,
                report_title=report_title,
                scope_label=scope_label,
                selected_majors=selected_majors if selected_majors else None,
                contact_df=contact_df,
                default_major=None if detected_majors else default_major,
                report_window=report_window,
                include_major_distribution_chart=include_major_distribution_chart,
                anchor_class_year=current_class_year,
                five_year_count=five_year_count,
            )
        except Exception as exc:
            st.error(str(exc))
            st.stop()

        st.success("Combined report generated successfully.")
        c1, c2, c3, c4, c5, c6, c7 = st.columns(7)
        c1.metric("Placements", f"{int(metrics['Placements']):,}")
        c2.metric("Companies", f"{int(metrics['Unique Companies']):,}")
        c3.metric("Majors", f"{int(metrics['Majors / Programs']):,}")
        c4.metric("Top Major", str(metrics["Top Major"]))
        c5.metric("Tier 1", f"{int(metrics['Tier 1 Employers']):,}")
        c6.metric("Tier 2", f"{int(metrics['Tier 2 Employers']):,}")
        c7.metric("Scope", scope_label)

        left, right = st.columns([2, 1])
        with left:
            st.markdown("### Top companies")
            st.dataframe(company_targets.head(25), width="stretch")
        with right:
            st.markdown("### Top 10 chart")
            if not company_targets.empty:
                st.bar_chart(company_targets.head(10).set_index("Company")["Total Placements"])

        clean_name = output_name.strip() or "Employer_Recruiting_Report.xlsx"
        if not clean_name.lower().endswith(".xlsx"):
            clean_name += ".xlsx"
        st.download_button(
            label="Download combined Excel report",
            data=report_bytes,
            file_name=clean_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    if make_per_major_zip and selected_majors:
        try:
            zip_bytes = build_reports_by_major_zip(
                placement_df,
                selected_majors,
                report_title_prefix=report_title,
                scope_label=scope_label,
                contact_df=contact_df,
                report_window=report_window,
                include_major_distribution_chart=include_major_distribution_chart,
                anchor_class_year=current_class_year,
                five_year_count=five_year_count,
            )
        except Exception as exc:
            st.error(f"Could not generate reports by major: {exc}")
            st.stop()
        st.success("Separate major reports generated successfully.")
        st.download_button(
            label="Download separate reports by major ZIP",
            data=zip_bytes,
            file_name=(
                f"Employer_Recruiting_Reports_By_Major_Class_of_{current_class_year}.zip"
                if report_window == "past_year"
                else f"Employer_Recruiting_Reports_By_Major_Class_of_{five_years[0]}_{five_years[-1]}.zip"
            ),
            mime="application/zip",
        )

st.caption(f"Last app render: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
