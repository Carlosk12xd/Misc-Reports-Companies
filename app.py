from __future__ import annotations

from datetime import datetime
from io import BytesIO
import zipfile

import pandas as pd
import streamlit as st

from report_builder import (
    build_report,
    build_reports_by_major_zip,
    clean_placement_data,
    detect_available_majors,
    excel_files_from_zip,
    read_excel,
    read_excel_sheets,
    read_zip_placements,
)

st.set_page_config(
    page_title="Employer Recruiting Report Builder",
    page_icon="📊",
    layout="wide",
)

st.title("Employer Recruiting Report Builder")
st.caption(
    "Upload one placement Excel file or a ZIP of major-specific Excel files. "
    "The app creates the same company-focused report format, but works for any major/program."
)

with st.sidebar:
    st.header("Report Settings")
    report_title = st.text_input("Report title", "Employer Recruiting Report")
    scope_label = st.text_input("Scope label", "Placement View")
    output_name = st.text_input("Combined report file name", "Employer_Recruiting_Report.xlsx")
    st.divider()
    st.write("Optional: upload a CRM/Handshake/contact export to enrich Company Targets with employer contact fields.")

placement_file = st.file_uploader("Upload placement Excel file or ZIP folder export", type=["xlsx", "xls", "zip"])
contact_file = st.file_uploader("Optional contact / Handshake / CRM file", type=["xlsx", "xls", "zip"], key="contact_file")

if not placement_file:
    st.info(
        "Upload a placement workbook or a ZIP like the Master Report folder. "
        "Expected fields include Company/Employer and Major/Program. Recommended fields include Job Title, Functional Area, Industry, Start Date, State, and Student ID."
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
        default_idx = lowered.index("sheet0") if "sheet0" in lowered else 0
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

if detected_majors:
    selected_majors = st.multiselect(
        "Majors/programs to include",
        detected_majors,
        default=detected_majors,
        help="Leave all selected to build one cross-program report. Select one or more programs to filter the report.",
    )
else:
    default_major = st.text_input("Major/program label", "UNSPECIFIED")
    selected_majors = []

make_per_major_zip = False
if detected_majors and len(selected_majors) > 1:
    make_per_major_zip = st.checkbox("Also generate one separate workbook per selected major", value=True)

st.subheader("Source Preview")
st.dataframe(placement_df.head(25), width="stretch")

try:
    preview_clean, _, _ = clean_placement_data(
        placement_df,
        selected_majors=selected_majors if selected_majors else None,
        default_major=None if detected_majors else default_major,
    )
    p1, p2, p3 = st.columns(3)
    p1.metric("Rows after filtering", f"{len(preview_clean):,}")
    p2.metric("Detected companies", f"{preview_clean['Company'].nunique():,}")
    p3.metric("Majors/programs", f"{preview_clean['Major'].nunique():,}")
except Exception as exc:
    st.warning(f"Preview could not be cleaned yet: {exc}")
    preview_clean = None

if st.button("Generate report", type="primary"):
    try:
        report_bytes, metrics, company_targets, cleaned = build_report(
            placement_df,
            report_title=report_title,
            scope_label=scope_label,
            selected_majors=selected_majors if selected_majors else None,
            contact_df=contact_df,
            default_major=None if detected_majors else default_major,
        )
    except Exception as exc:
        st.error(str(exc))
        st.stop()

    st.success("Report generated successfully.")

    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric("Placements", f"{int(metrics['Placements']):,}")
    c2.metric("Companies", f"{int(metrics['Unique Companies']):,}")
    c3.metric("Majors", f"{int(metrics['Majors / Programs']):,}")
    c4.metric("Top Major", str(metrics["Top Major"]))
    c5.metric("Tier 1", f"{int(metrics['Tier 1 Employers']):,}")
    c6.metric("Tier 2", f"{int(metrics['Tier 2 Employers']):,}")

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
        zip_bytes = build_reports_by_major_zip(
            placement_df,
            selected_majors,
            report_title_prefix=report_title,
            scope_label=scope_label,
            contact_df=contact_df,
        )
        st.download_button(
            label="Download separate reports by major ZIP",
            data=zip_bytes,
            file_name="Employer_Recruiting_Reports_By_Major.zip",
            mime="application/zip",
        )

st.caption(f"Last app render: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
