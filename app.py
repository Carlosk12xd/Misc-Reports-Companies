from __future__ import annotations

from datetime import datetime

import pandas as pd
import streamlit as st

from report_builder import build_report, read_excel, read_excel_sheets


st.set_page_config(
    page_title="IS/MISM Employer Recruiting Report Builder",
    page_icon="📊",
    layout="wide",
)

st.title("IS/MISM Employer Recruiting Report Builder")
st.caption("Upload a placement Excel export and generate the same company-focused report format used for the past-year and 5-year employer reports.")

with st.sidebar:
    st.header("Report Settings")
    report_title = st.text_input("Report title", "IS/MISM Employer Recruiting Report")
    scope_label = st.text_input("Scope label", "5-Year View")
    major_text = st.text_input("Majors to include", "BSIS, MISM")
    selected_majors = [m.strip().upper() for m in major_text.split(",") if m.strip()]
    output_name = st.text_input("Output file name", "IS_MISM_Company_Recruiting_Report.xlsx")
    st.divider()
    st.write("Optional: upload a CRM/Handshake/contact export to enrich Company Targets with available employer contact fields.")

placement_file = st.file_uploader("Upload placement Excel file", type=["xlsx", "xls"])
contact_file = st.file_uploader("Optional contact / Handshake / CRM Excel file", type=["xlsx", "xls"], key="contact_file")

if not placement_file:
    st.info("Upload a placement workbook to begin. Expected columns include Company, Major, Job Role, Job Title, Class of, Functional Area, Industry, Start Date, and State.")
    st.stop()

placement_bytes = placement_file.getvalue()
try:
    sheets = read_excel_sheets(placement_bytes)
except Exception as exc:
    st.error(f"Could not read the placement workbook: {exc}")
    st.stop()

# If the workbook already has a Placement Detail sheet, prefer it. Otherwise use the first sheet.
def default_sheet_index(names: list[str]) -> int:
    lowered = [name.lower() for name in names]
    if "placement detail" in lowered:
        return lowered.index("placement detail")
    if "sheet0" in lowered:
        return lowered.index("sheet0")
    return 0

sheet_name = st.selectbox("Placement sheet", sheets, index=default_sheet_index(sheets))

contact_df = None
if contact_file:
    contact_bytes = contact_file.getvalue()
    try:
        contact_sheets = read_excel_sheets(contact_bytes)
        contact_sheet = st.selectbox("Contact sheet", contact_sheets, index=0)
        contact_df = read_excel(contact_bytes, contact_sheet)
    except Exception as exc:
        st.warning(f"Contact workbook was uploaded but could not be read: {exc}")
        contact_df = None

try:
    placement_df = read_excel(placement_bytes, sheet_name)
except Exception as exc:
    st.error(f"Could not read selected placement sheet: {exc}")
    st.stop()

st.subheader("Source Preview")
st.dataframe(placement_df.head(20), use_container_width=True)

if st.button("Generate report", type="primary"):
    try:
        report_bytes, metrics, company_targets, cleaned = build_report(
            placement_df,
            report_title=report_title,
            scope_label=scope_label,
            selected_majors=selected_majors,
            contact_df=contact_df,
        )
    except Exception as exc:
        st.error(str(exc))
        st.stop()

    st.success("Report generated successfully.")

    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric("Placements", f"{metrics['Placements']:,}")
    c2.metric("Companies", f"{metrics['Unique Companies']:,}")
    c3.metric("BSIS", f"{metrics['BSIS Placements']:,}")
    c4.metric("MISM", f"{metrics['MISM Placements']:,}")
    c5.metric("Tier 1", f"{metrics['Tier 1 Employers']:,}")
    c6.metric("Tier 2", f"{metrics['Tier 2 Employers']:,}")

    left, right = st.columns([2, 1])
    with left:
        st.markdown("### Top companies")
        st.dataframe(company_targets.head(25), use_container_width=True)
    with right:
        st.markdown("### Top 10 chart")
        if not company_targets.empty:
            st.bar_chart(company_targets.head(10).set_index("Company")["Total Placements"])

    clean_name = output_name.strip() or "IS_MISM_Company_Recruiting_Report.xlsx"
    if not clean_name.lower().endswith(".xlsx"):
        clean_name += ".xlsx"

    st.download_button(
        label="Download Excel report",
        data=report_bytes,
        file_name=clean_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

st.caption(f"Last app render: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
