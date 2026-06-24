from __future__ import annotations

import math
import re
from collections import Counter
from datetime import datetime
from difflib import SequenceMatcher
from io import BytesIO
from typing import Dict, Iterable, List, Optional, Tuple

import pandas as pd


CORE_COLUMNS = [
    "Rank",
    "Company",
    "Total Placements",
    "BSIS",
    "MISM",
    "Major Mix",
    "Class Year(s)",
    "Primary Industry",
    "Primary Functional Area",
    "Job Role(s)",
    "Representative Job Title(s)",
    "State(s)",
    "First Start Date",
    "Latest Start Date",
    "Employer Tier",
    "Recruiting Priority",
]

CONTACT_COLUMNS = [
    "Handshake Link",
    "Website",
    "LinkedIn",
    "Employer Owner",
    "CSS Assigned",
    "Primary Contact",
    "Main Contact",
    "Engagement Status",
    "Contact Location",
    "Employee Count",
    "Contact Match Method",
]

PLACEMENT_DETAIL_COLUMNS = [
    "Company",
    "Major",
    "Job Role",
    "Job Title",
    "Class of",
    "Functional Area",
    "Industry",
    "Start Date",
    "State",
    "Student ID",
]

COLUMN_ALIASES = {
    "company": [
        "company", "company name", "employer", "employer name", "organization", "organization name",
        "account name", "company/account", "company / account",
    ],
    "major": ["major", "program", "degree", "academic program", "student major"],
    "job_role": ["job role", "role", "position category", "job category", "job family"],
    "job_title": ["job title", "title", "position", "offer title"],
    "class_year": ["class of", "class year", "graduation year", "grad year", "year"],
    "functional_area": ["functional area", "function", "career function", "business function"],
    "industry": ["industry", "industry name", "sector"],
    "start_date": ["start date", "job start date", "offer start date", "hire date", "employment start date"],
    "state": ["state", "state/province", "region", "location state", "work state", "province"],
    "student_id": ["student id", "student id.id", "student", "student identifier", "net id", "byu id"],
}

CONTACT_ALIASES = {
    "company": ["company", "company name", "employer", "employer name", "account name", "organization", "name"],
    "handshake": ["handshake", "handshake link", "handshake url", "handshake employer link", "employer handshake link"],
    "website": ["website", "company website", "web site", "url", "employer website"],
    "linkedin": ["linkedin", "linkedin url", "linkedin link", "company linkedin"],
    "owner": ["owner", "account owner", "employer owner", "record owner"],
    "css": ["css", "css assigned", "assigned css", "career success specialist"],
    "primary_contact": ["primary contact", "primary recruiter", "recruiter", "contact", "contact name"],
    "main_contact": ["main contact", "main recruiter", "main point of contact", "poc"],
    "engagement_status": ["engagement status", "status", "outreach status", "employer status"],
    "location": ["location", "headquarters", "hq", "city", "address"],
    "employee_count": ["employee count", "employees", "number of employees", "company size"],
}


def normalize_header(value: object) -> str:
    text = str(value or "").strip().lower()
    text = re.sub(r"[\n\r\t]+", " ", text)
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def normalize_company(value: object) -> str:
    text = str(value or "").lower().strip()
    text = text.replace("&", " and ")
    text = re.sub(r"\b(incorporated|inc|llc|l l c|ltd|limited|corp|corporation|co|company|llp|pllc|lp|the)\b", " ", text)
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def first_present(values: Iterable[object]) -> str:
    for value in values:
        if pd.notna(value) and str(value).strip():
            return str(value).strip()
    return ""


def top_value(values: Iterable[object], default: str = "Unspecified") -> str:
    cleaned = [str(v).strip() for v in values if pd.notna(v) and str(v).strip()]
    if not cleaned:
        return default
    return Counter(cleaned).most_common(1)[0][0]


def join_top(values: Iterable[object], limit: int = 6) -> str:
    cleaned = [str(v).strip() for v in values if pd.notna(v) and str(v).strip()]
    if not cleaned:
        return "Unspecified"
    ordered = [value for value, _ in Counter(cleaned).most_common(limit)]
    return ", ".join(ordered)


def join_unique_sorted(values: Iterable[object], limit: Optional[int] = None) -> str:
    cleaned: List[str] = []
    for value in values:
        if pd.isna(value) or not str(value).strip():
            continue
        text = str(value).strip()
        # Keep class years like 2026 clean when Excel parsed them as floats.
        if re.fullmatch(r"\d+\.0", text):
            text = text[:-2]
        cleaned.append(text)
    unique = sorted(set(cleaned), key=lambda x: (not x.isdigit(), x))
    if limit:
        unique = unique[:limit]
    return ", ".join(unique) if unique else "Unspecified"


def infer_columns(df: pd.DataFrame, aliases: Dict[str, List[str]]) -> Dict[str, str]:
    normalized_lookup = {normalize_header(col): col for col in df.columns}
    result: Dict[str, str] = {}

    # Exact alias match first.
    for canonical, candidates in aliases.items():
        for candidate in candidates:
            norm = normalize_header(candidate)
            if norm in normalized_lookup:
                result[canonical] = normalized_lookup[norm]
                break

    # Conservative fuzzy fallback for messy exports.
    all_norms = list(normalized_lookup.keys())
    for canonical, candidates in aliases.items():
        if canonical in result:
            continue
        best_score = 0.0
        best_col = None
        for candidate in candidates:
            cand = normalize_header(candidate)
            for norm_col in all_norms:
                score = SequenceMatcher(None, cand, norm_col).ratio()
                if score > best_score:
                    best_score = score
                    best_col = normalized_lookup[norm_col]
        if best_score >= 0.88 and best_col is not None:
            result[canonical] = best_col
    return result


def clean_placement_data(df: pd.DataFrame, selected_majors: Optional[List[str]] = None) -> Tuple[pd.DataFrame, Dict[str, str], List[str]]:
    mapping = infer_columns(df, COLUMN_ALIASES)
    required = ["company", "major"]
    missing = [field for field in required if field not in mapping]
    if missing:
        raise ValueError(
            "Could not identify required column(s): " + ", ".join(missing) +
            ". Make sure the upload includes company and major/program fields."
        )

    cleaned = pd.DataFrame()
    cleaned["Company"] = df[mapping["company"]].astype(str).str.strip()
    cleaned["Major"] = df[mapping["major"]].astype(str).str.strip().str.upper()

    optional_map = {
        "job_role": "Job Role",
        "job_title": "Job Title",
        "class_year": "Class of",
        "functional_area": "Functional Area",
        "industry": "Industry",
        "state": "State",
        "student_id": "Student ID",
    }
    for source_key, out_col in optional_map.items():
        if source_key in mapping:
            cleaned[out_col] = df[mapping[source_key]].apply(lambda x: "" if pd.isna(x) else str(x).strip())
        else:
            cleaned[out_col] = ""

    if "start_date" in mapping:
        cleaned["Start Date"] = pd.to_datetime(df[mapping["start_date"]], errors="coerce")
    else:
        cleaned["Start Date"] = pd.NaT

    cleaned = cleaned[cleaned["Company"].notna() & (cleaned["Company"].str.strip() != "")]
    cleaned = cleaned[~cleaned["Company"].str.lower().isin(["nan", "none"])]

    if selected_majors:
        selected = [m.upper().strip() for m in selected_majors if str(m).strip()]
        if selected:
            cleaned = cleaned[cleaned["Major"].isin(selected)]

    # Fill friendly defaults only after filtering.
    for col in ["Job Role", "Job Title", "Functional Area", "Industry", "State", "Class of"]:
        cleaned[col] = cleaned[col].replace({"nan": "", "None": ""}).fillna("")
        if col in ["Job Role", "Job Title", "Functional Area", "Industry"]:
            cleaned[col] = cleaned[col].replace("", "Unspecified")

    cleaned = cleaned[PLACEMENT_DETAIL_COLUMNS].copy()
    return cleaned.reset_index(drop=True), mapping, missing


def tier_thresholds(total_placements: int) -> Tuple[int, int]:
    tier1 = max(3, round(total_placements * 0.0125))
    tier2 = max(2, round(total_placements * 0.0050))
    if tier2 >= tier1:
        tier2 = max(2, tier1 - 1)
    return tier1, tier2


def make_company_targets(cleaned: pd.DataFrame) -> pd.DataFrame:
    total_placements = len(cleaned)
    tier1_threshold, tier2_threshold = tier_thresholds(total_placements)
    rows = []

    for company, group in cleaned.groupby("Company", dropna=False):
        majors = group["Major"].fillna("").str.upper()
        bsis = int((majors == "BSIS").sum())
        mism = int((majors == "MISM").sum())
        other = int(len(group) - bsis - mism)
        if bsis and mism:
            major_mix = "BSIS + MISM"
        elif bsis:
            major_mix = "BSIS only"
        elif mism:
            major_mix = "MISM only"
        elif other:
            major_mix = "Other"
        else:
            major_mix = "Unspecified"

        total = int(len(group))
        if total >= tier1_threshold:
            tier = "Tier 1 — Core employer"
            priority = "High"
        elif total >= tier2_threshold:
            tier = "Tier 2 — Relationship employer"
            priority = "Medium"
        else:
            tier = "Tier 3 — Emerging employer"
            priority = "Monitor"

        start_dates = pd.to_datetime(group["Start Date"], errors="coerce")
        rows.append({
            "Company": company,
            "Total Placements": total,
            "BSIS": bsis,
            "MISM": mism,
            "Major Mix": major_mix,
            "Class Year(s)": join_unique_sorted(group["Class of"]),
            "Primary Industry": top_value(group["Industry"]),
            "Primary Functional Area": top_value(group["Functional Area"]),
            "Job Role(s)": join_top(group["Job Role"], limit=6),
            "Representative Job Title(s)": join_top(group["Job Title"], limit=6),
            "State(s)": join_unique_sorted(group["State"], limit=8),
            "First Start Date": start_dates.min() if start_dates.notna().any() else pd.NaT,
            "Latest Start Date": start_dates.max() if start_dates.notna().any() else pd.NaT,
            "Employer Tier": tier,
            "Recruiting Priority": priority,
        })

    targets = pd.DataFrame(rows)
    if targets.empty:
        return pd.DataFrame(columns=CORE_COLUMNS)
    targets = targets.sort_values(["Total Placements", "Company"], ascending=[False, True]).reset_index(drop=True)
    targets.insert(0, "Rank", range(1, len(targets) + 1))
    return targets[CORE_COLUMNS]


def make_summary_tables(cleaned: pd.DataFrame, company_targets: pd.DataFrame) -> Dict[str, pd.DataFrame]:
    def vc_frame(series: pd.Series, name: str, count_name: str = "Placements", limit: Optional[int] = None) -> pd.DataFrame:
        vc = series.fillna("Unspecified").replace("", "Unspecified").value_counts(dropna=False)
        if limit:
            vc = vc.head(limit)
        frame = vc.rename_axis(name).reset_index(name=count_name)
        return frame

    summary = {
        "Top Companies": company_targets[["Company", "Total Placements", "BSIS", "MISM", "Employer Tier", "Recruiting Priority"]].head(20).copy(),
        "Major Mix": vc_frame(cleaned["Major"], "Major"),
        "Industry Mix": vc_frame(cleaned["Industry"], "Industry", limit=15),
        "Functional Area Mix": vc_frame(cleaned["Functional Area"], "Functional Area", limit=15),
        "Job Role Mix": vc_frame(cleaned["Job Role"], "Job Role", limit=15),
        "State Mix": vc_frame(cleaned["State"], "State", limit=15),
        "Class Year Mix": vc_frame(cleaned["Class of"], "Class of", limit=15),
        "Employer Tier Mix": vc_frame(company_targets["Employer Tier"], "Employer Tier", count_name="Companies"),
    }
    return summary


def prepare_contact_data(contact_df: Optional[pd.DataFrame]) -> Optional[pd.DataFrame]:
    if contact_df is None or contact_df.empty:
        return None
    mapping = infer_columns(contact_df, CONTACT_ALIASES)
    if "company" not in mapping:
        return None

    out = pd.DataFrame()
    out["Company"] = contact_df[mapping["company"]].apply(lambda x: "" if pd.isna(x) else str(x).strip())
    for key, out_col in [
        ("handshake", "Handshake Link"),
        ("website", "Website"),
        ("linkedin", "LinkedIn"),
        ("owner", "Employer Owner"),
        ("css", "CSS Assigned"),
        ("primary_contact", "Primary Contact"),
        ("main_contact", "Main Contact"),
        ("engagement_status", "Engagement Status"),
        ("location", "Contact Location"),
        ("employee_count", "Employee Count"),
    ]:
        if key in mapping:
            out[out_col] = contact_df[mapping[key]].apply(lambda x: "" if pd.isna(x) else str(x).strip())
        else:
            out[out_col] = ""
    out["_norm_company"] = out["Company"].map(normalize_company)
    out = out[out["_norm_company"] != ""].drop_duplicates("_norm_company", keep="first")
    return out


def enrich_with_contacts(company_targets: pd.DataFrame, contact_df: Optional[pd.DataFrame]) -> pd.DataFrame:
    contacts = prepare_contact_data(contact_df)
    if contacts is None or contacts.empty:
        return company_targets

    contact_by_norm = contacts.set_index("_norm_company")
    contact_norms = list(contact_by_norm.index)
    contact_cols = [col for col in CONTACT_COLUMNS if col != "Contact Match Method"]

    enriched_rows = []
    for _, row in company_targets.iterrows():
        norm = normalize_company(row["Company"])
        match_method = ""
        matched = None
        if norm in contact_by_norm.index:
            matched = contact_by_norm.loc[norm]
            match_method = "Exact normalized match"
        else:
            best_score = 0.0
            best_norm = None
            for candidate in contact_norms:
                score = SequenceMatcher(None, norm, candidate).ratio()
                if score > best_score:
                    best_score = score
                    best_norm = candidate
            if best_score >= 0.92 and best_norm:
                matched = contact_by_norm.loc[best_norm]
                match_method = f"Fuzzy match ({best_score:.0%})"

        new_row = row.to_dict()
        if matched is not None:
            for col in contact_cols:
                new_row[col] = matched.get(col, "")
            new_row["Contact Match Method"] = match_method
        else:
            for col in contact_cols:
                new_row[col] = ""
            new_row["Contact Match Method"] = ""
        enriched_rows.append(new_row)

    return pd.DataFrame(enriched_rows)


def build_report(
    placement_df: pd.DataFrame,
    report_title: str = "IS/MISM Employer Recruiting Report",
    scope_label: str = "5-Year View",
    selected_majors: Optional[List[str]] = None,
    contact_df: Optional[pd.DataFrame] = None,
) -> Tuple[bytes, Dict[str, int], pd.DataFrame, pd.DataFrame]:
    cleaned, _, _ = clean_placement_data(placement_df, selected_majors=selected_majors)
    company_targets = make_company_targets(cleaned)
    company_targets = enrich_with_contacts(company_targets, contact_df)
    summary = make_summary_tables(cleaned, company_targets)

    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter", datetime_format="m/d/yyyy", date_format="m/d/yyyy") as writer:
        workbook = writer.book

        # Palette aligned with the report style.
        navy = "#17365D"
        blue = "#1F4E78"
        teal = "#0F766E"
        light_blue = "#D9EAF7"
        pale_blue = "#EAF3F8"
        gray = "#F3F6F8"
        dark_gray = "#404040"
        white = "#FFFFFF"
        green = "#D9EAD3"
        yellow = "#FFF2CC"
        orange = "#FCE4D6"

        fmt_title = workbook.add_format({"bold": True, "font_size": 18, "font_color": white, "bg_color": navy, "align": "left", "valign": "vcenter"})
        fmt_subtitle = workbook.add_format({"font_size": 10, "font_color": dark_gray, "italic": True})
        fmt_section = workbook.add_format({"bold": True, "font_size": 12, "font_color": white, "bg_color": blue, "align": "left"})
        fmt_header = workbook.add_format({"bold": True, "font_color": white, "bg_color": blue, "border": 1, "border_color": "#B7B7B7", "align": "center", "valign": "vcenter", "text_wrap": True})
        fmt_text = workbook.add_format({"border": 1, "border_color": "#D9E2F3", "valign": "top", "text_wrap": True})
        fmt_num = workbook.add_format({"border": 1, "border_color": "#D9E2F3", "valign": "top", "num_format": "#,##0"})
        fmt_date = workbook.add_format({"border": 1, "border_color": "#D9E2F3", "valign": "top", "num_format": "m/d/yyyy"})
        fmt_kpi_label = workbook.add_format({"font_size": 9, "font_color": dark_gray, "align": "center", "valign": "vcenter", "bg_color": pale_blue})
        fmt_kpi_value = workbook.add_format({"bold": True, "font_size": 18, "font_color": navy, "align": "center", "valign": "vcenter", "bg_color": pale_blue})
        fmt_note = workbook.add_format({"font_size": 9, "font_color": dark_gray, "text_wrap": True, "valign": "top"})
        fmt_high = workbook.add_format({"bg_color": green, "font_color": "#274E13", "border": 1, "border_color": "#D9E2F3"})
        fmt_medium = workbook.add_format({"bg_color": yellow, "font_color": "#7F6000", "border": 1, "border_color": "#D9E2F3"})
        fmt_monitor = workbook.add_format({"bg_color": orange, "font_color": "#7F2A00", "border": 1, "border_color": "#D9E2F3"})
        fmt_link = workbook.add_format({"font_color": "#0563C1", "underline": 1, "border": 1, "border_color": "#D9E2F3", "text_wrap": True})

        # Sheets.
        dash = workbook.add_worksheet("Executive Dashboard")
        targets_ws = workbook.add_worksheet("Company Targets")
        summary_ws = workbook.add_worksheet("Summary Tables")
        detail_ws = workbook.add_worksheet("Placement Detail")
        writer.sheets["Executive Dashboard"] = dash
        writer.sheets["Company Targets"] = targets_ws
        writer.sheets["Summary Tables"] = summary_ws
        writer.sheets["Placement Detail"] = detail_ws

        for ws in [dash, targets_ws, summary_ws, detail_ws]:
            ws.hide_gridlines(2)
            ws.set_tab_color(blue)

        # Executive Dashboard.
        dash.set_column("A:A", 18)
        dash.set_column("B:C", 14)
        dash.set_column("D:D", 3)
        dash.set_column("E:G", 15)
        dash.set_column("H:H", 3)
        dash.set_column("I:K", 15)
        dash.set_column("L:M", 15)
        dash.set_row(0, 28)
        dash.merge_range("A1:M1", f"{report_title} — {scope_label}", fmt_title)
        dash.merge_range("A2:M2", f"Scope: {scope_label.lower()} placement file. Company focus; recruiter names/emails/phones are included only when provided in the upload/contact file.", fmt_subtitle)

        metrics = {
            "Placements": len(cleaned),
            "Unique Companies": int(company_targets["Company"].nunique()) if not company_targets.empty else 0,
            "BSIS Placements": int((cleaned["Major"] == "BSIS").sum()),
            "MISM Placements": int((cleaned["Major"] == "MISM").sum()),
            "Tier 1 Employers": int((company_targets["Employer Tier"] == "Tier 1 — Core employer").sum()) if not company_targets.empty else 0,
            "Tier 2 Employers": int((company_targets["Employer Tier"] == "Tier 2 — Relationship employer").sum()) if not company_targets.empty else 0,
        }
        kpi_cells = [("A4:B5", "Placements"), ("C4:D5", "Unique Companies"), ("E4:F5", "BSIS Placements"), ("G4:H5", "MISM Placements"), ("I4:J5", "Tier 1 Employers"), ("K4:M5", "Tier 2 Employers")]
        for cell_range, label in kpi_cells:
            dash.merge_range(cell_range, metrics[label], fmt_kpi_value)
            first_cell = cell_range.split(":")[0]
            row = int(re.findall(r"\d+", first_cell)[0]) + 1
            col_letters = re.findall(r"[A-Z]+", first_cell)[0]
            # Label in row 6 below the card.
            # We'll map manually because XlsxWriter row/col addressing is easier.
        labels_positions = [(5, 0, 5, 1, "Placements"), (5, 2, 5, 3, "Unique Companies"), (5, 4, 5, 5, "BSIS Placements"), (5, 6, 5, 7, "MISM Placements"), (5, 8, 5, 9, "Tier 1 Employers"), (5, 10, 5, 12, "Tier 2 Employers")]
        for r1, c1, r2, c2, label in labels_positions:
            dash.merge_range(r1, c1, r2, c2, label, fmt_kpi_label)

        dash.merge_range("A8:C8", "Top Companies by Placements", fmt_section)
        dash.merge_range("E8:G8", "Placement Mix by Industry", fmt_section)
        dash.merge_range("I8:K8", "Functional Area Mix", fmt_section)
        dash.merge_range("A31:C31", "Major Mix", fmt_section)
        dash.merge_range("E31:G31", "Employer Tier Mix", fmt_section)
        dash.merge_range("I31:M31", "Director Notes", fmt_section)
        dash.merge_range("I32:M37", "Use the Company Targets sheet as the working list for outreach. Tier 1 employers represent the highest placement volume in the uploaded file; Tier 2 employers are relationship-building targets; Tier 3 employers should be monitored for emerging recruiting potential.", fmt_note)

        # Summary Tables sheet.
        summary_ws.set_column("A:A", 28)
        summary_ws.set_column("B:K", 16)
        summary_ws.merge_range("A1:K1", f"Summary Tables — {scope_label} IS/MISM Placements", fmt_title)
        summary_ws.write("A2", "These tables feed the dashboard charts and provide quick drilldowns for director review.", fmt_subtitle)

        def write_df(ws, df: pd.DataFrame, start_row: int, start_col: int, title: str) -> Tuple[int, int, int, int]:
            ws.write(start_row, start_col, title, fmt_section)
            for c, col in enumerate(df.columns):
                ws.write(start_row + 1, start_col + c, col, fmt_header)
            for r, (_, record) in enumerate(df.iterrows(), start=start_row + 2):
                for c, col in enumerate(df.columns):
                    value = record[col]
                    cell_fmt = fmt_num if isinstance(value, (int, float)) and not isinstance(value, bool) else fmt_text
                    if pd.isna(value):
                        value = ""
                    ws.write(r, start_col + c, value, cell_fmt)
            return (start_row, start_col, start_row + len(df) + 1, start_col + max(0, len(df.columns) - 1))

        positions = {}
        positions["Top Companies"] = write_df(summary_ws, summary["Top Companies"], 3, 0, "Top 20 Companies")
        positions["Major Mix"] = write_df(summary_ws, summary["Major Mix"], 3, 7, "Major Mix")
        positions["Industry Mix"] = write_df(summary_ws, summary["Industry Mix"], 28, 0, "Industry Mix")
        positions["Functional Area Mix"] = write_df(summary_ws, summary["Functional Area Mix"], 28, 4, "Functional Area Mix")
        positions["Employer Tier Mix"] = write_df(summary_ws, summary["Employer Tier Mix"], 28, 8, "Employer Tier Mix")
        positions["Job Role Mix"] = write_df(summary_ws, summary["Job Role Mix"], 48, 0, "Job Role Mix")
        positions["State Mix"] = write_df(summary_ws, summary["State Mix"], 48, 4, "State Mix")
        positions["Class Year Mix"] = write_df(summary_ws, summary["Class Year Mix"], 48, 8, "Class Year Mix")

        # Charts.
        def add_bar_chart(title: str, sheet_range: Tuple[int, int, int, int], name_col_offset: int, value_col_offset: int):
            r1, c1, r2, c2 = sheet_range
            chart = workbook.add_chart({"type": "bar"})
            chart.add_series({
                "name": title,
                "categories": ["Summary Tables", r1 + 2, c1 + name_col_offset, r2, c1 + name_col_offset],
                "values": ["Summary Tables", r1 + 2, c1 + value_col_offset, r2, c1 + value_col_offset],
                "data_labels": {"value": True},
            })
            chart.set_title({"name": title})
            chart.set_legend({"none": True})
            chart.set_style(10)
            chart.set_size({"width": 430, "height": 250})
            return chart

        def add_pie_chart(title: str, sheet_range: Tuple[int, int, int, int], name_col_offset: int, value_col_offset: int):
            r1, c1, r2, c2 = sheet_range
            chart = workbook.add_chart({"type": "pie"})
            chart.add_series({
                "name": title,
                "categories": ["Summary Tables", r1 + 2, c1 + name_col_offset, r2, c1 + name_col_offset],
                "values": ["Summary Tables", r1 + 2, c1 + value_col_offset, r2, c1 + value_col_offset],
                "data_labels": {"percentage": True},
            })
            chart.set_title({"name": title})
            chart.set_style(10)
            chart.set_size({"width": 390, "height": 250})
            return chart

        top_chart = add_bar_chart("Top Companies", positions["Top Companies"], 0, 1)
        industry_chart = add_pie_chart("Industry Mix", positions["Industry Mix"], 0, 1)
        function_chart = add_bar_chart("Functional Area Mix", positions["Functional Area Mix"], 0, 1)
        major_chart = add_pie_chart("Major Mix", positions["Major Mix"], 0, 1)
        tier_chart = add_pie_chart("Employer Tier Mix", positions["Employer Tier Mix"], 0, 1)
        dash.insert_chart("A9", top_chart)
        dash.insert_chart("E9", industry_chart)
        dash.insert_chart("I9", function_chart)
        dash.insert_chart("A32", major_chart)
        dash.insert_chart("E32", tier_chart)

        # Company Targets sheet.
        targets_ws.set_column("A:A", 8)
        targets_ws.set_column("B:B", 32)
        targets_ws.set_column("C:E", 12)
        targets_ws.set_column("F:F", 16)
        targets_ws.set_column("G:G", 20)
        targets_ws.set_column("H:J", 22)
        targets_ws.set_column("K:K", 42)
        targets_ws.set_column("L:L", 28)
        targets_ws.set_column("M:N", 14)
        targets_ws.set_column("O:P", 24)
        if any(col in company_targets.columns for col in CONTACT_COLUMNS):
            targets_ws.set_column("Q:AA", 24)
        targets_ws.freeze_panes(1, 2)

        # Write data and add table.
        for col_idx, col in enumerate(company_targets.columns):
            targets_ws.write(0, col_idx, col, fmt_header)
        for row_idx, (_, record) in enumerate(company_targets.iterrows(), start=1):
            for col_idx, col in enumerate(company_targets.columns):
                value = record[col]
                if pd.isna(value):
                    value = ""
                if col in ["First Start Date", "Latest Start Date"] and value != "":
                    try:
                        value = pd.to_datetime(value).to_pydatetime()
                        targets_ws.write_datetime(row_idx, col_idx, value, fmt_date)
                    except Exception:
                        targets_ws.write(row_idx, col_idx, str(value), fmt_text)
                elif col in ["Rank", "Total Placements", "BSIS", "MISM", "Employee Count"]:
                    try:
                        targets_ws.write_number(row_idx, col_idx, float(value), fmt_num)
                    except Exception:
                        targets_ws.write(row_idx, col_idx, value, fmt_text)
                elif col in ["Handshake Link", "Website", "LinkedIn"] and str(value).startswith("http"):
                    targets_ws.write_url(row_idx, col_idx, str(value), fmt_link, string=str(value))
                elif col == "Recruiting Priority":
                    if value == "High":
                        targets_ws.write(row_idx, col_idx, value, fmt_high)
                    elif value == "Medium":
                        targets_ws.write(row_idx, col_idx, value, fmt_medium)
                    else:
                        targets_ws.write(row_idx, col_idx, value, fmt_monitor)
                else:
                    targets_ws.write(row_idx, col_idx, value, fmt_text)
        if len(company_targets) > 0:
            targets_ws.add_table(0, 0, len(company_targets), len(company_targets.columns) - 1, {
                "name": "CompanyTargets",
                "columns": [{"header": col} for col in company_targets.columns],
                "style": "Table Style Medium 2",
            })
            priority_col = list(company_targets.columns).index("Recruiting Priority")
            targets_ws.data_validation(1, priority_col, len(company_targets), priority_col, {"validate": "list", "source": ["High", "Medium", "Monitor"]})

        # Placement Detail sheet.
        detail_ws.set_column("A:A", 32)
        detail_ws.set_column("B:B", 12)
        detail_ws.set_column("C:F", 22)
        detail_ws.set_column("G:G", 24)
        detail_ws.set_column("H:H", 14)
        detail_ws.set_column("I:J", 16)
        detail_ws.freeze_panes(1, 0)
        for col_idx, col in enumerate(cleaned.columns):
            detail_ws.write(0, col_idx, col, fmt_header)
        for row_idx, (_, record) in enumerate(cleaned.iterrows(), start=1):
            for col_idx, col in enumerate(cleaned.columns):
                value = record[col]
                if pd.isna(value):
                    value = ""
                if col == "Start Date" and value != "":
                    try:
                        value = pd.to_datetime(value).to_pydatetime()
                        detail_ws.write_datetime(row_idx, col_idx, value, fmt_date)
                    except Exception:
                        detail_ws.write(row_idx, col_idx, str(value), fmt_text)
                else:
                    detail_ws.write(row_idx, col_idx, value, fmt_text)
        if len(cleaned) > 0:
            detail_ws.add_table(0, 0, len(cleaned), len(cleaned.columns) - 1, {
                "name": "PlacementDetail",
                "columns": [{"header": col} for col in cleaned.columns],
                "style": "Table Style Medium 2",
            })

        # Footer/data notes in Summary Tables.
        note_row = 68
        summary_ws.merge_range(note_row, 0, note_row, 10, "Data Notes", fmt_section)
        notes = [
            "Source: uploaded Excel placement file.",
            "Company Targets are grouped by company name after basic text normalization. Review unusual naming variants manually if needed.",
            "Employer tiers are dynamic: Tier 1 and Tier 2 thresholds scale with total placement volume so past-year and 5-year files remain comparable.",
            "Recruiter names, emails, and phone numbers are not generated. Contact fields appear only when a contact/enrichment workbook is uploaded and matched.",
        ]
        for i, note in enumerate(notes, start=note_row + 1):
            summary_ws.merge_range(i, 0, i, 10, note, fmt_note)

        dash.set_landscape()
        dash.fit_to_pages(1, 1)
        targets_ws.set_landscape()
        targets_ws.fit_to_pages(1, 0)
        summary_ws.set_landscape()
        summary_ws.fit_to_pages(1, 0)
        detail_ws.set_landscape()
        detail_ws.fit_to_pages(1, 0)

    output.seek(0)
    return output.getvalue(), metrics, company_targets, cleaned


def read_excel_sheets(file_bytes: bytes) -> List[str]:
    with pd.ExcelFile(BytesIO(file_bytes)) as xls:
        return xls.sheet_names


def read_excel(file_bytes: bytes, sheet_name: str) -> pd.DataFrame:
    return pd.read_excel(BytesIO(file_bytes), sheet_name=sheet_name)
