from __future__ import annotations

import re
import zipfile
from collections import Counter
from datetime import datetime
from difflib import SequenceMatcher
from io import BytesIO
from pathlib import PurePosixPath
from typing import Dict, Iterable, List, Optional, Tuple

import pandas as pd

BASE_TARGET_COLUMNS = [
    "Rank",
    "Company",
    "Total Placements",
]

POST_MAJOR_TARGET_COLUMNS = [
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
    "Main Contact",
    "Primary Contact",
    "Recruiter Email",
    "Recruiter Phone",
    "Engagement Status",
    "Outreach Status",
    "Last Contacted Date",
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
    "Source Program",
    "Source File",
]

COLUMN_ALIASES = {
    "company": [
        "company", "company name", "employer", "employer name", "organization", "organization name",
        "account name", "company/account", "company / account", "employer/company", "hiring company",
    ],
    "major": [
        "major", "majors", "program", "programs", "degree", "academic program", "student major",
        "relevant programs", "source program", "major/program", "major program",
    ],
    "job_role": ["job role", "job roles", "role", "position category", "job category", "job family"],
    "job_title": ["job title", "title", "position", "offer title", "position title"],
    "class_year": ["class of", "class year", "graduation year", "grad year", "year", "graduating class"],
    "functional_area": ["functional area", "function", "career function", "business function", "job function"],
    "industry": ["industry", "industry name", "sector", "employer industry", "company industry"],
    "start_date": ["start date", "job start date", "offer start date", "hire date", "employment start date", "date posted", "apply start date"],
    "state": ["state", "state/province", "region", "location state", "work state", "province", "job location"],
    "student_id": ["student id", "student id.id", "student", "student identifier", "net id", "byu id", "studentid", "studentid.id"],
}

CONTACT_ALIASES = {
    "company": ["company", "company name", "employer", "employer name", "employer name", "account name", "organization", "name"],
    "handshake": [
        "handshake", "handshake link", "handshake url", "handshake employer link",
        "employer handshake link", "handshake profile", "handshake link old", "handshake link (old)",
    ],
    "website": ["website", "company website", "web site", "url", "employer website", "company website"],
    "linkedin": ["linkedin", "linkedin url", "linkedin link", "company linkedin", "linkedin profile"],
    "owner": ["owner", "account owner", "employer owner", "record owner"],
    "css": ["css", "css assigned", "assigned css", "css assigned to", "career success specialist"],
    "primary_contact": ["primary contact", "primary recruiter", "recruiter", "contact", "contact name"],
    "main_contact": ["main contact", "main recruiter", "main point of contact", "poc", "main point of contact"],
    "email": ["email", "contact email", "recruiter email", "primary contact email", "main contact email"],
    "phone": ["phone", "phone number", "contact phone", "recruiter phone", "primary contact phone"],
    "engagement_status": ["engagement status", "ce engagement status", "status", "employer status", "recent contact status"],
    "outreach_status": ["outreach status", "recent outreach status"],
    "last_contacted": ["last contacted date", "last contact date", "last activity time", "last expressed engagement"],
    "location": ["location", "headquarters", "hq", "city", "address", "company location(s)", "state", "country"],
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


def normalize_major(value: object) -> str:
    text = str(value or "").strip()
    if not text or text.lower() in {"nan", "none", "null"}:
        return "UNSPECIFIED"
    # Keep compact program codes like BSFin, MAcc, MBA readable and consistent.
    text = re.sub(r"\s+", " ", text)
    return text.upper() if len(text) <= 12 else text.strip()


def first_present(values: Iterable[object]) -> str:
    for value in values:
        if pd.notna(value) and str(value).strip():
            return str(value).strip()
    return ""


def top_value(values: Iterable[object], default: str = "Unspecified") -> str:
    cleaned = [str(v).strip() for v in values if pd.notna(v) and str(v).strip() and str(v).strip().lower() not in {"nan", "none"}]
    if not cleaned:
        return default
    return Counter(cleaned).most_common(1)[0][0]


def join_top(values: Iterable[object], limit: int = 6) -> str:
    cleaned = [str(v).strip() for v in values if pd.notna(v) and str(v).strip() and str(v).strip().lower() not in {"nan", "none"}]
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
        if text.lower() in {"nan", "none"}:
            continue
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

    for canonical, candidates in aliases.items():
        for candidate in candidates:
            norm = normalize_header(candidate)
            if norm in normalized_lookup:
                result[canonical] = normalized_lookup[norm]
                break

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


def read_excel_sheets(file_bytes: bytes) -> List[str]:
    with pd.ExcelFile(BytesIO(file_bytes)) as xls:
        return xls.sheet_names


def preferred_sheet_name(file_bytes: bytes) -> str:
    sheets = read_excel_sheets(file_bytes)
    lowered = [s.lower() for s in sheets]
    for preferred in ["sheet0", "placement detail", "placements", "data"]:
        if preferred in lowered:
            return sheets[lowered.index(preferred)]
    return sheets[0]


def read_excel(file_bytes: bytes, sheet_name: Optional[str] = None) -> pd.DataFrame:
    sheet = sheet_name or preferred_sheet_name(file_bytes)
    return pd.read_excel(BytesIO(file_bytes), sheet_name=sheet)


def excel_files_from_zip(zip_bytes: bytes) -> List[Dict[str, object]]:
    files: List[Dict[str, object]] = []
    with zipfile.ZipFile(BytesIO(zip_bytes)) as zf:
        for info in zf.infolist():
            path = PurePosixPath(info.filename)
            name = path.name
            if info.is_dir() or name.startswith("._") or name == ".DS_Store" or "__MACOSX" in path.parts:
                continue
            if name.lower().endswith((".xlsx", ".xls")):
                folder_parts = [p for p in path.parts[:-1] if p not in {"", "."}]
                source_program = folder_parts[-1] if folder_parts else "Uploaded ZIP"
                files.append({
                    "path": info.filename,
                    "file_name": name,
                    "source_program": source_program,
                    "bytes": zf.read(info),
                })
    return files


def read_zip_placements(zip_bytes: bytes) -> Tuple[pd.DataFrame, pd.DataFrame]:
    excel_files = excel_files_from_zip(zip_bytes)
    if not excel_files:
        raise ValueError("No Excel files were found inside the uploaded ZIP.")

    frames = []
    manifest_rows = []
    for item in excel_files:
        try:
            sheet = preferred_sheet_name(item["bytes"])
            df = read_excel(item["bytes"], sheet)
            df["Source Program"] = item["source_program"]
            df["Source File"] = item["file_name"]
            df["Source Path"] = item["path"]
            frames.append(df)
            manifest_rows.append({
                "Source Program": item["source_program"],
                "Source File": item["file_name"],
                "Sheet Used": sheet,
                "Rows Imported": len(df),
                "Status": "Imported",
            })
        except Exception as exc:
            manifest_rows.append({
                "Source Program": item.get("source_program", ""),
                "Source File": item.get("file_name", ""),
                "Sheet Used": "",
                "Rows Imported": 0,
                "Status": f"Skipped: {exc}",
            })
    if not frames:
        raise ValueError("Excel files were found in the ZIP, but none could be imported.")
    return pd.concat(frames, ignore_index=True), pd.DataFrame(manifest_rows)


def detect_available_majors(df: pd.DataFrame) -> List[str]:
    mapping = infer_columns(df, COLUMN_ALIASES)
    if "major" in mapping:
        majors = df[mapping["major"]].dropna().map(normalize_major)
    elif "Source Program" in df.columns:
        majors = df["Source Program"].dropna().map(normalize_major)
    else:
        return []
    majors = [m for m in majors if m and m != "UNSPECIFIED"]
    return sorted(set(majors), key=lambda x: (x.lower()))


def clean_placement_data(
    df: pd.DataFrame,
    selected_majors: Optional[List[str]] = None,
    default_major: Optional[str] = None,
) -> Tuple[pd.DataFrame, Dict[str, str], List[str]]:
    mapping = infer_columns(df, COLUMN_ALIASES)
    missing = []
    if "company" not in mapping:
        missing.append("company")
    if "major" not in mapping and default_major is None and "Source Program" not in df.columns:
        missing.append("major")
    if missing:
        raise ValueError(
            "Could not identify required column(s): " + ", ".join(missing) +
            ". Make sure the upload includes company/employer and major/program fields."
        )

    cleaned = pd.DataFrame()
    cleaned["Company"] = df[mapping["company"]].apply(lambda x: "" if pd.isna(x) else str(x).strip())
    if "major" in mapping:
        cleaned["Major"] = df[mapping["major"]].apply(normalize_major)
    elif default_major:
        cleaned["Major"] = normalize_major(default_major)
    else:
        cleaned["Major"] = df["Source Program"].apply(normalize_major)

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

    cleaned["Source Program"] = df["Source Program"].apply(lambda x: "" if pd.isna(x) else str(x).strip()) if "Source Program" in df.columns else ""
    cleaned["Source File"] = df["Source File"].apply(lambda x: "" if pd.isna(x) else str(x).strip()) if "Source File" in df.columns else ""

    cleaned = cleaned[cleaned["Company"].notna() & (cleaned["Company"].str.strip() != "")]
    cleaned = cleaned[~cleaned["Company"].str.lower().isin(["nan", "none"])]

    if selected_majors:
        selected = {normalize_major(m) for m in selected_majors if str(m).strip()}
        if selected:
            cleaned = cleaned[cleaned["Major"].isin(selected)]

    for col in ["Job Role", "Job Title", "Functional Area", "Industry", "State", "Class of"]:
        cleaned[col] = cleaned[col].replace({"nan": "", "None": "", "NaT": ""}).fillna("")
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


def major_columns(cleaned: pd.DataFrame) -> List[str]:
    counts = cleaned["Major"].fillna("UNSPECIFIED").value_counts()
    return [m for m in counts.index.tolist() if m != "UNSPECIFIED"] + (["UNSPECIFIED"] if "UNSPECIFIED" in counts.index else [])


def format_major_mix(counts: Dict[str, int]) -> str:
    positive = [(major, count) for major, count in counts.items() if count > 0]
    if not positive:
        return "Unspecified"
    positive.sort(key=lambda x: (-x[1], x[0]))
    if len(positive) == 1:
        return f"{positive[0][0]} only"
    top = [m for m, _ in positive[:3]]
    if len(positive) <= 3:
        return " + ".join(top)
    return ", ".join(top) + f" + {len(positive) - 3} more"


def make_company_targets(cleaned: pd.DataFrame) -> Tuple[pd.DataFrame, List[str]]:
    total_placements = len(cleaned)
    tier1_threshold, tier2_threshold = tier_thresholds(total_placements)
    majors = major_columns(cleaned)
    rows = []

    for company, group in cleaned.groupby("Company", dropna=False):
        major_counts = {major: int((group["Major"] == major).sum()) for major in majors}
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
        row = {
            "Company": company,
            "Total Placements": total,
            **major_counts,
            "Major Mix": format_major_mix(major_counts),
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
        }
        rows.append(row)

    target_columns = BASE_TARGET_COLUMNS + majors + POST_MAJOR_TARGET_COLUMNS
    targets = pd.DataFrame(rows)
    if targets.empty:
        return pd.DataFrame(columns=target_columns), majors
    targets = targets.sort_values(["Total Placements", "Company"], ascending=[False, True]).reset_index(drop=True)
    targets.insert(0, "Rank", range(1, len(targets) + 1))
    return targets[target_columns], majors


def make_summary_tables(cleaned: pd.DataFrame, company_targets: pd.DataFrame, majors: List[str]) -> Dict[str, pd.DataFrame]:
    def vc_frame(series: pd.Series, name: str, count_name: str = "Placements", limit: Optional[int] = None) -> pd.DataFrame:
        vc = series.fillna("Unspecified").replace("", "Unspecified").value_counts(dropna=False)
        if limit:
            vc = vc.head(limit)
        return vc.rename_axis(name).reset_index(name=count_name)

    top_major_cols = [m for m in majors if m in company_targets.columns][:6]
    top_cols = ["Company", "Total Placements"] + top_major_cols + ["Employer Tier", "Recruiting Priority"]
    summary = {
        "Top Companies": company_targets[top_cols].head(20).copy() if not company_targets.empty else pd.DataFrame(columns=top_cols),
        "Major Mix": vc_frame(cleaned["Major"], "Major"),
        "Industry Mix": vc_frame(cleaned["Industry"], "Industry", limit=15),
        "Functional Area Mix": vc_frame(cleaned["Functional Area"], "Functional Area", limit=15),
        "Job Role Mix": vc_frame(cleaned["Job Role"], "Job Role", limit=15),
        "State Mix": vc_frame(cleaned["State"], "State", limit=15),
        "Class Year Mix": vc_frame(cleaned["Class of"], "Class of", limit=15),
        "Employer Tier Mix": vc_frame(company_targets["Employer Tier"], "Employer Tier", count_name="Companies") if not company_targets.empty else pd.DataFrame(columns=["Employer Tier", "Companies"]),
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

    field_map = [
        ("handshake", "Handshake Link"),
        ("website", "Website"),
        ("linkedin", "LinkedIn"),
        ("owner", "Employer Owner"),
        ("css", "CSS Assigned"),
        ("main_contact", "Main Contact"),
        ("primary_contact", "Primary Contact"),
        ("email", "Recruiter Email"),
        ("phone", "Recruiter Phone"),
        ("engagement_status", "Engagement Status"),
        ("outreach_status", "Outreach Status"),
        ("last_contacted", "Last Contacted Date"),
        ("location", "Contact Location"),
        ("employee_count", "Employee Count"),
    ]
    for key, out_col in field_map:
        if key in mapping:
            out[out_col] = contact_df[mapping[key]].apply(lambda x: "" if pd.isna(x) else str(x).strip())
        else:
            out[out_col] = ""

    # Build a cleaner location if City/State/Country exist separately.
    norm_cols = {normalize_header(c): c for c in contact_df.columns}
    loc_parts = []
    for key in ["city", "state", "country"]:
        if key in norm_cols:
            loc_parts.append(norm_cols[key])
    if loc_parts:
        location = contact_df[loc_parts].apply(lambda r: ", ".join([str(v).strip() for v in r if pd.notna(v) and str(v).strip()]), axis=1)
        out["Contact Location"] = out["Contact Location"].where(out["Contact Location"].astype(str).str.strip() != "", location)

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


def safe_table_name(name: str) -> str:
    safe = re.sub(r"[^A-Za-z0-9_]", "_", name)
    if not safe or safe[0].isdigit():
        safe = "T_" + safe
    return safe[:250]


def write_report_workbook(
    cleaned: pd.DataFrame,
    company_targets: pd.DataFrame,
    summary: Dict[str, pd.DataFrame],
    majors: List[str],
    metrics: Dict[str, object],
    report_title: str,
    scope_label: str,
) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter", datetime_format="m/d/yyyy", date_format="m/d/yyyy") as writer:
        workbook = writer.book

        navy = "#17365D"
        blue = "#1F4E78"
        pale_blue = "#EAF3F8"
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

        dash = workbook.add_worksheet("Executive Dashboard")
        targets_ws = workbook.add_worksheet("Company Targets")
        summary_ws = workbook.add_worksheet("Summary Tables")
        detail_ws = workbook.add_worksheet("Placement Detail")
        writer.sheets.update({
            "Executive Dashboard": dash,
            "Company Targets": targets_ws,
            "Summary Tables": summary_ws,
            "Placement Detail": detail_ws,
        })

        for ws in [dash, targets_ws, summary_ws, detail_ws]:
            ws.hide_gridlines(2)
            ws.set_tab_color(blue)

        dash.set_column("A:A", 18)
        dash.set_column("B:C", 14)
        dash.set_column("D:D", 3)
        dash.set_column("E:G", 15)
        dash.set_column("H:H", 3)
        dash.set_column("I:K", 15)
        dash.set_column("L:M", 15)
        dash.set_row(0, 28)
        dash.merge_range("A1:M1", f"{report_title} — {scope_label}", fmt_title)
        major_caption = ", ".join(majors[:8]) + (f" + {len(majors) - 8} more" if len(majors) > 8 else "")
        dash.merge_range("A2:M2", f"Scope: {scope_label}. Major/program coverage: {major_caption or 'All detected programs'}. Company focus; contact fields are included only when provided in the upload/contact file.", fmt_subtitle)

        kpi_items = [
            ("Placements", metrics.get("Placements", 0)),
            ("Unique Companies", metrics.get("Unique Companies", 0)),
            ("Majors / Programs", metrics.get("Majors / Programs", 0)),
            ("Top Major", metrics.get("Top Major", "—")),
            ("Tier 1 Employers", metrics.get("Tier 1 Employers", 0)),
            ("Tier 2 Employers", metrics.get("Tier 2 Employers", 0)),
        ]
        kpi_ranges = ["A4:B5", "C4:D5", "E4:F5", "G4:H5", "I4:J5", "K4:M5"]
        label_positions = [(5, 0, 5, 1), (5, 2, 5, 3), (5, 4, 5, 5), (5, 6, 5, 7), (5, 8, 5, 9), (5, 10, 5, 12)]
        for (label, value), cell_range, (r1, c1, r2, c2) in zip(kpi_items, kpi_ranges, label_positions):
            dash.merge_range(cell_range, value, fmt_kpi_value)
            dash.merge_range(r1, c1, r2, c2, label, fmt_kpi_label)

        dash.merge_range("A8:C8", "Top Companies by Placements", fmt_section)
        dash.merge_range("E8:G8", "Placement Mix by Industry", fmt_section)
        dash.merge_range("I8:K8", "Functional Area Mix", fmt_section)
        dash.merge_range("A31:C31", "Major Mix", fmt_section)
        dash.merge_range("E31:G31", "Employer Tier Mix", fmt_section)
        dash.merge_range("I31:M31", "Director Notes", fmt_section)
        dash.merge_range("I32:M37", "Use the Company Targets sheet as the working list for outreach. Tier 1 employers represent the highest placement volume in the uploaded file; Tier 2 employers are relationship-building targets; Tier 3 employers should be monitored for emerging recruiting potential. The report works for one major, multiple majors, or a ZIP of major-specific placement exports.", fmt_note)

        summary_ws.set_column("A:A", 28)
        summary_ws.set_column("B:K", 16)
        summary_ws.merge_range("A1:K1", f"Summary Tables — {scope_label} Employer Placements", fmt_title)
        summary_ws.write("A2", "These tables feed the dashboard charts and provide quick drilldowns for director review.", fmt_subtitle)

        def write_df(ws, df: pd.DataFrame, start_row: int, start_col: int, title: str) -> Tuple[int, int, int, int]:
            ws.write(start_row, start_col, title, fmt_section)
            for c, col in enumerate(df.columns):
                ws.write(start_row + 1, start_col + c, col, fmt_header)
            for r, (_, record) in enumerate(df.iterrows(), start=start_row + 2):
                for c, col in enumerate(df.columns):
                    value = record[col]
                    if pd.isna(value):
                        value = ""
                    cell_fmt = fmt_num if isinstance(value, (int, float)) and not isinstance(value, bool) else fmt_text
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

        if not company_targets.empty:
            dash.insert_chart("A9", add_bar_chart("Top Companies", positions["Top Companies"], 0, 1))
            dash.insert_chart("E9", add_pie_chart("Industry Mix", positions["Industry Mix"], 0, 1))
            dash.insert_chart("I9", add_bar_chart("Functional Area Mix", positions["Functional Area Mix"], 0, 1))
            dash.insert_chart("A32", add_pie_chart("Major Mix", positions["Major Mix"], 0, 1))
            dash.insert_chart("E32", add_pie_chart("Employer Tier Mix", positions["Employer Tier Mix"], 0, 1))

        targets_ws.set_column("A:A", 8)
        targets_ws.set_column("B:B", 32)
        targets_ws.set_column("C:C", 14)
        if majors:
            start_major_col = 3  # D, zero-based after Rank/Company/Total Placements
            end_major_col = start_major_col + len(majors) - 1
            targets_ws.set_column(start_major_col, end_major_col, 11)
            post_start = end_major_col + 1
        else:
            post_start = 3
        targets_ws.set_column(post_start, post_start, 18)  # Major Mix
        targets_ws.set_column(post_start + 1, post_start + 1, 20)  # Class years
        targets_ws.set_column(post_start + 2, post_start + 4, 22)
        targets_ws.set_column(post_start + 5, post_start + 5, 42)
        targets_ws.set_column(post_start + 6, post_start + 6, 28)
        targets_ws.set_column(post_start + 7, post_start + 8, 14)
        targets_ws.set_column(post_start + 9, post_start + 10, 24)
        if any(col in company_targets.columns for col in CONTACT_COLUMNS):
            contact_start = len(company_targets.columns) - len(CONTACT_COLUMNS)
            targets_ws.set_column(contact_start, len(company_targets.columns), 24)
        targets_ws.freeze_panes(1, 2)

        for col_idx, col in enumerate(company_targets.columns):
            targets_ws.write(0, col_idx, col, fmt_header)
        for row_idx, (_, record) in enumerate(company_targets.iterrows(), start=1):
            for col_idx, col in enumerate(company_targets.columns):
                value = record[col]
                if pd.isna(value):
                    value = ""
                if col in ["First Start Date", "Latest Start Date", "Last Contacted Date"] and value != "":
                    try:
                        value = pd.to_datetime(value).to_pydatetime()
                        targets_ws.write_datetime(row_idx, col_idx, value, fmt_date)
                    except Exception:
                        targets_ws.write(row_idx, col_idx, str(value), fmt_text)
                elif col in ["Rank", "Total Placements", "Employee Count"] + majors:
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

        detail_ws.set_column("A:A", 32)
        detail_ws.set_column("B:B", 14)
        detail_ws.set_column("C:F", 22)
        detail_ws.set_column("G:G", 24)
        detail_ws.set_column("H:H", 14)
        detail_ws.set_column("I:J", 16)
        detail_ws.set_column("K:L", 22)
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

        note_row = 68
        summary_ws.merge_range(note_row, 0, note_row, 10, "Data Notes", fmt_section)
        notes = [
            "Source: uploaded Excel placement file or ZIP of major-specific placement files.",
            "Company Targets are grouped by company name after basic text normalization. Review unusual naming variants manually if needed.",
            "Major columns are generated dynamically from the uploaded data; the app is not hardcoded to BSIS/MISM.",
            "Employer tiers are dynamic: Tier 1 and Tier 2 thresholds scale with total placement volume so past-year and five-year files remain comparable.",
            "Recruiter names, emails, and phone numbers are not generated. Contact fields appear only when a contact/enrichment workbook is uploaded and matched.",
        ]
        for i, note in enumerate(notes, start=note_row + 1):
            summary_ws.merge_range(i, 0, i, 10, note, fmt_note)

        dash.set_landscape(); dash.fit_to_pages(1, 1)
        targets_ws.set_landscape(); targets_ws.fit_to_pages(1, 0)
        summary_ws.set_landscape(); summary_ws.fit_to_pages(1, 0)
        detail_ws.set_landscape(); detail_ws.fit_to_pages(1, 0)

    output.seek(0)
    return output.getvalue()


def build_report(
    placement_df: pd.DataFrame,
    report_title: str = "Employer Recruiting Report",
    scope_label: str = "Placement View",
    selected_majors: Optional[List[str]] = None,
    contact_df: Optional[pd.DataFrame] = None,
    default_major: Optional[str] = None,
) -> Tuple[bytes, Dict[str, object], pd.DataFrame, pd.DataFrame]:
    cleaned, _, _ = clean_placement_data(placement_df, selected_majors=selected_majors, default_major=default_major)
    company_targets, majors = make_company_targets(cleaned)
    company_targets = enrich_with_contacts(company_targets, contact_df)
    summary = make_summary_tables(cleaned, company_targets, majors)

    major_counts = cleaned["Major"].value_counts()
    top_major = major_counts.index[0] if not major_counts.empty else "—"
    metrics: Dict[str, object] = {
        "Placements": int(len(cleaned)),
        "Unique Companies": int(company_targets["Company"].nunique()) if not company_targets.empty else 0,
        "Majors / Programs": int(len(majors)),
        "Top Major": top_major,
        "Top Major Placements": int(major_counts.iloc[0]) if not major_counts.empty else 0,
        "Tier 1 Employers": int((company_targets["Employer Tier"] == "Tier 1 — Core employer").sum()) if not company_targets.empty else 0,
        "Tier 2 Employers": int((company_targets["Employer Tier"] == "Tier 2 — Relationship employer").sum()) if not company_targets.empty else 0,
    }

    report_bytes = write_report_workbook(cleaned, company_targets, summary, majors, metrics, report_title, scope_label)
    return report_bytes, metrics, company_targets, cleaned


def build_reports_by_major_zip(
    placement_df: pd.DataFrame,
    majors: List[str],
    report_title_prefix: str = "Employer Recruiting Report",
    scope_label: str = "Placement View",
    contact_df: Optional[pd.DataFrame] = None,
) -> bytes:
    output = BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as zf:
        for major in majors:
            report_bytes, _, _, _ = build_report(
                placement_df,
                report_title=f"{major} Employer Recruiting Report",
                scope_label=scope_label,
                selected_majors=[major],
                contact_df=contact_df,
            )
            safe_major = re.sub(r"[^A-Za-z0-9_-]+", "_", major).strip("_") or "Major"
            zf.writestr(f"{safe_major}_Employer_Recruiting_Report.xlsx", report_bytes)
    output.seek(0)
    return output.getvalue()
