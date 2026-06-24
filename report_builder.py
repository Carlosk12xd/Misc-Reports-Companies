from __future__ import annotations

import re
import zipfile
from collections import Counter
from difflib import SequenceMatcher
from io import BytesIO
from pathlib import PurePosixPath
from typing import Dict, Iterable, List, Optional, Tuple

import pandas as pd

PLACEMENT_DETAIL_COLUMNS = [
    "Record ID",
    "Job Offer ID",
    "Company ID",
    "Company",
    "Major",
    "Job Role",
    "Job Title",
    "Class Year",
    "Functional Area",
    "Industry",
    "Start Date",
    "State",
]

BASE_TARGET_COLUMNS = ["Rank", "Company", "Total Placements"]
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

COLUMN_ALIASES = {
    "record_id": ["record id", "id", "placement id", "outcome id", "career outcomes id"],
    "job_offer_id": ["job offer id", "offer id", "job id", "posting id", "job posting id", "handshake job id"],
    "company_id": ["company id", "employer id", "account id", "organization id", "company.account id"],
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
    "class_year": ["class year", "class of", "graduation year", "grad year", "year", "graduating class"],
    "functional_area": ["functional area", "function", "career function", "business function", "job function"],
    "industry": ["industry", "industry name", "sector", "employer industry", "company industry"],
    "start_date": ["start date", "job start date", "offer start date", "hire date", "employment start date", "date posted", "apply start date"],
    "state": ["state", "state/province", "region", "location state", "work state", "province", "job location"],
}

CONTACT_ALIASES = {
    "company": ["company", "company name", "employer", "employer name", "account name", "organization", "name"],
    "handshake": ["handshake", "handshake link", "handshake url", "handshake employer link", "employer handshake link", "handshake profile", "handshake link old", "handshake link (old)"],
    "website": ["website", "company website", "web site", "url", "employer website"],
    "linkedin": ["linkedin", "linkedin url", "linkedin link", "company linkedin", "linkedin profile"],
    "owner": ["owner", "account owner", "employer owner", "record owner"],
    "css": ["css", "css assigned", "assigned css", "css assigned to", "career success specialist"],
    "primary_contact": ["primary contact", "primary recruiter", "recruiter", "contact", "contact name"],
    "main_contact": ["main contact", "main recruiter", "main point of contact", "poc"],
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
    text = re.sub(r"\s+", " ", text)
    return text.upper() if len(text) <= 12 else text


def display_value(value: object) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if text.lower() in {"nan", "none", "nat", "null"}:
        return ""
    if re.fullmatch(r"\d+\.0", text):
        return text[:-2]
    return text


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
    return sorted(set(majors), key=str.lower)


def top_value(values: Iterable[object], default: str = "Unspecified") -> str:
    cleaned = [display_value(v) for v in values]
    cleaned = [v for v in cleaned if v]
    if not cleaned:
        return default
    return Counter(cleaned).most_common(1)[0][0]


def join_top(values: Iterable[object], limit: int = 6) -> str:
    cleaned = [display_value(v) for v in values]
    cleaned = [v for v in cleaned if v and v != "Unspecified"]
    if not cleaned:
        return "Unspecified"
    ordered = [value for value, _ in Counter(cleaned).most_common(limit)]
    extra = len(set(cleaned)) - len(ordered)
    text = ", ".join(ordered)
    return text + (f" (+{extra} more)" if extra > 0 else "")


def join_unique_sorted(values: Iterable[object], limit: Optional[int] = None) -> str:
    cleaned = []
    for value in values:
        text = display_value(value)
        if text:
            cleaned.append(text)
    unique = sorted(set(cleaned), key=lambda x: (not x.isdigit(), x))
    if limit and len(unique) > limit:
        shown = unique[:limit]
        return ", ".join(shown) + f" (+{len(unique) - limit} more)"
    return ", ".join(unique) if unique else "Unspecified"


def clean_placement_data(
    df: pd.DataFrame,
    selected_majors: Optional[List[str]] = None,
    default_major: Optional[str] = None,
) -> Tuple[pd.DataFrame, Dict[str, str], List[str]]:
    mapping = infer_columns(df, COLUMN_ALIASES)
    missing = []
    if "company" not in mapping:
        missing.append("company/employer")
    if "major" not in mapping and default_major is None and "Source Program" not in df.columns:
        missing.append("major/program")
    if missing:
        raise ValueError("Could not identify required column(s): " + ", ".join(missing) + ".")

    cleaned = pd.DataFrame(index=df.index)
    for source_key, out_col in [
        ("record_id", "Record ID"),
        ("job_offer_id", "Job Offer ID"),
        ("company_id", "Company ID"),
    ]:
        cleaned[out_col] = df[mapping[source_key]].apply(display_value) if source_key in mapping else ""

    cleaned["Company"] = df[mapping["company"]].apply(display_value)
    if "major" in mapping:
        cleaned["Major"] = df[mapping["major"]].apply(normalize_major)
    elif default_major:
        cleaned["Major"] = normalize_major(default_major)
    else:
        cleaned["Major"] = df["Source Program"].apply(normalize_major)

    for source_key, out_col in [
        ("job_role", "Job Role"),
        ("job_title", "Job Title"),
        ("class_year", "Class Year"),
        ("functional_area", "Functional Area"),
        ("industry", "Industry"),
        ("state", "State"),
    ]:
        cleaned[out_col] = df[mapping[source_key]].apply(display_value) if source_key in mapping else ""

    cleaned["Start Date"] = pd.to_datetime(df[mapping["start_date"]], errors="coerce") if "start_date" in mapping else pd.NaT
    cleaned = cleaned[cleaned["Company"].astype(str).str.strip() != ""].copy()
    cleaned = cleaned[~cleaned["Company"].str.lower().isin(["nan", "none"])]

    if selected_majors:
        selected = {normalize_major(m) for m in selected_majors if str(m).strip()}
        if selected:
            cleaned = cleaned[cleaned["Major"].isin(selected)]

    for col in ["Job Role", "Job Title", "Functional Area", "Industry", "State", "Class Year"]:
        cleaned[col] = cleaned[col].fillna("").apply(display_value)
    for col in ["Job Role", "Job Title", "Functional Area", "Industry"]:
        cleaned[col] = cleaned[col].replace("", "Unspecified")

    return cleaned[PLACEMENT_DETAIL_COLUMNS].reset_index(drop=True), mapping, missing


def tier_thresholds(total_placements: int) -> Tuple[int, int]:
    tier1 = max(3, round(total_placements * 0.0125))
    tier2 = max(2, round(total_placements * 0.0050))
    if tier2 >= tier1:
        tier2 = max(2, tier1 - 1)
    return tier1, tier2


def major_columns(cleaned: pd.DataFrame) -> List[str]:
    counts = cleaned["Major"].fillna("UNSPECIFIED").value_counts()
    majors = [m for m in counts.index.tolist() if m != "UNSPECIFIED"]
    if "UNSPECIFIED" in counts.index:
        majors.append("UNSPECIFIED")
    return majors


def format_major_mix(counts: Dict[str, int]) -> str:
    positive = [(major, count) for major, count in counts.items() if count > 0]
    if not positive:
        return "Unspecified"
    positive.sort(key=lambda x: (-x[1], x[0]))
    if len(positive) == 1:
        return f"{positive[0][0]} only"
    if len(positive) <= 3:
        return " + ".join([m for m, _ in positive])
    return ", ".join([m for m, _ in positive[:3]]) + f" + {len(positive) - 3} more"


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
        rows.append({
            "Company": company,
            "Total Placements": total,
            **major_counts,
            "Major Mix": format_major_mix(major_counts),
            "Class Year(s)": join_unique_sorted(group["Class Year"]),
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
    target_columns = BASE_TARGET_COLUMNS + majors + POST_MAJOR_TARGET_COLUMNS
    targets = pd.DataFrame(rows)
    if targets.empty:
        return pd.DataFrame(columns=target_columns), majors
    targets = targets.sort_values(["Total Placements", "Company"], ascending=[False, True]).reset_index(drop=True)
    targets.insert(0, "Rank", range(1, len(targets) + 1))
    return targets[target_columns], majors


def summarize_dimension(cleaned: pd.DataFrame, dimension: str, label: str, majors: List[str], limit: Optional[int] = None) -> pd.DataFrame:
    use_majors = majors[:2]  # Preserve the exact A:E / G:K template width.
    rows = []
    for value, group in cleaned.groupby(cleaned[dimension].fillna("Unspecified").replace("", "Unspecified"), dropna=False):
        row = {
            label: value,
            "Placements": int(len(group)),
            "Unique Companies": int(group["Company"].nunique()),
        }
        for major in use_majors:
            row[major] = int((group["Major"] == major).sum())
        rows.append(row)
    cols = [label, "Placements", "Unique Companies"] + use_majors
    out = pd.DataFrame(rows, columns=cols)
    if out.empty:
        return pd.DataFrame(columns=cols)
    out = out.sort_values(["Placements", label], ascending=[False, True]).reset_index(drop=True)
    if limit:
        out = out.head(limit)
    return out


def make_summary_tables(cleaned: pd.DataFrame, company_targets: pd.DataFrame, majors: List[str]) -> Dict[str, pd.DataFrame]:
    employer_rows = []
    for tier, group in company_targets.groupby("Employer Tier", dropna=False):
        companies = int(group["Company"].nunique())
        placements = int(group["Total Placements"].sum())
        row = {"Employer Tier": tier, "Placements": placements, "Unique Companies": companies}
        for major in majors[:2]:
            row[major] = int(group[major].sum()) if major in group.columns else 0
        employer_rows.append(row)
    employer_cols = ["Employer Tier", "Placements", "Unique Companies"] + majors[:2]
    employer = pd.DataFrame(employer_rows, columns=employer_cols)
    if not employer.empty:
        order = {"Tier 1 — Core employer": 1, "Tier 2 — Relationship employer": 2, "Tier 3 — Emerging employer": 3}
        employer["_order"] = employer["Employer Tier"].map(order).fillna(99)
        employer = employer.sort_values("_order").drop(columns=["_order"])

    return {
        "Industry Summary": summarize_dimension(cleaned, "Industry", "Industry", majors),
        "Functional Area Summary": summarize_dimension(cleaned, "Functional Area", "Functional Area", majors),
        "Job Role Summary": summarize_dimension(cleaned, "Job Role", "Job Role", majors),
        "State Summary": summarize_dimension(cleaned, "State", "State", majors),
        "Class Year Summary": summarize_dimension(cleaned, "Class Year", "Class Year", majors),
        "Employer Tier Summary": employer,
    }


def prepare_contact_data(contact_df: Optional[pd.DataFrame]) -> Optional[pd.DataFrame]:
    if contact_df is None or contact_df.empty:
        return None
    mapping = infer_columns(contact_df, CONTACT_ALIASES)
    if "company" not in mapping:
        return None
    out = pd.DataFrame()
    out["Company"] = contact_df[mapping["company"]].apply(display_value)
    field_map = [
        ("handshake", "Handshake Link"), ("website", "Website"), ("linkedin", "LinkedIn"),
        ("owner", "Employer Owner"), ("css", "CSS Assigned"), ("main_contact", "Main Contact"),
        ("primary_contact", "Primary Contact"), ("email", "Recruiter Email"), ("phone", "Recruiter Phone"),
        ("engagement_status", "Engagement Status"), ("outreach_status", "Outreach Status"),
        ("last_contacted", "Last Contacted Date"), ("location", "Contact Location"), ("employee_count", "Employee Count"),
    ]
    for key, out_col in field_map:
        out[out_col] = contact_df[mapping[key]].apply(display_value) if key in mapping else ""
    norm_cols = {normalize_header(c): c for c in contact_df.columns}
    loc_parts = [norm_cols[k] for k in ["city", "state", "country"] if k in norm_cols]
    if loc_parts:
        location = contact_df[loc_parts].apply(lambda r: ", ".join([display_value(v) for v in r if display_value(v)]), axis=1)
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
    rows = []
    for _, row in company_targets.iterrows():
        norm = normalize_company(row["Company"])
        matched = None
        match_method = ""
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
        rows.append(new_row)
    return pd.DataFrame(rows)


def report_title_with_scope(report_title: str, scope_label: str) -> str:
    if scope_label and scope_label.strip() and scope_label.lower() not in report_title.lower():
        return f"{report_title} — {scope_label}"
    return report_title


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
        pale_blue = "#D9EAF7"
        border = "#B7B7B7"
        light_border = "#D9E2F3"
        white = "#FFFFFF"
        green = "#D9EAD3"
        yellow = "#FFF2CC"
        orange = "#FCE4D6"
        dark_gray = "#404040"

        fmt_title = workbook.add_format({"bold": True, "font_size": 18, "font_color": white, "bg_color": navy, "align": "left", "valign": "vcenter"})
        fmt_section = workbook.add_format({"bold": True, "font_size": 11, "font_color": white, "bg_color": blue, "border": 1, "border_color": border, "align": "left", "valign": "vcenter"})
        fmt_header = workbook.add_format({"bold": True, "font_color": white, "bg_color": blue, "border": 1, "border_color": border, "align": "center", "valign": "vcenter", "text_wrap": True})
        fmt_text = workbook.add_format({"border": 1, "border_color": light_border, "valign": "top", "text_wrap": True})
        fmt_num = workbook.add_format({"border": 1, "border_color": light_border, "valign": "top", "num_format": "#,##0"})
        fmt_date = workbook.add_format({"border": 1, "border_color": light_border, "valign": "top", "num_format": "m/d/yyyy"})
        fmt_kpi_label = workbook.add_format({"bold": True, "font_color": white, "bg_color": blue, "align": "center", "valign": "vcenter", "border": 1, "border_color": border})
        fmt_kpi_value = workbook.add_format({"bold": True, "font_size": 16, "font_color": navy, "bg_color": pale_blue, "align": "center", "valign": "vcenter", "border": 1, "border_color": border, "num_format": "#,##0"})
        fmt_kpi_text = workbook.add_format({"bold": True, "font_size": 13, "font_color": navy, "bg_color": pale_blue, "align": "center", "valign": "vcenter", "border": 1, "border_color": border})
        fmt_note = workbook.add_format({"font_size": 9, "font_color": dark_gray, "text_wrap": True, "valign": "top"})
        fmt_high = workbook.add_format({"bg_color": green, "font_color": "#274E13", "border": 1, "border_color": light_border})
        fmt_medium = workbook.add_format({"bg_color": yellow, "font_color": "#7F6000", "border": 1, "border_color": light_border})
        fmt_monitor = workbook.add_format({"bg_color": orange, "font_color": "#7F2A00", "border": 1, "border_color": light_border})
        fmt_link = workbook.add_format({"font_color": "#0563C1", "underline": 1, "border": 1, "border_color": light_border, "text_wrap": True})

        dash = workbook.add_worksheet("Executive Dashboard")
        targets_ws = workbook.add_worksheet("Company Targets")
        summary_ws = workbook.add_worksheet("Summary Tables")
        detail_ws = workbook.add_worksheet("Placement Detail")
        for ws in [dash, targets_ws, summary_ws, detail_ws]:
            ws.hide_gridlines(2)
            ws.set_tab_color(blue)

        # Executive Dashboard: matches uploaded 5-year IS/MISM format.
        dash.set_column("A:A", 28)
        for col in ["B", "C", "D", "E", "G", "H", "J", "K", "M"]:
            dash.set_column(f"{col}:{col}", 13)
        dash.set_column("F:F", 28)
        dash.set_column("I:I", 28)
        dash.set_column("L:L", 16)
        dash.set_row(0, 37.25)
        dash.set_row(2, 18)
        dash.merge_range("A1:M1", report_title_with_scope(report_title, scope_label), fmt_title)

        major_counts = cleaned["Major"].value_counts()
        first_major = majors[0] if majors else "Major"
        second_major = majors[1] if len(majors) > 1 else None
        kpis = [
            ("A2:C2", "A3:C3", "Total Placements", metrics.get("Placements", 0)),
            ("D2:F2", "D3:F3", "Unique Companies", metrics.get("Unique Companies", 0)),
            ("G2:I2", "G3:I3", f"{first_major} Placements", int(major_counts.get(first_major, 0))),
            ("J2:M2", "J3:M3", f"{second_major} Placements" if second_major else "Tier 1 Employers", int(major_counts.get(second_major, 0)) if second_major else metrics.get("Tier 1 Employers", 0)),
        ]
        for label_range, value_range, label, value in kpis:
            dash.merge_range(label_range, label, fmt_kpi_label)
            fmt = fmt_kpi_text if isinstance(value, str) else fmt_kpi_value
            dash.merge_range(value_range, value, fmt)

        dash.merge_range("A5:D5", "Top Companies by Placements", fmt_section)
        dash.merge_range("F5:G5", "Placement Mix by Industry", fmt_section)
        dash.merge_range("I5:J5", "Functional Area Mix", fmt_section)
        dash.merge_range("L5:M5", "Major Mix", fmt_section)

        top_companies = company_targets[["Company", "Total Placements"] + [m for m in majors[:2] if m in company_targets.columns]].head(12).copy()
        while len(top_companies.columns) < 4:
            top_companies[f" "] = ""
        top_companies = top_companies.iloc[:, :4]
        top_companies.columns = ["Company", "Placements"] + [m for m in majors[:2]] + ([""] * (4 - 2 - len(majors[:2])))

        industry_dash = summary["Industry Summary"].iloc[:, [0, 1]].head(12) if not summary["Industry Summary"].empty else pd.DataFrame(columns=["Industry", "Placements"])
        functional_dash = summary["Functional Area Summary"].iloc[:, [0, 1]].head(12) if not summary["Functional Area Summary"].empty else pd.DataFrame(columns=["Functional Area", "Placements"])
        major_dash = cleaned["Major"].value_counts().rename_axis("Major").reset_index(name="Placements").head(12)

        def write_dashboard_block(df: pd.DataFrame, start_row: int, start_col: int):
            for c, col in enumerate(df.columns):
                dash.write(start_row, start_col + c, col, fmt_header)
            for r, (_, record) in enumerate(df.iterrows(), start=start_row + 1):
                for c, col in enumerate(df.columns):
                    value = record[col]
                    dash.write(r, start_col + c, value, fmt_num if isinstance(value, (int, float)) and not isinstance(value, bool) else fmt_text)

        write_dashboard_block(top_companies, 5, 0)
        write_dashboard_block(industry_dash, 5, 5)
        write_dashboard_block(functional_dash, 5, 8)
        write_dashboard_block(major_dash, 5, 11)

        def add_dashboard_bar(title: str, sheet_name: str, first_data_row: int, first_col: int, last_data_row: int, chart_width: int = 650):
            chart = workbook.add_chart({"type": "bar"})
            chart.add_series({
                "name": title,
                "categories": [sheet_name, first_data_row, first_col, last_data_row, first_col],
                "values": [sheet_name, first_data_row, first_col + 1, last_data_row, first_col + 1],
                "data_labels": {"value": True},
            })
            chart.set_title({"name": title})
            chart.set_legend({"none": True})
            chart.set_style(10)
            chart.set_size({"width": chart_width, "height": 360})
            chart.set_x_axis({"major_gridlines": {"visible": False}})
            return chart

        # Chart anchors mirror the uploaded format: two charts across, two charts down.
        dash.insert_chart("A20", add_dashboard_bar("Top Companies by Placements", "Executive Dashboard", 6, 0, min(17, 6 + len(top_companies) - 1), 650))
        dash.insert_chart("H20", add_dashboard_bar("Placement Mix by Industry", "Executive Dashboard", 6, 5, min(17, 6 + len(industry_dash) - 1), 560))
        dash.insert_chart("A40", add_dashboard_bar("Functional Area Mix", "Executive Dashboard", 6, 8, min(17, 6 + len(functional_dash) - 1), 650))
        dash.insert_chart("H40", add_dashboard_bar("Major Mix", "Executive Dashboard", 6, 11, min(17, 6 + len(major_dash) - 1), 560))

        # Company Targets sheet.
        targets_ws.freeze_panes(1, 2)
        widths = [8, 30, 12]
        widths += [13] * len(majors)
        widths += [14, 22, 24, 13, 30, 42, 28, 14, 13, 25, 18]
        widths += [24] * max(0, len(company_targets.columns) - len(widths))
        for i, width in enumerate(widths[:len(company_targets.columns)]):
            targets_ws.set_column(i, i, width)
        for c, col in enumerate(company_targets.columns):
            targets_ws.write(0, c, col, fmt_header)
        for r, (_, record) in enumerate(company_targets.iterrows(), start=1):
            for c, col in enumerate(company_targets.columns):
                value = record[col]
                if pd.isna(value):
                    value = ""
                if col in ["First Start Date", "Latest Start Date", "Last Contacted Date"] and value != "":
                    try:
                        targets_ws.write_datetime(r, c, pd.to_datetime(value).to_pydatetime(), fmt_date)
                    except Exception:
                        targets_ws.write(r, c, display_value(value), fmt_text)
                elif col in ["Rank", "Total Placements", "Employee Count"] + majors:
                    try:
                        targets_ws.write_number(r, c, float(value), fmt_num)
                    except Exception:
                        targets_ws.write(r, c, display_value(value), fmt_text)
                elif col in ["Handshake Link", "Website", "LinkedIn"] and str(value).startswith("http"):
                    targets_ws.write_url(r, c, str(value), fmt_link, string=str(value))
                elif col == "Recruiting Priority":
                    targets_ws.write(r, c, value, fmt_high if value == "High" else fmt_medium if value == "Medium" else fmt_monitor)
                else:
                    targets_ws.write(r, c, value, fmt_text)
        if len(company_targets) > 0:
            targets_ws.add_table(0, 0, len(company_targets), len(company_targets.columns) - 1, {
                "name": "CompanyTargets5Year",
                "columns": [{"header": col} for col in company_targets.columns],
                "style": "Table Style Medium 2",
            })
            if "Recruiting Priority" in company_targets.columns:
                priority_col = list(company_targets.columns).index("Recruiting Priority")
                targets_ws.data_validation(1, priority_col, len(company_targets), priority_col, {"validate": "list", "source": ["High", "Medium", "Monitor"]})

        # Summary Tables: exact same block locations as the IS/MISM 5-year template.
        summary_ws.set_column("A:A", 30)
        summary_ws.set_column("B:B", 16)
        summary_ws.set_column("C:E", 13)
        summary_ws.set_column("F:F", 4)
        summary_ws.set_column("G:G", 30)
        summary_ws.set_column("H:H", 16)
        summary_ws.set_column("I:K", 13)
        summary_ws.set_row(0, 32)
        summary_ws.merge_range("A1:K1", f"Summary Tables — {scope_label} {('/'.join(majors[:2]) if majors else 'Employer')} Placements", fmt_title)

        def write_summary_block(title: str, df: pd.DataFrame, top_row: int, left_col: int, width_cols: int = 5):
            summary_ws.merge_range(top_row, left_col, top_row, left_col + width_cols - 1, title, fmt_section)
            for c in range(width_cols):
                header = df.columns[c] if c < len(df.columns) else ""
                summary_ws.write(top_row + 1, left_col + c, header, fmt_header)
            for r, (_, record) in enumerate(df.iterrows(), start=top_row + 2):
                for c in range(width_cols):
                    if c < len(df.columns):
                        val = record[df.columns[c]]
                    else:
                        val = ""
                    if pd.isna(val):
                        val = ""
                    summary_ws.write(r, left_col + c, val, fmt_num if isinstance(val, (int, float)) and not isinstance(val, bool) else fmt_text)

        write_summary_block("Industry Summary", summary["Industry Summary"], 2, 0)
        write_summary_block("Functional Area Summary", summary["Functional Area Summary"], 2, 6)
        write_summary_block("Job Role Summary", summary["Job Role Summary"], 18, 0)
        write_summary_block("State Summary", summary["State Summary"], 18, 6)
        write_summary_block("Class Year Summary", summary["Class Year Summary"], 58, 0)
        write_summary_block("Employer Tier Summary", summary["Employer Tier Summary"], 58, 6)

        # Placement detail sheet: exact same column order as the template.
        detail_widths = [14, 14, 14, 30, 14, 20, 36, 12, 22, 26, 14, 18]
        for i, width in enumerate(detail_widths):
            detail_ws.set_column(i, i, width)
        detail_ws.freeze_panes(1, 0)
        for c, col in enumerate(cleaned.columns):
            detail_ws.write(0, c, col, fmt_header)
        for r, (_, record) in enumerate(cleaned.iterrows(), start=1):
            for c, col in enumerate(cleaned.columns):
                value = record[col]
                if pd.isna(value):
                    value = ""
                if col == "Start Date" and value != "":
                    try:
                        detail_ws.write_datetime(r, c, pd.to_datetime(value).to_pydatetime(), fmt_date)
                    except Exception:
                        detail_ws.write(r, c, display_value(value), fmt_text)
                else:
                    detail_ws.write(r, c, value, fmt_text)
        if len(cleaned) > 0:
            detail_ws.add_table(0, 0, len(cleaned), len(cleaned.columns) - 1, {
                "name": "PlacementDetail5Year",
                "columns": [{"header": col} for col in cleaned.columns],
                "style": "Table Style Medium 2",
            })

        # Print settings.
        dash.set_landscape(); dash.fit_to_pages(1, 1)
        targets_ws.set_landscape(); targets_ws.fit_to_pages(1, 0)
        summary_ws.set_landscape(); summary_ws.fit_to_pages(1, 0)
        detail_ws.set_landscape(); detail_ws.fit_to_pages(1, 0)
    output.seek(0)
    return output.getvalue()


def build_report(
    placement_df: pd.DataFrame,
    report_title: str = "Employer Recruiting Report",
    scope_label: str = "Year View",
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
    scope_label: str = "Year View",
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
