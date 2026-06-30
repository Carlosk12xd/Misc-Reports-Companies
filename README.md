# Major Company Recruiting Report Builder

This Streamlit app converts placement Excel exports into the same four-sheet company recruiting report format as the IS/MISM 5-year report.

## Output format

Every generated workbook uses this structure:

1. **Executive Dashboard**
   - Row 1 title bar
   - KPI blocks in rows 2–3
   - Dashboard data blocks starting in row 5
   - Four bar charts in the same positions as the IS/MISM template
2. **Company Targets**
   - Ranked employer list
   - Total placements
   - Dynamic major/program columns
   - Major mix, industry, function, roles, job titles, states, date range, employer tier, recruiting priority
3. **Summary Tables**
   - Industry Summary
   - Functional Area Summary
   - Job Role Summary
   - State Summary
   - Class Year Summary
   - Employer Tier Summary
4. **Placement Detail**
   - Record ID, Job Offer ID, Company ID, Company, Major, Job Role, Job Title, Class Year, Functional Area, Industry, Start Date, State

## Uploads supported

- A single `.xlsx` or `.xls` placement export
- A `.zip` containing multiple major-specific Excel exports
- Optional CRM/Handshake/contact export for company contact enrichment

## Deploy on Streamlit Cloud

Repo root should contain:

```text
app.py
report_builder.py
requirements.txt
README.md
sanity_check.py
```

Streamlit main file path:

```text
app.py
```

## Local run

```bash
pip install -r requirements.txt
streamlit run app.py
```


## Reporting period / class year option

The sidebar includes a **Report Period** section that filters by **Class Year / Class Of** rather than Start Date.

Default settings:

- **Current class year:** `2026`
- **Past year report:** includes only `Class of 2026`
- **5-year report:** includes `Class of 2026, 2025, 2024, 2023, and 2022`

You can change the current class year in the sidebar later. For example, changing it to `2027` makes the past-year report use `Class of 2027` and the 5-year report use `2027–2023`.

The selected label is written into the generated Excel workbook title, dashboard, summary areas, and default download filenames, so directors can immediately tell whether they are reviewing the 5-year employer view or the past-year employer view.


## Major distribution graph option

The sidebar includes an **Include major distribution graph** checkbox. Keep it on for combined cross-major reports. Turn it off when the uploaded file is already for a single major; the Excel report keeps the same template structure but replaces the Major Mix chart area with a short note instead of a redundant one-bar chart.
