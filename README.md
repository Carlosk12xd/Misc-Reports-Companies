# Employer Recruiting Report Builder

This Streamlit app converts placement exports into a director-ready, company-focused recruiting workbook. It uses the same style and structure as the five-year company report, but it is **not hardcoded to IS/MISM**.

## What it handles

- A single `.xlsx` / `.xls` placement export.
- A `.zip` containing multiple major/program folders, each with its own Excel placement export, such as:
  - `MBA/Full_Time_Placements_2026_06_24.xlsx`
  - `Finance/Full_Time_Placements_2026_06_24.xlsx`
  - `ACC/Full_Time_Placements_2026_06_24.xlsx`
- Any major/program code found in the data, including MBA, BSFin, MAcc, BSGSCM, BSHRM, BSBusM, BSEDM, MPA, BSStrat, BSEnt, BSIS, MISM, or future programs.
- Optional CRM/Handshake/contact exports for company contact enrichment.

## Generated workbook

The generated report includes:

1. **Executive Dashboard** — KPIs and charts.
2. **Company Targets** — ranked employer list with dynamic major/program count columns.
3. **Summary Tables** — top companies, major mix, industry mix, functional area mix, role mix, state mix, class-year mix, and tier mix.
4. **Placement Detail** — cleaned row-level placement data with source file/program fields.

## Key behavior

- The app dynamically detects majors/programs from the uploaded data.
- Company-level columns change depending on the majors included in the file.
- The dashboard KPI block is generic: placements, companies, number of majors/programs, top major, Tier 1 employers, and Tier 2 employers.
- Employer tiers scale with placement volume so past-year and five-year reports stay comparable.
- Contact fields are only filled when present in an uploaded CRM/Handshake/contact file. The app does not invent recruiter names, emails, or phone numbers.
- When uploading a ZIP with multiple majors, the app can generate:
  - one combined cross-major report, and
  - a ZIP containing one separate report per selected major.

## Expected placement columns

Required:

- Company / Employer
- Major / Program

Recommended:

- Job Role
- Job Title
- Class of
- Functional Area
- Industry
- Start Date
- State
- Student ID

The app auto-detects common variations of these column names.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy to Streamlit Community Cloud

1. Push these files to a GitHub repository.
2. Go to Streamlit Community Cloud.
3. Select the repository.
4. Set `app.py` as the entry point.
5. Deploy.

No secrets are required unless you later add direct API connections.
