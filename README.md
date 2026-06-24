# IS/MISM Employer Recruiting Report Builder

This Streamlit app converts a placement Excel export into a director-ready IS/MISM employer recruiting workbook with the same structure as the past-year and five-year company reports.

## What it creates

The generated Excel file includes four sheets:

1. **Executive Dashboard** — KPIs and charts.
2. **Company Targets** — ranked employer list with placement counts, BSIS/MISM mix, tier, and recruiting priority.
3. **Summary Tables** — source tables for industries, functional areas, job roles, states, class years, and tiers.
4. **Placement Detail** — cleaned row-level placement data.

If you upload a separate contact/Handshake/CRM export, the app attempts to match by company name and enriches the Company Targets sheet with available fields such as Handshake link, website, LinkedIn, owner, CSS assigned, primary contact, engagement status, location, and employee count.

## Expected placement columns

The app auto-detects common column names. The placement file should include at least:

- Company
- Major

Recommended columns:

- Job Role
- Job Title
- Class of
- Functional Area
- Industry
- Start Date
- State
- Student ID

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy to Streamlit Community Cloud

1. Push these files to a GitHub repository.
2. Go to Streamlit Community Cloud.
3. Choose the repository and set `app.py` as the entry point.
4. Add no secrets unless your future version connects directly to CRM/Handshake APIs.

## Notes

- The app does not invent recruiter emails, phone numbers, or contacts. It only adds contact data when those fields are present in an uploaded contact workbook.
- Employer tiers scale dynamically with the size of the uploaded placement file so a past-year report and five-year report are comparable.
