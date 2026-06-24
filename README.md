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

