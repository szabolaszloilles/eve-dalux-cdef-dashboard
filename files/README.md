# CDEF Defects Dashboard

A dashboard + Excel report for **Construction Defects (CDEF)**, built directly
from the Dalux Excel export.

## Files
- `app.py` — Streamlit dashboard
- `cdef_data.py` — shared data-processing module (reads Excel or JSON)
- `build_excel.py` — generates the Excel summary report
- `CDEF_Report.xlsx` — pre-built report (Summary, By Contractor, Weekly New/Accepted/Total, Raw CDEF)
- `requirements.txt`

## Run the dashboard
```bash
pip install -r requirements.txt
streamlit run app.py
```
Then upload your **Dalux CDEF Excel export** in the sidebar. (A JSON export is
also accepted for backward compatibility.) The file is parsed in-session and is
not stored.

The loader auto-detects the header row and, if the export contains all process
types, automatically filters to `Type = Construction Defect`.

## Regenerate the Excel report
```bash
python build_excel.py <your_dalux_export>.xlsx CDEF_Report.xlsx
```

## Definitions
- **Contractor** = Work Package (Dalux column I) — populated for every CDEF record, including closed ones, so there is no "unassigned" bucket.
- **New defects / week** = count by *date created*, grouped into ISO weeks (Mon–Sun).
- **Accepted defects / week** = status *Approved* / *Approved, follow-up*, dated by *last-modified* (approval) date. Dalux does not expose a dedicated approval date in the export, so the modified date is used as the closure timestamp.
- **Total defects / week** = cumulative count of all defects raised up to and including that week.
- **Δ vs prev** = change against the previous week (shown on every metric).
