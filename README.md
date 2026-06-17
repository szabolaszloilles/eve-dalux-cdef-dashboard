# CDEF Defects Dashboard

A separate dashboard + Excel report for **Construction Defects (CDEF)**, built from the
Dalux JSON export.

## Files
- `app.py` — Streamlit dashboard
- `cdef_data.py` — shared data-processing module (used by both app and report)
- `build_excel.py` — generates the Excel report
- `CDEF_Report.xlsx` — pre-built report (Summary, By Contractor, Weekly New/Accepted/Total, Raw CDEF)
- `requirements.txt`

## Run the dashboard
```bash
pip install -r requirements.txt
streamlit run app.py
```
Then upload the Dalux JSON export in the sidebar, or place it next to `app.py`
as `Dalux_Dashboard_TS_Meeting.json` and it loads automatically.

## Regenerate the Excel report
```bash
python build_excel.py Dalux_Dashboard_TS_Meeting.json CDEF_Report.xlsx
```

## Definitions
- **Contractor** = Work Package (Dalux column I) — populated for every CDEF record, including closed ones, so there is no "unassigned" bucket.
- **New defects / week** = count by *date created*, grouped into ISO weeks (Mon–Sun).
- **Accepted defects / week** = status *Approved* / *Approved, follow-up*, dated by *last-modified* (approval) date. Dalux does not expose a dedicated approval date in the export, so the modified date is used as the closure timestamp.
- **Total defects / week** = cumulative count of all defects raised up to and including that week.
- **Δ vs prev** = change against the previous week (shown on every metric).
