# CDEF daily contractor reports — server deployment

## What this job does
Reads the two Dalux exports, builds a per-contractor HTML email body and Excel
report, and writes a `manifest.csv` listing which email goes to whom.
**It does not send any email.** Sending is handled separately (e.g. Power Automate
reading the manifest).

## Files to place on the server (same folder)
| File | Purpose |
|---|---|
| `contractor_reports.py` | the job that is scheduled |
| `cdef_data.py` | data loading + metrics (required) |
| `build_excel.py` | Excel report builder (required) |
| `contractors.csv` | contractor → email mapping (edit this to change recipients) |
| `defect_type_overrides.csv` | manual defect-type categorisation for older defects |

## Python packages
    pip install -r requirements-server.txt
Only `pandas` and `openpyxl`. Streamlit and Plotly are **not** required.

## How to run it manually (to test)
    python contractor_reports.py <CDEF export>.xlsx <CDEF-SG export>.xlsx --out <output folder>

Optional arguments:
    --contractors path\to\contractors.csv   (default: contractors.csv beside the script)
    --date YYYY-MM-DD                       (override the run date, for testing)

## Scheduling — this is what makes it automatic
Copying the files to the server does **not** make anything run. A scheduled task
must be created.

### Windows (Task Scheduler)
- Trigger: daily, e.g. 01:00 (after the Dalux export has been produced)
- Action: Start a program
  - Program:   `C:\Python\python.exe`   (full path to python.exe)
  - Arguments: `contractor_reports.py "L:\...\CDEF.xlsx" "L:\...\CDEF_SG.xlsx" --out "L:\...\daily_output"`
  - Start in:  the folder containing the scripts  ← important, or imports fail
- Run whether the user is logged on or not
- Use a service account that has read/write access to the L: drive path

### Linux (cron)
    0 1 * * *  cd /opt/cdef && /usr/bin/python3 contractor_reports.py /mnt/l/CDEF.xlsx /mnt/l/CDEF_SG.xlsx --out /mnt/l/daily_output

## Checking it worked (no server access needed)
Each run writes to the output folder, under a dated subfolder:
- `run.log`      — what happened, including any error
- `manifest.csv` — one row per email; the `data_fresh` column is `NO` if the
                   source export looks stale (older than 3 days). Do not send
                   when this is `NO`.

## Note on drive letters
A mapped drive such as `L:` may not be visible to a scheduled task running as a
service account. If the job fails to find the files, use the full UNC path
(`\\server\share\folder\...`) instead of `L:\...`.
