# Dalux scraper — how to finish and run it

The scaffold is complete except for the **navigation selectors**, which depend on
Dalux's actual page markup and must be discovered on the live site.

## Step 1 — install

    pip install playwright pandas openpyxl
    playwright install chromium

## Step 2 — discover the selectors (the only manual part)

    playwright codegen https://build.dalux.com/client/login

A browser opens next to a recorder window. Log in and click through to the
export **exactly as you would by hand**. For every click, codegen prints the
selector it would use, e.g.

    page.get_by_role("button", name="Export").click()

Copy each one into `dalux_config.ini` under `[selectors]`. You need:

| Config key | What to click when recording |
|---|---|
| `username_input`, `password_input`, `login_button` | the login form |
| `login_success_marker` | anything visible only after login (e.g. the project list) |
| `project_link` | the EVE project |
| `quality_control_link`, `tasks_link` | navigating to the task list |
| `cdef_view` / `sg_view` | the saved filter/view for each export |
| `export_button`, `export_excel_option`, `export_confirm_button` | requesting the export |
| `export_ready_marker` | the "file is ready" message that appears ~1 min later |
| `export_open_button` | the button that actually downloads it (blank if the ready message is itself the button) |

Tip: prefer text-based selectors (`text=Export`) over generated CSS paths —
they survive layout changes far better.

### The export is asynchronous
Dalux shows "file is being prepared", then about a minute later "file is ready"
with a prompt to open it. The scraper handles this in two phases: it requests the
export, waits for `export_ready_marker` to appear (up to `export_wait_seconds`,
default 300), then clicks to collect the download.

When recording, **wait for the ready message and click it** so codegen captures
that selector too — it is the one the scraper depends on most.

## Step 3 — credentials (never in the repo or on a shared drive)

Either set environment variables:

    setx DALUX_USER "you@ceh.hu"
    setx DALUX_PASSWORD "..."

or create the file named by `credentials_file` in `dalux_config.ini`:

    [dalux]
    username = you@ceh.hu
    password = ...

Restrict it to the service account (Properties > Security). Add `*credentials*`
and `*.ini` to `.gitignore`.

## Step 4 — test with the browser visible

    python dalux_scraper.py --headful

Watch it drive the UI. Fix any selector that fails, then run headless:

    python dalux_scraper.py

## Step 5 — schedule it before the report job

    01:00  dalux_scraper.py        -> writes CDEF.xlsx and CDEF_SG.xlsx
    01:30  contractor_reports.py   -> builds the per-contractor emails

Leave a gap so a slow or retried scrape still finishes first.

## Safety behaviour already built in

* **Verification before replacement** — a download must be a real .xlsx, contain
  the expected Dalux columns, have enough rows, and be the right export type.
  An HTML error page or a truncated export is rejected.
* **Both-or-neither** — files are only placed once *both* exports verify, so you
  never get a fresh CDEF next to a stale CDEF-SG.
* **Previous file kept** as `*.previous` before each replacement.
* **Retries** with increasing backoff (default 3 attempts).
* **On total failure the old files are left untouched**, and
  `contractor_reports.py` then flags `data_fresh=NO` in the manifest — so the
  sending step can hold delivery rather than mailing stale numbers.

## If Dalux changes and it breaks

Only `dalux_config.ini` should need editing — re-run codegen, replace the broken
selector, done. No Python change, no redeploy.
