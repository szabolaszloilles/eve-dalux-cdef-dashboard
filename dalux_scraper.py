"""Download the Dalux CDEF and CDEF-SG exports with a headless browser.

Runs unattended (scheduled at 01:00). It logs in, selects each saved filter,
requests an Excel export, waits for Dalux to prepare it, collects the download,
verifies it, and places both files in the target folder.

Everything site-specific — URLs, selectors, the saved-filter names, the tree
path — lives in dalux_config.ini, so a Dalux layout change is fixed by editing
that file rather than this one.

Safety behaviour
----------------
* A download must be a real .xlsx, contain the expected Dalux columns, have
  enough rows, and be the RIGHT export type (exact match, so a CDEF-SG export
  cannot masquerade as CDEF).
* Files are placed only once BOTH exports verify, so you never get a fresh
  CDEF beside a stale CDEF-SG.
* The previous file is kept as *.previous before each replacement.
* On total failure the old files are left untouched, and contractor_reports.py
  then flags data_fresh=NO so the sending step can hold delivery.
* Any failure saves a screenshot next to this script showing what the browser
  was looking at.

Credentials
-----------
Never in this file, never in the repo, never on a shared drive. Either the
environment variables DALUX_USER / DALUX_PASSWORD, or an ini file named by
`credentials_file` in dalux_config.ini:

    [dalux]
    username = someone@ceh.hu
    password = ...

Usage
-----
    python dalux_scraper.py                       # uses dalux_config.ini
    python dalux_scraper.py --config other.ini
    python dalux_scraper.py --headful             # watch it, for debugging
"""
import argparse
import configparser
import datetime as dt
import os
import re
import shutil
import sys
import tempfile
import time
import traceback

import pandas as pd

REQUIRED_COLUMNS = {"No.", "Type", "Date created", "Work package", "Status"}


# ---------------------------------------------------------------------------
# config + credentials
# ---------------------------------------------------------------------------
def load_config(path="dalux_config.ini"):
    if not os.path.exists(path):
        raise FileNotFoundError(f"config not found: {path}")
    cfg = configparser.ConfigParser()
    cfg.read(path, encoding="utf-8")
    return cfg


def load_credentials(cfg):
    """Environment variables win; otherwise read the file named in the config."""
    user = os.environ.get("DALUX_USER")
    pwd = os.environ.get("DALUX_PASSWORD")
    if user and pwd:
        return user, pwd

    path = cfg.get("login", "credentials_file", fallback="").strip()
    if path and os.path.exists(path):
        c = configparser.ConfigParser()
        c.read(path, encoding="utf-8")
        return c.get("dalux", "username"), c.get("dalux", "password")

    raise RuntimeError(
        "No Dalux credentials found. Set DALUX_USER and DALUX_PASSWORD "
        f"environment variables, or create {path!r} with a [dalux] section.")


# ---------------------------------------------------------------------------
# verification — never let a bad download replace a good file
# ---------------------------------------------------------------------------
def verify_export(path, min_rows, log, expect_type=None):
    """Check a downloaded file really is the expected Dalux export.

    Guards against: HTML error pages saved as .xlsx, empty exports, exports
    accidentally filtered down to a handful of rows, and wrong export type.
    """
    if not os.path.exists(path):
        log(f"  VERIFY FAIL: file does not exist: {path}")
        return False
    size = os.path.getsize(path)
    if size < 5000:
        log(f"  VERIFY FAIL: file suspiciously small ({size} bytes) — "
            f"probably an error page, not an export")
        return False
    with open(path, "rb") as f:
        if f.read(2) != b"PK":
            log("  VERIFY FAIL: not a valid .xlsx (missing ZIP signature)")
            return False
    try:
        probe = pd.read_excel(path, sheet_name=0, header=None, nrows=15)
    except Exception as e:
        log(f"  VERIFY FAIL: cannot read the workbook: {e}")
        return False

    header_row = None
    for i in range(len(probe)):
        vals = {str(v).strip() for v in probe.iloc[i].tolist()}
        if len(REQUIRED_COLUMNS & vals) >= 4:
            header_row = i
            break
    if header_row is None:
        log(f"  VERIFY FAIL: expected Dalux columns not found "
            f"({', '.join(sorted(REQUIRED_COLUMNS))})")
        return False

    df = pd.read_excel(path, sheet_name=0, header=header_row)
    rows = len(df)
    if rows < min_rows:
        log(f"  VERIFY FAIL: only {rows} rows (expected at least {min_rows}) — "
            f"the export may be filtered or incomplete")
        return False

    if expect_type and "Type" in df.columns:
        types = {str(t).strip() for t in df["Type"].dropna().unique()}
        # Exact match, not substring: "Construction Defect - SG" CONTAINS
        # "Construction Defect", so a substring test would wrongly accept an SG
        # export as a plain CDEF one (and vice versa) if a selector misfired.
        if not any(t.casefold() == expect_type.casefold() for t in types):
            log(f"  VERIFY FAIL: expected a '{expect_type}' export but found "
                f"types: {sorted(types)}")
            return False
        if len(types) > 1:
            log(f"  NOTE: export contains more than one Type: {sorted(types)}")

    log(f"  verified OK: {rows} rows, header on row {header_row + 1}")
    return True


def place_file(temp_path, final_path, log):
    """Move a verified download into place, keeping one backup of the previous."""
    os.makedirs(os.path.dirname(final_path), exist_ok=True)
    if os.path.exists(final_path):
        backup = final_path + ".previous"
        try:
            if os.path.exists(backup):
                os.remove(backup)
            shutil.copy2(final_path, backup)
            log(f"  previous file kept as {os.path.basename(backup)}")
        except OSError as e:
            log(f"  WARNING: could not keep a backup: {e}")
    shutil.move(temp_path, final_path)
    log(f"  placed: {final_path}")


# ---------------------------------------------------------------------------
# the browser work
# ---------------------------------------------------------------------------
def do_login(page, cfg, user, pwd, log):
    """Fill in and submit the Dalux login form, then wait for it to settle.

    Dalux does a MULTI-STEP redirect after login, ending in a tempSessionToken
    exchange. Navigating away while that is in flight kills the session, so we
    wait until the URL has stopped changing before returning.
    """
    sel = cfg["selectors"]
    timeout = cfg.getint("run", "timeout_seconds", fallback=120) * 1000
    settle = cfg.getint("run", "login_settle_seconds", fallback=20)

    page.fill(sel["username_input"], user, timeout=timeout)
    page.fill(sel["password_input"], pwd, timeout=timeout)
    page.click(sel["login_button"], timeout=timeout)

    deadline = time.time() + timeout / 1000
    last_url, stable_since = None, None
    while time.time() < deadline:
        cur = page.url
        if cur != last_url:
            last_url, stable_since = cur, time.time()
        elif (time.time() - stable_since) >= 3 and "tempSessionToken" not in cur:
            break
        page.wait_for_timeout(500)

    try:
        page.wait_for_load_state("networkidle", timeout=timeout)
    except Exception:
        pass
    page.wait_for_timeout(settle * 1000)
    return "/login" not in page.url


def goto_authenticated(page, url, cfg, user, pwd, log):
    """Open `url`, signing in if Dalux redirects to the login page.

    Dalux preserves the destination in ?returnUrl=..., so logging in at that
    point delivers us straight to the page we asked for. This is far more
    reliable than logging in first and navigating afterwards, and it avoids
    rendering the Field dashboard (whose charts peg a GPU-less server's CPU).
    """
    timeout = cfg.getint("run", "timeout_seconds", fallback=120) * 1000
    for attempt in (1, 2, 3):
        page.goto(url, timeout=timeout, wait_until="domcontentloaded")
        try:
            page.wait_for_load_state("networkidle", timeout=timeout)
        except Exception:
            pass

        if "/login" not in page.url:
            return True

        log(f"  redirected to login (try {attempt}/3) — signing in; "
            f"Dalux's returnUrl will bring us back")
        if do_login(page, cfg, user, pwd, log):
            if "/login" not in page.url:
                return True
        page.wait_for_timeout(5000)
    return False


def download_export(page, cfg, which, download_dir, log, user=None, pwd=None):
    """Drive the UI to produce one export and return the downloaded file path.

    `which` is "cdef" or "sg".

    Flow confirmed by recording the real UI:
      1. select the saved filter (by URL if configured, else by clicking it)
      2. More > Export  (Export is a MENU ITEM, not a button)
      3. untick "Overview drawing" so only the Excel table is produced
      4. Export  -> Dalux prepares the file asynchronously (~1 minute)
      5. "The file is ready" -> Open  -> this opens a POPUP which delivers the
         download; the popup is then closed
    """
    sel = cfg["selectors"]
    timeout = cfg.getint("run", "timeout_seconds", fallback=120) * 1000
    ready_wait = cfg.getint("run", "export_wait_seconds", fallback=300) * 1000

    # --- 1. select the saved filter ---------------------------------------
    # PREFERRED: go straight to the saved filter's own URL.
    #   * No guessing at tree structure or icon positions.
    #   * Avoids rendering the Field module's dashboard, whose charts peg the
    #     CPU at 100% on a server with no GPU (software rendering).
    # An unauthenticated request is redirected to /login?returnUrl=<target>,
    # so goto_authenticated() signs in there and Dalux delivers us straight to
    # the filter.
    log(f"  selecting the {which.upper()} filter")
    view_url = sel.get("cdef_url" if which == "cdef" else "sg_url", "").strip()

    arrived = False
    if view_url:
        arrived = goto_authenticated(page, view_url, cfg, user, pwd, log)
        if arrived:
            log(f"  opened the {which.upper()} filter by direct URL")
        else:
            log("  direct URL did not work — falling back to clicking")

    # FALLBACK: click through the UI. Heavier (it renders the Field dashboard),
    # so only used if no URL is configured or the URL route failed.
    if not arrived:
        if "/login" in page.url and user and pwd:
            log("  signing in before trying the click route")
            do_login(page, cfg, user, pwd, log)

        field_name = sel.get("field_module_name", "").strip() or "Field"
        field_link = sel.get("field_link_text", "").strip() or "Go to Field"
        if "/quality-control" not in page.url:
            for desc, loc in (
                    (f"left rail '{field_name}'",
                     page.get_by_role("button",
                                      name=re.compile(rf"^{re.escape(field_name)}\b"))),
                    (f"'{field_link}' link", page.get_by_text(field_link)),
                    (f"any '{field_name}' button",
                     page.get_by_role("button", name=field_name)),
            ):
                try:
                    loc.first.click(timeout=15000)
                    page.wait_for_url("**/quality-control/**", timeout=20000)
                    log(f"  opened Field via {desc}")
                    break
                except Exception:
                    log(f"  '{desc}' did not open Field, trying next")
        if "/quality-control" not in page.url:
            raise RuntimeError(
                f"could not reach the task list — still at {page.url}")

        want_name = (sel.get("cdef_view_name", "").strip() or "CDEF") \
            if which == "cdef" else (sel.get("sg_view_name", "").strip() or "CDEF-SG")
        target = page.get_by_text(want_name, exact=True).first

        def _visible(loc, ms=1500):
            try:
                loc.wait_for(state="visible", timeout=ms)
                return True
            except Exception:
                return False

        def _expand(label):
            node = (page.locator("dlx-treeview-node")
                    .filter(has=page.get_by_text(label, exact=True)).last)
            icon = node.locator(".mask-icon").first
            if _visible(icon, 4000):
                icon.click(timeout=10000)
                page.wait_for_timeout(900)
                return True
            return False

        chain = [x.strip() for x in sel.get("filter_tree_path", "").split(">")
                 if x.strip()] or ["Tasks", "Saved filters", "CÉH ADMIN"]
        for parent in chain:
            if _visible(target, 1200):
                break
            try:
                if _expand(parent):
                    log(f"  expanded '{parent}'")
            except Exception:
                pass

        if not _visible(target, 5000):
            raise RuntimeError(
                f"could not find the '{want_name}' saved filter in the tree")
        target.click(timeout=timeout)
        try:
            page.wait_for_load_state("networkidle", timeout=timeout)
        except Exception:
            pass

    page.wait_for_timeout(2000)
    if "/login" in page.url:
        raise RuntimeError(f"session lost before exporting ({page.url})")

    # --- 2. More > Export --------------------------------------------------
    log("  opening the Export dialog (More > Export)")
    page.click(sel["more_menu_button"], timeout=timeout)
    page.click(sel["export_button"], timeout=timeout)

    # --- 3. export options -------------------------------------------------
    # Whole list, not just selected tasks.
    if sel.get("export_scope_entire_list"):
        try:
            page.click(sel["export_scope_entire_list"], timeout=15000)
            log("  scope set to 'Entire list'")
        except Exception:
            log("  NOTE: could not set 'Entire list'; using the dialog default")

    # "Overview drawing" is pre-ticked and would make Dalux render drawings —
    # much slower, and not the plain workbook we want.
    ov_label = sel.get("overview_drawing_label", "").strip() or "Overview drawing"
    try:
        box = (page.locator("dlx-field-print-button")
               .filter(has_text=ov_label)
               .get_by_label("", exact=True))
        if box.is_checked(timeout=10000):
            box.uncheck(timeout=10000)
            log(f"  unticked '{ov_label}'")
    except Exception:
        log(f"  NOTE: '{ov_label}' already unticked or not found")

    # --- 4. request the export --------------------------------------------
    log("  requesting the export")
    page.click(sel["export_confirm_button"], timeout=timeout)

    ready_sel = sel.get("export_ready_marker", "").strip()
    log(f"  waiting up to {ready_wait // 1000}s for Dalux to prepare the file")
    started = time.time()
    if ready_sel:
        page.wait_for_selector(ready_sel, timeout=ready_wait)
        log(f"  export ready after {time.time() - started:.0f}s")

    # --- 5. Open -> popup -> download --------------------------------------
    # Clicking Open opens a new tab which serves the file. The download event
    # is on the ORIGINAL page, so expect_download wraps expect_popup.
    popup = None
    with page.expect_download(timeout=ready_wait) as dl_info:
        with page.expect_popup(timeout=timeout) as popup_info:
            page.click(sel["export_open_button"], timeout=timeout)
        popup = popup_info.value
    download = dl_info.value

    tmp = os.path.join(download_dir, f"{which}_{int(time.time())}.xlsx")
    download.save_as(tmp)
    log(f"  downloaded to {tmp} ({os.path.getsize(tmp):,} bytes)")

    if popup:
        try:
            popup.close()
            log("  download tab closed")
        except Exception:
            pass

    # Dismiss the ready dialog if it is still open, so the next export can start
    if sel.get("export_close_button"):
        try:
            page.click(sel["export_close_button"], timeout=10000)
            log("  dialog closed")
        except Exception:
            pass

    return tmp


def scrape_once(cfg, log, headful=False):
    """One full attempt: log in, export both files, verify, place them."""
    from playwright.sync_api import sync_playwright

    user, pwd = load_credentials(cfg)
    sel = cfg["selectors"]
    timeout = cfg.getint("run", "timeout_seconds", fallback=120) * 1000
    min_rows = cfg.getint("run", "min_rows", fallback=50)
    target = cfg.get("output", "target_folder")

    produced = {}
    with tempfile.TemporaryDirectory() as tmpdir:
        with sync_playwright() as p:
            # A server has no GPU, so Chromium falls back to software
            # rendering. Dalux's charts then peg the CPU at 100% and can
            # freeze the machine. These flags disable GPU paths and other
            # work we do not need for a headless export.
            args = [
                "--disable-gpu",
                "--disable-software-rasterizer",
                "--disable-dev-shm-usage",
                "--disable-extensions",
                "--disable-background-timer-throttling",
                "--disable-backgrounding-occluded-windows",
                "--disable-renderer-backgrounding",
                "--js-flags=--max-old-space-size=512",
            ]
            if cfg.getboolean("run", "block_images", fallback=True):
                # Images are pure rendering cost for an export job.
                args.append("--blink-settings=imagesEnabled=false")
            browser = p.chromium.launch(headless=not headful, args=args)
            context = browser.new_context(
                accept_downloads=True,
                viewport={"width": 1280, "height": 900},
                # Skip retina-quality rendering.
                device_scale_factor=1,
            )

            if cfg.getboolean("run", "block_media", fallback=True):
                # Never fetch images, fonts or media at all. Big CPU/network
                # saving; nothing we click depends on them.
                context.route(
                    "**/*",
                    lambda route: route.abort()
                    if route.request.resource_type in
                    ("image", "media", "font") else route.continue_())

            page = context.new_page()
            try:
                # We do NOT log in first and then navigate. Instead each export
                # goes straight to its saved-filter URL; Dalux redirects to
                # /login?returnUrl=... and signing in there delivers us to the
                # right page. This avoids both the session-timing problem and
                # rendering the Field dashboard, whose charts saturate the CPU
                # on a server with no GPU.
                log(f"will sign in as {user} when Dalux asks")

                for which, key, expect in (
                        ("cdef", "cdef_filename", "Construction Defect"),
                        ("sg", "sg_filename", "Construction Defect - SG")):
                    tmp = download_export(page, cfg, which, tmpdir, log,
                                          user=user, pwd=pwd)
                    if not verify_export(tmp, min_rows if which == "cdef" else 1,
                                         log, expect_type=expect):
                        raise RuntimeError(
                            f"{which.upper()} export failed verification — "
                            f"the existing file has NOT been replaced")
                    produced[which] = (tmp, os.path.join(target, cfg.get("output", key)))
            except Exception:
                # Capture what the browser was looking at — essential when this
                # runs headless on a server and nobody can watch it.
                try:
                    # Save next to the script so it is easy to find (the temp
                    # folder path Windows reports uses 8.3 short names).
                    shot_dir = os.path.dirname(os.path.abspath(__file__))
                    shot = os.path.join(
                        shot_dir, f"dalux_failure_{dt.datetime.now():%Y%m%d_%H%M%S}.png")
                    page.screenshot(path=shot, full_page=True)
                    log(f"  FAILURE SCREENSHOT: {shot}")
                    log(f"  page URL at failure: {page.url}")
                except Exception:
                    pass
                raise
            finally:
                context.close()
                browser.close()

        # only place files once BOTH verified — avoids a half-updated pair
        for which, (tmp, final) in produced.items():
            place_file(tmp, final, log)

    return True


# ---------------------------------------------------------------------------
# entry point
# ---------------------------------------------------------------------------
def main(argv=None):
    ap = argparse.ArgumentParser(description="Download Dalux CDEF exports.")
    ap.add_argument("--config", default="dalux_config.ini")
    ap.add_argument("--headful", action="store_true",
                    help="show the browser (debugging)")
    ap.add_argument("--log", default=None, help="write the log here as well")
    a = ap.parse_args(argv)

    lines = []

    def log(msg):
        line = f"[{dt.datetime.now():%Y-%m-%d %H:%M:%S}] {msg}"
        lines.append(line)
        print(line, flush=True)

    cfg = load_config(a.config)
    headful = a.headful or cfg.getboolean("run", "headful", fallback=False)
    attempts = cfg.getint("run", "max_attempts", fallback=3)

    ok = False
    for attempt in range(1, attempts + 1):
        log(f"=== attempt {attempt} of {attempts} ===")
        try:
            scrape_once(cfg, log, headful=headful)
            log("SUCCESS: both exports downloaded, verified and placed")
            ok = True
            break
        except Exception as e:
            log(f"attempt {attempt} FAILED: {e}")
            for ln in traceback.format_exc().splitlines():
                log("  " + ln)
            if attempt < attempts:
                wait = 60 * attempt
                log(f"retrying in {wait}s")
                time.sleep(wait)

    if not ok:
        log("GIVING UP — previous export files left untouched. "
            "The downstream report job will flag the data as stale.")

    if a.log:
        os.makedirs(os.path.dirname(a.log), exist_ok=True)
        with open(a.log, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))

    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
