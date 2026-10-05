"""Generate per-contractor daily CDEF reports.

For each active contractor this writes:
  <outdir>/<run-date>/<contractor>/body.html      email body (Outlook-safe HTML)
  <outdir>/<run-date>/<contractor>/<name>.xlsx    their own Excel report
  <outdir>/<run-date>/manifest.csv                contractor, email, subject, paths
  <outdir>/<run-date>/run.log                     what happened, for diagnosis

The manifest is the handoff to whatever sends the mail (e.g. Power Automate):
one row per email to send. Nothing here sends anything itself.

CLI:
    python contractor_reports.py CDEF.xlsx [CDEF_SG.xlsx] --out L:/CDEF_daily
"""
import argparse
import datetime as dt
import html
import os
import sys
import traceback

import pandas as pd

import cdef_data as cd
import build_excel as be

# How old the newest defect may be before the data is treated as stale.
STALE_AFTER_DAYS = 3

BRAND = "#1B2A4A"
ACCENT = "#3E6DB5"
GREEN = "#2E7D46"
RED = "#C0392B"
AMBER = "#E0A200"
INK = "#2B2F36"
MUTED = "#6B7280"
LINE = "#E6EAF0"


# --------------------------------------------------------------------------
# config
# --------------------------------------------------------------------------
def load_contractors(path):
    """Read the contractor -> email mapping. Only rows with active=yes are used."""
    cfg = pd.read_csv(path, dtype=str).fillna("")
    cfg.columns = [c.strip().lower() for c in cfg.columns]
    if "contractor" not in cfg.columns or "email" not in cfg.columns:
        raise ValueError(f"{path} must have at least 'contractor' and 'email' columns")
    if "active" in cfg.columns:
        cfg = cfg[cfg["active"].str.strip().str.lower().isin(["yes", "y", "true", "1"])]
    cfg["contractor"] = cfg["contractor"].str.strip()
    cfg["email"] = cfg["email"].str.strip()
    if "cc" not in cfg.columns:
        cfg["cc"] = ""
    return cfg[cfg["email"] != ""].reset_index(drop=True)


# --------------------------------------------------------------------------
# metrics
# --------------------------------------------------------------------------
def contractor_metrics(df, contractor, run_date=None):
    """Weekly, daily, by-type and overall figures for one contractor."""
    sub = df[df["contractor"] == contractor]
    if sub.empty:
        return None

    closed_out = sub["status"].isin(["Discontinued", "Rejected"])
    open_now = int((~sub["is_accepted"] & ~closed_out).sum())

    wk_end = cd.default_period_end(7, run_date)
    dy_end = cd.default_period_end(1, run_date)
    week = cd.week_resolution_rates(sub, end_date=wk_end, period_days=7)
    day = cd.week_resolution_rates(sub, end_date=dy_end, period_days=1)

    # accepted counts inside each window (by modified date)
    def accepted_in(start, end):
        acc = sub[sub["is_accepted"]]
        m = acc["modified"]
        return int((m.notna() & (m.dt.date >= start) & (m.dt.date <= end)).sum())

    wk_start = wk_end - dt.timedelta(days=6)
    total = len(sub)
    resolved_total = int(sub["is_accepted"].sum())

    return {
        "contractor": contractor,
        "total": total,
        "open": open_now,
        "resolved_total": resolved_total,
        "closing_pct": (resolved_total / total * 100) if total else 0.0,
        "reported_ready": int((sub["status"] == "Reported ready").sum()),
        "week": {
            "label": week["label"],
            "new": week["new"],
            "accepted": accepted_in(wk_start, wk_end),
            "resolved_of_intake": week["resolved_todate"],
            "pct": week["pct_todate"],
            "by_type": week["by_type"],
        },
        "day": {
            "label": day["label"],
            "new": day["new"],
            "accepted": accepted_in(dy_end, dy_end),
            "resolved_of_intake": day["resolved_todate"],
            "pct": day["pct_todate"],
        },
    }


# --------------------------------------------------------------------------
# Outlook-safe HTML
# --------------------------------------------------------------------------
def _cell(content, align="left", bold=False, colour=INK, size="14px", pad="8px 10px"):
    weight = "700" if bold else "400"
    return (f'<td style="padding:{pad};border-bottom:1px solid {LINE};'
            f'font-family:Segoe UI,Arial,sans-serif;font-size:{size};'
            f'color:{colour};font-weight:{weight};text-align:{align};">{content}</td>')


def _kpi_box(label, value, note="", colour=BRAND):
    return (
        f'<td width="25%" valign="top" style="padding:6px;">'
        f'<table role="presentation" width="100%" cellpadding="0" cellspacing="0" '
        f'style="border:1px solid {LINE};border-radius:6px;background:#FFFFFF;">'
        f'<tr><td style="padding:12px 14px;font-family:Segoe UI,Arial,sans-serif;">'
        f'<div style="font-size:12px;color:{MUTED};text-transform:uppercase;'
        f'letter-spacing:0.4px;font-weight:700;">{html.escape(label)}</div>'
        f'<div style="font-size:28px;font-weight:700;color:{colour};'
        f'padding-top:4px;">{value}</div>'
        f'<div style="font-size:12px;color:{MUTED};padding-top:2px;">{note}</div>'
        f'</td></tr></table></td>')


def _pct_colour(p):
    return GREEN if p >= 60 else (AMBER if p >= 25 else RED)


def _section(title):
    return (f'<tr><td style="padding:22px 6px 8px 6px;font-family:Segoe UI,Arial,sans-serif;'
            f'font-size:17px;font-weight:700;color:{BRAND};">{html.escape(title)}</td></tr>')


def render_email_html(m, m_sg=None, project="EVE Factory Project Debrecen",
                      run_date=None):
    """Build an Outlook-compatible HTML email body for one contractor.

    Uses nested tables and inline styles only — Outlook renders with Word's
    engine, which ignores flexbox/grid and most modern CSS.
    """
    run_date = run_date or dt.date.today()
    name = html.escape(m["contractor"])
    wk, dy = m["week"], m["day"]

    parts = []
    parts.append(
        f'<table role="presentation" width="100%" cellpadding="0" cellspacing="0" '
        f'style="background:#F4F6FA;padding:20px 0;"><tr><td align="center">'
        f'<table role="presentation" width="680" cellpadding="0" cellspacing="0" '
        f'style="width:680px;max-width:680px;background:#FFFFFF;border:1px solid {LINE};'
        f'border-radius:8px;">')

    # header
    parts.append(
        f'<tr><td style="background:{BRAND};padding:20px 24px;border-radius:8px 8px 0 0;'
        f'font-family:Segoe UI,Arial,sans-serif;">'
        f'<div style="color:#FFFFFF;font-size:20px;font-weight:700;">'
        f'Construction Defects — {name}</div>'
        f'<div style="color:#C8D4E8;font-size:13px;padding-top:4px;">'
        f'{html.escape(project)} &nbsp;·&nbsp; report generated {run_date:%d %b %Y}</div>'
        f'</td></tr>')

    parts.append('<tr><td style="padding:14px 18px;"><table role="presentation" '
                 'width="100%" cellpadding="0" cellspacing="0">')

    # current position
    parts.append(_section("Current position"))
    parts.append('<tr><td><table role="presentation" width="100%" cellpadding="0" '
                 'cellspacing="0"><tr>')
    parts.append(_kpi_box("Open defects", f'{m["open"]:,}', "awaiting action", RED
                          if m["open"] else GREEN))
    parts.append(_kpi_box("Reported ready", f'{m["reported_ready"]:,}',
                          "awaiting CÉH approval", ACCENT))
    parts.append(_kpi_box("Total to date", f'{m["total"]:,}', "all defects", BRAND))
    parts.append(_kpi_box("Closing rate", f'{m["closing_pct"]:.2f}%', "overall",
                          _pct_colour(m["closing_pct"])))
    parts.append('</tr></table></td></tr>')

    # yesterday
    parts.append(_section(f'Yesterday — {html.escape(dy["label"])}'))
    parts.append('<tr><td><table role="presentation" width="100%" cellpadding="0" '
                 'cellspacing="0"><tr>')
    parts.append(_kpi_box("New defects", f'{dy["new"]:,}', "raised", RED if dy["new"] else GREEN))
    parts.append(_kpi_box("Accepted", f'{dy["accepted"]:,}', "approved/closed", GREEN))
    parts.append(_kpi_box("Of which resolved", f'{dy["resolved_of_intake"]:,}',
                          "of yesterday's intake", ACCENT))
    parts.append(_kpi_box("Resolved rate", f'{dy["pct"]:.2f}%', "yesterday's intake",
                          _pct_colour(dy["pct"])))
    parts.append('</tr></table></td></tr>')

    # this reporting week
    parts.append(_section(f'Reporting week — {html.escape(wk["label"])} (Fri–Thu)'))
    parts.append('<tr><td><table role="presentation" width="100%" cellpadding="0" '
                 'cellspacing="0"><tr>')
    parts.append(_kpi_box("New defects", f'{wk["new"]:,}',
                          f'raised {wk["label"]}', RED if wk["new"] else GREEN))
    parts.append(_kpi_box("Accepted", f'{wk["accepted"]:,}',
                          f'approved {wk["label"]}', GREEN))
    parts.append(_kpi_box("Of which resolved", f'{wk["resolved_of_intake"]:,}',
                          f'of the {wk["label"]} intake', ACCENT))
    parts.append(_kpi_box("Resolved rate", f'{wk["pct"]:.2f}%',
                          f'of the {wk["label"]} intake', _pct_colour(wk["pct"])))
    parts.append('</tr></table></td></tr>')

    # by defect type
    if wk["by_type"]:
        parts.append(_section("This week's defects by type"))
        rows = [
            '<tr><td colspan="4" style="padding:0 6px;">'
            '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" '
            f'style="border:1px solid {LINE};border-radius:6px;border-collapse:separate;">'
            '<tr style="background:#F7F9FC;">'
            + _cell("Defect type", bold=True, colour=MUTED, size="12px")
            + _cell("Raised", "right", True, MUTED, "12px")
            + _cell("Resolved", "right", True, MUTED, "12px")
            + _cell("Resolved %", "right", True, MUTED, "12px")
            + '</tr>']
        for t in wk["by_type"]:
            rows.append(
                "<tr>"
                + _cell(html.escape(str(t["type"])))
                + _cell(f'{t["new"]:,}', "right")
                + _cell(f'{t["resolved_todate"]:,}', "right")
                + _cell(f'{t["pct_todate"]:.2f}%', "right", True,
                        _pct_colour(t["pct_todate"]))
                + "</tr>")
        rows.append("</table></td></tr>")
        parts.append("".join(rows))

    # SG block
    if m_sg:
        sgw, sgd = m_sg["week"], m_sg["day"]
        parts.append(_section("CDEF-SG (separate workflow)"))
        parts.append('<tr><td><table role="presentation" width="100%" cellpadding="0" '
                     'cellspacing="0"><tr>')
        parts.append(_kpi_box("Open (SG)", f'{m_sg["open"]:,}', "awaiting action",
                              RED if m_sg["open"] else GREEN))
        parts.append(_kpi_box("New yesterday", f'{sgd["new"]:,}', sgd["label"], ACCENT))
        parts.append(_kpi_box("New (week)", f'{sgw["new"]:,}', sgw["label"], ACCENT))
        parts.append(_kpi_box("Closing rate", f'{m_sg["closing_pct"]:.2f}%', "overall",
                              _pct_colour(m_sg["closing_pct"])))
        parts.append('</tr></table></td></tr>')

    # footer
    parts.append(
        f'<tr><td style="padding:22px 6px 6px 6px;font-family:Segoe UI,Arial,sans-serif;'
        f'font-size:12px;color:{MUTED};line-height:1.6;border-top:1px solid {LINE};">'
        f'The attached Excel file contains the full breakdown for {name}.<br>'
        f'<b>Resolved</b> = status <i>Approved</i> or <i>Approved, follow-up</i>. '
        f'<b>Open</b> = not yet approved, discontinued or rejected. '
        f'Reporting weeks run Friday 00:00 – Thursday 24:00.<br>'
        f'This is an automated message from CÉH Zrt. — please do not reply to this address.'
        f'</td></tr>')

    parts.append('</table></td></tr></table></td></tr></table>')
    return "".join(parts)


# --------------------------------------------------------------------------
# orchestration
# --------------------------------------------------------------------------
def _safe(name):
    keep = "".join(ch if ch.isalnum() or ch in " -_" else "_" for ch in str(name))
    return keep.strip().replace(" ", "_")


def check_freshness(df, run_date, log):
    """Warn if the newest defect is suspiciously old (scrape may have failed)."""
    if df is None or df.empty:
        return False
    newest = df["created"].dropna().max()
    if pd.isna(newest):
        log("WARNING: no valid 'created' dates in the data")
        return False
    age = (run_date - newest.date()).days
    log(f"newest defect created {newest.date()} ({age} days before run date)")
    if age > STALE_AFTER_DAYS:
        log(f"WARNING: data looks STALE (older than {STALE_AFTER_DAYS} days). "
            f"The export may not have refreshed.")
        return False
    return True


def generate(cdef_path, sg_path=None, contractors_csv="contractors.csv",
             outdir="output", run_date=None, project="EVE Factory Project Debrecen"):
    """Build per-contractor HTML + Excel and a manifest. Returns the manifest path."""
    run_date = run_date or dt.date.today()
    root = os.path.join(outdir, run_date.strftime("%Y-%m-%d"))
    os.makedirs(root, exist_ok=True)

    log_path = os.path.join(root, "run.log")
    log_lines = []

    def log(msg):
        stamp = dt.datetime.now().strftime("%H:%M:%S")
        line = f"[{stamp}] {msg}"
        log_lines.append(line)
        print(line)

    manifest_rows = []
    try:
        log(f"run date {run_date}")
        log(f"reading CDEF export: {cdef_path}")
        df = cd.load_cdef_any(cdef_path)
        log(f"  {len(df)} CDEF records")
        fresh = check_freshness(df, run_date, log)

        df_sg = None
        if sg_path:
            log(f"reading CDEF-SG export: {sg_path}")
            df_sg = cd.load_cdef_any(sg_path)
            log(f"  {len(df_sg)} CDEF-SG records")

        cfg = load_contractors(contractors_csv)
        log(f"{len(cfg)} active contractors in {contractors_csv}")

        failures = []
        for _, row in cfg.iterrows():
            name = row["contractor"]
            # Each contractor is isolated: one failure must not cost the whole
            # night's reporting for everyone else.
            try:
                m = contractor_metrics(df, name, run_date)
                if m is None:
                    log(f"SKIP {name}: no defects in the export")
                    continue
                m_sg = contractor_metrics(df_sg, name, run_date) if df_sg is not None else None

                cdir = os.path.join(root, _safe(name))
                os.makedirs(cdir, exist_ok=True)

                body = render_email_html(m, m_sg, project=project, run_date=run_date)
                body_path = os.path.join(cdir, "body.html")
                with open(body_path, "w", encoding="utf-8") as f:
                    f.write(body)

                sub = df[df["contractor"] == name]
                sub_sg = (df_sg[df_sg["contractor"] == name]
                          if df_sg is not None else None)
                xlsx_path = os.path.join(cdir, f"{_safe(name)}_CDEF_{run_date:%Y%m%d}.xlsx")
                be.build_workbook(sub, df_sg=sub_sg).save(xlsx_path)

                subject = (f"CDEF daily report — {name} — {run_date:%d %b %Y} "
                           f"({m['open']} open)")
                manifest_rows.append({
                    "contractor": name,
                    "email": row["email"],
                    "cc": row.get("cc", ""),
                    "subject": subject,
                    "body_html": os.path.abspath(body_path),
                    "attachment": os.path.abspath(xlsx_path),
                    "open_defects": m["open"],
                    "new_yesterday": m["day"]["new"],
                    "data_fresh": "yes" if fresh else "NO",
                })
                log(f"OK  {name}: open={m['open']} new_yesterday={m['day']['new']}")
            except Exception as e:
                failures.append(name)
                log(f"ERROR {name}: {type(e).__name__}: {e}")
                for ln in traceback.format_exc().splitlines()[-4:]:
                    log("    " + ln)
                log(f"      -> skipping {name}; the other contractors continue")

        # Reconcile the CSV against what is actually in the Dalux export, so a
        # name mismatch is visible instead of silently producing no email.
        in_data = set(df["contractor"].dropna().unique())
        if df_sg is not None:
            in_data |= set(df_sg["contractor"].dropna().unique())
        in_csv = set(cfg["contractor"])
        missing_csv = sorted(in_data - in_csv)
        missing_data = sorted(in_csv - in_data)
        if missing_csv:
            log(f"WARNING: work packages in Dalux with no row in the contractor "
                f"file (nobody will be emailed): {', '.join(missing_csv)}")
        if missing_data:
            log(f"NOTE: contractors in the contractor file with no defects in "
                f"this export: {', '.join(missing_data)}")
        if failures:
            log(f"WARNING: {len(failures)} contractor(s) failed: {', '.join(failures)}")

    except Exception:
        log("FAILED with an exception:")
        for line in traceback.format_exc().splitlines():
            log("  " + line)
        with open(log_path, "w", encoding="utf-8") as f:
            f.write("\n".join(log_lines))
        raise

    manifest_path = os.path.join(root, "manifest.csv")
    pd.DataFrame(manifest_rows).to_csv(manifest_path, index=False, encoding="utf-8-sig")
    log(f"manifest written: {manifest_path} ({len(manifest_rows)} emails to send)")
    with open(log_path, "w", encoding="utf-8") as f:
        f.write("\n".join(log_lines))
    return manifest_path


def main(argv=None):
    p = argparse.ArgumentParser(description="Build per-contractor daily CDEF reports.")
    p.add_argument("cdef", help="Dalux CDEF export (.xlsx)")
    p.add_argument("sg", nargs="?", default=None, help="Dalux CDEF-SG export (.xlsx)")
    p.add_argument("--contractors", default="contractors.csv",
                   help="contractor -> email mapping CSV")
    p.add_argument("--out", default="output", help="output root folder")
    p.add_argument("--date", default=None, help="override run date (YYYY-MM-DD)")
    a = p.parse_args(argv)
    run_date = dt.date.fromisoformat(a.date) if a.date else None
    generate(a.cdef, a.sg, a.contractors, a.out, run_date)
    return 0


if __name__ == "__main__":
    sys.exit(main())
