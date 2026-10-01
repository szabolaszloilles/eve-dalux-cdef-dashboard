"""CDEF data processing: load JSON, build weekly + contractor aggregations."""
import json
import os
import datetime as dt
import pandas as pd

APPROVED_STATUSES = {"Approved", "Approved, follow-up"}


def _parse(s):
    if not s:
        return None
    try:
        return dt.datetime.fromisoformat(s)
    except ValueError:
        return None


def _week_start(d):
    """Monday of the ISO week containing date d (as date)."""
    return (d - dt.timedelta(days=d.weekday()))


def load_cdef(json_path):
    """Return a tidy DataFrame of CDEF records from the Dalux JSON export."""
    with open(json_path, encoding="utf-8") as f:
        data = json.load(f)
    return _load_cdef_from_dict(data)


def _load_cdef_from_dict(data):
    rows = []
    for c in data.get("CDEF", []):
        created = _parse(c.get("xlsCreationDate")) or _parse(c.get("forwardedDate"))
        modified = _parse(c.get("dateModified"))
        status = (c.get("status") or "").strip()
        rows.append({
            "id": c.get("id"),
            "subject": c.get("subject", ""),
            "contractor": (c.get("workPackage") or "(none)").strip(),
            "status": status,
            "statusDetailed": c.get("statusDetailed", ""),
            "discipline": c.get("discipline", ""),
            "role": c.get("role", ""),
            "responsibleCompany": c.get("responsibleCompany", ""),
            "created": created,
            "modified": modified,
            "is_accepted": status in APPROVED_STATUSES,
        })
    return _finalize(pd.DataFrame(rows))


# Columns the Excel loader needs, by their Dalux header label.
# Each value is a list of accepted aliases (Dalux export variants differ slightly).
_XL_COLS = {
    "id": ["No."],
    "type": ["Type"],
    "subject": ["Subject"],
    "created": ["Date created"],
    "modified": ["Date modified"],
    "role": ["Role"],
    "contractor": ["Work package"],
    "discipline": ["Szakág / Discipline"],
    "status": ["Status"],
    "statusDetailed": ["Status detailed"],
    "responsibleCompany": ["Responsible (company)", "Responsible"],
    "defectType": ["Defect Type / Hiba típusa", "Defect Type", "Hiba típusa"],
    # Days from creation to approval; only filled for status "Approved".
    # Used by the deadline follow-up to date each approval exactly.
    "resolutionTime": ["Resolution time"],
    # "Name, Company" of whoever last changed the CDEF. For a Reported-ready
    # item last changed by the contractor, Date modified = the report-ready date.
    "modifiedBy": ["Modified by"],
}

# Only these must be present for the dashboard to work. Others degrade gracefully.
_XL_REQUIRED = ["id", "created", "modified", "contractor", "status"]


def _resolve(present, key):
    """Return the actual column label present for a logical key, or None."""
    for alias in _XL_COLS[key]:
        if alias in present:
            return alias
    return None


def _find_header_row(src, sheet, scan=15):
    """Locate the row index whose cells contain the known Dalux headers.

    `src` may be a path or a seekable file-like buffer (BytesIO).
    """
    if hasattr(src, "seek"):
        src.seek(0)
    probe = pd.read_excel(src, sheet_name=sheet, header=None, nrows=scan)
    targets = {"No.", "Type", "Subject", "Date created", "Work package", "Status"}
    for i in range(len(probe)):
        rowvals = {str(v).strip() for v in probe.iloc[i].tolist()}
        if len(targets & rowvals) >= 4:
            return i
    return 0  # assume header on first row


def _clean_discipline(v):
    """'Infrastruktúra / Infrastructure' -> 'Infrastruktúra' (Hungarian side)."""
    if not isinstance(v, str) or not v.strip():
        return ""
    return v.split(" / ")[0].strip()


NOT_SPECIFIED = "Pre-categorisation (closed)"

# Manual categorisation of older defects (Dalux only started collecting the
# Defect Type field partway through the project). Applied automatically to any
# defect whose type is blank in the export.
OVERRIDES_FILENAME = "defect_type_overrides.csv"


def load_type_overrides(path_or_buffer=None):
    """Return {defect_id: defect_type} from the overrides CSV.

    With no argument, looks for OVERRIDES_FILENAME next to this module.
    Returns an empty dict if the file is absent or unreadable, so the app
    still works without it.
    """
    if path_or_buffer is None:
        path_or_buffer = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                      OVERRIDES_FILENAME)
        if not os.path.exists(path_or_buffer):
            return {}
    try:
        if hasattr(path_or_buffer, "read"):
            head = path_or_buffer.read(4)
            path_or_buffer.seek(0)
            is_xlsx = head[:2] == b"PK"
        else:
            is_xlsx = str(path_or_buffer).lower().endswith((".xlsx", ".xls", ".xlsm"))
        if is_xlsx:
            hdr = _find_header_row(path_or_buffer, 0)
            if hasattr(path_or_buffer, "seek"):
                path_or_buffer.seek(0)
            ov = pd.read_excel(path_or_buffer, header=hdr, dtype=str)
        else:
            ov = pd.read_csv(path_or_buffer, dtype=str)
    except Exception:
        return {}
    cols = {c.lower().strip(): c for c in ov.columns}
    id_col = cols.get("id") or cols.get("no.") or list(ov.columns)[0]
    ty_col = (cols.get("defecttype") or cols.get("defect type")
              or list(ov.columns)[-1])
    ov = ov[[id_col, ty_col]].dropna()
    ov[id_col] = ov[id_col].astype(str).str.strip()
    ov[ty_col] = ov[ty_col].astype(str).str.strip()
    ov = ov[(ov[ty_col] != "") & (ov[id_col] != "")]
    return dict(zip(ov[id_col], ov[ty_col].apply(_clean_defect_type)))


def apply_type_overrides(df, overrides=None):
    """Fill blank defect types from the manual categorisation lookup.

    Dalux-supplied types always win; overrides only fill gaps.
    """
    if overrides is None:
        overrides = load_type_overrides()
    if not overrides or df.empty:
        return df
    blank = df["defectType"].isin([NOT_SPECIFIED, "", None]) | df["defectType"].isna()
    mapped = df.loc[blank, "id"].astype(str).str.strip().map(overrides)
    df.loc[blank, "defectType"] = mapped.fillna(NOT_SPECIFIED).values
    return df


def _clean_defect_type(v):
    """'Quality defect / Minőségi hiba' -> 'Quality defect'.

    Blank/missing values become NOT_SPECIFIED so they can still be reported on
    (the field was introduced partway through the project, so older defects
    have no type).
    """
    if not isinstance(v, str) or not v.strip():
        return NOT_SPECIFIED
    return v.split(" / ")[0].strip()


def _to_dt(v):
    if pd.isna(v):
        return None
    if isinstance(v, dt.datetime):
        return v
    try:
        return pd.to_datetime(v).to_pydatetime()
    except (ValueError, TypeError):
        return None


def load_cdef_excel(src, sheet=0):
    """Return a tidy DataFrame of CDEF records from the raw Dalux Excel export.

    `src` may be a file path or a seekable file-like buffer (BytesIO). Auto-detects
    the header row, filters to Construction Defect rows (if a Type column is
    present), and maps Dalux columns to the same schema as load_cdef.
    """
    def _seek():
        if hasattr(src, "seek"):
            src.seek(0)

    # Resolve sheet: prefer one literally named 'Data' if present.
    _seek()
    xl = pd.ExcelFile(src)
    if isinstance(sheet, int):
        sheet = "Data" if "Data" in xl.sheet_names else xl.sheet_names[0]
    hdr = _find_header_row(src, sheet)
    _seek()
    raw = pd.read_excel(src, sheet_name=sheet, header=hdr, dtype=str)

    present = set(raw.columns)
    resolved = {key: _resolve(present, key) for key in _XL_COLS}

    missing_required = [key for key in _XL_REQUIRED if resolved[key] is None]
    if missing_required:
        pretty = {"id": "No.", "created": "Date created", "modified": "Date modified",
                  "contractor": "Work package", "status": "Status"}
        names = ", ".join(pretty.get(k, k) for k in missing_required)
        raise ValueError(
            "The uploaded Excel is missing required Dalux columns: " + names
            + ". Make sure this is a Dalux CDEF export (not the summary report).")

    # Re-read date columns without forcing them to strings
    date_labels = [resolved[k] for k in ("created", "modified") if resolved[k]]
    _seek()
    dates = pd.read_excel(src, sheet_name=sheet, header=hdr,
                          usecols=lambda c: c in date_labels)

    # Filter to CDEF if a Type column exists
    if resolved["type"]:
        # Accept every Construction Defect variant, e.g. plain "Construction
        # Defect" and sub-types such as "Construction Defect - SG". Anything
        # else in a mixed export (MAF, RFI, MS/QCP, HDEF...) is filtered out.
        _t = raw[resolved["type"]].astype(str).str.strip()
        keep = _t.str.lower().str.startswith("construction defect")
        raw = raw[keep].copy()
        dates = dates.loc[raw.index]

    def col(key):
        lbl = resolved[key]
        return raw[lbl] if lbl else pd.Series([""] * len(raw), index=raw.index)

    status = col("status").fillna("").astype(str).str.strip()
    created = dates[resolved["created"]].apply(_to_dt) if resolved["created"] else None
    modified = dates[resolved["modified"]].apply(_to_dt) if resolved["modified"] else None

    out = pd.DataFrame({
        "id": col("id").fillna("").astype(str).str.strip(),
        "subject": col("subject").fillna("").astype(str),
        "contractor": col("contractor").fillna("(none)").astype(str).str.strip().replace("", "(none)"),
        "status": status,
        "statusDetailed": col("statusDetailed").fillna("").astype(str),
        "discipline": col("discipline").apply(_clean_discipline),
        "role": col("role").fillna("").astype(str),
        "responsibleCompany": col("responsibleCompany").fillna("").astype(str),
        "daluxType": col("type").fillna("").astype(str).str.strip(),
        "defectType": col("defectType").apply(_clean_defect_type),
        "resolution_days": pd.to_numeric(col("resolutionTime"), errors="coerce"),
        "modified_by": col("modifiedBy").fillna("").astype(str).str.strip(),
        "created": list(created) if created is not None else None,
        "modified": list(modified) if modified is not None else None,
    })
    out["is_accepted"] = out["status"].isin(APPROVED_STATUSES)
    return _finalize(out.reset_index(drop=True))


def load_cdef_any(path_or_buffer, filename=None):
    """Dispatch by file type.

    `path_or_buffer` may be a path string or a file-like buffer. When a buffer is
    passed, supply `filename` so the extension can be detected.
    """
    name = (filename or (path_or_buffer if isinstance(path_or_buffer, str) else "")).lower()
    if name.endswith(".json"):
        if isinstance(path_or_buffer, str):
            return load_cdef(path_or_buffer)
        import json
        path_or_buffer.seek(0)
        data = json.load(path_or_buffer)
        return _load_cdef_from_dict(data)
    if name.endswith((".xlsx", ".xls", ".xlsm")):
        return load_cdef_excel(path_or_buffer)
    # Fall back: assume Excel for buffers without a clear extension
    if not isinstance(path_or_buffer, str):
        return load_cdef_excel(path_or_buffer)
    raise ValueError(f"Unsupported file type: {path_or_buffer}")


def _responsibility(role):
    """Classify who a defect currently sits with, based on its Role.

    Roles starting with 'CÉH' or 'EVE' are with us (CÉH/EVE) for review/action;
    everything else (including blank roles) is with the contractor.
    """
    r = (role or "").strip()
    if r.startswith("CÉH") or r.startswith("EVE"):
        return "With CÉH/EVE"
    return "With Contractor"


TOP10_DONE_WORDS = ("approved", "resolved")
TOP10_EXCLUDE_WORDS = ("waiting", "awaiting")


def _top10_is_done(value):
    """True if a status cell means the issue has been replied to / closed.

    'Waiting for approval' contains 'approval' but means NOT done, so it is
    explicitly excluded.
    """
    s = str(value).strip().lower()
    if not s or s == "nan":
        return False
    if any(w in s for w in TOP10_EXCLUDE_WORDS):
        return False
    return any(w in s for w in TOP10_DONE_WORDS)


def _top10_sheet_contractor(sheet_name):
    """Map a 'Top 10 Issues XYZ' sheet name to the Work Package contractor name."""
    s = sheet_name.lower().replace("top 10 issues", "").strip()
    lookup = {
        "huake": "Huake",
        "crec mep": "CREC MEP",
        "synergy": "Synergy",
        "cscec5 mep": "CSCEC5 MEP",
        "cscec5 infra": "CSCEC5 INFRA",
        "minimax": "Minimax",
        "whb": "WHB",
        "general force": "General Force",
    }
    key = " ".join(s.split())
    for k, v in lookup.items():
        if k in key:
            return v
    return sheet_name.strip()


def load_top10_ratios(path_or_buffer):
    """Parse the Top-10 issues tracker workbook.

    One sheet per contractor. An issue counts as 'replied' when either the
    Deadline (G) or Status (H) column says approved/resolved — the status is
    inconsistently placed between the two columns. The denominator is the number
    of rows that actually have content (not always 10).

    Returns a DataFrame: contractor, listed, replied, reply_pct.
    """
    xl = pd.ExcelFile(path_or_buffer)
    rows = []
    for sheet in xl.sheet_names:
        if "top 10" not in sheet.lower():
            continue
        if hasattr(path_or_buffer, "seek"):
            path_or_buffer.seek(0)
        d = pd.read_excel(path_or_buffer, sheet_name=sheet, header=0)
        listed = replied = 0
        for _, r in d.iterrows():
            # a row counts as listed if it has content beyond the item number
            cells = [r.iloc[c] for c in range(1, min(8, len(r)))]
            has_content = any(str(c).strip() not in ("", "nan") for c in cells)
            if not has_content:
                continue
            listed += 1
            g = r.iloc[6] if len(r) > 6 else None
            h = r.iloc[7] if len(r) > 7 else None
            if _top10_is_done(g) or _top10_is_done(h):
                replied += 1
        if listed:
            rows.append({
                "contractor": _top10_sheet_contractor(sheet),
                "listed": listed,
                "replied": replied,
                "reply_pct": replied / listed * 100,
            })
    return pd.DataFrame(rows)


def reported_ready_ceh(df):
    """Count defects with status 'Reported ready' whose Role starts with 'CÉH'
    (contractor has reported ready; now waiting on CÉH to approve/close)."""
    role = df["role"].fillna("").astype(str).str.strip()
    status = df["status"].fillna("").astype(str).str.strip()
    return int(((status == "Reported ready") & (role.str.startswith("CÉH"))).sum())


def _finalize(df):
    """Attach ISO-week columns and responsibility. Shared by both loaders."""
    if df.empty:
        df["created_week"] = []
        df["modified_week"] = []
        df["responsibility"] = []
        return df
    # 'Awaiting approval' is the contractor-side label for the same state CÉH
    # sees as 'Reported ready' — normalise so both count identically (as Open,
    # and in the Reported-ready-CÉH figure).
    df["status"] = (df["status"].fillna("").astype(str).str.strip()
                    .replace({"Awaiting approval": "Reported ready",
                              "Awaiting Approval": "Reported ready"}))
    df["created"] = pd.to_datetime(df["created"], errors="coerce")
    df["modified"] = pd.to_datetime(df["modified"], errors="coerce")
    df["created_week"] = df["created"].apply(
        lambda d: _week_start(d.date()) if pd.notna(d) else None)
    df["modified_week"] = df["modified"].apply(
        lambda d: _week_start(d.date()) if pd.notna(d) else None)
    df["responsibility"] = df["role"].apply(_responsibility)
    if "daluxType" not in df.columns:
        df["daluxType"] = ""
    df["daluxType"] = df["daluxType"].fillna("").astype(str).str.strip()
    df["is_sg"] = df["daluxType"].str.lower().str.contains(r"\bsg\b", regex=True)
    if "defectType" not in df.columns:
        df["defectType"] = NOT_SPECIFIED
    df["defectType"] = df["defectType"].fillna(NOT_SPECIFIED).replace("", NOT_SPECIFIED)
    if "resolution_days" not in df.columns:
        df["resolution_days"] = float("nan")
    if "modified_by" not in df.columns:
        df["modified_by"] = ""
    df = apply_type_overrides(df)
    return df


def _full_week_index(df):
    weeks = pd.concat([
        df["created_week"].dropna(),
        df["modified_week"].dropna(),
    ])
    start, end = weeks.min(), weeks.max()
    idx = pd.date_range(start=start, end=end, freq="W-MON").date
    # date_range W-MON gives Mondays; ensure list of date objects covering range
    out, cur = [], start
    while cur <= end:
        out.append(cur)
        cur = cur + dt.timedelta(days=7)
    return out


def weekly_metrics(df):
    """Weekly New / Accepted / Total with week-over-week deltas."""
    weeks = _full_week_index(df)
    new_counts = df.dropna(subset=["created_week"]).groupby("created_week").size()
    acc_counts = df[df["is_accepted"]].dropna(subset=["modified_week"]).groupby("modified_week").size()

    recs = []
    cumulative = 0
    for w in weeks:
        new_n = int(new_counts.get(w, 0))
        acc_n = int(acc_counts.get(w, 0))
        cumulative += new_n
        recs.append({
            "Week starting": w,
            "ISO week": f"{w.isocalendar()[0]}-W{w.isocalendar()[1]:02d}",
            "New defects": new_n,
            "Accepted defects": acc_n,
            "Total defects (cumulative)": cumulative,
        })
    w = pd.DataFrame(recs)
    for col, dname, pctname in [
        ("New defects", "New Δ vs prev", "New Δ %"),
        ("Accepted defects", "Accepted Δ vs prev", "Accepted Δ %"),
        ("Total defects (cumulative)", "Total Δ vs prev", "Total Δ %"),
    ]:
        w[dname] = w[col].diff()
        prev = w[col].shift(1)
        w[pctname] = (w[col].diff() / prev.replace(0, pd.NA)) * 100
    return w


def _last_complete_weekend_day(today=None):
    """Return the end date of the most recent complete reporting week.

    Reporting weeks run Friday 00:00 → Thursday 24:00, so a week ends on a
    Thursday. If today is Friday, yesterday's Thursday has just finished and is
    used — that is the case the Friday presentation relies on. If today IS
    Thursday, the day is not over yet, so the previous Thursday is used.
    """
    if today is None:
        today = dt.date.today()
    # weekday(): Mon=0, Tue=1, Wed=2, Thu=3, Fri=4, Sat=5, Sun=6
    offset = (today.weekday() - 3) % 7
    if offset == 0:  # today IS Thursday -> not finished yet, take previous
        offset = 7
    return today - dt.timedelta(days=offset)


# Backwards-compatible alias (the anchor day changed from Friday to Thursday)
_last_complete_friday = _last_complete_weekend_day


def _last_complete_day(today=None):
    """Return yesterday — the most recent day that has fully finished."""
    if today is None:
        today = dt.date.today()
    return today - dt.timedelta(days=1)


def default_period_end(period_days=7, today=None):
    """Pick the right anchor for a reporting period.

    period_days == 1  -> yesterday (the last complete day)
    otherwise         -> the last complete Fri–Thu reporting week
    """
    if period_days == 1:
        return _last_complete_day(today)
    return _last_complete_weekend_day(today)


def week_resolution_rates(df, end_date=None, period_days=7):
    """Resolution rates for the cohort of defects CREATED in the reporting week.

    Returns a dict with:
      new           - defects created during the week
      resolved_todate / pct_todate    - of those, how many are now Approved
                                        (or Approved, follow-up), regardless of when
      resolved_inweek / pct_inweek    - of those, how many were approved within the
                                        same Fri-Thu week (by last-modified date)
      label         - the week label, e.g. '27 Jun-03 Jul'
    """
    if end_date is None:
        end_date = default_period_end(period_days)
    start = end_date - dt.timedelta(days=period_days - 1)

    created = df["created"]
    cohort = df[created.notna() & (created.dt.date >= start) & (created.dt.date <= end_date)]

    n_new = len(cohort)
    resolved = cohort[cohort["is_accepted"]]
    n_todate = len(resolved)

    mod = resolved["modified"]
    n_inweek = int((mod.notna() & (mod.dt.date >= start) & (mod.dt.date <= end_date)).sum())

    def pct(n):
        return (n / n_new * 100) if n_new else 0.0

    def _breakdown(col, key_name):
        """Group the week's cohort by `col` and compute resolution rates."""
        out = []
        if not n_new:
            return out
        for val, grp in cohort.groupby(col):
            g_new = len(grp)
            g_res = grp[grp["is_accepted"]]
            g_todate = len(g_res)
            gmod = g_res["modified"]
            g_inweek = int((gmod.notna() & (gmod.dt.date >= start)
                            & (gmod.dt.date <= end_date)).sum())
            out.append({
                key_name: val,
                "new": g_new,
                "resolved_todate": g_todate,
                "pct_todate": (g_todate / g_new * 100) if g_new else 0.0,
                "resolved_inweek": g_inweek,
                "pct_inweek": (g_inweek / g_new * 100) if g_new else 0.0,
            })
        out.sort(key=lambda x: -x["new"])
        return out

    by_type = _breakdown("defectType", "type")
    by_contractor = _breakdown("contractor", "contractor")

    # Contractor x defect type cross-breakdown
    by_contractor_type = []
    if n_new:
        for (ctr, dtype), grp in cohort.groupby(["contractor", "defectType"]):
            g_new = len(grp)
            g_res = grp[grp["is_accepted"]]
            g_todate = len(g_res)
            gmod = g_res["modified"]
            g_inweek = int((gmod.notna() & (gmod.dt.date >= start)
                            & (gmod.dt.date <= end_date)).sum())
            by_contractor_type.append({
                "contractor": ctr,
                "type": dtype,
                "new": g_new,
                "resolved_todate": g_todate,
                "pct_todate": (g_todate / g_new * 100) if g_new else 0.0,
                "resolved_inweek": g_inweek,
                "pct_inweek": (g_inweek / g_new * 100) if g_new else 0.0,
            })
        by_contractor_type.sort(key=lambda x: (-x["new"], x["contractor"]))

    return {
        "label": _period_label(start, end_date),
        "new": n_new,
        "resolved_todate": n_todate,
        "pct_todate": pct(n_todate),
        "resolved_inweek": n_inweek,
        "pct_inweek": pct(n_inweek),
        "by_type": by_type,
        "by_contractor": by_contractor,
        "by_contractor_type": by_contractor_type,
    }


_METRIC_COLUMNS = [
    "Week starting", "ISO week", "New defects", "New Δ vs prev", "New Δ %",
    "Accepted defects", "Accepted Δ vs prev", "Accepted Δ %",
    "Total defects (cumulative)", "Total Δ vs prev", "Total Δ %"]


def safe_int(value, default=0):
    """int() that tolerates NaN/None.

    The delta columns are produced with .diff(), so the FIRST period has no
    predecessor and its delta is NaN. A contractor whose defects all fall in
    one reporting period therefore has NaN in the only row there is — and
    int(NaN) raises ValueError. Treat "no previous period" as a zero change.
    """
    try:
        if value is None or pd.isna(value):
            return default
        return int(value)
    except (TypeError, ValueError):
        return default


def _empty_metrics():
    """A correctly-shaped but empty metrics frame."""
    return pd.DataFrame(columns=_METRIC_COLUMNS)


def rolling_metrics(df, end_date=None, period_days=7):
    """New / Accepted / Total over consecutive `period_days`-day windows.

    Windows are fixed Saturday→Friday reporting weeks. By default the latest
    window is the last COMPLETE Fri–Thu week (ending the most recent Thursday that
    has fully passed); earlier windows step back by `period_days`.

    Returns the same column shape as weekly_metrics so the app can render it
    unchanged: the 'Week starting' / 'ISO week' columns become the period start
    and a human label like '28 Jun–04 Jul'.
    """
    if end_date is None:
        end_date = default_period_end(period_days)

    created = df["created"].dropna()
    if created.empty:
        return _empty_metrics()

    earliest = created.min().date()
    # Build period end-dates stepping back from end_date until we cover earliest.
    period_ends = []
    cur = end_date
    while cur >= earliest:
        period_ends.append(cur)
        cur = cur - dt.timedelta(days=period_days)
    period_ends = period_ends[::-1]  # chronological

    # A contractor whose FIRST defect was created after the last complete
    # period ended (e.g. newly mobilised this week) produces no periods at all.
    # Return the same well-formed empty frame rather than letting the caller
    # hit a KeyError on a column-less DataFrame.
    if not period_ends:
        return _empty_metrics()

    created_dates = df["created"].dropna().dt.date
    acc_mask = df["is_accepted"]
    modified_dates = df.loc[acc_mask, "modified"].dropna().dt.date

    recs = []
    for pe in period_ends:
        ps = pe - dt.timedelta(days=period_days - 1)  # inclusive start
        new_n = int(((created_dates >= ps) & (created_dates <= pe)).sum())
        acc_n = int(((modified_dates >= ps) & (modified_dates <= pe)).sum())
        cumulative = int((created_dates <= pe).sum())
        label = _period_label(ps, pe)
        recs.append({
            "Week starting": ps,
            "ISO week": label,
            "New defects": new_n,
            "Accepted defects": acc_n,
            "Total defects (cumulative)": cumulative,
        })
    w = pd.DataFrame(recs)
    for col, dname, pctname in [
        ("New defects", "New Δ vs prev", "New Δ %"),
        ("Accepted defects", "Accepted Δ vs prev", "Accepted Δ %"),
        ("Total defects (cumulative)", "Total Δ vs prev", "Total Δ %"),
    ]:
        w[dname] = w[col].diff()
        prev = w[col].shift(1)
        w[pctname] = (w[col].diff() / prev.replace(0, pd.NA)) * 100
    return w


def _period_label(start, end):
    """'17–23 Jun', '29 Jun–05 Jul', or just '06 Aug' for a single day."""
    if start == end:
        return f"{start:%d %b}"
    if start.month == end.month:
        return f"{start.day}–{end.day} {end:%b}"
    return f"{start:%d %b}–{end:%d %b}"


def contractor_status_matrix(df):
    """Contractor x status counts, with totals."""
    mat = pd.crosstab(df["contractor"], df["status"], margins=True, margins_name="Total")
    # move Total row to bottom (crosstab already does), sort contractors by Total desc
    body = mat.drop(index="Total")
    body = body.sort_values("Total", ascending=False)
    return pd.concat([body, mat.loc[["Total"]]])


def contractor_type_summary(df):
    """All-time contractor x defect type: totals, resolved, open, resolution %."""
    if df.empty:
        return pd.DataFrame(columns=["contractor", "defectType", "Total",
                                     "Resolved", "Open", "Resolved %"])
    closed_out = df["status"].isin(["Discontinued", "Rejected"])
    work = df.assign(_open=(~df["is_accepted"] & ~closed_out))
    g = (work.groupby(["contractor", "defectType"])
         .agg(Total=("id", "count"),
              Resolved=("is_accepted", "sum"),
              Open=("_open", "sum"))
         .reset_index())
    g["Resolved %"] = (g["Resolved"] / g["Total"] * 100).round(2)
    g = g.sort_values(["contractor", "Total"], ascending=[True, False])
    return g.rename(columns={"contractor": "Contractor",
                             "defectType": "Defect type"}).reset_index(drop=True)


def contractor_summary(df):
    g = df.groupby("contractor").agg(
        Total=("id", "count"),
        Accepted=("is_accepted", "sum"),
    )
    g["Open"] = g["Total"] - g["Accepted"]
    g["Acceptance %"] = (g["Accepted"] / g["Total"] * 100).round(2)
    return g.sort_values("Total", ascending=False).reset_index()


# ---------------------------------------------------------------------------
# Contractor deadline follow-up
# ---------------------------------------------------------------------------
# A contractor has fulfilled a committed deadline when the CDEF is in one of
# these statuses. "Awaiting approval" is already normalised to "Reported ready"
# by _finalize. A CDEF that was reported ready and then rejected is back to
# "Rejected", so it counts as not fulfilled.
FULFILLED_STATUSES = {"Approved", "Approved, follow-up", "Reported ready"}
WITHDRAWN_STATUSES = {"Discontinued"}

DL_MISSED = "Missed"
DL_DUE_SOON = "Due within 7 days"
DL_NOT_DUE = "Not yet due"
DL_ON_TIME = "Met on time"
DL_LATE = "Met late"
DL_UNCONFIRMED = "Met, timing unconfirmed"
DL_WITHDRAWN = "Discontinued"
DL_NOT_FOUND = "Not in Dalux export"
DL_ORDER = [DL_MISSED, DL_DUE_SOON, DL_NOT_DUE, DL_ON_TIME, DL_LATE, DL_UNCONFIRMED,
            DL_WITHDRAWN, DL_NOT_FOUND]

_DL_ID_COLS = ["CDEF No.", "CDEF No", "CDEF", "CDEF number", "No."]
_DL_DATE_COLS = ["Parsed deadline", "Deadline", "Committed deadline"]
_DL_CONTR_COLS = ["Contractor", "Worksheet", "Work package"]
_DL_SHEET_PREFERENCE = ["Source deadlines", "CDEF Follow-up"]


def cdef_key(value, sg=None):
    """Normalise a CDEF number for matching: 'cdef 1234', 'CDEF-1234' ->
    'CDEF1234'; 'CDEF-SG82', 'CDEF SG 82' -> 'CDEF-SG82'.

    `sg=True/False` forces the prefix (used for rows of the SG / main export,
    whatever their own No. format looks like). Returns None if no number."""
    import re
    txt = str(value or "").strip().upper()
    m = re.search(r"(\d+)\s*$", txt)
    if not m:
        return None
    is_sg = bool(re.search(r"\bSG|-SG|SG\s*\d", txt)) if sg is None else sg
    return f"CDEF-SG{int(m.group(1))}" if is_sg else f"CDEF{int(m.group(1))}"


def _dl_find_table(xl, sheet):
    """Return (header_row, id_col, date_col) for a deadline table, or None."""
    probe = pd.read_excel(xl, sheet_name=sheet, header=None, nrows=20, dtype=str)
    for i in range(len(probe)):
        vals = [str(v).strip() for v in probe.iloc[i].tolist()]
        id_col = next((c for c in _DL_ID_COLS if c in vals), None)
        date_col = next((c for c in _DL_DATE_COLS if c in vals), None)
        if id_col and date_col:
            return i, id_col, date_col
    return None


def load_deadlines(src):
    """Read the contractor deadline workbook.

    Accepts the full CDEF_Status_Summary workbook (its 'Source deadlines'
    sheet is used; only rows marked 'Dated register' when that column exists)
    or any simple sheet with a CDEF number column and a deadline column.

    Returns (deadlines, review):
      deadlines - one row per CDEF: key, cdef_no, contractor, deadline, source
      review    - rows whose deadline is not a usable date (text notes etc.)
    """
    if hasattr(src, "seek"):
        src.seek(0)
    xl = pd.ExcelFile(src)
    sheets = ([s for s in _DL_SHEET_PREFERENCE if s in xl.sheet_names]
              + [s for s in xl.sheet_names if s not in _DL_SHEET_PREFERENCE])
    found = None
    for sh in sheets:
        t = _dl_find_table(xl, sh)
        if t:
            found = (sh,) + t
            break
    if not found:
        raise ValueError(
            "No deadline table found. The workbook needs a sheet with a CDEF "
            "number column (e.g. 'CDEF No.') and a deadline column "
            "(e.g. 'Deadline' or 'Parsed deadline').")
    sheet, hdr_row, id_col, date_col = found
    raw = pd.read_excel(xl, sheet_name=sheet, header=hdr_row)
    raw = raw[raw[id_col].notna()].copy()

    use_col = "Use in summary" if "Use in summary" in raw.columns else None
    contr_col = next((c for c in _DL_CONTR_COLS if c in raw.columns), None)
    orig_col = "Original deadline" if "Original deadline" in raw.columns else None

    raw["key"] = raw[id_col].apply(cdef_key)
    raw["deadline"] = pd.to_datetime(raw[date_col], errors="coerce").dt.normalize()
    raw["cdef_no"] = raw[id_col].astype(str).str.strip()
    raw["contractor"] = (raw[contr_col].astype(str).str.strip()
                         .str.replace(r"\s+scope$", "", regex=True)
                         if contr_col else "")
    raw["original"] = raw[orig_col] if orig_col else raw[date_col]
    raw["source"] = sheet + (": " + raw["Source cell"].astype(str)
                             if "Source cell" in raw.columns else "")

    # Contractor from the curated follow-up register wins over the raw
    # worksheet name (e.g. 'Synergy scope' rows belong to 'Synergy').
    if sheet == "Source deadlines" and "CDEF Follow-up" in xl.sheet_names:
        t = _dl_find_table(xl, "CDEF Follow-up")
        if t:
            fu = pd.read_excel(xl, sheet_name="CDEF Follow-up", header=t[0])
            if "Contractor" in fu.columns:
                fu = fu[fu[t[1]].notna()]
                cmap = dict(zip(fu[t[1]].apply(cdef_key),
                                fu["Contractor"].astype(str).str.strip()))
                raw["contractor"] = raw["key"].map(cmap).fillna(raw["contractor"])

    if use_col:
        selected = raw[raw[use_col].astype(str).str.strip() == "Dated register"]
        review_mask = raw[use_col].astype(str).str.strip() == "Review"
    else:
        selected = raw[raw["deadline"].notna()]
        review_mask = raw["deadline"].isna()

    review = raw[review_mask][["cdef_no", "contractor", "original", "source"]]
    selected = selected[selected["deadline"].notna() & selected["key"].notna()]
    # If a CDEF still has several dates, keep the latest (the current
    # commitment) and say how many dates the file held.
    n_dates = selected.groupby("key")["deadline"].transform("nunique")
    selected = (selected.assign(n_dates=n_dates)
                .sort_values("deadline")
                .drop_duplicates("key", keep="last"))
    deadlines = selected[["key", "cdef_no", "contractor", "deadline",
                          "n_dates", "source"]].reset_index(drop=True)
    return deadlines, review.reset_index(drop=True)


def export_snapshot_date(df):
    """Latest 'Date modified' in the export: the moment its statuses describe."""
    m = df["modified"].dropna()
    return m.max().date() if not m.empty else dt.date.today()


def _changed_by_contractor(modified_by):
    """True when 'Name, Company' names a company other than CÉH or EVE."""
    company = str(modified_by or "").rsplit(",", 1)[-1].strip().upper()
    return bool(company) and not ("CÉH" in company or "CEH" in company
                                  or company.startswith("EVE"))


def _fulfilled_date(row):
    """(date, exact) - when the CDEF reached its fulfilled status.

    Reported ready, last changed by the contractor: 'Date modified' is the
    moment they reported it ready (exact).
    Approved: creation + Dalux 'Resolution time' = the approval moment. The
    contractor reported ready on or before that, so the date is only an upper
    bound (not exact).
    Anything else: 'Date modified', also an upper bound.
    """
    if row["status"] == "Reported ready" and _changed_by_contractor(row.get("modified_by")):
        return row["modified"], True
    if (row["status"] == "Approved" and pd.notna(row.get("resolution_days"))
            and pd.notna(row["created"])):
        return row["created"] + pd.to_timedelta(row["resolution_days"], unit="D"), False
    return row["modified"], False


def deadline_followup(df, df_sg, deadlines, as_of):
    """Classify every committed deadline against the Dalux status.

    `as_of` (date): deadlines strictly before it are due. The deadline day
    itself still counts as in time. Returns one row per committed CDEF.
    """
    frames = []
    for data, is_sg in ((df, False), (df_sg, True)):
        if data is not None and not data.empty:
            d = data.copy()
            d["key"] = d["id"].apply(lambda v: cdef_key(v, sg=is_sg))
            frames.append(d)
    dal = (pd.concat(frames, ignore_index=True).drop_duplicates("key")
           if frames else pd.DataFrame(columns=["key"]))
    keep = ["key", "id", "subject", "contractor", "status", "role",
            "created", "modified", "modified_by", "resolution_days"]
    dal = dal[[c for c in keep if c in dal.columns]].rename(
        columns={"contractor": "work_package"})
    m = deadlines.merge(dal, on="key", how="left")

    as_of = pd.Timestamp(as_of)
    m["fulfilled_by"] = pd.NaT
    m["date_exact"] = False
    ful = m["status"].isin(FULFILLED_STATUSES)
    if ful.any():
        fd = m[ful].apply(_fulfilled_date, axis=1)
        m.loc[ful, "fulfilled_by"] = [d for d, _ in fd]
        m.loc[ful, "date_exact"] = [e for _, e in fd]
    m["fulfilled_by"] = pd.to_datetime(m["fulfilled_by"]).dt.normalize()

    def classify(r):
        if pd.isna(r["status"]):
            return DL_NOT_FOUND
        if r["status"] in WITHDRAWN_STATUSES:
            return DL_WITHDRAWN
        if r["status"] in FULFILLED_STATUSES:
            if r["fulfilled_by"] <= r["deadline"]:
                return DL_ON_TIME
            return DL_LATE if r["date_exact"] else DL_UNCONFIRMED
        if r["deadline"] < as_of:
            return DL_MISSED
        if r["deadline"] <= as_of + pd.Timedelta(days=7):
            return DL_DUE_SOON
        return DL_NOT_DUE

    m["result"] = m.apply(classify, axis=1)
    m["days_overdue"] = [
        int((as_of - d).days) if res == DL_MISSED else None
        for d, res in zip(m["deadline"], m["result"])]
    m["days_left"] = [
        int((d - as_of).days) if res in (DL_DUE_SOON, DL_NOT_DUE) else None
        for d, res in zip(m["deadline"], m["result"])]
    m["result"] = pd.Categorical(m["result"], categories=DL_ORDER, ordered=True)
    return m.sort_values(["result", "deadline", "contractor"]).reset_index(drop=True)


def deadline_contractor_summary(fu, as_of):
    """Per contractor, every committed deadline in exactly one column, so the
    columns add up to 'Committed'. Fulfilled items count as met even when the
    deadline is still ahead. 'Deadline passed' and 'Missed %' cover only the
    deadlines whose day is over."""
    as_of = pd.Timestamp(as_of)
    live = fu[~fu["result"].isin([DL_WITHDRAWN, DL_NOT_FOUND])]

    def n(mask):
        return live[mask].groupby("contractor").size()

    g = pd.DataFrame({
        "Committed": live.groupby("contractor").size(),
        "Met on time": n(live["result"] == DL_ON_TIME),
        "Met late": n(live["result"] == DL_LATE),
        "Met, timing unconfirmed": n(live["result"] == DL_UNCONFIRMED),
        "Missed": n(live["result"] == DL_MISSED),
        "Open, due within 7 days": n(live["result"] == DL_DUE_SOON),
        "Open, due later": n(live["result"] == DL_NOT_DUE),
        "Deadline passed": n(live["deadline"] < as_of),
    }).fillna(0).astype(int)
    g["Missed %"] = (g["Missed"] / g["Deadline passed"].where(g["Deadline passed"] > 0)
                     * 100).round(1).fillna(0.0)
    g = g.sort_values(["Missed", "Committed"], ascending=False)
    total = g.drop(columns="Missed %").sum()
    total["Missed %"] = (round(total["Missed"] / total["Deadline passed"] * 100, 1)
                         if total["Deadline passed"] else 0.0)
    g.loc["Total"] = total
    ints = [c for c in g.columns if c != "Missed %"]
    g[ints] = g[ints].astype(int)
    g.index.name = "Contractor"
    return g.reset_index()
