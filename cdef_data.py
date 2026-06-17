"""CDEF data processing: load JSON, build weekly + contractor aggregations."""
import json
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
_XL_COLS = {
    "id": "No.",
    "type": "Type",
    "subject": "Subject",
    "created": "Date created",
    "modified": "Date modified",
    "role": "Role",
    "contractor": "Work package",
    "discipline": "Szakág / Discipline",
    "status": "Status",
    "statusDetailed": "Status detailed",
    "responsibleCompany": "Responsible (company)",
}


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
    # Re-read dates without forcing str for the date columns
    _seek()
    dates = pd.read_excel(src, sheet_name=sheet, header=hdr,
                          usecols=lambda c: c in (_XL_COLS["created"], _XL_COLS["modified"]))

    present = set(raw.columns)
    missing = [lbl for key, lbl in _XL_COLS.items()
               if key not in ("type",) and lbl not in present]
    if missing:
        raise ValueError(
            "The uploaded Excel is missing expected Dalux columns: "
            + ", ".join(missing)
            + ". Make sure this is a Dalux export (not the summary report).")

    # Filter to CDEF if a Type column exists
    if _XL_COLS["type"] in present:
        raw = raw[raw[_XL_COLS["type"]].astype(str).str.strip() == "Construction Defect"].copy()
        dates = dates.loc[raw.index]

    def col(key):
        return raw[_XL_COLS[key]] if _XL_COLS[key] in raw.columns else pd.Series([""] * len(raw))

    status = col("status").fillna("").astype(str).str.strip()
    created = dates[_XL_COLS["created"]].apply(_to_dt) if _XL_COLS["created"] in dates else None
    modified = dates[_XL_COLS["modified"]].apply(_to_dt) if _XL_COLS["modified"] in dates else None

    out = pd.DataFrame({
        "id": col("id").fillna("").astype(str).str.strip(),
        "subject": col("subject").fillna("").astype(str),
        "contractor": col("contractor").fillna("(none)").astype(str).str.strip().replace("", "(none)"),
        "status": status,
        "statusDetailed": col("statusDetailed").fillna("").astype(str),
        "discipline": col("discipline").apply(_clean_discipline),
        "role": col("role").fillna("").astype(str),
        "responsibleCompany": col("responsibleCompany").fillna("").astype(str),
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


def _finalize(df):
    """Attach ISO-week columns. Shared by JSON and Excel loaders."""
    if df.empty:
        df["created_week"] = []
        df["modified_week"] = []
        return df
    df["created"] = pd.to_datetime(df["created"], errors="coerce")
    df["modified"] = pd.to_datetime(df["modified"], errors="coerce")
    df["created_week"] = df["created"].apply(
        lambda d: _week_start(d.date()) if pd.notna(d) else None)
    df["modified_week"] = df["modified"].apply(
        lambda d: _week_start(d.date()) if pd.notna(d) else None)
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


def contractor_status_matrix(df):
    """Contractor x status counts, with totals."""
    mat = pd.crosstab(df["contractor"], df["status"], margins=True, margins_name="Total")
    # move Total row to bottom (crosstab already does), sort contractors by Total desc
    body = mat.drop(index="Total")
    body = body.sort_values("Total", ascending=False)
    return pd.concat([body, mat.loc[["Total"]]])


def contractor_summary(df):
    g = df.groupby("contractor").agg(
        Total=("id", "count"),
        Accepted=("is_accepted", "sum"),
    )
    g["Open"] = g["Total"] - g["Accepted"]
    g["Acceptance %"] = (g["Accepted"] / g["Total"] * 100).round(1)
    return g.sort_values("Total", ascending=False).reset_index()
