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
    """Return a tidy DataFrame of CDEF records."""
    with open(json_path, encoding="utf-8") as f:
        data = json.load(f)
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
    df = pd.DataFrame(rows)
    df["created_week"] = df["created"].apply(lambda d: _week_start(d.date()) if pd.notna(d) else None)
    df["modified_week"] = df["modified"].apply(lambda d: _week_start(d.date()) if pd.notna(d) else None)
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
