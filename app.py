"""CDEF Defects Dashboard — Streamlit.

Run:  streamlit run app.py
Then upload the Dalux JSON, or place it next to this file as
'Dalux_Dashboard_TS_Meeting.json'.
"""
import datetime as dt
import io
import pandas as pd
import streamlit as st
import plotly.graph_objects as go
from plotly.subplots import make_subplots

import cdef_data as cd

st.set_page_config(page_title="CDEF Defects Dashboard", page_icon="🏗️", layout="wide")

NAVY = "#1F3864"
BLUE = "#2E5496"
GREEN = "#2E7D32"
RED = "#C0392B"
GREY = "#8A8A8A"

st.markdown("""
<style>
    .block-container {padding-top: 2rem;}
    h1 {color:#1F3864;}
    div[data-testid="stMetricValue"] {font-size: 2rem;}
</style>
""", unsafe_allow_html=True)


# ---------------- data loading ----------------
@st.cache_data(show_spinner=False)
def load(file_bytes=None, filename=None, path=None):
    if file_bytes is not None:
        buffer = io.BytesIO(file_bytes)
        return cd.load_cdef_any(buffer, filename=filename)
    return cd.load_cdef_any(path)


st.sidebar.title("CDEF Dashboard")
uploaded = st.sidebar.file_uploader(
    "Dalux CDEF export (Excel)", type=["xlsx", "xls", "xlsm", "json"])

df = None
if uploaded is not None:
    try:
        df = load(file_bytes=uploaded.getvalue(), filename=uploaded.name)
    except Exception as e:
        st.title("🏗️ CDEF — Construction Defects Dashboard")
        st.error(f"Couldn't read that file: {e}")
        st.stop()
else:
    import os
    defaults = ["Reports.xlsx", "Dalux_Dashboard_TS_Meeting.json"]
    local = next((d for d in defaults if os.path.exists(d)), None)
    if local:
        df = load(path=local)
        st.sidebar.caption(f"Loaded local {local}")
    else:
        st.title("🏗️ CDEF — Construction Defects Dashboard")
        st.info("👈 **Upload your Dalux CDEF Excel export** in the sidebar to load the dashboard.")
        st.caption("Your file is processed in-session and is not stored or shared.")
        st.stop()

# ---------------- filters ----------------
st.sidebar.markdown("### Filters")
st.sidebar.caption("All items start selected. Remove tags (×) to narrow the view.")

contractors = sorted(df["contractor"].unique())
sel_contr = st.sidebar.multiselect(
    "Contractor (Work Package)", contractors, default=contractors)

disciplines = sorted(d for d in df["discipline"].unique() if d)
has_blank_disc = (df["discipline"] == "").any()
disc_options = disciplines + (["(none)"] if has_blank_disc else [])
sel_disc = st.sidebar.multiselect(
    "Discipline", disc_options, default=disc_options)

valid_dates = df["created"].dropna()
dmin, dmax = valid_dates.min().date(), valid_dates.max().date()
date_range = st.sidebar.date_input("Created between", (dmin, dmax),
                                   min_value=dmin, max_value=dmax)
if isinstance(date_range, tuple) and len(date_range) == 2:
    d0, d1 = date_range
else:
    d0, d1 = dmin, dmax

# Map the "(none)" pseudo-option back to blank disciplines
disc_real = [d for d in sel_disc if d != "(none)"]
disc_mask = df["discipline"].isin(disc_real)
if "(none)" in sel_disc:
    disc_mask = disc_mask | (df["discipline"] == "")

mask = df["contractor"].isin(sel_contr) & disc_mask
mask &= df["created"].dt.date.between(d0, d1)
fdf = df[mask].copy()

st.sidebar.markdown(f"**{len(fdf)}** of {len(df)} defects shown")

if fdf.empty:
    st.warning("No defects match the current filters. Add some tags back, or widen the date range.")
    st.stop()

weekly = cd.weekly_metrics(fdf)
last, prev = weekly.iloc[-1], weekly.iloc[-2] if len(weekly) > 1 else weekly.iloc[-1]

# ---------------- header ----------------
st.title("🏗️ CDEF — Construction Defects Dashboard")
st.caption(f"EVE Factory Project Debrecen · latest week {last['ISO week']} "
           f"(w/c {last['Week starting']:%d %b %Y}) · {len(fdf)} defects in view")


# ---------------- reusable render helpers ----------------
def render_kpis(_df, _weekly):
    _last = _weekly.iloc[-1]
    wk = _last["ISO week"]
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Total defects (to date)", int(_last["Total defects (cumulative)"]),
              f"{int(_last['Total Δ vs prev']):+d} added in {wk}",
              help="Cumulative count of all defects raised so far. The delta is how many were added during the latest week.")
    k2.metric(f"New defects in {wk}", int(_last["New defects"]),
              f"{int(_last['New Δ vs prev']):+d} vs week before",
              help=f"Defects raised during {wk}. The delta compares this week's count of new defects to the previous week's count of new defects.")
    k3.metric(f"Accepted in {wk}", int(_last["Accepted defects"]),
              f"{int(_last['Accepted Δ vs prev']):+d} vs week before",
              help=f"Defects approved/closed during {wk}. The delta compares this week's count to the previous week's.")
    open_n = int((~_df["is_accepted"] & ~_df["status"].isin(["Discontinued", "Rejected"])).sum())
    k4.metric("Open defects (now)", open_n,
              help="Defects not yet approved, discontinued, or rejected — i.e. currently outstanding.")


def trend_block(_weekly, title, valcol, dcol, pcol, color, key=""):
    st.markdown(f"**{title}**")
    g1, g2 = st.columns([3, 2])
    with g1:
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        fig.add_bar(x=_weekly["ISO week"], y=_weekly[valcol], name=valcol,
                    marker_color=color, opacity=0.85)
        fig.add_scatter(x=_weekly["ISO week"], y=_weekly[dcol], name="Δ vs prev week",
                        mode="lines+markers", line=dict(color="#444", width=2),
                        secondary_y=True)
        fig.update_layout(height=300, margin=dict(l=10, r=10, t=10, b=10),
                          legend=dict(orientation="h", y=-0.25), plot_bgcolor="white")
        fig.update_yaxes(title_text=valcol, secondary_y=False)
        fig.update_yaxes(title_text="Δ", secondary_y=True, showgrid=False)
        st.plotly_chart(fig, width="stretch", key=f"chart_{key}_{valcol}")
    with g2:
        tbl = _weekly[["ISO week", valcol, dcol, pcol]].tail(10).iloc[::-1].copy()

        def style_delta(v):
            if pd.isna(v):
                return ""
            return f"color: {GREEN}" if v > 0 else (f"color: {RED}" if v < 0 else "")
        sty = (tbl.style
               .format({dcol: "{:+.0f}", pcol: "{:+.1f}%", valcol: "{:.0f}"}, na_rep="–")
               .map(style_delta, subset=[dcol]))
        st.dataframe(sty, hide_index=True, width="stretch", height=300,
                     key=f"tbl_{key}_{valcol}")


def all_trends(_weekly, key=""):
    trend_block(_weekly, "New defects per week", "New defects",
                "New Δ vs prev", "New Δ %", BLUE, key=key)
    trend_block(_weekly, "Accepted defects per week", "Accepted defects",
                "Accepted Δ vs prev", "Accepted Δ %", GREEN, key=key)
    trend_block(_weekly, "Total defects (cumulative) per week",
                "Total defects (cumulative)", "Total Δ vs prev", "Total Δ %", NAVY, key=key)


# ---------------- focus-contractor control ----------------
focus_contractors = sorted(fdf["contractor"].unique())
fc1, fc2 = st.columns([3, 2])
with fc1:
    focus = st.selectbox("🔍 Focus on a single contractor",
                         ["All contractors"] + focus_contractors)
with fc2:
    show_overall = st.toggle("Also show overall view", value=True,
                             help="Keep the all-contractors dashboard visible below the focused one.")

st.divider()

# ---------------- focused contractor section ----------------
if focus != "All contractors":
    cdf = fdf[fdf["contractor"] == focus].copy()
    cweekly = cd.weekly_metrics(cdf)
    st.subheader(f"📌 {focus} — {len(cdf)} defects")
    render_kpis(cdf, cweekly)
    st.markdown("")
    # status breakdown for this contractor
    sc = cdf["status"].value_counts()
    chips = "  ·  ".join(f"**{v}** {k}" for k, v in sc.items())
    st.caption(chips)
    st.markdown("**Weekly trends — " + focus + "**")
    all_trends(cweekly, key="focus")
    st.divider()
    if not show_overall:
        st.stop()

# ---------------- overall KPIs ----------------
if focus != "All contractors":
    st.subheader("📊 Overall — all contractors in view")
render_kpis(fdf, weekly)

st.divider()

# ---------------- Report by Contractor ----------------
st.subheader("Report by Contractor")
csum = cd.contractor_summary(fdf)
cc1, cc2 = st.columns([3, 2])

with cc1:
    order = csum.sort_values("Total")["contractor"].tolist()
    pivot = (fdf.groupby(["contractor", "status"]).size()
             .unstack(fill_value=0).reindex(order))
    fig = go.Figure()
    palette = {"Ongoing": BLUE, "Reported ready": "#7FA6D9", "Approved": GREEN,
               "Approved, follow-up": "#66BB6A", "Rejected": RED, "Discontinued": GREY}
    for status in pivot.columns:
        fig.add_bar(y=pivot.index, x=pivot[status], name=status,
                    orientation="h", marker_color=palette.get(status, "#BBBBBB"))
    fig.update_layout(barmode="stack", height=max(300, 40 * len(pivot)),
                      margin=dict(l=10, r=10, t=10, b=10),
                      legend=dict(orientation="h", y=-0.15),
                      xaxis_title="Defects", plot_bgcolor="white")
    st.plotly_chart(fig, width="stretch")

with cc2:
    disp = csum.rename(columns={"contractor": "Contractor"})
    st.dataframe(
        disp,
        hide_index=True, width="stretch",
        height=max(300, 40 * len(disp)),
        column_config={
            "Total": st.column_config.ProgressColumn(
                "Total", format="%d",
                min_value=0, max_value=int(disp["Total"].max())),
            "Acceptance %": st.column_config.NumberColumn(
                "Acceptance %", format="%.1f%%"),
        })

st.divider()

# ---------------- Weekly trends (overall) ----------------
st.subheader("Weekly trends — with previous-week comparison")
all_trends(weekly, key="overall")

st.divider()

# ---------------- Export ----------------
st.subheader("Export")
st.caption("Download the full multi-sheet Excel report (Summary, By Contractor, "
           "Weekly New/Accepted/Total, Raw CDEF) for **all loaded data** — "
           "the sidebar filters and contractor focus do not affect this report.")

import build_excel as be


@st.cache_data(show_spinner="Building Excel report…")
def _report_bytes(token):
    return be.build_report_bytes(df)

# Cache key tied to the dataset size + latest week so it rebuilds when data changes
_token = f"{len(df)}-{weekly.iloc[-1]['ISO week']}"
report_bytes = _report_bytes(_token)

st.download_button(
    "⬇️ Download Excel report",
    data=report_bytes,
    file_name=f"CDEF_Report_{dt.date.today():%Y%m%d}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

with st.expander("Definitions & method"):
    st.markdown("""
- **Contractor** = Work Package (Dalux column I) — populated for all CDEF records, including closed ones.
- **New defects** = count by *date created*, grouped into ISO weeks (Mon–Sun).
- **Accepted defects** = status *Approved* / *Approved, follow-up*, dated by *last-modified* (approval) date.
  Dalux does not expose a dedicated approval date in the export, so the modified date is used as the closure timestamp.
- **Total defects** = cumulative count of all defects raised up to and including that week.
- **Δ vs prev** = change against the previous week.
""")
