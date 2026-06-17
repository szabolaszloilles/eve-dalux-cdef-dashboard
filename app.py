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
def load(file_bytes=None, path=None):
    if file_bytes is not None:
        tmp = io.BytesIO(file_bytes)
        import json
        data = json.load(tmp)
        # write to a temp path is overkill; replicate load_cdef on dict
        import tempfile, os
        with tempfile.NamedTemporaryFile("w", suffix=".json", delete=False, encoding="utf-8") as f:
            json.dump(data, f)
            tmppath = f.name
        df = cd.load_cdef(tmppath)
        os.unlink(tmppath)
        return df
    return cd.load_cdef(path)


st.sidebar.title("CDEF Dashboard")
uploaded = st.sidebar.file_uploader("Dalux JSON export", type=["json"])

df = None
if uploaded is not None:
    df = load(file_bytes=uploaded.getvalue())
else:
    import os
    default = "Dalux_Dashboard_TS_Meeting.json"
    if os.path.exists(default):
        df = load(path=default)
        st.sidebar.caption(f"Loaded local {default}")
    else:
        st.title("🏗️ CDEF — Construction Defects Dashboard")
        st.info("👈 **Upload your Dalux JSON export** in the sidebar to load the dashboard.")
        st.caption("Your file is processed in-session and is not stored or shared.")
        st.stop()

# ---------------- filters ----------------
st.sidebar.markdown("### Filters")
contractors = sorted(df["contractor"].unique())
sel_contr = st.sidebar.multiselect("Contractor (Work Package)", contractors, default=contractors)
disciplines = sorted(d for d in df["discipline"].unique() if d)
sel_disc = st.sidebar.multiselect("Discipline", disciplines, default=disciplines)

valid_dates = df["created"].dropna()
dmin, dmax = valid_dates.min().date(), valid_dates.max().date()
date_range = st.sidebar.date_input("Created between", (dmin, dmax), min_value=dmin, max_value=dmax)
if isinstance(date_range, tuple) and len(date_range) == 2:
    d0, d1 = date_range
else:
    d0, d1 = dmin, dmax

mask = df["contractor"].isin(sel_contr) & (df["discipline"].isin(sel_disc) | (df["discipline"] == ""))
mask &= df["created"].dt.date.between(d0, d1)
fdf = df[mask].copy()

if fdf.empty:
    st.warning("No defects match the current filters.")
    st.stop()

weekly = cd.weekly_metrics(fdf)
last, prev = weekly.iloc[-1], weekly.iloc[-2] if len(weekly) > 1 else weekly.iloc[-1]

# ---------------- header + KPIs ----------------
st.title("🏗️ CDEF — Construction Defects Dashboard")
st.caption(f"EVE Factory Project Debrecen · latest week {last['ISO week']} "
           f"(w/c {last['Week starting']:%d %b %Y}) · {len(fdf)} defects in view")

k1, k2, k3, k4 = st.columns(4)
k1.metric("Total defects", int(last["Total defects (cumulative)"]),
          f"{int(last['Total Δ vs prev']):+d} vs prev wk")
k2.metric("New this week", int(last["New defects"]),
          f"{int(last['New Δ vs prev']):+d} vs prev wk")
k3.metric("Accepted this week", int(last["Accepted defects"]),
          f"{int(last['Accepted Δ vs prev']):+d} vs prev wk")
open_n = int((~fdf["is_accepted"] & ~fdf["status"].isin(["Discontinued", "Rejected"])).sum())
k4.metric("Open defects", open_n)

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

# ---------------- Weekly trends ----------------
st.subheader("Weekly trends — with previous-week comparison")


def trend_block(title, valcol, dcol, pcol, color):
    st.markdown(f"**{title}**")
    g1, g2 = st.columns([3, 2])
    with g1:
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        fig.add_bar(x=weekly["ISO week"], y=weekly[valcol], name=valcol,
                    marker_color=color, opacity=0.85)
        fig.add_scatter(x=weekly["ISO week"], y=weekly[dcol], name="Δ vs prev week",
                        mode="lines+markers", line=dict(color="#444", width=2),
                        secondary_y=True)
        fig.update_layout(height=300, margin=dict(l=10, r=10, t=10, b=10),
                          legend=dict(orientation="h", y=-0.25), plot_bgcolor="white")
        fig.update_yaxes(title_text=valcol, secondary_y=False)
        fig.update_yaxes(title_text="Δ", secondary_y=True, showgrid=False)
        st.plotly_chart(fig, width="stretch")
    with g2:
        tbl = weekly[["ISO week", valcol, dcol, pcol]].tail(10).iloc[::-1].copy()

        def style_delta(v):
            if pd.isna(v):
                return ""
            return f"color: {GREEN}" if v > 0 else (f"color: {RED}" if v < 0 else "")
        sty = (tbl.style
               .format({dcol: "{:+.0f}", pcol: "{:+.1f}%", valcol: "{:.0f}"}, na_rep="–")
               .map(style_delta, subset=[dcol]))
        st.dataframe(sty, hide_index=True, width="stretch", height=300)


trend_block("New defects per week", "New defects", "New Δ vs prev", "New Δ %", BLUE)
trend_block("Accepted defects per week", "Accepted defects", "Accepted Δ vs prev", "Accepted Δ %", GREEN)
trend_block("Total defects (cumulative) per week", "Total defects (cumulative)",
            "Total Δ vs prev", "Total Δ %", NAVY)

st.divider()

# ---------------- Export note ----------------
st.subheader("Export")
st.caption("A pre-built multi-sheet Excel report (`CDEF_Report.xlsx`) is delivered alongside this app. "
           "To regenerate it for a new JSON export, run:  `python build_excel.py <export>.json CDEF_Report.xlsx`")

with st.expander("Definitions & method"):
    st.markdown("""
- **Contractor** = Work Package (Dalux column I) — populated for all CDEF records, including closed ones.
- **New defects** = count by *date created*, grouped into ISO weeks (Mon–Sun).
- **Accepted defects** = status *Approved* / *Approved, follow-up*, dated by *last-modified* (approval) date.
  Dalux does not expose a dedicated approval date in the export, so the modified date is used as the closure timestamp.
- **Total defects** = cumulative count of all defects raised up to and including that week.
- **Δ vs prev** = change against the previous week.
""")
