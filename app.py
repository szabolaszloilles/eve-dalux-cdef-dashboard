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

st.set_page_config(page_title="CDEF Defects Dashboard", page_icon="🏗️",
                   layout="wide", initial_sidebar_state="expanded",
                   menu_items={"About": None, "Get Help": None, "Report a bug": None})

# ---- palette ----
NAVY = "#1B2A4A"      # headers / primary
BLUE = "#3E6DB5"      # bars / accents
GREEN = "#2E7D46"     # good
RED = "#C0392B"       # bad
INK = "#2B2F36"       # body text
MUTED = "#8A94A6"     # secondary text / neutral deltas
CARD_BG = "#FFFFFF"
LINE = "#E6EAF0"

st.markdown(f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

    html, body, [class*="css"] {{ font-family: 'Inter','Segoe UI',sans-serif; }}
    .block-container {{ padding-top: 2.2rem; max-width: 1400px; }}

    /* Hide Streamlit chrome: top-right menu, source-code icon, footer, deploy */
    #MainMenu {{ visibility: hidden; }}
    header [data-testid="stToolbar"] {{ display: none; }}
    header [data-testid="stActionButtonIcon"] {{ display: none; }}
    [data-testid="stToolbarActions"] {{ display: none; }}
    footer {{ visibility: hidden; }}
    .stDeployButton {{ display: none; }}
    a[href*="github.com"] {{ display: none !important; }}

    h1 {{ color: {NAVY}; font-weight: 700; letter-spacing: -0.5px; font-size: 2.4rem; }}
    h2, h3 {{ color: {NAVY}; font-weight: 600; }}
    h3 {{ font-size: 1.55rem; }}
    h4 {{ color: {NAVY}; font-weight: 600; font-size: 1.3rem; }}
    .stCaption, div[data-testid="stCaptionContainer"] p {{ font-size: 0.95rem; }}
    section[data-testid="stSidebar"] label p {{ font-size: 1rem; font-weight: 600; }}

    /* KPI cards */
    div[data-testid="stMetric"] {{
        background: {CARD_BG};
        border: 1px solid {LINE};
        border-radius: 12px;
        padding: 14px 16px 12px 16px;
        box-shadow: 0 1px 2px rgba(16,24,40,0.04);
    }}
    div[data-testid="stMetricLabel"] p {{
        font-size: 0.95rem; font-weight: 600; color: {MUTED};
        letter-spacing: 0.2px;
    }}
    div[data-testid="stMetricValue"] {{ font-size: 2.4rem; font-weight: 700; color: {NAVY}; }}
    div[data-testid="stMetricDelta"] {{ font-size: 1rem; font-weight: 600; }}

    /* Section row labels (#### Total etc.) */
    .stMarkdown h4 {{ margin-top: 0.6rem; margin-bottom: 0.4rem; }}

    hr {{ border-color: {LINE}; }}
    section[data-testid="stSidebar"] {{ background: #F7F9FC; border-right: 1px solid {LINE}; }}
</style>
""", unsafe_allow_html=True)


# ---------------- data loading ----------------
@st.cache_data(show_spinner=False)
def load(file_bytes=None, filename=None, path=None, override_bytes=None):
    ov = None
    if override_bytes is not None:
        ov = cd.load_type_overrides(io.BytesIO(override_bytes))
    if file_bytes is not None:
        buffer = io.BytesIO(file_bytes)
        d = cd.load_cdef_any(buffer, filename=filename)
    else:
        d = cd.load_cdef_any(path)
    if ov:
        # Re-apply with the uploaded mapping (fills any remaining blanks)
        d = cd.apply_type_overrides(d, ov)
    return d


st.sidebar.title("CDEF Dashboard")
uploaded = st.sidebar.file_uploader(
    "Dalux CDEF export (Excel)", type=["xlsx", "xls", "xlsm", "json"])

with st.sidebar.expander("Defect-type categorisation", expanded=False):
    st.caption(
        "Older defects have no Defect Type in Dalux. A built-in categorisation "
        "file fills those gaps automatically. Upload an updated one here to "
        "refresh it (columns: `id`, `defectType` — or a Dalux export with a "
        "'Defect type' column).")
    ov_file = st.file_uploader("Updated categorisation (optional)",
                               type=["csv", "xlsx"], key="ovfile")

with st.sidebar.expander("Top 10 issues tracker", expanded=False):
    st.caption(
        "Upload the *Top 10 Project Issues* workbook (one sheet per contractor) "
        "to show the reply ratio alongside the Dalux closing ratios.")
    top10_file = st.file_uploader("Top 10 issues (optional)",
                                  type=["xlsx"], key="top10file")

_ov_bytes = ov_file.getvalue() if ov_file is not None else None

df = None
if uploaded is not None:
    try:
        df = load(file_bytes=uploaded.getvalue(), filename=uploaded.name,
                  override_bytes=_ov_bytes)
    except Exception as e:
        st.title("🏗️ CDEF — Construction Defects Dashboard")
        st.error(f"Couldn't read that file: {e}")
        st.stop()
else:
    import os
    defaults = ["Reports.xlsx", "Dalux_Dashboard_TS_Meeting.json"]
    local = next((d for d in defaults if os.path.exists(d)), None)
    if local:
        df = load(path=local, override_bytes=_ov_bytes)
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

dtypes = sorted(df["defectType"].unique())
sel_dtype = st.sidebar.multiselect("Defect type", dtypes, default=dtypes)

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
mask &= df["defectType"].isin(sel_dtype)
mask &= df["created"].dt.date.between(d0, d1)
fdf = df[mask].copy()

st.sidebar.markdown(f"**{len(fdf)}** of {len(df)} defects shown")

if fdf.empty:
    st.warning("No defects match the current filters. Add some tags back, or widen the date range.")
    st.stop()

weekly = cd.rolling_metrics(fdf)
last, prev = weekly.iloc[-1], weekly.iloc[-2] if len(weekly) > 1 else weekly.iloc[-1]

# ---------------- header ----------------
st.title("🏗️ CDEF — Construction Defects Dashboard")
st.caption(f"EVE Factory Project Debrecen · latest reporting week {last['ISO week']} (Fri–Thu) "
           f"· {len(fdf)} defects in view")


# ---------------- reusable render helpers ----------------
def render_kpis(_df, _weekly, extra_rr_ceh=False):
    _last = _weekly.iloc[-1]
    wk = _last["ISO week"]
    n_cols = 6 if extra_rr_ceh else 4
    cols = st.columns(n_cols)
    k1, k2, k3, k4 = cols[0], cols[1], cols[2], cols[3]
    k1.metric("Total defects (to date)", int(_last["Total defects (cumulative)"]),
              f"{int(_last['Total Δ vs prev']):+d} in latest week", delta_color="off",
              help="Cumulative count of all defects raised so far. The delta is how many were added in the latest reporting week.")
    k2.metric(f"New defects ({wk})", int(_last["New defects"]),
              f"{int(_last['New Δ vs prev']):+d} vs previous week", delta_color="inverse",
              help=f"Defects raised in the reporting week {wk} (Fri–Thu). More new defects is worse (red); fewer is better (green).")
    k3.metric(f"Accepted ({wk})", int(_last["Accepted defects"]),
              f"{int(_last['Accepted Δ vs prev']):+d} vs previous week", delta_color="normal",
              help=f"Defects approved/closed in the reporting week {wk} (Fri–Thu). More accepted is better (green); fewer is worse (red).")
    open_n = int((~_df["is_accepted"] & ~_df["status"].isin(["Discontinued", "Rejected"])).sum())
    k4.metric("Open defects (now)", open_n,
              help="Defects not yet approved, discontinued, or rejected — i.e. currently outstanding.")
    if extra_rr_ceh:
        rr = cd.reported_ready_ceh(_df)
        cols[4].metric("Reported ready · CÉH", rr,
                       help="Defects the contractor has reported ready (shown as 'Awaiting approval' in the contractor's Dalux) that are with CÉH, awaiting approval/closure.")
        rates = cd.week_resolution_rates(_df)
        cols[5].metric(
            f"Resolved of {wk} intake",
            f"{rates['pct_todate']:.0f}%",
            f"{rates['pct_inweek']:.0f}% within the week", delta_color="off",
            help=(f"Of the {rates['new']} defects raised in {wk}, "
                  f"{rates['resolved_todate']} are now approved ({rates['pct_todate']:.1f}%). "
                  f"{rates['resolved_inweek']} were approved within that same week "
                  f"({rates['pct_inweek']:.1f}%). 'Resolved' = status Approved or Approved, follow-up."))


def trend_block(_weekly, title, valcol, dcol, pcol, color, key="", good="down"):
    st.markdown(f"**{title}**")
    g1, g2 = st.columns([3, 2])
    with g1:
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        fig.add_bar(x=_weekly["ISO week"], y=_weekly[valcol], name=valcol,
                    marker_color=color, opacity=0.9)
        fig.add_scatter(x=_weekly["ISO week"], y=_weekly[dcol], name="Δ vs prev week",
                        mode="lines+markers", line=dict(color="#8A94A6", width=2),
                        marker=dict(size=4), secondary_y=True)
        fig.update_layout(height=300, margin=dict(l=10, r=10, t=10, b=10),
                          legend=dict(orientation="h", y=-0.25),
                          plot_bgcolor="white", paper_bgcolor="rgba(0,0,0,0)",
                          font=dict(family="Inter, Segoe UI, sans-serif", size=12, color=INK))
        fig.update_yaxes(title_text=valcol, secondary_y=False,
                         gridcolor="#EEF1F5", zerolinecolor="#E2E6EC")
        fig.update_yaxes(title_text="Δ", secondary_y=True, showgrid=False)
        fig.update_xaxes(showgrid=False)
        st.plotly_chart(fig, width="stretch", key=f"chart_{key}_{valcol}")
    with g2:
        tbl = _weekly[["ISO week", valcol, dcol, pcol]].tail(10).iloc[::-1].copy()
        tbl = tbl.rename(columns={"ISO week": "Week"})

        def style_delta(v):
            if pd.isna(v) or v == 0:
                return f"color: {MUTED}"
            if good == "neutral":
                return f"color: {MUTED}"
            improving = (v < 0) if good == "down" else (v > 0)
            return f"color: {GREEN}" if improving else f"color: {RED}"
        sty = (tbl.style
               .format({dcol: "{:+.0f}", pcol: "{:+.1f}%", valcol: "{:.0f}"}, na_rep="–")
               .map(style_delta, subset=[dcol]))
        st.dataframe(sty, hide_index=True, width="stretch", height=300,
                     key=f"tbl_{key}_{valcol}")


def all_trends(_weekly, key=""):
    trend_block(_weekly, "New defects per week (Fri–Thu)", "New defects",
                "New Δ vs prev", "New Δ %", BLUE, key=key, good="down")
    trend_block(_weekly, "Accepted defects per week (Fri–Thu)", "Accepted defects",
                "Accepted Δ vs prev", "Accepted Δ %", GREEN, key=key, good="up")
    trend_block(_weekly, "Total defects (cumulative) per week (Fri–Thu)",
                "Total defects (cumulative)", "Total Δ vs prev", "Total Δ %", NAVY,
                key=key, good="neutral")


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
    cweekly = cd.rolling_metrics(cdf)
    st.subheader(f"📌 {focus} — {len(cdf)} defects")
    render_kpis(cdf, cweekly)
    st.markdown("")
    # status breakdown for this contractor
    sc = cdf["status"].value_counts()
    chips = "  ·  ".join(f"**{v}** {k}" for k, v in sc.items())
    st.caption(chips)

    # defect-type breakdown for this contractor (all defects, not just this week)
    _ftypes = (cdf.groupby("defectType")
               .agg(total=("id", "count"), resolved=("is_accepted", "sum"))
               .reset_index())
    if not _ftypes.empty:
        _ftypes["open"] = _ftypes.apply(
            lambda r: int(((cdf["defectType"] == r["defectType"])
                           & ~cdf["is_accepted"]
                           & ~cdf["status"].isin(["Discontinued", "Rejected"])).sum()),
            axis=1)
        _ftypes["pct"] = (_ftypes["resolved"] / _ftypes["total"] * 100)
        _ftypes = _ftypes.sort_values("total", ascending=False)

        _tc_map = {"Quality defect": "#3E6DB5",
                   "Permit-related issue": "#E0A200",
                   "Appearance defect": "#7E57C2",
                   "Pre-categorisation (closed)": "#B6BECC"}

        st.markdown(f"**Defect types — {focus}**")
        _fcols = st.columns(min(len(_ftypes), 4))
        for _i, (_, _r) in enumerate(_ftypes.iterrows()):
            _col = _tc_map.get(_r["defectType"], "#B6BECC")
            _pct = _r["pct"]
            with _fcols[_i % len(_fcols)]:
                st.markdown(
                    f"""
                    <div style="background:#FFFFFF;border:1px solid {LINE};
                                border-left:6px solid {_col};border-radius:12px;
                                padding:14px 18px;margin-bottom:12px;
                                box-shadow:0 1px 3px rgba(16,24,40,0.05);">
                      <div style="font-size:1rem;font-weight:700;color:{INK};
                                  line-height:1.3;min-height:2.6em;">
                        {_r['defectType']}</div>
                      <div style="display:flex;align-items:baseline;gap:10px;
                                  margin:6px 0 8px 0;">
                        <span style="font-size:2rem;font-weight:800;color:{_col};
                                     line-height:1;">{int(_r['total'])}</span>
                        <span style="font-size:0.92rem;color:{MUTED};">defects</span>
                      </div>
                      <div style="background:#EEF1F5;border-radius:999px;height:9px;
                                  overflow:hidden;margin-bottom:8px;">
                        <div style="width:{max(_pct, 2):.1f}%;background:{_col};
                                    height:100%;border-radius:999px;"></div>
                      </div>
                      <div style="font-size:0.95rem;color:{MUTED};">
                        <b style="color:{INK};">{_pct:.0f}%</b> resolved ·
                        <b style="color:{INK};">{int(_r['open'])}</b> open
                      </div>
                    </div>
                    """, unsafe_allow_html=True)

    st.markdown("**7-day trends — " + focus + "**")
    all_trends(cweekly, key="focus")
    st.divider()
    if not show_overall:
        st.stop()

# ---------------- overall KPIs: 3 stacked rows ----------------
if focus != "All contractors":
    st.subheader("📊 Overall — all contractors in view")

contractor_df = fdf[fdf["responsibility"] == "With Contractor"]
cehve_df = fdf[fdf["responsibility"] == "With CÉH/EVE"]

st.markdown("#### Total")
render_kpis(fdf, weekly, extra_rr_ceh=True)

st.markdown("#### 🏗️ With Contractor")
if contractor_df.empty:
    st.caption("No defects currently with the contractor in this view.")
else:
    render_kpis(contractor_df, cd.rolling_metrics(contractor_df))

st.markdown("#### 🏢 With CÉH/EVE")
st.caption("Defects whose Role starts with *CÉH* or *EVE* — with us for review/action.")
if cehve_df.empty:
    st.caption("No defects currently with CÉH/EVE in this view.")
else:
    render_kpis(cehve_df, cd.rolling_metrics(cehve_df))

st.divider()

# ---------------- Resolution by defect type ----------------
_rates = cd.week_resolution_rates(fdf)
st.subheader(f"Resolution by defect type — {_rates['label']} intake")

if not _rates["by_type"]:
    st.info("No defects were raised in the latest reporting week.")
else:
    # Headline summary strip
    _tot_pct = _rates["pct_todate"]
    _bar_col = GREEN if _tot_pct >= 60 else ("#E0A200" if _tot_pct >= 25 else "#C0392B")
    _tint = ("#EAF5EE" if _tot_pct >= 60 else
             ("#FDF4E3" if _tot_pct >= 25 else "#FBEDEB"))
    st.markdown(
        f"""
        <div style="background:{_tint};border:1px solid {LINE};
                    border-left:7px solid {_bar_col};
                    border-radius:14px;padding:22px 26px;margin-bottom:22px;">
          <div style="font-size:0.95rem;color:{MUTED};text-transform:uppercase;
                      letter-spacing:0.8px;font-weight:700;">Week intake resolved</div>
          <div style="display:flex;align-items:baseline;gap:22px;margin-top:10px;
                      flex-wrap:wrap;">
            <div style="font-size:3.6rem;font-weight:800;color:{_bar_col};
                        line-height:1;">{_tot_pct:.0f}%</div>
            <div style="font-size:1.25rem;color:{INK};line-height:1.5;">
              <b style="font-size:1.5rem;">{_rates['resolved_todate']}</b>
              of <b style="font-size:1.5rem;">{_rates['new']}</b> defects raised
              this week are resolved<br>
              <span style="color:{MUTED};font-size:1.05rem;">
                {_rates['pct_inweek']:.0f}% were closed within the week itself</span>
            </div>
          </div>
        </div>
        """, unsafe_allow_html=True)

    # One card per defect type
    _cards = _rates["by_type"]
    _cols = st.columns(min(len(_cards), 4))
    _wk_total = max(_rates["new"], 1)
    for _i, _t in enumerate(_cards):
        _p = _t["pct_todate"]
        _c = GREEN if _p >= 60 else ("#E0A200" if _p >= 25 else "#C0392B")
        _ct = ("#EAF5EE" if _p >= 60 else ("#FDF4E3" if _p >= 25 else "#FBEDEB"))
        _share = _t["new"] / _wk_total * 100
        with _cols[_i % len(_cols)]:
            st.markdown(
                f"""
                <div style="background:#FFFFFF;border:1px solid {LINE};
                            border-radius:14px;overflow:hidden;height:100%;
                            box-shadow:0 2px 6px rgba(16,24,40,0.06);">
                  <div style="background:{_ct};border-bottom:1px solid {LINE};
                              padding:14px 18px;">
                    <div style="font-size:1.1rem;font-weight:700;color:{INK};
                                line-height:1.3;min-height:2.6em;">{_t['type']}</div>
                    <div style="font-size:0.92rem;color:{MUTED};margin-top:2px;">
                      {_share:.0f}% of week intake</div>
                  </div>
                  <div style="padding:18px;">
                    <div style="font-size:3rem;font-weight:800;color:{_c};
                                line-height:1;">{_p:.0f}%</div>
                    <div style="font-size:0.95rem;color:{MUTED};margin-top:4px;">
                      resolved to date</div>
                    <div style="background:#EEF1F5;border-radius:999px;height:12px;
                                overflow:hidden;margin:14px 0 14px 0;">
                      <div style="width:{max(_p,2):.1f}%;background:{_c};height:100%;
                                  border-radius:999px;"></div>
                    </div>
                    <div style="font-size:1.15rem;color:{INK};font-weight:600;">
                      {_t['resolved_todate']} <span style="color:{MUTED};
                      font-weight:400;">of</span> {_t['new']}
                      <span style="color:{MUTED};font-weight:400;">raised</span>
                    </div>
                    <div style="font-size:0.95rem;color:{MUTED};margin-top:6px;
                                padding-top:10px;border-top:1px solid {LINE};">
                      {_t['pct_inweek']:.0f}% closed within the week
                    </div>
                  </div>
                </div>
                """, unsafe_allow_html=True)

    st.markdown(
        f"""<div style="font-size:0.95rem;color:{MUTED};margin-top:16px;">
        Resolved = status <i>Approved</i> or <i>Approved, follow-up</i>.
        Colour bands: <b style="color:{GREEN};">green ≥60%</b> ·
        <b style="color:#E0A200;">amber ≥25%</b> ·
        <b style="color:#C0392B;">red below 25%</b>.
        </div>""", unsafe_allow_html=True)

st.divider()

# ---------------- Resolution by contractor ----------------
st.subheader(f"Resolution by contractor — {_rates['label']} intake")

if not _rates.get("by_contractor_type"):
    st.info("No defects were raised in the latest reporting week.")
else:
    _bct = pd.DataFrame(_rates["by_contractor_type"])
    _bc = pd.DataFrame(_rates["by_contractor"]).sort_values("new", ascending=False)

    _type_colours = {"Quality defect": "#3E6DB5",
                     "Permit-related issue": "#E0A200",
                     "Appearance defect": "#7E57C2",
                     "Pre-categorisation (closed)": "#B6BECC"}

    st.markdown(
        f"""<div style="font-size:1.05rem;color:{INK};margin-bottom:16px;">
        For each contractor, the share of <b>{_rates['label']}</b> defects now resolved,
        split by defect type. Legend: """
        + " &nbsp;·&nbsp; ".join(
            f'<span style="color:{c};font-weight:700;">■</span> {t}'
            for t, c in _type_colours.items() if t in set(_bct["type"]))
        + "</div>", unsafe_allow_html=True)

    _ncols = 2
    _rows = [_bc.iloc[i:i + _ncols] for i in range(0, len(_bc), _ncols)]
    for _chunk in _rows:
        _cols = st.columns(_ncols)
        for _j, (_, _c) in enumerate(_chunk.iterrows()):
            _name = _c["contractor"]
            _cp = _c["pct_todate"]
            _hdr_col = GREEN if _cp >= 60 else ("#E0A200" if _cp >= 25 else "#C0392B")
            _hdr_tint = ("#EAF5EE" if _cp >= 60 else
                         ("#FDF4E3" if _cp >= 25 else "#FBEDEB"))
            _sub = _bct[_bct["contractor"] == _name].sort_values("new", ascending=False)

            _rows_html = ""
            for _, _t in _sub.iterrows():
                _tc = _type_colours.get(_t["type"], "#B6BECC")
                _tp = _t["pct_todate"]
                _rows_html += f"""
                  <div style="margin-bottom:14px;">
                    <div style="display:flex;justify-content:space-between;
                                align-items:baseline;margin-bottom:5px;">
                      <span style="font-size:1rem;font-weight:600;color:{INK};">
                        <span style="color:{_tc};">■</span> {_t['type']}</span>
                      <span style="font-size:1.15rem;font-weight:700;color:{_tc};">
                        {_tp:.0f}%</span>
                    </div>
                    <div style="background:#EEF1F5;border-radius:999px;height:11px;
                                overflow:hidden;">
                      <div style="width:{max(_tp, 2):.1f}%;background:{_tc};
                                  height:100%;border-radius:999px;"></div>
                    </div>
                    <div style="font-size:0.92rem;color:{MUTED};margin-top:4px;">
                      {int(_t['resolved_todate'])} resolved of
                      <b style="color:{INK};">{int(_t['new'])}</b> raised
                    </div>
                  </div>"""

            with _cols[_j]:
                st.markdown(
                    f"""
                    <div style="background:#FFFFFF;border:1px solid {LINE};
                                border-radius:14px;overflow:hidden;margin-bottom:18px;
                                box-shadow:0 2px 6px rgba(16,24,40,0.06);">
                      <div style="background:{_hdr_tint};border-bottom:1px solid {LINE};
                                  padding:14px 20px;display:flex;
                                  justify-content:space-between;align-items:baseline;">
                        <div>
                          <div style="font-size:1.25rem;font-weight:700;color:{INK};">
                            {_name}</div>
                          <div style="font-size:0.95rem;color:{MUTED};margin-top:2px;">
                            {int(_c['new'])} raised this week</div>
                        </div>
                        <div style="text-align:right;">
                          <div style="font-size:2.2rem;font-weight:800;
                                      color:{_hdr_col};line-height:1;">
                            {_cp:.0f}%</div>
                          <div style="font-size:0.85rem;color:{MUTED};">resolved</div>
                        </div>
                      </div>
                      <div style="padding:18px 20px 6px 20px;">{_rows_html}</div>
                    </div>
                    """, unsafe_allow_html=True)

    with st.expander("Full contractor × defect type table"):
        _tbl = _bct[["contractor", "type", "new", "resolved_todate",
                     "pct_todate", "pct_inweek"]].copy()
        _tbl.columns = ["Contractor", "Defect type", "Raised", "Resolved",
                        "Resolved %", "In-week %"]
        st.dataframe(
            _tbl, hide_index=True, width="stretch",
            height=min(520, max(200, 40 * (len(_tbl) + 1))),
            column_config={
                "Raised": st.column_config.NumberColumn("Raised", format="%d"),
                "Resolved": st.column_config.NumberColumn("Resolved", format="%d"),
                "Resolved %": st.column_config.ProgressColumn(
                    "Resolved %", format="%.0f%%", min_value=0, max_value=100),
                "In-week %": st.column_config.NumberColumn("In-week %", format="%.0f%%"),
            })

    st.markdown(
        f"""<div style="font-size:0.95rem;color:{MUTED};margin-top:6px;">
        Resolved = status <i>Approved</i> or <i>Approved, follow-up</i>.
        Header % is the contractor's overall rate for the week; the rows below split it
        by defect type. Check the counts — a single defect reads as 0% or 100%.
        </div>""", unsafe_allow_html=True)

st.divider()

# ---------------- Contractor performance overview (3 ratios) ----------------
st.subheader("Contractor performance overview")

_t10 = None
if top10_file is not None:
    try:
        _t10 = cd.load_top10_ratios(io.BytesIO(top10_file.getvalue()))
    except Exception as e:
        st.warning(f"Couldn't read the Top 10 workbook: {e}")

_csum_all = cd.contractor_summary(fdf)
_wk_map = {c["contractor"]: c for c in _rates.get("by_contractor", [])}

_perf = []
for _, _row in _csum_all.iterrows():
    _c = _row["contractor"]
    _wk = _wk_map.get(_c)
    _t10row = None
    if _t10 is not None and not _t10.empty:
        _m = _t10[_t10["contractor"] == _c]
        if not _m.empty:
            _t10row = _m.iloc[0]
    # only show contractors that have something meaningful in at least one metric
    if _t10row is None and _wk is None and _row["Total"] < 5:
        continue
    _perf.append({
        "Contractor": _c,
        "Top 10 reply %": (float(_t10row["reply_pct"]) if _t10row is not None else None),
        f"{_rates['label']} closing %": (_wk["pct_todate"] if _wk else None),
        "Overall closing %": float(_row["Acceptance %"]),
        "_total": int(_row["Total"]),
    })

if not _perf:
    st.info("No contractor data available for this view.")
else:
    _pf = pd.DataFrame(_perf).sort_values("_total", ascending=False)
    _wk_col = f"{_rates['label']} closing %"
    _series = [("Top 10 reply %", "#2E5C99"),
               (_wk_col, "#E07B39"),
               ("Overall closing %", "#2E7D46")]

    figp = go.Figure()
    for _cn, _col in _series:
        _vals = _pf[_cn]
        if _vals.isna().all():
            continue
        figp.add_bar(
            x=_pf["Contractor"], y=_vals, name=_cn, marker_color=_col,
            text=[("" if pd.isna(v) else f"{v:.0f}%") for v in _vals],
            textposition="outside", cliponaxis=False,
            hovertemplate="<b>%{x}</b><br>" + _cn + ": %{y:.0f}%<extra></extra>")
    figp.update_layout(
        barmode="group", bargap=0.25, bargroupgap=0.06,
        height=460, margin=dict(l=10, r=10, t=30, b=10),
        yaxis_title="", legend=dict(orientation="h", y=-0.16),
        plot_bgcolor="white", paper_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter, Segoe UI, sans-serif", size=13, color=INK))
    figp.update_yaxes(range=[0, 112], ticksuffix="%",
                      gridcolor="#EEF1F5", zerolinecolor="#E2E6EC")
    figp.update_xaxes(showgrid=False)
    st.plotly_chart(figp, width="stretch", key="chart_contractor_performance")

    _disp = _pf[["Contractor", "Top 10 reply %", _wk_col, "Overall closing %"]]
    st.dataframe(
        _disp, hide_index=True, width="stretch",
        column_config={c: st.column_config.NumberColumn(c, format="%.0f%%")
                       for c in _disp.columns if c != "Contractor"})

    if _t10 is None:
        st.caption("💡 Upload the *Top 10 issues* workbook in the sidebar to add the "
                   "reply-ratio bars.")
    st.markdown(
        f"""<div style="font-size:0.95rem;color:{MUTED};margin-top:8px;">
        <b>Top 10 reply %</b> — issues in the Top-10 tracker marked approved/resolved,
        out of those listed. &nbsp;
        <b>{_rates['label']} closing %</b> — defects raised in the reporting week that
        are now resolved. &nbsp;
        <b>Overall closing %</b> — all defects to date that are resolved.
        Resolved = <i>Approved</i> / <i>Approved, follow-up</i>.
        </div>""", unsafe_allow_html=True)

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
    # Meaningful, calm status palette: open shades of blue, resolved greens, muted for closed-out
    palette = {"Ongoing": "#3E6DB5", "Reported ready": "#8FB0DC",
               "Approved": "#2E7D46", "Approved, follow-up": "#66A97F",
               "Rejected": "#C0392B", "Discontinued": "#B6BECC"}
    for status in pivot.columns:
        fig.add_bar(y=pivot.index, x=pivot[status], name=status,
                    orientation="h", marker_color=palette.get(status, "#B6BECC"))
    fig.update_layout(barmode="stack", height=max(320, 42 * len(pivot)),
                      margin=dict(l=10, r=10, t=6, b=10),
                      legend=dict(orientation="h", y=-0.15),
                      xaxis_title="Defects", plot_bgcolor="white",
                      paper_bgcolor="rgba(0,0,0,0)",
                      font=dict(family="Inter, Segoe UI, sans-serif", size=12, color=INK))
    fig.update_xaxes(gridcolor="#EEF1F5", zerolinecolor="#E2E6EC")
    fig.update_yaxes(showgrid=False)
    st.plotly_chart(fig, width="stretch")

with cc2:
    disp = csum.rename(columns={"contractor": "Contractor"})
    st.dataframe(
        disp,
        hide_index=True, width="stretch",
        height=max(320, 42 * len(disp)),
        column_config={
            "Total": st.column_config.NumberColumn("Total", format="%d"),
            "Accepted": st.column_config.NumberColumn("Accepted", format="%d"),
            "Open": st.column_config.NumberColumn("Open", format="%d"),
            "Acceptance %": st.column_config.ProgressColumn(
                "Acceptance %", format="%.1f%%", min_value=0, max_value=100),
        })

st.divider()

# ---------------- Rolling trends (overall) ----------------
st.subheader("Weekly trends (Fri–Thu) — with previous-week comparison")
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
    return be.build_report_bytes(df, top10=_t10)

# Cache key tied to the dataset size + latest week so it rebuilds when data changes
_token = f"{len(df)}-{weekly.iloc[-1]['ISO week']}-{0 if _t10 is None else len(_t10)}"
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
- **7-day periods** are anchored on **today**: the latest window covers the last 7 days ending today, the one before it the 7 days before that, and so on — not fixed Monday–Sunday calendar weeks.
- **Reporting weeks** run **Friday 00:00 → Thursday 24:00**. The latest window shown is the last *complete* Fri–Thu week.
- **New defects** = count by *date created*, within each Fri–Thu week.
- **Accepted defects** = status *Approved* / *Approved, follow-up*, dated by *last-modified* (approval) date, within each Fri–Thu week.
  Dalux does not expose a dedicated approval date in the export, so the modified date is used as the closure timestamp.
- **Total defects** = cumulative count of all defects raised up to and including the end of that window.
- **Δ vs prev** = change against the previous Fri–Thu week.
""")
