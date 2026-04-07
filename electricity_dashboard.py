import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import re
from datetime import datetime

st.set_page_config(page_title="Electricity Usage Overview", layout="wide")

# ─── Thai electricity tariff (PEA TOU rates, Baht/kWh) ───────────────────────
ON_PEAK_RATE  = 4.1839   # Mon–Fri 09:00–22:00
OFF_PEAK_RATE = 2.6037   # All other times
FT_ADJ        = 0.0972   # Ft surcharge (per kWh)

# ─── Helpers ─────────────────────────────────────────────────────────────────
DAY_TH = {"อา": 6, "จ": 0, "อ": 1, "พ": 2, "พฤ": 3, "ศ": 4, "ส": 5}  # weekday (Mon=0)

def parse_date_col(raw: str):
    """Parse '01/03\n(อา)' → (datetime, weekday_int)"""
    match = re.match(r"(\d{2}/\d{2})\n\((.+)\)", raw)
    if not match:
        return None, None
    date_str, th_day = match.group(1), match.group(2)
    # Determine year from month context – simple heuristic using current year
    d, m = map(int, date_str.split("/"))
    year = datetime.now().year
    if m > datetime.now().month + 1:   # probably last year
        year -= 1
    try:
        dt = datetime(year, m, d)
    except ValueError:
        return None, None
    return dt, DAY_TH.get(th_day, dt.weekday())


@st.cache_data
def load_data(file_bytes: bytes):
    """Parse the Excel file and return a tidy DataFrame."""
    import io
    raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name="Clean Data", header=None)

    # Row 0 = dates, Row 1 = column headers (On/Off/Total), Rows 2+ = meter data
    date_cols = []     # list of dicts: {col_idx, date, weekday}
    for i in range(4, raw.shape[1], 3):
        val = raw.iloc[0, i]
        if pd.notna(val):
            dt, wd = parse_date_col(str(val))
            if dt:
                date_cols.append({"col_idx": i, "date": dt, "weekday": wd})

    # Build tidy rows
    records = []
    for row_i in range(2, raw.shape[0]):
        meter  = raw.iloc[row_i, 0]
        group  = raw.iloc[row_i, 1]
        subgrp = raw.iloc[row_i, 2]
        if pd.isna(meter) or pd.isna(group):
            continue
        for dc in date_cols:
            ci = dc["col_idx"]
            on  = pd.to_numeric(raw.iloc[row_i, ci],     errors="coerce")
            off = pd.to_numeric(raw.iloc[row_i, ci + 1], errors="coerce")
            tot = pd.to_numeric(raw.iloc[row_i, ci + 2], errors="coerce")
            records.append({
                "meter": str(meter),
                "department": str(group),
                "sub_group": str(subgrp),
                "date": dc["date"],
                "weekday": dc["weekday"],
                "on_peak": on if pd.notna(on) else 0.0,
                "off_peak": off if pd.notna(off) else 0.0,
                "total": tot if pd.notna(tot) else 0.0,
            })

    df = pd.DataFrame(records)
    df["date"] = pd.to_datetime(df["date"])
    df["week_num"] = df["date"].dt.isocalendar().week.astype(int)
    df["year"]     = df["date"].dt.isocalendar().year.astype(int)
    df["year_week"] = df["year"].astype(str) + "-W" + df["week_num"].astype(str).str.zfill(2)
    return df


def get_complete_weeks(df: pd.DataFrame):
    """Return list of year_week strings that have all 7 days (Sun–Sat)."""
    # A complete week has 7 distinct dates
    wk_days = df.groupby("year_week")["date"].nunique()
    return wk_days[wk_days == 7].index.tolist()


def week_agg(df: pd.DataFrame, year_week: str, dept_filter=None):
    """Aggregate on_peak / off_peak / total for a given year_week."""
    sub = df[df["year_week"] == year_week]
    if dept_filter and dept_filter != "Factory (All)":
        sub = sub[sub["department"] == dept_filter]
    return sub[["on_peak", "off_peak", "total"]].sum()


def dept_week_agg(df: pd.DataFrame, year_week: str):
    """Aggregate by department for a given year_week."""
    sub = df[df["year_week"] == year_week]
    return sub.groupby("department")[["on_peak", "off_peak", "total"]].sum().reset_index()


# ─── Custom CSS ──────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main-title {font-size:26px; font-weight:700; text-align:center; color:#1a237e; margin-bottom:20px;}
    .kpi-card {
        background:#fff; border-radius:10px; padding:18px 22px;
        box-shadow:0 2px 8px rgba(0,0,0,0.1); text-align:center;
    }
    .kpi-label {font-size:13px; color:#555; font-weight:600; margin-bottom:4px;}
    .kpi-value {font-size:28px; font-weight:800; color:#1a237e;}
    .kpi-sub   {font-size:13px; color:#888; margin-top:3px;}
    .kpi-badge-on  {color:#e65100; font-weight:700;}
    .kpi-badge-off {color:#2e7d32; font-weight:700;}
    .section-header {
        font-size:16px; font-weight:700; color:#1a237e;
        border-left:4px solid #1565c0; padding-left:10px; margin:24px 0 12px;
    }
</style>
""", unsafe_allow_html=True)

# ─── Sidebar / Upload ─────────────────────────────────────────────────────────
with st.sidebar:
    st.image("https://img.icons8.com/fluency/60/lightning-bolt.png", width=60)
    st.markdown("## ⚡ Electricity Dashboard")
    uploaded = st.file_uploader("อัปโหลดไฟล์ Clean Data (.xlsx)", type=["xlsx"])
    st.markdown("---")
    st.caption("อัตราค่าไฟ (MEA TOU)")
    st.caption(f"• On Peak+FT : {ON_PEAK_RATE + FT_ADJ:.4f} ฿/kWh")
    st.caption(f"• Off Peak+FT: {OFF_PEAK_RATE + FT_ADJ:.4f} ฿/kWh")

if not uploaded:
    st.markdown('<div class="main-title">Electricity Usage Overview</div>', unsafe_allow_html=True)
    st.info("👈 กรุณาอัปโหลดไฟล์ Excel (Clean Data) ที่แถบซ้ายมือเพื่อเริ่มต้นใช้งาน")
    st.stop()

# ─── Load & validate ─────────────────────────────────────────────────────────
df = load_data(uploaded.read())
complete_weeks = sorted(get_complete_weeks(df))

if len(complete_weeks) < 1:
    st.error("ไม่พบสัปดาห์ที่ครบ 7 วัน กรุณาตรวจสอบไฟล์ข้อมูล")
    st.stop()

latest_week  = complete_weeks[-1]
prev_week    = complete_weeks[-2] if len(complete_weeks) >= 2 else None

# ─── Section 1: KPI Cards ────────────────────────────────────────────────────
st.markdown('<div class="main-title">⚡ Electricity Usage Overview</div>', unsafe_allow_html=True)

agg_latest = week_agg(df, latest_week)
on_kwh  = agg_latest["on_peak"]
off_kwh = agg_latest["off_peak"]
total_kwh = on_kwh + off_kwh
cost = on_kwh * (ON_PEAK_RATE + FT_ADJ) + off_kwh * (OFF_PEAK_RATE + FT_ADJ)

on_pct  = on_kwh  / total_kwh * 100 if total_kwh else 0
off_pct = off_kwh / total_kwh * 100 if total_kwh else 0

# Compare with previous week
if prev_week:
    agg_prev = week_agg(df, prev_week)
    prev_total = agg_prev["on_peak"] + agg_prev["off_peak"]
    chg_total = (total_kwh - prev_total) / prev_total * 100 if prev_total else 0
    chg_on    = (on_kwh  - agg_prev["on_peak"])  / agg_prev["on_peak"]  * 100 if agg_prev["on_peak"]  else 0
    chg_off   = (off_kwh - agg_prev["off_peak"]) / agg_prev["off_peak"] * 100 if agg_prev["off_peak"] else 0

    def arrow(v): return ("🔺" if v >= 0 else "🔻") + f" {abs(v):.1f}%"
else:
    chg_total = chg_on = chg_off = 0
    def arrow(v): return ""

c1, c2, c3, c4 = st.columns(4)
with c1:
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-label">Total Energy This Week</div>
        <div class="kpi-value">{total_kwh:,.0f} <span style="font-size:16px">kWh</span></div>
        <div class="kpi-sub">สัปดาห์ {latest_week} &nbsp; {arrow(chg_total)}</div>
    </div>""", unsafe_allow_html=True)
with c2:
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-label">On Peak Usage</div>
        <div class="kpi-value kpi-badge-on">{on_kwh:,.0f} <span style="font-size:15px">kWh</span></div>
        <div class="kpi-sub">({on_pct:.0f}%) &nbsp; {arrow(chg_on)}</div>
    </div>""", unsafe_allow_html=True)
with c3:
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-label">Off Peak Usage</div>
        <div class="kpi-value kpi-badge-off">{off_kwh:,.0f} <span style="font-size:15px">kWh</span></div>
        <div class="kpi-sub">({off_pct:.0f}%) &nbsp; {arrow(chg_off)}</div>
    </div>""", unsafe_allow_html=True)
with c4:
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-label">Cost Estimate</div>
        <div class="kpi-value">{cost:,.0f} <span style="font-size:16px">บาท</span></div>
        <div class="kpi-sub">อัตรา TOU (รวม Ft)</div>
    </div>""", unsafe_allow_html=True)

# ─── Section 2: Weekly Usage Comparison ─────────────────────────────────────
st.markdown('<div class="section-header">📊 Weekly Usage Comparison</div>', unsafe_allow_html=True)

departments = sorted(df["department"].unique().tolist())
filter_options = ["Factory (All)"] + departments

col_filter, col_chart = st.columns([1, 3])
with col_filter:
    dept_sel = st.selectbox("🔍 เลือกแผนก / Factory", filter_options, index=0)

agg_cur  = week_agg(df, latest_week, dept_sel)
agg_prev_f = week_agg(df, prev_week, dept_sel) if prev_week else None

fig_weekly = go.Figure()
weeks_label = [f"สัปดาห์ที่แล้ว\n({prev_week})", f"สัปดาห์นี้\n({latest_week})"]

on_vals  = [agg_prev_f["on_peak"]  if agg_prev_f is not None else 0, agg_cur["on_peak"]]
off_vals = [agg_prev_f["off_peak"] if agg_prev_f is not None else 0, agg_cur["off_peak"]]

fig_weekly.add_trace(go.Bar(
    name="On Peak", x=weeks_label, y=on_vals,
    marker_color="#e65100", text=[f"{v:,.0f}" for v in on_vals], textposition="outside"
))
fig_weekly.add_trace(go.Bar(
    name="Off Peak", x=weeks_label, y=off_vals,
    marker_color="#1565c0", text=[f"{v:,.0f}" for v in off_vals], textposition="outside"
))
fig_weekly.update_layout(
    barmode="group", height=380,
    yaxis_title="kWh",
    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    margin=dict(t=40, b=20),
    plot_bgcolor="white",
    paper_bgcolor="white",
    title_text=f"On Peak vs Off Peak — {dept_sel}",
    title_font_size=14,
)
fig_weekly.update_yaxes(gridcolor="#f0f0f0")

with col_chart:
    st.plotly_chart(fig_weekly, use_container_width=True)

# ─── Section 3: Department Usage Breakdown ───────────────────────────────────
st.markdown('<div class="section-header">🏭 Department Usage Breakdown</div>', unsafe_allow_html=True)

dept_cur  = dept_week_agg(df, latest_week).set_index("department")
dept_prev = dept_week_agg(df, prev_week).set_index("department") if prev_week else None

all_depts = sorted(dept_cur.index.tolist())

# Build grouped horizontal bar chart per department
fig_dept = make_subplots(
    rows=len(all_depts), cols=1,
    shared_xaxes=True,
    subplot_titles=all_depts,
    vertical_spacing=0.04,
)

for i, dep in enumerate(all_depts, start=1):
    cur_on  = dept_cur.loc[dep, "on_peak"]  if dep in dept_cur.index  else 0
    cur_off = dept_cur.loc[dep, "off_peak"] if dep in dept_cur.index  else 0
    prv_on  = dept_prev.loc[dep, "on_peak"]  if (dept_prev is not None and dep in dept_prev.index) else 0
    prv_off = dept_prev.loc[dep, "off_peak"] if (dept_prev is not None and dep in dept_prev.index) else 0

    total_cur = cur_on + cur_off
    total_prv = prv_on + prv_off
    pct_chg = ((total_cur - total_prv) / total_prv * 100) if total_prv else 0
    arrow_str = f"▲{pct_chg:.0f}%" if pct_chg >= 0 else f"▼{abs(pct_chg):.0f}%"
    clr_arrow = "#e53935" if pct_chg >= 0 else "#43a047"

    show_legend = (i == 1)

    # Previous week bars (lighter)
    fig_dept.add_trace(go.Bar(
        name=f"สัปดาห์ก่อน On", x=[prv_on], y=["สัปดาห์ก่อน"],
        orientation="h", marker_color="#ffb74d",
        legendgroup="prev_on", showlegend=show_legend,
        hovertemplate=f"On Peak (ก่อน): %{{x:,.0f}} kWh<extra></extra>",
    ), row=i, col=1)
    fig_dept.add_trace(go.Bar(
        name=f"สัปดาห์ก่อน Off", x=[prv_off], y=["สัปดาห์ก่อน"],
        orientation="h", marker_color="#90caf9",
        legendgroup="prev_off", showlegend=show_legend,
        hovertemplate=f"Off Peak (ก่อน): %{{x:,.0f}} kWh<extra></extra>",
    ), row=i, col=1)

    # Current week bars
    fig_dept.add_trace(go.Bar(
        name="สัปดาห์นี้ On", x=[cur_on], y=["สัปดาห์นี้"],
        orientation="h", marker_color="#e65100",
        legendgroup="cur_on", showlegend=show_legend,
        hovertemplate=f"On Peak (นี้): %{{x:,.0f}} kWh<extra></extra>",
    ), row=i, col=1)
    fig_dept.add_trace(go.Bar(
        name="สัปดาห์นี้ Off", x=[cur_off], y=["สัปดาห์นี้"],
        orientation="h", marker_color="#1565c0",
        legendgroup="cur_off", showlegend=show_legend,
        hovertemplate=f"Off Peak (นี้): %{{x:,.0f}} kWh<extra></extra>",
    ), row=i, col=1)

    # Annotation: % change
    fig_dept.add_annotation(
        x=max(cur_on + cur_off, prv_on + prv_off) * 1.02,
        y=0.5, yref=f"y{i}", xref=f"x{i}",
        text=f"<b>{arrow_str}</b>",
        showarrow=False, font=dict(color=clr_arrow, size=11),
        xanchor="left",
    )

fig_dept.update_layout(
    barmode="stack",
    height=max(120 * len(all_depts), 600),
    legend=dict(orientation="h", yanchor="bottom", y=1.01, xanchor="right", x=1),
    margin=dict(l=100, r=100, t=60, b=20),
    plot_bgcolor="white",
    paper_bgcolor="white",
)
for i in range(1, len(all_depts) + 1):
    fig_dept.update_xaxes(gridcolor="#f0f0f0", row=i, col=1)

st.plotly_chart(fig_dept, use_container_width=True)

# ─── Footer ──────────────────────────────────────────────────────────────────
st.markdown("---")
st.caption(f"📅 ข้อมูลสัปดาห์ล่าสุด: **{latest_week}** | สัปดาห์ก่อนหน้า: **{prev_week or 'N/A'}** | ข้อมูลทั้งหมด {df['date'].nunique()} วัน")
