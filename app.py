"""
app.py
======
AI Schedule Evaluation Dashboard — main Streamlit application.

Upload an AI Schedule (predicted generation) file and a Meter Data (actual
generation) file for a single day of 15-minute blocks. The app matches the
two datasets block-wise, computes deviation and DSM penalty using piecewise
slab logic, and presents KPIs, interactive charts, a searchable/sortable
table, and a downloadable, professionally formatted Excel report.

Run with:  streamlit run app.py
"""

from __future__ import annotations

import datetime as _dt

import pandas as pd
import streamlit as st

from config import PLANT_CONFIGS, DEFAULT_PLANT, PlantConfig
from utils import parse_energy_file, merge_datasets, FileValidationError
from calculator import evaluate_schedule, build_summary, day_summary_metrics
from graphs import all_figures
from excel_export import build_excel_report

st.set_page_config(
    page_title="AI Schedule Evaluation Dashboard",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---------------------------------------------------------------------------
# Green & white theme
# ---------------------------------------------------------------------------
DARK_GREEN = "#0F3D2E"
GREEN = "#16A34A"
LIGHT_GREEN = "#F0FDF4"
MID_GREEN = "#BBF7D0"
AMBER = "#F59E0B"

st.markdown(
    f"""
    <style>
        .block-container {{padding-top: 1.2rem; padding-bottom: 2rem; max-width: 1300px;}}
        html, body, [class*="css"] {{ font-family: 'Segoe UI', 'Inter', sans-serif; }}

        /* ---- Hero banner ---- */
        .hero-banner {{
            background: linear-gradient(120deg, {DARK_GREEN} 0%, {GREEN} 100%);
            border-radius: 16px;
            padding: 28px 32px;
            margin-bottom: 24px;
            box-shadow: 0 8px 24px rgba(15, 61, 46, 0.18);
        }}
        .hero-banner h1 {{
            color: #FFFFFF; margin: 0; font-size: 2rem; font-weight: 800;
            letter-spacing: -0.5px;
        }}
        .hero-banner p {{
            color: {MID_GREEN}; margin: 6px 0 0 0; font-size: 0.98rem;
        }}
        .hero-badge {{
            display: inline-block; background: rgba(255,255,255,0.18);
            color: #FFFFFF; padding: 3px 12px; border-radius: 999px;
            font-size: 0.78rem; font-weight: 700; letter-spacing: 0.4px;
            margin-top: 10px;
        }}

        /* ---- Section headers ---- */
        .section-header {{
            display: flex; align-items: center; gap: 8px;
            font-size: 1.25rem; font-weight: 700; color: {DARK_GREEN};
            margin: 6px 0 14px 0; padding-bottom: 8px;
            border-bottom: 2px solid {MID_GREEN};
        }}

        /* ---- KPI metric cards ---- */
        div[data-testid="stMetric"] {{
            background: #FFFFFF;
            border: 1px solid {MID_GREEN};
            border-left: 4px solid {GREEN};
            border-radius: 12px;
            padding: 14px 16px 10px 16px;
            box-shadow: 0 2px 8px rgba(15, 61, 46, 0.06);
            transition: transform 0.15s ease, box-shadow 0.15s ease;
        }}
        div[data-testid="stMetric"]:hover {{
            transform: translateY(-2px);
            box-shadow: 0 6px 16px rgba(15, 61, 46, 0.14);
        }}
        div[data-testid="stMetricLabel"] {{ font-weight: 600; color: {DARK_GREEN}; }}
        div[data-testid="stMetricValue"] {{ color: #0F172A; }}

        h1, h2, h3 {{ color: {DARK_GREEN}; }}

        /* ---- Sidebar ---- */
        section[data-testid="stSidebar"] {{
            background: linear-gradient(180deg, {LIGHT_GREEN} 0%, #FFFFFF 220px);
            border-right: 1px solid {MID_GREEN};
        }}
        section[data-testid="stSidebar"] h2 {{ color: {DARK_GREEN}; font-size: 1.05rem; }}

        /* ---- Buttons ---- */
        .stButton > button, .stDownloadButton > button {{
            border-radius: 10px; font-weight: 700; border: none;
        }}
        .stButton > button[kind="primary"], .stDownloadButton > button[kind="primary"] {{
            background: linear-gradient(120deg, {GREEN} 0%, {DARK_GREEN} 100%);
            box-shadow: 0 4px 12px rgba(22, 163, 74, 0.28);
        }}

        /* ---- Tabs ---- */
        .stTabs [data-baseweb="tab-list"] {{ gap: 4px; border-bottom: 1px solid {MID_GREEN}; }}
        .stTabs [aria-selected="true"] {{ color: {DARK_GREEN} !important; font-weight: 700; }}
        .stTabs [data-baseweb="tab-highlight"] {{ background-color: {GREEN} !important; }}

        /* ---- Expanders ---- */
        details {{ border: 1px solid {MID_GREEN} !important; border-radius: 10px !important; }}
        summary {{ color: {DARK_GREEN}; font-weight: 600; }}

        /* ---- Dataframe header tint ---- */
        [data-testid="stDataFrame"] {{ border: 1px solid {MID_GREEN}; border-radius: 8px; }}

        /* ---- Footer ---- */
        .footer-bar {{
            text-align: center; color: #FFFFFF; background: {DARK_GREEN};
            border-radius: 10px; padding: 10px 16px; font-size: 0.85rem; margin-top: 12px;
        }}
    </style>
    """,
    unsafe_allow_html=True,
)


# ---------------------------------------------------------------------------
# Sidebar — plant configuration & file upload
# ---------------------------------------------------------------------------
with st.sidebar:
    st.markdown("## ⚡ Plant Configuration")
    plant_key = st.selectbox("Plant", options=list(PLANT_CONFIGS.keys()),
                              index=list(PLANT_CONFIGS.keys()).index(DEFAULT_PLANT))
    base_plant = PLANT_CONFIGS[plant_key]

    capacity_mw = st.number_input(
        "Installed Capacity (MW)", min_value=0.01,
        value=float(base_plant.installed_capacity_mw), step=0.1, format="%.3f",
    )
    plant = PlantConfig(
        name=base_plant.name,
        installed_capacity_mw=capacity_mw,
        block_minutes=base_plant.block_minutes,
        dsm_slabs=base_plant.dsm_slabs,
    )
    st.caption(f"Scheduling block: {plant.block_minutes} minutes "
               f"({plant.blocks_per_day} blocks/day)")

    with st.expander("DSM Slab Structure"):
        slab_rows = [
            {"Band": (f"{s.lower:.0f}% – {s.upper:.0f}%" if s.upper is not None
                      else f"Above {s.lower:.0f}%"),
             "Rate (Rs/kWh)": s.rate}
            for s in plant.dsm_slabs
        ]
        st.dataframe(pd.DataFrame(slab_rows), hide_index=True, use_container_width=True)

    st.markdown("---")
    st.markdown("## 📂 Upload Data")
    schedule_file = st.file_uploader(
        "AI Schedule File (Predicted Generation)", type=["csv", "xlsx", "xls"],
        help="Must contain Date/Time and Predicted MW (or kW) for each block.",
    )
    meter_file = st.file_uploader(
        "Meter Data File (Actual Generation)", type=["csv", "xlsx", "xls"],
        help="Must contain Date/Time and Actual MW (or kW) for each block.",
    )

    with st.expander("Advanced: timestamp alignment"):
        st.caption(
            "If your files use a Time/TimeStamp column instead of an explicit "
            "Block number, choose whether that timestamp marks the START or "
            "END of its 15-minute block."
        )
        sched_ts_mode = st.radio("AI Schedule file timestamp marks:", ["start", "end"],
                                  index=0, horizontal=True, key="sched_ts_mode")
        meter_ts_mode = st.radio("Meter Data file timestamp marks:", ["start", "end"],
                                  index=0, horizontal=True, key="meter_ts_mode")

    process_clicked = st.button("🚀 Process & Evaluate", type="primary", use_container_width=True,
                                 disabled=not (schedule_file and meter_file))

st.markdown(
    f"""
    <div class="hero-banner">
        <h1>⚡ AI Schedule Evaluation Dashboard</h1>
        <p>Evaluate AI-generated power schedules against meter data and auto-generate the DSM penalty report.</p>
        <span class="hero-badge">🌱 {plant.name} &nbsp;·&nbsp; {plant.installed_capacity_mw:g} MW &nbsp;·&nbsp; {plant.block_minutes}-min blocks</span>
    </div>
    """,
    unsafe_allow_html=True,
)

if not (schedule_file and meter_file):
    st.info(
        "👈 Upload both the **AI Schedule File** and the **Meter Data File** in the sidebar, "
        "then click **Process & Evaluate** to run the full DSM analysis."
    )
    st.caption("Once files are processed, an expandable **DSM Formula Explainer** tab walks through every formula with a worked example.")
    st.stop()


# ---------------------------------------------------------------------------
# Processing pipeline
# ---------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def _run_pipeline(sched_bytes, sched_name, meter_bytes, meter_name,
                   block_minutes, sched_ts_mode, meter_ts_mode,
                   capacity_mw, plant_name, slabs_key):
    import io
    sched_file = io.BytesIO(sched_bytes)
    sched_file.name = sched_name
    meter_file_ = io.BytesIO(meter_bytes)
    meter_file_.name = meter_name

    sched_res = parse_energy_file(sched_file, role="schedule",
                                   block_minutes=block_minutes, timestamp_marks=sched_ts_mode)
    meter_res = parse_energy_file(meter_file_, role="meter",
                                   block_minutes=block_minutes, timestamp_marks=meter_ts_mode)
    merged, warnings = merge_datasets(sched_res, meter_res)
    return merged, warnings


if process_clicked or "result_df" in st.session_state:
    try:
        if process_clicked:
            with st.spinner("Reading files, matching blocks and calculating DSM penalties..."):
                schedule_file.seek(0)
                meter_file.seek(0)
                sched_bytes = schedule_file.read()
                meter_bytes = meter_file.read()
                merged, warnings = _run_pipeline(
                    sched_bytes, schedule_file.name, meter_bytes, meter_file.name,
                    plant.block_minutes, st.session_state.sched_ts_mode, st.session_state.meter_ts_mode,
                    plant.installed_capacity_mw, plant.name,
                    tuple((s.lower, s.upper, s.rate) for s in plant.dsm_slabs),
                )
                result_df = evaluate_schedule(merged, plant)
                summary = build_summary(result_df, plant)

            # Build the Excel report ONLY here, once per "Process & Evaluate"
            # click -- NOT on every script rerun. Streamlit reruns the whole
            # script on every widget interaction (filters, search, sort,
            # pagination, etc.), so if this lived outside the process_clicked
            # branch it would rebuild the whole report on every single click.
            # All 6 charts are native Excel chart objects built straight from
            # worksheet formulas (see excel_export.py) -- no image rendering
            # / kaleido involved, so this step is now fast and has no
            # external-process dependency.
            with st.spinner("Building Excel report..."):
                excel_bytes = build_excel_report(result_df, summary, plant)

            st.session_state["result_df"] = result_df
            st.session_state["merged_df"] = merged
            st.session_state["warnings"] = warnings
            st.session_state["summary"] = summary
            st.session_state["excel_bytes"] = excel_bytes
        else:
            result_df = st.session_state["result_df"]
            merged = st.session_state["merged_df"]
            warnings = st.session_state["warnings"]
            summary = st.session_state["summary"]
            excel_bytes = st.session_state["excel_bytes"]
    except FileValidationError as exc:
        st.error(f"❌ {exc}")
        st.stop()
    except Exception as exc:  # pragma: no cover
        st.error(f"❌ Unexpected error while processing files: {exc}")
        st.stop()
else:
    st.stop()

if warnings:
    with st.expander(f"⚠️ {len(warnings)} data validation note(s)", expanded=False):
        for w in warnings:
            st.warning(w)

# ---------------------------------------------------------------------------
# KPI Cards
# ---------------------------------------------------------------------------
st.markdown('<div class="section-header">📊 Summary</div>', unsafe_allow_html=True)
r1 = st.columns(4)
r1[0].metric("Total Blocks", summary["Total Blocks"])
r1[1].metric("Total Scheduled Energy", f"{summary['Total Scheduled Energy (MWh)']:.3f} MWh")
r1[2].metric("Total Actual Energy", f"{summary['Total Actual Energy (MWh)']:.3f} MWh")
r1[3].metric("Total DSM Penalty", f"₹{summary['Total DSM Penalty (Rs)']:.2f}")

r2 = st.columns(4)
r2[0].metric("Max Positive Deviation", f"{summary['Maximum Positive Deviation (MW)']:.3f} MW")
r2[1].metric("Max Negative Deviation", f"{summary['Maximum Negative Deviation (MW)']:.3f} MW")
r2[2].metric("Average Deviation", f"{summary['Average Deviation (MW)']:.3f} MW")
r2[3].metric("Average Penalty / Block", f"₹{summary['Average Penalty per Block (Rs)']:.2f}")

mpb = summary["Maximum Penalty Block"]
r3 = st.columns(4)
r3[0].metric("Over Generation Blocks", summary["Over Generation Blocks"])
r3[1].metric("Under Generation Blocks", summary["Under Generation Blocks"])
r3[2].metric(
    "Maximum Penalty Block",
    f"Block {mpb['Block']}" if mpb["Block"] is not None else "N/A",
    help=f"{mpb['Time_Label']} — ₹{mpb['Penalty']:.2f}" if mpb["Block"] is not None else None,
)
r3[3].metric("Max Penalty Amount", f"₹{mpb['Penalty']:.2f}")

# ---------------------------------------------------------------------------
# Filters
# ---------------------------------------------------------------------------
st.markdown('<div class="section-header">🔍 Filters</div>', unsafe_allow_html=True)
f1, f2, f3, f4 = st.columns([1.2, 2, 1.2, 1.6])

dates_available = sorted([d for d in result_df["Date"].dropna().unique()])
with f1:
    if dates_available:
        selected_dates = st.multiselect("Date", options=dates_available, default=dates_available)
    else:
        selected_dates = None
        st.caption("No date info in data")

min_block, max_block = int(result_df["Block"].min()), int(result_df["Block"].max())
with f2:
    block_range = st.slider("Block Range", min_value=min_block, max_value=max_block,
                             value=(min_block, max_block))

max_penalty_val = float(result_df["Total_Block_Penalty"].max()) if len(result_df) else 0.0
with f3:
    penalty_threshold = st.number_input("Min Penalty (₹)", min_value=0.0,
                                         max_value=max(max_penalty_val, 0.01),
                                         value=0.0, step=1.0)

with f4:
    dev_filter = st.radio("Deviation Type", ["All", "Only Over Generation", "Only Under Generation"],
                           horizontal=False)

filtered = result_df.copy()
if selected_dates:
    filtered = filtered[filtered["Date"].isin(selected_dates)]
filtered = filtered[(filtered["Block"] >= block_range[0]) & (filtered["Block"] <= block_range[1])]
filtered = filtered[filtered["Total_Block_Penalty"] >= penalty_threshold]
if dev_filter == "Only Over Generation":
    filtered = filtered[filtered["Deviation_Type"] == "Over Generation"]
elif dev_filter == "Only Under Generation":
    filtered = filtered[filtered["Deviation_Type"] == "Under Generation"]

st.caption(f"Showing **{len(filtered)}** of **{len(result_df)}** blocks after filters.")

st.markdown("---")

# ---------------------------------------------------------------------------
# Tabs: Charts / Table / Downloads / Formulas
# ---------------------------------------------------------------------------
tab_charts, tab_table, tab_report, tab_downloads, tab_formulas = st.tabs(
    ["📈 Charts", "📋 Block-wise Table", "📄 Report (Excel Format)",
     "⬇️ Downloads", "🧮 DSM Formula Explainer"]
)

with tab_charts:
    if filtered.empty:
        st.warning("No blocks match the current filters.")
    else:
        figs = all_figures(filtered)
        c1, c2 = st.columns(2)
        with c1:
            st.plotly_chart(figs["actual_vs_predicted"], use_container_width=True)
            st.plotly_chart(figs["penalty_bar"], use_container_width=True)
            st.plotly_chart(figs["over_under_pie"], use_container_width=True)
        with c2:
            st.plotly_chart(figs["deviation_bar"], use_container_width=True)
            st.plotly_chart(figs["cumulative_penalty"], use_container_width=True)
            st.plotly_chart(figs["deviation_histogram"], use_container_width=True)

with tab_table:
    search_col, page_col, size_col = st.columns([2, 1, 1])
    with search_col:
        search_term = st.text_input("🔎 Search table", placeholder="Search any value...")
    with size_col:
        page_size = st.selectbox("Rows per page", [10, 25, 50, 96, "All"], index=1)

    display_cols = [
        "Date", "Block", "Time_Label", "Scheduled_MW", "Actual_MW", "Deviation_MW",
        "Deviation_%", "Abs_Deviation_%", "Block_Energy_kWh", "Total_Block_Penalty",
        "Deviation_Type",
    ]
    table_df = filtered[display_cols].rename(columns={
        "Time_Label": "Time Block", "Scheduled_MW": "Scheduled MW", "Actual_MW": "Actual MW",
        "Deviation_MW": "Deviation MW", "Deviation_%": "Deviation %",
        "Abs_Deviation_%": "Abs Deviation %", "Block_Energy_kWh": "Block Energy (kWh)",
        "Total_Block_Penalty": "Total Penalty (₹)", "Deviation_Type": "Deviation Type",
    })

    if search_term:
        mask = table_df.astype(str).apply(
            lambda col: col.str.contains(search_term, case=False, na=False)
        ).any(axis=1)
        table_df = table_df[mask]

    sort_col = st.selectbox("Sort by", table_df.columns, index=list(table_df.columns).index("Block"))
    sort_desc = st.checkbox("Descending", value=False)
    table_df = table_df.sort_values(sort_col, ascending=not sort_desc).reset_index(drop=True)

    total_rows = len(table_df)
    if page_size != "All" and total_rows > page_size:
        n_pages = max(1, -(-total_rows // page_size))
        with page_col:
            page_num = st.number_input("Page", min_value=1, max_value=n_pages, value=1, step=1)
        start = (page_num - 1) * page_size
        page_df = table_df.iloc[start:start + page_size]
    else:
        page_df = table_df

    max_penalty = table_df["Total Penalty (₹)"].max() if total_rows else 0
    max_dev = table_df["Abs Deviation %"].max() if total_rows else 0

    def _highlight(row):
        styles = [""] * len(row)
        if row["Total Penalty (₹)"] == max_penalty and max_penalty > 0:
            styles = ["background-color:#FCA5A5"] * len(row)
        if row["Abs Deviation %"] == max_dev and max_dev > 0:
            styles = [s + ";background-color:#FDE68A" if s else "background-color:#FDE68A" for s in styles]
        return styles

    styled = page_df.style.apply(_highlight, axis=1).format({
        "Scheduled MW": "{:.3f}", "Actual MW": "{:.3f}", "Deviation MW": "{:.3f}",
        "Deviation %": "{:.2f}%", "Abs Deviation %": "{:.2f}%",
        "Block Energy (kWh)": "{:.1f}", "Total Penalty (₹)": "₹{:.2f}",
    })
    st.caption("🔴 Highest penalty block   🟡 Highest absolute deviation block")
    st.dataframe(styled, use_container_width=True, hide_index=True)

with tab_report:
    st.caption(
        "This mirrors the exact layout of the downloadable Excel report "
        "(\"Schedule vs Meter + Penalty\") — the full day's data, not affected "
        "by the filters above."
    )

    report_cols = ["Block", "Time_Label", "Scheduled_MW", "Actual_MW",
                   "Deviation_MW", "Deviation_%", "Total_Block_Penalty"]
    report_df = result_df[report_cols].rename(columns={
        "Time_Label": "Time", "Scheduled_MW": "AI Schedule (MW)",
        "Actual_MW": "Actual (MW)", "Deviation_MW": "Deviation (MW)",
        "Deviation_%": "Deviation % (Capacity)", "Total_Block_Penalty": "Penalty (Rs)",
    })
    if "Generated_At" in result_df.columns:
        report_df["Scheduled at"] = result_df["Generated_At"].fillna("")
    else:
        report_df["Scheduled at"] = ""

    def _report_style(row):
        styles = [""] * len(row)
        dev = row["Deviation (MW)"]
        dev_color = "#C00000" if abs(dev) > 0.5 else "#1F4E9C"
        for col in ["Deviation (MW)", "Deviation % (Capacity)"]:
            idx = row.index.get_loc(col)
            styles[idx] = f"color:{dev_color}; font-weight:600;"
        if row["Penalty (Rs)"] > 0:
            idx = row.index.get_loc("Penalty (Rs)")
            styles[idx] = "color:#C00000; font-weight:700;"
        return styles

    styled_report = report_df.style.apply(_report_style, axis=1).format({
        "AI Schedule (MW)": "{:.3f}", "Actual (MW)": "{:.3f}", "Deviation (MW)": "{:+.3f}",
        "Deviation % (Capacity)": "{:+.2f}", "Penalty (Rs)": "{:.2f}",
    })
    st.caption("🔴 Deviation > 0.5 MW or a block with a penalty   🔵 Deviation ≤ 0.5 MW")
    st.dataframe(styled_report, use_container_width=True, hide_index=True, height=420)

    st.markdown('<div class="section-header" style="font-size:1.05rem;">📑 Day Summary — Accuracy and DSM Penalty</div>',
                unsafe_allow_html=True)
    metrics = day_summary_metrics(result_df, plant)
    fmt_map = {
        "int": lambda v: f"{v:.0f}",
        "mw": lambda v: f"{v:.3f} MW",
        "mw_signed": lambda v: f"{v:+.3f} MW",
        "pct": lambda v: f"{v:.2f}%",
        "currency": lambda v: f"₹{v:,.2f}",
    }
    m1, m2 = st.columns(2)
    half = (len(metrics) + 1) // 2
    for col, chunk in zip([m1, m2], [metrics[:half], metrics[half:]]):
        with col:
            for m in chunk:
                is_total = "TOTAL DSM PENALTY" in m["label"]
                value_str = fmt_map[m["kind"]](m["value"])
                if is_total:
                    st.markdown(
                        f'<div style="background:#FCE4E4; border-radius:8px; padding:8px 12px; '
                        f'margin-bottom:6px; font-weight:700; color:#C00000; display:flex; '
                        f'justify-content:space-between;"><span>{m["label"]}</span><span>{value_str}</span></div>',
                        unsafe_allow_html=True,
                    )
                else:
                    st.markdown(
                        f'<div style="display:flex; justify-content:space-between; padding:5px 4px; '
                        f'border-bottom:1px solid #F0FDF4; font-size:0.92rem;">'
                        f'<span style="color:#475569;">{m["label"]}</span>'
                        f'<span style="font-weight:600;">{value_str}</span></div>',
                        unsafe_allow_html=True,
                    )

    st.markdown("")
    st.download_button(
        "⬇️ Download this exact report as Excel (.xlsx)",
        data=excel_bytes,
        file_name=f"{plant.name}_Schedule_vs_Meter_Penalty_{_dt.date.today()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary", use_container_width=True, key="download_report_tab",
    )

with tab_downloads:
    st.markdown("Download the processed data and the full formatted DSM penalty report.")
    d1, d2, d3 = st.columns(3)

    with d1:
        st.download_button(
            "⬇️ Merged Dataset (CSV)",
            data=merged.to_csv(index=False).encode("utf-8"),
            file_name=f"{plant.name}_merged_dataset.csv", mime="text/csv",
            use_container_width=True,
        )
    with d2:
        st.download_button(
            "⬇️ Full Block-wise Report (CSV)",
            data=result_df.to_csv(index=False).encode("utf-8"),
            file_name=f"{plant.name}_dsm_blockwise_report.csv", mime="text/csv",
            use_container_width=True,
        )
    with d3:
        st.download_button(
            "⬇️ Full Excel DSM Report (.xlsx)",
            data=excel_bytes,
            file_name=f"{plant.name}_DSM_Penalty_Report_{_dt.date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary", use_container_width=True,
        )

with tab_formulas:
    st.markdown(f"""
#### Deviation

- **Deviation MW** = Actual MW − Scheduled MW
- **Deviation %** = (Deviation MW ÷ Installed Capacity) × 100  *(Installed Capacity = {plant.installed_capacity_mw} MW)*
- **Absolute Deviation %** is used to decide which DSM slab(s) apply. The sign of Deviation MW is preserved separately to label the block as **Over Generation** or **Under Generation**.

#### Block Energy

- **Block Energy (kWh)** = Deviation MW × {plant.block_hours} hour × 1000

#### DSM Slabs (current plant configuration)
""")
    slab_md = "| Band | Penalty Rate |\n|---|---|\n"
    for s in plant.dsm_slabs:
        band = f"{s.lower:.0f}% – {s.upper:.0f}%" if s.upper is not None else f"Above {s.lower:.0f}%"
        slab_md += f"| {band} | ₹{s.rate:.2f} per kWh |\n"
    st.markdown(slab_md)

    st.markdown("""
#### Piecewise Slab Logic (important!)

Penalty is **not** applied to the full deviation at a single rate. Only the portion of the
deviation that falls *inside* each slab is charged at that slab's rate.

**Worked example** — Deviation % = 19.60%, Block Energy = 250 kWh:

| Slab | Width charged | Energy in slab | Rate | Penalty |
|---|---|---|---|---|
| 0% – 10% | 10.00% | 250 × 10/19.60 = 127.55 kWh | ₹0.00 | ₹0.00 |
| 10% – 15% | 5.00% | 250 × 5/19.60 = 63.78 kWh | ₹0.50 | ₹31.89 |
| 15% – 19.60% | 4.60% | 250 × 4.60/19.60 = 58.67 kWh | ₹0.75 | ₹44.00 |
| **Total** | 19.60% | 250.00 kWh | — | **₹75.89** |

General formula for a slab with lower bound *L* and upper bound *U* (deviation % = *D*, block energy = *E*):

```
slab_width   = overlap between [L, U] and [0, D]
Energy(slab) = E × slab_width / D
Penalty(slab) = Energy(slab) × slab_rate

Total Block Penalty = Σ Penalty(slab) over all slabs
```
""")

st.markdown(
    '<div class="footer-bar">⚡ AI Schedule Evaluation Dashboard &nbsp;·&nbsp; '
    'Built for renewable energy scheduling &amp; DSM penalty evaluation &nbsp;·&nbsp; '
    'All calculations run automatically — no manual work required.</div>',
    unsafe_allow_html=True,
)
