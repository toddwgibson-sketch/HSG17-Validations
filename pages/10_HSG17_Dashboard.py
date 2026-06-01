#!/usr/bin/env python3
"""
HSG17 Validation Error Dashboard
Clean, executive view for the HSG17 site.

- Current-state only (latest per Block + category)
- Beautiful widget-style cards with deltas
- Reuses the exact central log format as the T0-to-Host tool
- Same delta + "get to zero" philosophy as the JPB15/SYD20 dashboard
"""

import streamlit as st
import pandas as pd
import plotly.express as px
from pathlib import Path
from datetime import datetime

from utils.auth import require_login

require_login()

st.set_page_config(
    page_title="HSG17 Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ====================== BEAUTIFUL C-SUITE STYLING ======================
st.markdown("""
<style>
    .stApp .main-header,
    .main-header {
        font-size: 3.5rem !important;
        font-weight: 700;
        margin-bottom: 0.1rem;
        line-height: 1.1;
    }
    .sub-header {
        font-size: 1.05rem;
        margin-bottom: 1.8rem;
    }
    .section-header {
        font-size: 1.35rem;
        font-weight: 600;
        margin-top: 1.8rem;
        margin-bottom: 0.6rem;
    }
</style>
""", unsafe_allow_html=True)

st.markdown('<p class="main-header">HSG17 Dashboard</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-header">Current State • DH Blocks • Progress to Zero</p>', unsafe_allow_html=True)

# ====================== DATA ======================
DATA_FILE = Path(__file__).parent.parent / "data" / "validation_error_log.xlsx"

@st.cache_data(ttl=30)
def load_data():
    if not DATA_FILE.exists():
        return pd.DataFrame(columns=[
            "timestamp", "hall", "rack_type", "building", 
            "error_category", "count", "source_file", "processed_by"
        ])
    df = pd.read_excel(DATA_FILE)
    df['timestamp'] = pd.to_datetime(df['timestamp'], errors='coerce')
    return df.dropna(subset=['timestamp'])

df = load_data()

# Filter strictly to HSG17
hsg17_df = df[df['hall'] == "HSG17"].copy()

if hsg17_df.empty:
    abs_path = DATA_FILE.resolve()
    st.warning("No HSG17 data logged yet.")
    st.info("Process files using the **HSG17 T0-to-Host** tool in this app to populate the dashboard.")
    st.markdown("### Central log location (this repo)")
    st.code(str(abs_path))
    st.stop()

# ====================== HELPER FUNCTIONS (current state + deltas) ======================
def get_latest_snapshot(dataframe: pd.DataFrame) -> pd.DataFrame:
    """Only the most recent entry per (building/block, error_category)."""
    if dataframe.empty:
        return dataframe
    return (
        dataframe.sort_values('timestamp')
        .groupby(['hall', 'building', 'error_category'], as_index=False)
        .last()
    )

def get_latest_with_deltas(dataframe: pd.DataFrame) -> pd.DataFrame:
    """Current vs previous run + delta for each block+category."""
    if dataframe.empty:
        return pd.DataFrame(columns=['hall', 'building', 'error_category', 'current', 'previous', 'delta'])

    records = []
    for (hall, bldg, cat), group in dataframe.sort_values('timestamp').groupby(['hall', 'building', 'error_category']):
        group = group.sort_values('timestamp')
        current = int(group.iloc[-1]['count'])

        if len(group) >= 2:
            previous = int(group.iloc[-2]['count'])
            delta = current - previous
        else:
            previous = None
            delta = None

        if pd.isna(delta):
            delta = None

        records.append({
            'hall': hall,
            'building': bldg,
            'error_category': cat,
            'current': current,
            'previous': previous,
            'delta': delta
        })
    return pd.DataFrame(records)

current = get_latest_snapshot(hsg17_df)
current_with_deltas = get_latest_with_deltas(hsg17_df)

# ====================== STORAGE INFO (user always asks for this) ======================
with st.expander("📍 Where is the HSG17 dashboard data stored?", expanded=True):
    abs_path = DATA_FILE.resolve()
    st.code(str(abs_path))

    if DATA_FILE.exists():
        st.success("✅ Log file exists")
        st.caption(f"Size: {DATA_FILE.stat().st_size:,} bytes")

        with open(DATA_FILE, "rb") as f:
            st.download_button(
                "📥 Download live validation_error_log.xlsx",
                data=f,
                file_name="HSG17_validation_error_log.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                width="stretch"
            )
    else:
        st.warning("Log file does not exist yet.")

    st.markdown(f"""
    **Local Windows path:**
    ```
    C:\\Users\\toddy\\Documents\\GitHub\\HSG17-Validations\\data\\validation_error_log.xlsx
    ```
    """)

st.divider()

# ====================== EXECUTIVE KPIs ======================
st.markdown('<div class="section-header">Executive Snapshot</div>', unsafe_allow_html=True)

col1, col2, col3, col4 = st.columns(4)

total_errors = int(current['count'].sum())
unique_blocks = current['building'].nunique()
active_rack_types = current['rack_type'].nunique()

with col1:
    st.metric("Total Open Issues (HSG17)", f"{total_errors:,}")
with col2:
    st.metric("Blocks with Issues", unique_blocks)
with col3:
    st.metric("Rack Types Active", active_rack_types)
with col4:
    st.metric("Last Updated", current['timestamp'].max().strftime("%Y-%m-%d %H:%M") if not current.empty else "—")

st.divider()

# ====================== ERROR BREAKDOWN BY BLOCK (the widget cards the user loves) ======================
st.markdown('<div class="section-header">Error Breakdown by Block</div>', unsafe_allow_html=True)

# HSG17 category colors (professional)
CAT_COLORS = {
    "LLDP Mismatch + Link Down": "#e74c3c",
    "Optic Errors": "#3498db",
    "FEC_BER Errors": "#9b59b6",
    "Interface Down Errors": "#e67e22",
}
CAT_LABELS = {
    "LLDP Mismatch + Link Down": "LLDP Mismatch",
    "Optic Errors": "Optics",
    "FEC_BER Errors": "FEC BER",
    "Interface Down Errors": "Interface Down",
}

if not current.empty:
    building_order = sorted(current['building'].unique())
    CARDS_PER_ROW = 5

    category_order = [c for c in CAT_LABELS.keys() if c in current['error_category'].unique()]

    for start_idx in range(0, len(building_order), CARDS_PER_ROW):
        row_buildings = building_order[start_idx : start_idx + CARDS_PER_ROW]
        cols = st.columns(CARDS_PER_ROW)

        for i, bldg in enumerate(row_buildings):
            with cols[i]:
                bldg_deltas = current_with_deltas[current_with_deltas['building'] == bldg]

                cat_current = {}
                cat_delta = {}
                for _, row in bldg_deltas.iterrows():
                    cat = row['error_category']
                    cat_current[cat] = row['current']
                    cat_delta[cat] = row['delta']

                bldg_total = sum(cat_current.values())

                valid_deltas = [d for d in cat_delta.values() if pd.notna(d)]
                total_delta = sum(valid_deltas) if valid_deltas else None

                with st.container(border=True):
                    # Block name
                    st.markdown(f"<div style='font-size:1.05rem; font-weight:600; margin-bottom:2px'>{bldg}</div>", unsafe_allow_html=True)

                    # Big total + delta
                    total_str = str(bldg_total)
                    if pd.notna(total_delta):
                        delta_int = int(total_delta)
                        delta_sign = f"({delta_int:+d})" if delta_int != 0 else ""
                        delta_color = "green" if delta_int < 0 else "red"
                        total_str += f" <span style='font-size:0.9rem; color:{delta_color};'>{delta_sign}</span>"
                    st.markdown(f"<div style='font-size:1.9rem; font-weight:700; line-height:1.1; margin-bottom:6px'>{total_str}</div>", unsafe_allow_html=True)

                    # Mini stacked bar
                    bar_data = []
                    for cat in category_order:
                        val = cat_current.get(cat, 0)
                        if val > 0:
                            bar_data.append({
                                "Category": CAT_LABELS.get(cat, cat),
                                "Count": val,
                                "Color": CAT_COLORS.get(cat, "#7f8c8d")
                            })

                    if bar_data:
                        bar_df = pd.DataFrame(bar_data)
                        fig = px.bar(
                            bar_df,
                            x="Count",
                            y=[""] * len(bar_df),
                            color="Category",
                            orientation="h",
                            color_discrete_map={d["Category"]: d["Color"] for d in bar_data},
                            height=42
                        )
                        fig.update_layout(
                            barmode="stack",
                            margin=dict(l=0, r=0, t=0, b=0),
                            xaxis_visible=False,
                            yaxis_visible=False,
                            showlegend=False,
                            height=42
                        )
                        fig.update_traces(marker_line_width=0)
                        st.plotly_chart(fig, width="stretch", key=f"hsg17_bar_{bldg}", config={"displayModeBar": False})

                    # Compact list with deltas
                    st.markdown("<div style='margin-top:4px; font-size:0.82rem; line-height:1.25'>", unsafe_allow_html=True)
                    for cat in category_order:
                        label = CAT_LABELS.get(cat, cat)
                        val = cat_current.get(cat, 0)
                        d = cat_delta.get(cat)
                        color = CAT_COLORS.get(cat, "#7f8c8d")

                        delta_html = ""
                        if pd.notna(d):
                            d_int = int(d)
                            delta_color = "green" if d_int < 0 else "red"
                            delta_str = f"({d_int:+d})"
                            delta_html = f" <span style='color:{delta_color}; font-size:0.75rem;'>{delta_str}</span>"

                        st.markdown(
                            f"<span style='color:{color}; font-weight:600'>■</span> {label}: <b>{val}</b>{delta_html}",
                            unsafe_allow_html=True
                        )
                    st.markdown("</div>", unsafe_allow_html=True)
else:
    st.info("No HSG17 block data yet. Run the T0-to-Host tool to start logging issues.")

st.divider()

# ====================== PIVOT TABLE ======================
st.markdown('<div class="section-header">Errors by Category × Block</div>', unsafe_allow_html=True)

if not current.empty:
    pivot = (
        current.pivot_table(
            index="error_category",
            columns="building",
            values="count",
            aggfunc="sum",
            fill_value=0
        )
        .astype(int)
    )
    pivot["Total"] = pivot.sum(axis=1)
    pivot = pivot.sort_values("Total", ascending=False)
    pivot.loc["TOTAL"] = pivot.sum()

    st.dataframe(
        pivot,
        width="stretch",
        column_config={
            col: st.column_config.NumberColumn(col, format="%d") 
            for col in pivot.columns
        }
    )

    # Simple bar of totals
    cat_totals = pivot.drop("TOTAL")["Total"].reset_index()
    cat_totals.columns = ["Category", "Errors"]
    fig = px.bar(cat_totals, x="Category", y="Errors", height=280)
    fig.update_layout(margin=dict(l=0, r=0, t=20, b=0))
    st.plotly_chart(fig, width="stretch", key="hsg17_cat_totals", config={"displayModeBar": False})

st.caption("Data source: HSG17 validation_error_log.xlsx (inside this repo). Re-uploading the same blocks overwrites previous counts — dashboard always shows current state.")