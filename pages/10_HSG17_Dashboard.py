#!/usr/bin/env python3

import streamlit as st
import pandas as pd
import plotly.express as px
from pathlib import Path
from datetime import datetime

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
st.markdown('<p class="sub-header">Current State • Placement Groups • Progress to Zero</p>', unsafe_allow_html=True)

DATA_FILE = Path(__file__).parent.parent / "data" / "validation_error_log.xlsx"

@st.cache_data(ttl=30)
def load_data():
    if not DATA_FILE.exists():
        return pd.DataFrame(columns=[
            "timestamp", "hall", "rack_type", "building", "rack",
            "error_category", "count", "source_file", "processed_by"
        ])
    df = pd.read_excel(DATA_FILE)
    df['timestamp'] = pd.to_datetime(df['timestamp'], errors='coerce')
    return df.dropna(subset=['timestamp'])

df = load_data()

hsg17_df = df[df['hall'] == "HSG17"].copy()

if hsg17_df.empty:
    st.warning("No HSG17 data logged yet.")
    st.info("Run the T0-to-Host tool (in the other tab) to log some issues from an LV export + cutsheets. The dashboard will then show current state + deltas per Placement Group.")
    st.stop()

# --- Sidebar Filters (more interactive development for the dashboard) ---
with st.sidebar:
    st.header("🔍 Filters")
    all_buildings = sorted(hsg17_df['building'].unique())
    selected_buildings = st.multiselect(
        "Placement Groups (Buildings)", 
        all_buildings, 
        default=all_buildings,
        help="Filter the view to specific Placement Groups (e.g. PG14 for rack 3110)"
    )
    all_cats = sorted(hsg17_df['error_category'].unique())
    selected_cats = st.multiselect(
        "Error Categories", 
        all_cats, 
        default=all_cats
    )

    min_date = hsg17_df['timestamp'].min().date()
    max_date = hsg17_df['timestamp'].max().date()
    date_range = st.date_input(
        "Consider logs from / to",
        value=(min_date, max_date),
        min_value=min_date,
        max_value=max_date,
        help="Only include log entries within this date range when computing current state and deltas."
    )

# Apply filters to the raw log data before computing snapshots/deltas
filtered_df = hsg17_df[
    hsg17_df['building'].isin(selected_buildings) &
    hsg17_df['error_category'].isin(selected_cats) &
    (hsg17_df['timestamp'].dt.date >= date_range[0]) &
    (hsg17_df['timestamp'].dt.date <= date_range[1])
].copy()

if filtered_df.empty:
    st.warning("No data matches the current filters. Adjust the Placement Groups, Categories or Date Range in the sidebar.")
    st.stop()

def get_latest_snapshot(dataframe: pd.DataFrame) -> pd.DataFrame:
    """Only the most recent entry per (building/placement group, error_category)."""
    if dataframe.empty:
        return dataframe
    return (
        dataframe.sort_values('timestamp')
        .groupby(['hall', 'building', 'error_category'], as_index=False)
        .last()
    )

def get_latest_with_deltas(dataframe: pd.DataFrame) -> pd.DataFrame:
    """Current vs previous run + delta for each placement group+category."""
    if dataframe.empty:
        return pd.DataFrame(columns=['hall', 'building', 'error_category', 'rack', 'current', 'previous', 'delta'])

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

        rack = str(group.iloc[-1].get('rack', '')) if 'rack' in group.columns else ''

        records.append({
            'hall': hall,
            'building': bldg,
            'error_category': cat,
            'rack': rack,
            'current': current,
            'previous': previous,
            'delta': delta
        })
    return pd.DataFrame(records)

current = get_latest_snapshot(filtered_df)
current_with_deltas = get_latest_with_deltas(filtered_df)

if DATA_FILE.exists():
    col1, col2 = st.columns(2)
    with col1:
        with open(DATA_FILE, "rb") as f:
            st.download_button(
                "📥 Download live validation_error_log.xlsx",
                data=f,
                file_name="HSG17_validation_error_log.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                width="stretch"
            )
    with col2:
        # Reset button for dashboard data (user requested for clearing test data)
        if st.button("🗑️ Reset Dashboard Data (clear HSG17 log)", type="secondary", width="stretch", help="Removes all HSG17 entries from the log so you can start fresh with real data. This only affects the dashboard feed."):
            try:
                if DATA_FILE.exists():
                    df = pd.read_excel(DATA_FILE)
                    if 'hall' in df.columns:
                        df_clean = df[df['hall'] != "HSG17"].copy()
                    else:
                        df_clean = df.copy()
                    if df_clean.empty:
                        # Recreate with proper headers
                        df_clean = pd.DataFrame(columns=[
                            "timestamp", "hall", "rack_type", "building", 
                            "error_category", "count", "source_file", "processed_by"
                        ])
                    df_clean.to_excel(DATA_FILE, index=False)
                    st.success("HSG17 dashboard data has been cleared. The page will refresh with empty state.")
                    st.rerun()
                else:
                    st.info("No data file found to clear.")
            except Exception as e:
                st.error(f"Failed to clear data: {e}")

st.divider()

st.markdown('<div class="section-header">Executive Snapshot (respects sidebar filters)</div>', unsafe_allow_html=True)

col1, col2, col3, col4 = st.columns(4)

total_errors = int(current['count'].sum())
unique_blocks = current['building'].nunique()
active_rack_types = current['rack_type'].nunique()

with col1:
    st.metric("Total Open Issues (HSG17)", f"{total_errors:,}")
with col2:
    st.metric("Placement Groups with Issues", unique_blocks)
with col3:
    st.metric("Rack Types Active", active_rack_types)
with col4:
    st.metric("Last Updated", current['timestamp'].max().strftime("%Y-%m-%d %H:%M") if not current.empty else "—")

st.divider()

st.markdown('<div class="section-header">Error Breakdown by Placement Group (respects sidebar filters)</div>', unsafe_allow_html=True)

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
                    st.markdown(f"<div style='font-size:1.05rem; font-weight:600; margin-bottom:2px'>{bldg}</div>", unsafe_allow_html=True)

                    total_str = str(bldg_total)
                    if pd.notna(total_delta):
                        delta_int = int(total_delta)
                        delta_sign = f"({delta_int:+d})" if delta_int != 0 else ""
                        delta_color = "green" if delta_int < 0 else "red"
                        total_str += f" <span style='font-size:0.9rem; color:{delta_color};'>{delta_sign}</span>"
                    st.markdown(f"<div style='font-size:1.9rem; font-weight:700; line-height:1.1; margin-bottom:6px'>{total_str}</div>", unsafe_allow_html=True)

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
    st.info("No matching Placement Group data after filters. Try broadening your sidebar selections.")

st.divider()

st.markdown('<div class="section-header">Errors by Category × Placement Group (respects sidebar filters)</div>', unsafe_allow_html=True)

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

    cat_totals = pivot.drop("TOTAL")["Total"].reset_index()
    cat_totals.columns = ["Category", "Errors"]
    fig = px.bar(cat_totals, x="Category", y="Errors", height=280)
    fig.update_layout(margin=dict(l=0, r=0, t=20, b=0))
    st.plotly_chart(fig, width="stretch", key="hsg17_cat_totals", config={"displayModeBar": False})

st.divider()

st.markdown('<div class="section-header">Current Issues Detail</div>', unsafe_allow_html=True)

if not current.empty:
    detail = current_with_deltas[['building', 'rack', 'error_category', 'current', 'previous', 'delta']].copy()
    detail = detail.rename(columns={
        'building': 'Placement Group',
        'rack': 'Rack',
        'error_category': 'Category',
        'current': 'Current',
        'previous': 'Previous',
        'delta': 'Delta'
    })
    detail = detail.sort_values(['Placement Group', 'Category'])
    st.dataframe(detail, width="stretch", hide_index=True, use_container_width=True)

    # CSV export for the current filtered view
    csv = detail.to_csv(index=False).encode('utf-8')
    st.download_button(
        "📥 Download current filtered view (CSV)",
        data=csv,
        file_name="hsg17_current_issues.csv",
        mime="text/csv",
        width="stretch"
    )
else:
    st.info("No current issues in filtered view.")

st.divider()

# Progress Trend moved to bottom as requested
st.markdown('<div class="section-header">Progress Trend (Total Open Issues Over Time)</div>', unsafe_allow_html=True)

if not filtered_df.empty:
    # Aggregate total issues per logged run time (using the raw filtered log entries)
    trend = (
        filtered_df.groupby(filtered_df['timestamp'].dt.strftime('%Y-%m-%d %H:%M'))['count']
        .sum()
        .reset_index()
    )
    trend.columns = ['Run Time', 'Total Issues']
    trend = trend.sort_values('Run Time')

    fig = px.line(
        trend, 
        x='Run Time', 
        y='Total Issues', 
        markers=True,
        title='Total Open Issues per Processing Run (filtered view)'
    )
    fig.update_layout(
        height=320, 
        margin=dict(l=0, r=0, t=30, b=0),
        xaxis_title='Run Timestamp',
        yaxis_title='Sum of Counts (all categories)'
    )
    fig.update_traces(line=dict(width=3))
    st.plotly_chart(fig, width="stretch", key="hsg17_trend", config={"displayModeBar": False})

    st.caption("Each point represents the total issues logged in one run of the T0-to-Host tool (within your current filters). Re-runs update the 'current' cards above but the full history stays in the log for trending.")
else:
    st.info("No data for trend chart with current filters.")
