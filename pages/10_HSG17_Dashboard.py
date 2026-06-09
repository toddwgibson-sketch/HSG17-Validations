#!/usr/bin/env python3

import sys
from pathlib import Path

# Ensure project root is on sys.path for "from utils.xxx" when this page
# is exec'd by st.navigation (fixes ImportError on Streamlit Cloud).
_here = Path(__file__).resolve()
_root = _here.parent.parent if _here.parent.name == "pages" else _here.parent
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from pathlib import Path
from datetime import datetime

from utils.hsg17_models import is_gpu_rack

st.set_page_config(
    page_title="HSG17 Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("📊 HSG17 Dashboard")
st.caption("Current State • Placement Groups • Progress to Zero")

# ====================== DASHBOARD STYLING (cards + panels) ======================
st.markdown("""
<style>
    .section-header {
        font-size: 1.25rem;
        font-weight: 600;
        margin-top: 1.4rem;
        margin-bottom: 0.5rem;
        color: #e2e8f0;
        border-left: 4px solid #22d3ee;
        padding-left: 10px;
    }
    /* Fancy metric cards */
    .hsg17-metric-card {
        border-radius: 14px;
        padding: 16px 18px;
        color: white;
        box-shadow: 0 10px 15px -3px rgb(0 0 0 / 0.3), 0 4px 6px -4px rgb(0 0 0 / 0.3);
        min-height: 118px;
        position: relative;
        overflow: hidden;
    }
    .hsg17-metric-card .label {
        font-size: 0.72rem;
        opacity: 0.85;
        font-weight: 500;
        letter-spacing: 0.6px;
        text-transform: uppercase;
    }
    .hsg17-metric-card .value {
        font-size: 1.95rem;
        font-weight: 700;
        line-height: 1.05;
        margin-top: 2px;
    }
    .hsg17-metric-card .icon {
        font-size: 1.9rem;
        opacity: 0.85;
    }
    .hsg17-metric-card .progress {
        margin-top: 10px;
        height: 3px;
        background: rgba(255,255,255,0.22);
        border-radius: 999px;
        position: relative;
    }
    .hsg17-metric-card .progress-fill {
        height: 3px;
        background: rgba(255,255,255,0.95);
        border-radius: 999px;
    }
    .hsg17-metric-card .progress-dot {
        position: absolute;
        top: -1.5px;
        width: 6px;
        height: 6px;
        background: white;
        border-radius: 999px;
        box-shadow: 0 0 0 2px rgba(255,255,255,0.4);
    }
    .progress {
        margin: 4px 0 8px;
        height: 3px;
        background: rgba(255,255,255,0.25);
        border-radius: 999px;
        position: relative;
    }
    .progress-fill {
        height: 3px;
        background: rgba(255,255,255,0.95);
        border-radius: 999px;
    }
    .progress-dot {
        position: absolute;
        top: -1.5px;
        width: 6px;
        height: 6px;
        background: white;
        border-radius: 999px;
        box-shadow: 0 0 0 2px rgba(255,255,255,0.4);
    }
    .hsg17-metric-card .sub {
        font-size: 0.62rem;
        opacity: 0.65;
        margin-top: 3px;
    }
    /* PG breakdown cards - full gradient like exec snapshot cards, with outline */
    .hsg17-pg-card {
        border-radius: 10px;
        padding: 10px 12px;
        margin-bottom: 4px;
        color: white;
        box-shadow: 0 10px 15px -3px rgb(0 0 0 / 0.3), 0 4px 6px -4px rgb(0 0 0 / 0.3);
    }
    .hsg17-pg-card .pg-pill {
        background-color:#0f172a; 
        border:2px solid #22d3ee; 
        border-radius:9999px; 
        padding:2px 8px; 
        font-size:0.7rem; 
        font-weight:500; 
        color:#e0f2fe; 
        letter-spacing:0.5px;
    }
    /* Rack table panels */
    .rack-panel {
        background: #1e2937;
        border: 1px solid #334155;
        border-radius: 10px;
        padding: 8px 12px;
        margin-bottom: 12px;
    }
    /* Bottom panels style */
    .dashboard-panel {
        background: #1e2937;
        border: 1px solid #334155;
        border-radius: 12px;
        padding: 12px;
        margin-bottom: 8px;
    }

    /* Tone down bright red / harsh colors on sidebar filters (multiselect tags, date inputs) */
    .stSidebar .stMultiSelect [data-baseweb="tag"] {
        background-color: #1e3a8a !important;  /* muted blue to match dark theme */
        color: #e0f2fe !important;
        border: 1px solid #22d3ee !important;
        border-radius: 999px !important;
    }
    .stSidebar .stMultiSelect [data-baseweb="tag"]:hover {
        background-color: #1e40af !important;
    }
    .stSidebar .stMultiSelect [data-baseweb="tag"] > span {
        color: #e0f2fe !important;
    }
    /* Date input focus to use cyan instead of red */
    .stSidebar .stDateInput > div > div > div > input:focus {
        border-color: #22d3ee !important;
        box-shadow: 0 0 0 1px #22d3ee !important;
    }
    /* Mute sidebar filter labels and header for consistency */
    .stSidebar .stMultiSelect label, .stSidebar .stDateInput label {
        color: #94a3b8 !important;
    }

    /* Cool gradient buttons for download error log and reset dashboard */
    .stDownloadButton button, .stButton button {
        background: linear-gradient(135deg, #1e3a8a, #0369a1) !important;
        color: #e0f2fe !important;
        border: 1px solid #22d3ee !important;
        border-radius: 8px !important;
        padding: 0.6rem 1rem !important;
        font-weight: 500 !important;
        transition: all 0.2s ease !important;
    }
    .stDownloadButton button:hover, .stButton button:hover {
        background: linear-gradient(135deg, #0369a1, #1e3a8a) !important;
        box-shadow: 0 0 0 2px #22d3ee !important;
        transform: translateY(-1px) !important;
    }
    .stButton button[kind="secondary"] {
        background: linear-gradient(135deg, #334155, #1e2937) !important;
        border: 1px solid #64748b !important;
    }
</style>
""", unsafe_allow_html=True)

DATA_FILE = Path(__file__).parent.parent / "data" / "validation_error_log.xlsx"

@st.cache_data(ttl=30)
def load_data():
    if not DATA_FILE.exists():
        return pd.DataFrame(columns=[
            "timestamp", "hall", "rack_type", "building", "rack",
            "error_category", "count", "source_file", "processed_by"
        ])
    df = pd.read_excel(DATA_FILE)
    # Upgrade schema if 'rack' column missing (for old logs without rack)
    if 'rack' not in df.columns:
        cols = list(df.columns)
        idx = cols.index('building') + 1 if 'building' in cols else len(cols)
        df.insert(idx, 'rack', '')
        df['rack'] = df['rack'].astype('object').fillna('').astype(str)
        try:
            df.to_excel(DATA_FILE, index=False)
            print(f"[DASHBOARD] Upgraded log schema with 'rack' column")
        except Exception as e:
            print(f"[DASHBOARD] Could not save upgraded log: {e}")
    df['timestamp'] = pd.to_datetime(df['timestamp'], errors='coerce')
    if not df.empty:
        # Treat stored timestamps as UTC (we log in UTC now) and convert to local tz of this computer for display.
        # This makes "Last Updated" etc. match your local time. Historical data may shift if previously logged in another tz.
        local_tz = datetime.now().astimezone().tzinfo
        df['timestamp'] = df['timestamp'].dt.tz_localize('UTC').dt.tz_convert(local_tz)
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

        rack_val = group.iloc[-1].get('rack', '') if 'rack' in group.columns else ''
        rack = str(rack_val) if pd.notna(rack_val) and rack_val != '' else ''

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
    st.markdown('<div class="dashboard-panel">', unsafe_allow_html=True)
    st.markdown("<div style='font-size:0.85rem; font-weight:600; color:#94a3b8; margin-bottom:4px;'>Data Management (unified — 01 LV Portal + 02 Slack tool)</div>", unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    with col1:
        with open(DATA_FILE, "rb") as f:
            st.download_button(
                "📥 Download Error Log",
                data=f,
                file_name="HSG17_validation_error_log.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                width="stretch",
                help="Download the full validation error log (all halls)"
            )
    with col2:
        # Reset button for dashboard data (user requested for clearing test data)
        if st.button("🗑️ Reset HSG17 Data", type="secondary", width="stretch", help="Removes all HSG17 entries from the log so you can start fresh with real data. This only affects the dashboard feed."):
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
    st.markdown('</div>', unsafe_allow_html=True)

st.divider()

st.markdown('<div class="section-header">Executive Snapshot (respects sidebar filters)</div>', unsafe_allow_html=True)

col1, col2, col3, col4 = st.columns(4)

total_errors = int(current['count'].sum())
unique_blocks = current['building'].nunique()
active_rack_types = current['rack_type'].nunique()
last_ts = "—"
if not current.empty:
    ts = current['timestamp'].max()
    if pd.notna(ts):
        ts = pd.Timestamp(ts)
        if ts.tz is None:
            ts = ts.tz_localize('UTC')
        local_ts = ts.tz_convert(datetime.now().astimezone().tzinfo)
        last_ts = local_ts.strftime("%Y-%m-%d %H:%M")

def _metric_card(label, value, icon, c1, c2, sub=""):
    return f'''
    <div class="hsg17-metric-card" style="background: linear-gradient(135deg, {c1}, {c2});">
        <div style="display:flex; justify-content:space-between; align-items:flex-start;">
            <div>
                <div class="label">{label}</div>
                <div class="value">{value}</div>
            </div>
            <div class="icon">{icon}</div>
        </div>
        <div class="progress">
            <div class="progress-fill" style="width: {min(92, 35 + (hash(str(value)) % 40))}%;"></div>
            <div class="progress-dot" style="left: {min(92, 35 + (hash(str(value)) % 40))}%;"></div>
        </div>
        <div class="sub">{sub}</div>
    </div>
    '''

with col1:
    st.markdown(_metric_card("TOTAL OPEN ISSUES (HSG17)", f"{total_errors:,}", "⚠️", "#0ea5e9", "#0369a1", "Current snapshot • filtered"), unsafe_allow_html=True)
with col2:
    st.markdown(_metric_card("PLACEMENT GROUPS WITH ISSUES", str(unique_blocks), "📍", "#f43f5e", "#9f1239", "Active PGs in view"), unsafe_allow_html=True)
with col3:
    st.markdown(_metric_card("RACK TYPES ACTIVE", str(active_rack_types), "🖥️", "#f59e0b", "#c2410f", "Unique rack types"), unsafe_allow_html=True)
with col4:
    st.markdown(_metric_card("LAST UPDATED", last_ts, "🕒", "#a855f7", "#6b21a8", "Latest processing run"), unsafe_allow_html=True)

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

# Gradient palette for PG cards - cycle like the exec snapshot cards
GRADIENTS = [
    ("#0ea5e9", "#0369a1"),  # blue
    ("#f43f5e", "#9f1239"),  # rose
    ("#f59e0b", "#c2410f"),  # orange
    ("#a855f7", "#6b21a8"),  # purple
    ("#10b981", "#047857"),  # green
    ("#06b6d4", "#0e7490"),  # cyan
    ("#8b5cf6", "#5b21b6"),  # violet
]

if not current.empty:
    building_order = sorted(current['building'].unique())
    CARDS_PER_ROW = 5

    category_order = [c for c in CAT_LABELS.keys() if c in current['error_category'].unique()]

    for start_idx in range(0, len(building_order), CARDS_PER_ROW):
        row_buildings = building_order[start_idx : start_idx + CARDS_PER_ROW]
        cols = st.columns(CARDS_PER_ROW)

        for i, bldg in enumerate(row_buildings):
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

            total_str = str(bldg_total)
            if pd.notna(total_delta):
                delta_int = int(total_delta)
                delta_sign = f"({delta_int:+d})" if delta_int != 0 else ""
                delta_color = "green" if delta_int < 0 else "red"
                total_str += f" <span style='font-size:0.9rem; color:{delta_color};'>{delta_sign}</span>"

            bar_data = []
            for cat in category_order:
                val = cat_current.get(cat, 0)
                if val > 0:
                    bar_data.append({
                        "Category": CAT_LABELS.get(cat, cat),
                        "Count": val,
                        "Color": CAT_COLORS.get(cat, "#7f8c8d")
                    })

            grad_idx = (start_idx + i) % len(GRADIENTS)
            g1, g2 = GRADIENTS[grad_idx]

            # build the list html (white text for gradient bg)
            list_html = "<div style='margin-top:4px; font-size:0.82rem; line-height:1.25; color:#f8fafc;'>"
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

                list_html += f"<span style='color:{color}; font-weight:600'>■</span> {label}: <b>{val}</b>{delta_html}<br>"
            list_html += "</div>"

            with cols[i]:
                st.markdown(f'<div class="hsg17-pg-card" style="background: linear-gradient(135deg, {g1}, {g2}); color: white; border: none; box-shadow: 0 10px 15px -3px rgb(0 0 0 / 0.3), 0 4px 6px -4px rgb(0 0 0 / 0.3);">', unsafe_allow_html=True)

                st.markdown(f'<span class="pg-pill">{bldg}</span>', unsafe_allow_html=True)

                st.markdown(f"<div style='font-size:1.9rem; font-weight:700; line-height:1.1; margin-bottom:6px; color:#f8fafc;'>{total_str}</div>", unsafe_allow_html=True)

                if bar_data:
                    total_for_bar = sum(d['Count'] for d in bar_data)
                    bar_html = '<div style="height:16px; background: rgba(255,255,255,0.15); border-radius:999px; overflow:hidden; display:flex; margin:4px 0 6px; border:1px solid rgba(255,255,255,0.3);">'
                    for d in bar_data:
                        pct = (d['Count'] / total_for_bar * 100) if total_for_bar > 0 else 0
                        bar_html += f'<div style="width:{pct}%; background:{d["Color"]}; height:100%;"></div>'
                    bar_html += '</div>'
                    st.markdown(bar_html, unsafe_allow_html=True)

                st.markdown('<div style="background: rgba(0,0,0,0.25); border-radius: 6px; padding: 4px; margin-top: 4px;">', unsafe_allow_html=True)
                st.markdown(list_html, unsafe_allow_html=True)
                st.markdown('</div>', unsafe_allow_html=True)

                st.markdown('</div>', unsafe_allow_html=True)
else:
    st.info("No matching Placement Group data after filters. Try broadening your sidebar selections.")

st.divider()

st.markdown('<div class="section-header">Errors by Category × Rack (per Placement Group, respects sidebar filters)</div>', unsafe_allow_html=True)

if not current.empty:
    # New design: one table per Placement Group
    # rows: error_category, columns: rack (within the PG)
    # Only GPU racks (from your HSG17 - Placement Groups.txt)
    for bldg in sorted(current['building'].dropna().unique()):
        sub = current[current['building'] == bldg]
        if 'rack' in sub.columns:
            sub = sub[sub['rack'].apply(is_gpu_rack)]
        if sub.empty:
            continue
        sub_pivot = (
            sub.pivot_table(
                index="error_category",
                columns="rack",
                values="count",
                aggfunc="sum",
                fill_value=0
            )
            .astype(int)
        )
        sub_pivot["Total"] = sub_pivot.sum(axis=1)
        sub_pivot = sub_pivot.sort_values("Total", ascending=False)
        sub_pivot.loc["TOTAL"] = sub_pivot.sum()

        st.markdown('<div class="rack-panel">', unsafe_allow_html=True)
        st.markdown(f"<div style='font-weight:600; color:#e0f2fe; margin: 2px 0 6px;'>{bldg}</div>", unsafe_allow_html=True)
        st.dataframe(
            sub_pivot,
            width="stretch",
            column_config={
                col: st.column_config.NumberColumn(col, format="%d") 
                for col in sub_pivot.columns
            }
        )
        st.markdown('</div>', unsafe_allow_html=True)

st.divider()

# Progress Trend moved to bottom as requested (full width, pie removed)
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

    # Full-width trend panel
    st.markdown('<div class="dashboard-panel">', unsafe_allow_html=True)
    st.markdown("<div style='font-size:0.85rem; font-weight:600; color:#94a3b8; margin-bottom:4px;'>Issues Over Time (filtered runs)</div>", unsafe_allow_html=True)
    # Trend chart - combo bar + line
    fig = go.Figure()
    # Bars for each run
    fig.add_trace(go.Bar(
        x=trend['Run Time'],
        y=trend['Total Issues'],
        name='Issues per Run',
        marker_color='#67e8f9',
        opacity=0.7
    ))
    # Line + area on top (smooth trend)
    fig.add_trace(go.Scatter(
        x=trend['Run Time'],
        y=trend['Total Issues'],
        mode='lines+markers',
        line=dict(color='#22d3ee', width=3, shape='spline'),
        fill='tozeroy',
        fillcolor='rgba(34, 211, 238, 0.15)',
        marker=dict(size=6, color='#67e8f9', line=dict(width=1, color='#0b1120')),
        name='Trend'
    ))
    fig.update_layout(
        height=340,
        margin=dict(l=0, r=0, t=8, b=0),
        xaxis_title='Run Timestamp',
        yaxis_title='Open Issues',
        plot_bgcolor='#0b1120',
        paper_bgcolor='#0b1120',
        font_color='#e2e8f0',
        xaxis=dict(gridcolor='#1e2937', zerolinecolor='#334155'),
        yaxis=dict(gridcolor='#1e2937', zerolinecolor='#334155'),
        hovermode='x unified',
        barmode='overlay',
        showlegend=False
    )
    fig.update_traces(hovertemplate='%{x}<br>%{y} issues<extra></extra>')

    # Export / Print pills for this panel
    tcols = st.columns([0.75, 0.25])
    with tcols[1]:
        bb1, bb2 = st.columns(2)
        with bb1:
            csv = trend.to_csv(index=False).encode('utf-8')
            st.download_button("⤵", data=csv, file_name="hsg17_trend.csv", mime="text/csv", key="trend_csv_dl", use_container_width=True, help="Export trend data")
        with bb2:
            if st.button("🖨", key="trend_print", use_container_width=True, help="Print chart"):
                st.toast("Use browser Print (Ctrl+P)")
    st.plotly_chart(fig, width="stretch", key="hsg17_trend", config={"displayModeBar": False})
    st.markdown('</div>', unsafe_allow_html=True)

    st.caption("Total open issues over time (bars = per run, line/area = trend). Re-runs update the cards above; history is preserved for trends.")
else:
    st.info("No data for trend chart with current filters.")
