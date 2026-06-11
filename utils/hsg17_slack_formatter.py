#!/usr/bin/env python3
r'''
HSG17 T1-to-T0 Slack Formatter — Core Logic (v2)
Sourced exactly from the updated reference (slack_formatter_T1toT0.v2).

This module contains ONLY the pure formatting logic (no tkinter, no Streamlit, no HSG17-specific logging).
It is intended to be imported by the Streamlit page (02_HSG17_T1_to_T0_Slack.py) and/or other tools.

The output Excel it produces must be identical to what the v2 reference script produces
on the same inputs (cutsheet Installation Sheet + Slack-style validation reports).

Key v2 updates applied:
- A-side columns in Downlinks/Mismatches use "Act." prefix (Act. Hostname, Act. Interface, Act. L/R, Act. Rack, Act. Elevation).
- Generalized mismatch grouping via Union-Find on (Act. side <-> Possible side) to handle pairs *and* groups of 3+ interconnected mismatches.
- Removed all orange/yellow fill coloring for mismatch groups.
- Groups are delineated with bold (medium) grid lines: thick top border on first row of group, thick bottom border on last row of group (thin borders otherwise).
- New optional "Mismatch↔Downlink" cross-reference tab (mismatches whose Expected side appears as a downlink).
- Updated L/R, fill, anchor, and rack-collection logic for the Act. column names.
- All other formatting, pink Possible/Active-Z columns, Summary, grey-out, autofit, NOTE+filter, tab order, and final top-2-rack rename preserved.
'''

import os
import shutil
from collections import Counter

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ── Tab-name resolution ──────────────────────────────────────────────────────
# The source-report generator periodically renames tabs (e.g. adding a
# "_with_pp" suffix). To survive these renames, look up tabs by ANY of their
# known aliases instead of hard-coding a single name.

TAB_ALIASES = {
    'lldp':         ('lldp_sp', 'full_path_lldp_with_int_down'),
    'optics':       ('optics_rx_tx_threshold', 'optics_rx_tx_threshold_with_pp'),
    'interfaces':   ('interfaces_sp', 'interfaces_sp_with_pp'),
    'combined_fec': ('combined_fec', 'combined_fec_with_pp'),
}

# Tabs that should always be dropped from the output (case-insensitive match).
# Add or remove names here as the source report evolves.
TABS_TO_REMOVE = (
    'device_reporting_failure',
    'bgp_sp',
    'spectrum_health',
    'sp_power',
    'sp_fans',
    'optics_temp',
    'pre_fec_ber_threshold_with_pp',
    'unknown_test_sp',
    'summary',    # source report summary; our Summary tab replaces it
)

# Columns to strip from every tab in the final output. Exact-match,
# case-sensitive. Edit this list to drop more columns globally.
COLUMNS_TO_REMOVE = (
    'Building',
    'Act. Building',
    'Exp. Building',
    'PP_A',
    'PP_Z',
    # combined_fec source noise:
    'Remote Host',
    'Remote Interface',
    'Mapped Remote Host',
    'Mapped Remote Interface',
    'Mapped Remote Rack',
    'Mapped Remote Elevation',
    'Remote Host Match',
    'Remote Interface Match',
    'Remote End Match',
    'Z_end_host',
    'Z_end_intf',
    'rack_z',
    'Z_Rack',
    'Z_Elevation',
    'Index',
    'Source Sheet',
    'Placement Group',
)

# Tabs that should receive Z-side info (Z Hostname, Z Interface,
# Z Rack, Z Elevation) pulled from the cutsheet by Hostname+Interface match.
# Columns are inserted right after Destination_port. Add more tab names
# here if you want Z-side info in additional tabs.
Z_FILL_TABS = ('Optics', 'combined_fec')


def find_tab(wb_or_sheetnames, key):
    """Return the actual tab name in the workbook for the given logical key,
    or None if no alias is present."""
    names = (wb_or_sheetnames.sheetnames
             if hasattr(wb_or_sheetnames, 'sheetnames')
             else list(wb_or_sheetnames))
    for alias in TAB_ALIASES[key]:
        if alias in names:
            return alias
    return None


# ── Style helpers ────────────────────────────────────────────────────────────

PINK   = 'FFB6C1'
YELLOW = 'FFFF00'
ORANGE = 'FFA500'   # reciprocal-swap pair highlight


def thin_border():
    s = Side(style='thin', color='000000')
    return Border(left=s, right=s, top=s, bottom=s)


def clear_and_border(ws, pink_cols=None):
    """Remove all fills (preserve pink cols); yellow-highlight row 1; apply black border everywhere."""
    bd        = thin_border()
    no_fill   = PatternFill(fill_type=None)
    pink_fill = PatternFill('solid', start_color=PINK)
    yellow_fill = PatternFill('solid', start_color=YELLOW)
    pink_cols = set(pink_cols or [])
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row,
                             min_col=1, max_col=ws.max_column):
        for cell in row:
            if cell.row == 1:
                cell.fill = yellow_fill
            elif cell.column in pink_cols:
                cell.fill = pink_fill
            else:
                cell.fill = no_fill
            cell.border = bd
            if cell.font:
                cell.font = Font(
                    bold=cell.font.bold,
                    name=cell.font.name or 'Arial',
                    size=cell.font.size or 10,
                    color='FF000000'
                )


def header_cell(cell, value, fill=None):
    cell.value     = value
    cell.font      = Font(bold=True, name='Arial', size=10)
    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=False)
    cell.border    = thin_border()
    cell.fill      = fill if fill else PatternFill('solid', start_color=YELLOW)


def data_cell(cell, value, fill=None):
    cell.value  = value
    cell.border = thin_border()
    if fill:
        cell.fill = fill


def autofit_sheet(ws, header_row_height=24, data_row_height=20, max_col_width=80):
    """Expand columns to fit content and give rows a comfortable height.

    - Column width = longest cell content in that column (+ padding), capped
      at `max_col_width` so a single huge value can't blow up the layout.
    - Cells inside a merged range are ignored when measuring column width.
    - Row 1 gets `header_row_height`; remaining rows get `data_row_height`.
    """
    # Cells that are part of a merge — skip them when measuring widths
    merged = set()
    for mrange in ws.merged_cells.ranges:
        for r in range(mrange.min_row, mrange.max_row + 1):
            for c in range(mrange.min_col, mrange.max_col + 1):
                merged.add((r, c))

    col_max = {}
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row,
                             min_col=1, max_col=ws.max_column):
        for cell in row:
            if cell.value is None:
                continue
            if (cell.row, cell.column) in merged:
                continue
            # Handle multi-line values: width = longest line
            longest_line = max(
                (len(line) for line in str(cell.value).splitlines()),
                default=0
            )
            letter = get_column_letter(cell.column)
            if longest_line > col_max.get(letter, 0):
                col_max[letter] = longest_line

    for letter, length in col_max.items():
        ws.column_dimensions[letter].width = min(length + 4, max_col_width)

    # Row heights
    if ws.max_row >= 1:
        ws.row_dimensions[1].height = header_row_height
    for r in range(2, ws.max_row + 1):
        ws.row_dimensions[r].height = data_row_height


def write_sheet(wb, name, df):
    ws = wb.create_sheet(name)
    bd = thin_border()
    for c, col in enumerate(df.columns, 1):
        header_cell(ws.cell(row=1, column=c), col)
    for r, (_, row) in enumerate(df.iterrows(), 2):
        for c, col in enumerate(df.columns, 1):
            val = row[col]
            cell = ws.cell(row=r, column=c, value=None if pd.isna(val) else val)
            cell.border = bd
    for c, col in enumerate(df.columns, 1):
        mx = max([len(str(col))] + [len(str(v)) for v in df[col].dropna()])
        ws.column_dimensions[get_column_letter(c)].width = min(mx + 2, 40)
    ws.freeze_panes = 'A2'
    return ws


# ── Reference-file loaders ───────────────────────────────────────────────────

def load_cutsheet(path):
    """Load the allconnections cutsheet.

    Supports two schemas:
      • Legacy  — sheet named 'Installation Sheet' with columns Hostname,
                  Interface, Z Hostname, Z Interface already split out.
      • New     — sheet named 'Sheet1' where DeviceA/DeviceB are combined
                  'hostname interface' strings. This function normalises both
                  into the same frame with canonical column names so the rest
                  of the pipeline is unchanged.
    """
    import re as _re
    xl = pd.ExcelFile(path)

    if 'Installation Sheet' in xl.sheet_names:
        return pd.read_excel(path, sheet_name='Installation Sheet')

    # ── New schema: Sheet1 with DeviceA / DeviceB combined ──────────────────
    df = pd.read_excel(path, sheet_name=xl.sheet_names[0])

    def _split_device(col):
        """'hostname swpXsY' → (hostname, swpXsY)"""
        parts = df[col].str.rsplit(' ', n=1, expand=True)
        return parts[0].str.strip(), parts[1].str.strip()

    df['Hostname'],   df['Interface']   = _split_device('DeviceA')
    df['Z Hostname'], df['Z Interface'] = _split_device('DeviceB')

    # L/R: last character of DeviceA/B Physical Port ('2L' → 'L')
    def _lr(col):
        return df[col].apply(
            lambda v: str(v).strip()[-1] if pd.notna(v) and str(v).strip() else '')

    df['L/R']   = _lr('DeviceA Physical Port')
    df['Z L/R'] = _lr('DeviceB Physical Port')

    # Rack / Elevation: extracted from 'Rack NNNN UN'
    def _rack(col):
        return (df[col].str.extract(r'Rack\s+(\d+)')[0]
                .astype(float).fillna(0).astype(int))

    def _elev(col):
        return (df[col].str.extract(r'U(\d+)')[0]
                .astype(float).fillna(0).astype(int))

    df['Rack']        = _rack('RackA')
    df['Elevation']   = _elev('RackA')
    df['Z Rack']      = _rack('RackB')
    df['Z Elevation'] = _elev('RackB')

    return df


def build_cutsheet_lookup(cut_df):
    """Key: (Hostname, Interface) → row dict for Source_port etc.

    Adapts to schema changes: only carries forward the fill columns that
    actually exist in the cutsheet. DMARC1/DMARC2 were dropped in later
    cutsheet revisions; the lookup still works without them.
    """
    candidate_cols = ['Source_port', 'DMARC1', 'DMARC2', 'Destination_port']
    fill_cols = [c for c in candidate_cols if c in cut_df.columns]
    lookup = {}
    for _, row in cut_df.iterrows():
        key = (
            str(row['Hostname']).strip(),
            str(row['Interface']).strip(),
        )
        lookup[key] = {c: row[c] for c in fill_cols}
    # Stash which columns were actually available so step 6 can use the same set
    lookup['__fill_cols__'] = fill_cols
    return lookup


def build_z_lookup(cut_df):
    """Key: (Z Hostname, Z Interface) → full row. Z-side is 1:1 unique in
    the cutsheet (same as A-side), so the simpler 2-tuple key is reliable
    and matches the (Hostname, Interface) pattern used elsewhere."""
    lookup = {}
    for _, row in cut_df.iterrows():
        key = (
            str(row['Z Hostname']).strip(),
            str(row['Z Interface']).strip(),
        )
        lookup[key] = row
    return lookup


def paired_subport(iface):
    """Return the paired sub-port interface name.

    s0 ↔ s1 are a pair, s2 ↔ s3 are a pair.
    e.g. 'swp4s0' → 'swp4s1', 'swp15s3' → 'swp15s2'.
    Returns None if iface doesn't end in s0/s1/s2/s3.
    """
    pairs = {'s0': 's1', 's1': 's0', 's2': 's3', 's3': 's2'}
    for suffix, mate in pairs.items():
        if iface.endswith(suffix):
            return iface[:-len(suffix)] + mate
    return None


# ── Core processor ───────────────────────────────────────────────────────────

def process_file(input_path, output_path, cut_df, log):
    shutil.copy2(input_path, output_path)
    wb = load_workbook(output_path)

    cutsheet_lookup = build_cutsheet_lookup(cut_df)
    z_lookup        = build_z_lookup(cut_df)

    # ── 1. Split lldp tab → Downlinks / Mismatches ──────────────────────────
    mis_orig_df = None   # keep for step 6b
    lldp_tab = find_tab(wb, 'lldp')
    if lldp_tab:
        log(f"  · Splitting {lldp_tab} → Downlinks / Mismatches")
        df = pd.read_excel(input_path, sheet_name=lldp_tab)
        down_df     = df[df['Act. Interface'] == 'interface down'].copy()
        mis_orig_df = df[df['Act. Interface'].str.startswith('swp', na=False)].copy()
        # Downlinks: drop only the Active-side columns; keep the Expected-side
        # columns and order them right after Elevation. Source_port and
        # Destination_port will be inserted between Elevation and Expected*
        # by step 6, giving the final order:
        #   Act. Hostname, Act. Interface, Act. Rack, Act. Elevation,
        #   Source_port, Destination_port,
        #   Expected Hostname, Exp. Interface, Exp. Rack, Exp. Elevation
        drop = ['Active Host', 'Act. Interface', 'Act. Rack', 'Act. Elevation']
        down_df.drop(columns=[c for c in drop if c in down_df.columns], inplace=True)
        # Rename A-side columns to Act. prefix (v2)
        down_df.rename(columns={
            'Hostname':  'Act. Hostname',
            'Interface': 'Act. Interface',
            'Rack':      'Act. Rack',
            'Elevation': 'Act. Elevation',
        }, inplace=True)
        # Pull Expected* to the end (relative order preserved) so they sit
        # after Elevation in the written sheet.
        exp_cols = ['Expected Hostname', 'Exp. Interface', 'Exp. Rack', 'Exp. Elevation']
        present_exp = [c for c in exp_cols if c in down_df.columns]
        if present_exp:
            rest = [c for c in down_df.columns if c not in present_exp]
            down_df = down_df[rest + present_exp]
        del wb[lldp_tab]
        write_sheet(wb, 'Downlinks', down_df)
        # Rename A-side columns in Mismatches too before writing (v2).
        # Drop the source report's Active-side columns first to avoid duplicate
        # column names (the source tab already has Act. Interface etc).
        mis_write_df = mis_orig_df.copy()
        act_src_drop = ['Active Host', 'Act. Interface', 'Act. Rack', 'Act. Elevation']
        mis_write_df.drop(columns=[c for c in act_src_drop if c in mis_write_df.columns],
                          inplace=True)
        mis_write_df.rename(columns={
            'Hostname':  'Act. Hostname',
            'Interface': 'Act. Interface',
            'Rack':      'Act. Rack',
            'Elevation': 'Act. Elevation',
        }, inplace=True)
        write_sheet(wb, 'Mismatches', mis_write_df)

    # ── 2. Optics tab ───────────────────────────────────────────────────────
    optics_src = find_tab(wb, 'optics')
    if optics_src:
        log(f"  · Processing Optics tab ({optics_src})")
        # Extra cols introduced by the *_with_pp variant — drop them too.
        drop_cols = {'Transceiver', 'Channel',
                     'Min Threshold (dBm)', 'Max Threshold (dBm)',
                     'PP_A', 'PP_Z', 'Z_end_host', 'Z_end_intf',
                     'rack_z', 'Z_Rack', 'Z_Elevation', 'Index',
                     'Status', 'Placement Group'}
        optics_df = pd.read_excel(input_path, sheet_name=optics_src)
        optics_df.drop(columns=[c for c in drop_cols if c in optics_df.columns], inplace=True)
        # Put 'Metric' first, then 'Measured (dBm)', so both columns can be
        # frozen and stay visible while scrolling horizontally.
        leading = [c for c in ('Metric', 'Measured (dBm)') if c in optics_df.columns]
        if leading:
            rest = [c for c in optics_df.columns if c not in leading]
            optics_df = optics_df[leading + rest]
        del wb[optics_src]
        write_sheet(wb, 'Optics', optics_df)
        # Freeze row 1 + the two leading columns (Metric, Measured (dBm)).
        wb['Optics'].freeze_panes = 'C2' if len(leading) >= 2 else 'B2'

    # ── 3. Remove interfaces tab ────────────────────────────────────────────
    interfaces_tab = find_tab(wb, 'interfaces')
    if interfaces_tab:
        log(f"  · Removing {interfaces_tab}")
        del wb[interfaces_tab]

    # ── 3a. Remove unwanted source tabs ─────────────────────────────────────
    drop_lower = {t.lower() for t in TABS_TO_REMOVE}
    for existing in list(wb.sheetnames):
        if existing.lower() in drop_lower:
            log(f"  · Removing {existing}")
            del wb[existing]

    # ── 3b. combined_fec: move Lock Status + Pre-FEC BER to the front ───────
    fec_tab = find_tab(wb, 'combined_fec')
    if fec_tab:
        log(f"  · Reordering {fec_tab} (Lock Status, Pre-FEC BER first)")
        fec_df = pd.read_excel(input_path, sheet_name=fec_tab)

        def _norm(s):
            # Normalize hyphen variants and whitespace for tolerant matching
            return (str(s)
                    .replace('\u2011', '-')   # non-breaking hyphen
                    .replace('\u2013', '-')   # en dash
                    .replace('\u2014', '-')   # em dash
                    .strip()
                    .lower())

        wanted = ['lock status', 'pre-fec ber']
        front = []
        for target in wanted:
            for col in fec_df.columns:
                if _norm(col) == target and col not in front:
                    front.append(col)
                    break

        if front:
            rest = [c for c in fec_df.columns if c not in front]
            fec_df = fec_df[front + rest]
            del wb[fec_tab]
            write_sheet(wb, 'combined_fec', fec_df)
        else:
            log("    ⚠ Lock Status / Pre-FEC BER not found — leaving combined_fec as-is")

    # ── 4. Reorder tabs ─────────────────────────────────────────────────────
    # Put the four primary output tabs at the END in the documented order,
    # everything else stays in its current relative order. Rewriting _sheets
    # directly is simpler and more robust than chained move_sheet() calls
    # whose offset arithmetic depended on the count of intermediate tabs.
    # cannot_decode_sp sits between Summary and Downlinks in the target layout
    desired  = ['Downlinks', 'Mismatches', 'Optics', 'combined_fec', 'cannot_decode_sp']
    existing = [s for s in desired if s in wb.sheetnames]
    others   = [s for s in wb.sheetnames if s not in desired]
    wb._sheets = [wb[n] for n in others + existing]

    # ── 5. Insert L/R columns (sourced from cutsheet) ───────────────────────
    log("  · Adding L/R mapped columns (from cutsheet)")
    # Build interface→L/R map from both A-side and Z-side cutsheet columns.
    lr_from_cut = {}
    for _, row in cut_df.iterrows():
        iface = str(row.get('Interface', '') or '').strip()
        lr    = str(row.get('L/R', '')    or '').strip()
        if iface and lr:
            lr_from_cut[iface] = lr
        z_iface = str(row.get('Z Interface', '') or '').strip()
        z_lr    = str(row.get('Z L/R', '')      or '').strip()
        if z_iface and z_lr:
            lr_from_cut[z_iface] = z_lr

    lr_name_for = {
        'Act. Interface': 'Act. L/R',
        'Interface':      'L/R',
        'Z Interface':    'Z L/R',
        'Exp. Interface': 'Exp. L/R',
    }
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        header = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
        targets = [(i+1, h) for i, h in enumerate(header) if h in lr_name_for]
        for col_idx, col_name in sorted(targets, reverse=True):
            new_name = lr_name_for[col_name]
            ws.insert_cols(col_idx + 1)
            header_cell(ws.cell(row=1, column=col_idx + 1), new_name)
            for r in range(2, ws.max_row + 1):
                val = str(ws.cell(row=r, column=col_idx).value or '').strip()
                ws.cell(row=r, column=col_idx + 1, value=lr_from_cut.get(val, ''))
                ws.cell(row=r, column=col_idx + 1).border = thin_border()
            ws.column_dimensions[get_column_letter(col_idx + 1)].width = 10

    # ── 5b. Populate Source_port / DMARC1 / DMARC2 / Destination_port ────────
    # Match each row's (Hostname, Interface) against the cutsheet and fill
    # the available cutsheet columns. New fill columns get inserted right
    # after Elevation when the target sheet doesn't already have them.
    fill_cols = cutsheet_lookup.get('__fill_cols__',
                                    ['Source_port', 'DMARC1', 'DMARC2', 'Destination_port'])
    log(f"  · Filling {', '.join(fill_cols) or '(no cutsheet fill cols available)'} (match on Hostname + Interface)")
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        header = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
        if not all(c in header for c in ['Act. Hostname', 'Act. Interface']) and \
           not all(c in header for c in ['Hostname', 'Interface']):
            continue
        host_col = 'Act. Hostname' if 'Act. Hostname' in header else 'Hostname'
        int_col  = 'Act. Interface' if 'Act. Interface' in header else 'Interface'
        anchor_col = next((c for c in ['Act. Elevation', 'Elevation'] if c in header), None)
        anchor = (header.index(anchor_col) + 1) if anchor_col else len(header)
        insert_at = anchor + 1
        for col_name in fill_cols:
            if col_name in header:
                continue
            ws.insert_cols(insert_at)
            header_cell(ws.cell(row=1, column=insert_at), col_name)
            for r in range(2, ws.max_row + 1):
                ws.cell(row=r, column=insert_at).border = thin_border()
            ws.column_dimensions[get_column_letter(insert_at)].width = max(len(col_name)+2, 14)
            insert_at += 1
        header = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
        host_c = header.index(host_col)+1
        int_c  = header.index(int_col)+1
        fill_idx = {c: header.index(c)+1 for c in fill_cols}
        for r in range(2, ws.max_row + 1):
            host  = str(ws.cell(row=r, column=host_c).value or '').strip()
            iface = str(ws.cell(row=r, column=int_c).value or '').strip()
            match = cutsheet_lookup.get((host, iface))
            if match:
                for col_name, col_idx in fill_idx.items():
                    val = match.get(col_name)
                    if val is not None and not (isinstance(val, float) and pd.isna(val)):
                        ws.cell(row=r, column=col_idx, value=val)

    # ── 6c. Fill Z-side info in designated tabs (default: Optics) ───────────
    Z_COLS = ['Z Hostname', 'Z Interface', 'Z L/R', 'Z Rack', 'Z Elevation']
    z_available = [c for c in Z_COLS if c in cut_df.columns]
    # Build unconditionally — also used by step 6b Mismatches act_lookup for Z L/R
    z_by_host_int = {}
    for _, row in cut_df.iterrows():
        k = (str(row['Hostname']).strip(), str(row['Interface']).strip())
        z_by_host_int[k] = {c: row[c] for c in z_available}

    if z_available and Z_FILL_TABS:
        for tab in Z_FILL_TABS:
            if tab not in wb.sheetnames:
                continue
            ws_z = wb[tab]
            header = [ws_z.cell(row=1, column=c).value
                      for c in range(1, ws_z.max_column + 1)]
            if not all(c in header for c in ['Act. Hostname', 'Act. Interface']) and \
               not all(c in header for c in ['Hostname', 'Interface']):
                continue
            log(f"  · Filling Z-side info in {tab}: {', '.join(z_available)}")
            # Anchor: right after Destination_port if present, else after
            # Act. Elevation / Elevation, else at the end of the sheet.
            if 'Destination_port' in header:
                anchor = header.index('Destination_port') + 1
            elif 'Act. Elevation' in header:
                anchor = header.index('Act. Elevation') + 1
            elif 'Elevation' in header:
                anchor = header.index('Elevation') + 1
            else:
                anchor = len(header)
            insert_at = anchor + 1
            for col_name in z_available:
                if col_name in header:
                    continue
                ws_z.insert_cols(insert_at)
                header_cell(ws_z.cell(row=1, column=insert_at), col_name)
                for r in range(2, ws_z.max_row + 1):
                    ws_z.cell(row=r, column=insert_at).border = thin_border()
                ws_z.column_dimensions[get_column_letter(insert_at)].width = max(len(col_name)+2, 14)
                insert_at += 1
            header = [ws_z.cell(row=1, column=c).value
                      for c in range(1, ws_z.max_column + 1)]
            z_host_col = 'Act. Hostname' if 'Act. Hostname' in header else 'Hostname'
            z_int_col  = 'Act. Interface' if 'Act. Interface' in header else 'Interface'
            host_c = header.index(z_host_col)+1
            int_c  = header.index(z_int_col)+1
            fill_idx = {c: header.index(c)+1 for c in z_available}
            for r in range(2, ws_z.max_row + 1):
                host  = str(ws_z.cell(row=r, column=host_c).value or '').strip()
                iface = str(ws_z.cell(row=r, column=int_c).value or '').strip()
                match = z_by_host_int.get((host, iface))
                if match:
                    for col_name, col_idx in fill_idx.items():
                        val = match.get(col_name)
                        if val is not None and not (isinstance(val, float) and pd.isna(val)):
                            ws_z.cell(row=r, column=col_idx, value=val)

    # ── 6b. Mismatches: Possible columns + Active Z columns (pink) ───────────
    if 'Mismatches' in wb.sheetnames:
        log("  · Building Possible + Active Z columns in Mismatches")
        pink_fill   = PatternFill('solid', start_color=PINK)
        yellow_fill = PatternFill('solid', start_color=YELLOW)
        bd          = thin_border()

        # Build act_lookup from original lldp tab (alias-aware)
        # z_by_host_int built earlier for Z-fill is reused here to get Z L/R from cutsheet
        act_lookup = {}
        src_sheets = pd.ExcelFile(input_path).sheet_names
        src_lldp = find_tab(src_sheets, 'lldp')
        if src_lldp:
            orig_df  = pd.read_excel(input_path, sheet_name=src_lldp)
            mis_rows = orig_df[orig_df['Act. Interface'].str.startswith('swp', na=False)]
            for _, row in mis_rows.iterrows():
                key       = (str(row['Hostname']).strip(), str(row['Interface']).strip())
                act_host  = str(row['Active Host']).strip()
                act_iface = str(row['Act. Interface']).strip()
                # Pull Z L/R from cutsheet: Active Host/Interface is the Z-side,
                # so look it up in z_lookup (keyed by Z Hostname, Z Interface).
                # Fallback to paired sub-port (s0↔s1, s2↔s3) if exact miss.
                cut_z_row = z_lookup.get((act_host, act_iface))
                if cut_z_row is None:
                    mate = paired_subport(act_iface)
                    if mate:
                        cut_z_row = z_lookup.get((act_host, mate))
                z_lr_val = ''
                if cut_z_row is not None:
                    raw = cut_z_row.get('Z L/R')
                    if raw is not None and not (isinstance(raw, float) and pd.isna(raw)):
                        z_lr_val = str(raw).strip()
                act_lookup[key] = {
                    'Z Hostname' : act_host,
                    'Z Interface': act_iface,
                    'Z L/R'      : z_lr_val,
                    'Z Rack'     : int(float(str(row['Act. Rack']))),
                    'Z Elevation': int(float(str(row['Act. Elevation']))),
                }

        ws_m   = wb['Mismatches']
        header = [ws_m.cell(row=1, column=c).value for c in range(1, ws_m.max_column + 1)]

        # In v2 the write step already stripped most Active cols and renamed A-side to Act.*.
        # Only drop any lingering 'Active Host' here.
        act_drop = {'Active Host'}
        for idx in sorted([i+1 for i, h in enumerate(header) if h in act_drop], reverse=True):
            ws_m.delete_cols(idx)
        header = [ws_m.cell(row=1, column=c).value for c in range(1, ws_m.max_column + 1)]

        h_idx = header.index('Act. Hostname') + 1
        i_idx = header.index('Act. Interface') + 1

        # Collect act data per row
        act_rows = []
        for r in range(2, ws_m.max_row + 1):
            hn    = str(ws_m.cell(row=r, column=h_idx).value or '').strip()
            iface = str(ws_m.cell(row=r, column=i_idx).value or '').strip()
            act_rows.append(act_lookup.get((hn, iface), {}))

        # Possible columns: match act Z key against cutsheet Z side.
        # Filter out ones whose source column doesn't exist in this cutsheet
        # (e.g. DMARC1 / DMARC2 were dropped in later cutsheet revisions).
        possible_cols_all = [
            ('Possible Hostname',         'Hostname'),
            ('Possible Interface',        'Interface'),
            ('Possible L/R',             'L/R'),
            ('Possible Rack',             'Rack'),
            ('Possible Elevation',        'Elevation'),
            ('Possible Source_port',      'Source_port'),
            ('Possible DMARC1',           'DMARC1'),
            ('Possible DMARC2',           'DMARC2'),
            ('Possible Destination_port', 'Destination_port'),
        ]
        cut_cols = set(cut_df.columns)
        possible_cols = [
            (out_col, src) for out_col, src in possible_cols_all
            if src in cut_cols
        ]

        possible_data = {col: [] for col, _ in possible_cols}
        for act in act_rows:
            zh   = act.get('Z Hostname', '')
            zi   = act.get('Z Interface', '')
            match = z_lookup.get((zh, zi)) if zh else None
            # Fallback: if exact sub-port not in cutsheet, try its pair.
            # s0↔s1 are a pair, s2↔s3 are a pair.
            if match is None and zh and zi:
                mate = paired_subport(zi)
                if mate:
                    match = z_lookup.get((zh, mate))
            for col, src in possible_cols:
                if match is not None:
                    val = match.get(src, '')
                else:
                    val = ''
                possible_data[col].append(val)

        # Write Possible columns
        pink_col_indices = []
        start = ws_m.max_column + 1
        for c_off, (col_name, _) in enumerate(possible_cols):
            col_idx = start + c_off
            pink_col_indices.append(col_idx)
            hdr = ws_m.cell(row=1, column=col_idx, value=col_name)
            hdr.font = Font(bold=True, name='Arial', size=10)
            hdr.alignment = Alignment(horizontal='center', vertical='center', wrap_text=False)
            hdr.fill  = yellow_fill
            hdr.border = bd
            ws_m.column_dimensions[get_column_letter(col_idx)].width = max(len(col_name)+3, 14)
            for r_off, val in enumerate(possible_data[col_name]):
                cell = ws_m.cell(row=r_off+2, column=col_idx,
                                 value=val if val != '' else None)
                cell.fill   = pink_fill
                cell.border = bd

        # Write Active Z columns
        act_z_cols = ['Z Hostname', 'Z Interface', 'Z L/R', 'Z Rack', 'Z Elevation']
        start2 = ws_m.max_column + 1
        for c_off, col_name in enumerate(act_z_cols):
            col_idx = start2 + c_off
            pink_col_indices.append(col_idx)
            hdr = ws_m.cell(row=1, column=col_idx, value=col_name)
            hdr.font = Font(bold=True, name='Arial', size=10)
            hdr.alignment = Alignment(horizontal='center', vertical='center', wrap_text=False)
            hdr.fill  = yellow_fill
            hdr.border = bd
            ws_m.column_dimensions[get_column_letter(col_idx)].width = max(len(col_name)+3, 14)
            for r_off, act in enumerate(act_rows):
                val = act.get(col_name, '')
                cell = ws_m.cell(row=r_off+2, column=col_idx,
                                 value=val if val != '' else None)
                cell.fill   = pink_fill
                cell.border = bd

    # ── 6c. Strip unwanted columns across every tab ─────────────────────────
    if COLUMNS_TO_REMOVE:
        log(f"  · Stripping columns: {', '.join(COLUMNS_TO_REMOVE)}")
        drop_set = set(COLUMNS_TO_REMOVE)
        for sheet_name in wb.sheetnames:
            ws_x = wb[sheet_name]
            header = [ws_x.cell(row=1, column=c).value
                      for c in range(1, ws_x.max_column + 1)]
            # Find 1-based indices to drop, delete right-to-left so earlier
            # indices stay valid as columns shift left.
            to_drop = [i + 1 for i, h in enumerate(header) if h in drop_set]
            for idx in sorted(to_drop, reverse=True):
                ws_x.delete_cols(idx)

    # ── 7. Summary tab (per Rack breakdown) ─────────────────────────────────
    log("  · Creating Summary tab")

    # Gather rack counts per sheet from the workbook data
    tab_rack  = {}
    all_racks = set()
    no_fill_s   = PatternFill(fill_type=None)
    yellow_fill_s = PatternFill('solid', start_color=YELLOW)
    center_s  = Alignment(horizontal='center', vertical='center', wrap_text=False)
    bd_s      = thin_border()

    def _s(cell, value, bold=False, header=False):
        cell.value     = value
        cell.font      = Font(bold=bold, name='Arial', size=10)
        cell.alignment = center_s
        cell.border    = bd_s
        cell.fill      = yellow_fill_s if header else no_fill_s

    for sname in wb.sheetnames:
        ws_tmp = wb[sname]
        hdr = [ws_tmp.cell(row=1, column=c).value for c in range(1, ws_tmp.max_column+1)]
        if 'Rack' not in hdr:
            tab_rack[sname] = {}
            continue
        rack_col = hdr.index('Rack') + 1
        counts = {}
        for r in range(2, ws_tmp.max_row + 1):
            val = ws_tmp.cell(row=r, column=rack_col).value
            if val is not None:
                try:
                    k = int(float(str(val)))
                    counts[k] = counts.get(k, 0) + 1
                    all_racks.add(k)
                except ValueError:
                    pass
        tab_rack[sname] = counts

    racks      = sorted(all_racks)
    tabs_order = [n for n in wb.sheetnames]
    total_cols = 1 + len(racks) + 1  # Tab Name + per rack + Total

    # Excel treats sheet names case-insensitively. The new report ships a
    # lowercase 'summary' tab; delete any case-variant before creating ours.
    for existing in list(wb.sheetnames):
        if existing.lower() == 'summary':
            del wb[existing]
    wb.create_sheet('Summary', 0)
    ws_s = wb['Summary']

    # Title
    ws_s.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_cols)
    title_c = ws_s.cell(row=1, column=1, value='Tab Summary by Rack')
    title_c.font = Font(bold=True, name='Arial', size=13)
    title_c.alignment = center_s
    title_c.border    = bd_s
    title_c.fill      = yellow_fill_s
    ws_s.row_dimensions[1].height = 28

    # Header row
    _s(ws_s.cell(row=2, column=1), 'Tab Name', bold=True, header=True)
    for c, rack in enumerate(racks, start=2):
        _s(ws_s.cell(row=2, column=c), str(rack), bold=True, header=True)
    _s(ws_s.cell(row=2, column=total_cols), 'Total', bold=True, header=True)

    # Data rows (exclude the Summary tab itself — case-insensitive, since the
    # source report may ship a lowercase 'summary' tab that we delete later)
    rack_totals = {r: 0 for r in racks}
    data_tabs   = [n for n in tabs_order if n.lower() != 'summary']
    for i, tab_name in enumerate(data_tabs, start=3):
        _s(ws_s.cell(row=i, column=1), tab_name)
        row_total = 0
        for c, rack in enumerate(racks, start=2):
            count = tab_rack.get(tab_name, {}).get(rack, 0)
            _s(ws_s.cell(row=i, column=c), count if count > 0 else '')
            rack_totals[rack] += count
            row_total += count
        _s(ws_s.cell(row=i, column=total_cols), row_total, bold=True)

    # Grand total row
    tot_r = 3 + len(data_tabs)
    _s(ws_s.cell(row=tot_r, column=1), 'TOTAL', bold=True)
    grand = 0
    for c, rack in enumerate(racks, start=2):
        _s(ws_s.cell(row=tot_r, column=c), rack_totals[rack], bold=True)
        grand += rack_totals[rack]
    _s(ws_s.cell(row=tot_r, column=total_cols), grand, bold=True)

    # Column widths
    ws_s.column_dimensions['A'].width = 20
    for c in range(2, total_cols + 1):
        ws_s.column_dimensions[get_column_letter(c)].width = 14

    # ── 8. No fill + borders (preserve pink in Mismatches) ──────────────────
    log("  · Removing fills and applying borders")
    # The earlier pink_col_indices was captured in step 6b, but step 6c (column
    # strip) deletes columns from the middle of Mismatches afterwards, which
    # shifts all Possible/Z columns left and invalidates those stale indices.
    # Recompute by name from the current header so every pink column gets
    # filled correctly regardless of how many columns were stripped.
    if 'Mismatches' in wb.sheetnames:
        ws_m = wb['Mismatches']
        m_header = [ws_m.cell(row=1, column=c).value
                    for c in range(1, ws_m.max_column + 1)]
        Z_NAMES = {'Z Hostname', 'Z Interface', 'Z L/R', 'Z Rack', 'Z Elevation'}
        pink_col_indices = [
            i + 1 for i, h in enumerate(m_header)
            if (h and (str(h).startswith('Possible ') or h in Z_NAMES))
        ]

    for sheet_name in wb.sheetnames:
        pcols = pink_col_indices if sheet_name == 'Mismatches' else []
        clear_and_border(wb[sheet_name], pink_cols=pcols)

    # ── 8c. Centre-align all cells across all tabs ──────────────────────────
    log("  · Aligning all cells to middle-centre")
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=False)
    for sheet_name in wb.sheetnames:
        for row in wb[sheet_name].iter_rows():
            for cell in row:
                cell.alignment = center_align

    # ── 8b. Add NOTE column + autofilter to all tabs ───────────────────────
    log("  · Adding NOTE column and filters to all tabs")
    no_fill     = PatternFill(fill_type=None)
    yellow_fill = PatternFill('solid', start_color=YELLOW)
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        for col_name in ['NOTE']:
            col_idx = ws.max_column + 1
            hdr = ws.cell(row=1, column=col_idx, value=col_name)
            hdr.font      = Font(bold=True, name='Arial', size=10)
            hdr.alignment = Alignment(horizontal='center', vertical='center', wrap_text=False)
            hdr.fill      = yellow_fill
            hdr.border    = thin_border()
            ws.column_dimensions[get_column_letter(col_idx)].width = 16
            for r in range(2, ws.max_row + 1):
                cell = ws.cell(row=r, column=col_idx)
                cell.fill   = no_fill
                cell.border = thin_border()
        if ws.max_row > 1 and ws.max_column > 0:
            ws.auto_filter.ref = ws.dimensions

    # ── 8d. Grey-out Optics rows that are matched in Downlinks ──────────────
    if 'Optics' in wb.sheetnames and 'Downlinks' in wb.sheetnames:
        log("  · Greying out matched Optics rows")

        # Use Act. variants for Downlinks (v2 naming) + the enriched columns.
        MATCH_COLS = [
            'Act. Hostname', 'Act. Interface', 'Act. L/R', 'Act. Rack', 'Act. Elevation',
            'Source_port', 'DMARC1', 'DMARC2', 'Destination_port',
            'Z Hostname', 'Z Interface', 'Z L/R', 'Z Rack', 'Z Elevation',
        ]
        GREY_FONT_COLOR = 'FFD3D3D3'  # light grey

        ws_dl = wb['Downlinks']
        dl_header = [ws_dl.cell(row=1, column=c).value
                     for c in range(1, ws_dl.max_column + 1)]

        # Build a set of tuples from Downlinks for all match columns present
        dl_match_cols = [c for c in MATCH_COLS if c in dl_header]
        dl_col_idx    = {c: dl_header.index(c) + 1 for c in dl_match_cols}

        dl_keys = set()
        for r in range(2, ws_dl.max_row + 1):
            key = tuple(
                str(ws_dl.cell(row=r, column=dl_col_idx[c]).value or '').strip()
                for c in dl_match_cols
            )
            dl_keys.add(key)

        ws_op = wb['Optics']
        op_header = [ws_op.cell(row=1, column=c).value
                     for c in range(1, ws_op.max_column + 1)]

        # Only match on columns present in both sheets
        common_cols  = [c for c in dl_match_cols if c in op_header]
        op_col_idx   = {c: op_header.index(c) + 1 for c in common_cols}
        dl_col_idx_c = {c: dl_header.index(c) + 1 for c in common_cols}

        # Rebuild dl_keys using only common columns
        dl_keys_common = set()
        for r in range(2, ws_dl.max_row + 1):
            key = tuple(
                str(ws_dl.cell(row=r, column=dl_col_idx_c[c]).value or '').strip()
                for c in common_cols
            )
            dl_keys_common.add(key)

        for r in range(2, ws_op.max_row + 1):
            op_key = tuple(
                str(ws_op.cell(row=r, column=op_col_idx[c]).value or '').strip()
                for c in common_cols
            )
            if op_key in dl_keys_common:
                for c in range(1, ws_op.max_column + 1):
                    cell = ws_op.cell(row=r, column=c)
                    cell.font = Font(
                        bold=cell.font.bold if cell.font else False,
                        name=(cell.font.name if cell.font else None) or 'Arial',
                        size=(cell.font.size if cell.font else None) or 10,
                        color=GREY_FONT_COLOR,
                    )

    # ── 8e. Expand all columns and rows on every sheet ──────────────────────
    log("  · Expanding all columns and rows")
    for sheet_name in wb.sheetnames:
        autofit_sheet(wb[sheet_name])

    # ── 8f. Highlight mismatch groups (any size, incl. 3+) with bold grid lines ─
    log("  · Highlighting mismatch groups (bold borders)")
    highlight_mismatch_pairs(wb, log)

    # ── 8g. Build Mismatch↔Downlink cross-reference tab (v2) ─────────────────
    log("  · Building Mismatch↔Downlink tab")
    build_downlink_mismatch_tab(wb, log)

    # Re-sort tabs now that Mismatch↔Downlink may have been added (primary
    # tabs stay in the documented order; the cross-ref tab goes with the rest).
    existing2 = [s for s in desired if s in wb.sheetnames]
    others2   = [s for s in wb.sheetnames if s not in desired]
    wb._sheets = [wb[n] for n in others2 + existing2]

    wb.save(output_path)

    # ── 9. Rename by top-2 Rack numbers ─────────────────────────────────────
    try:
        all_racks = []
        for sheet_name in wb.sheetnames:
            ws     = wb[sheet_name]
            header = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column+1)]
            if 'Rack' in header or 'Act. Rack' in header:
                rc = (header.index('Act. Rack') + 1 if 'Act. Rack' in header
                      else header.index('Rack') + 1)
                for r in range(2, ws.max_row + 1):
                    val = ws.cell(row=r, column=rc).value
                    if val is not None:
                        try: all_racks.append(int(float(str(val))))
                        except ValueError: pass
        if all_racks:
            top2     = [str(r) for r, _ in Counter(all_racks).most_common(2)]
            new_name = '+'.join(top2) + '.xlsx'
            new_path = os.path.join(os.path.dirname(output_path), new_name)
            load_workbook(output_path).save(new_path)
            if new_path != output_path:
                os.remove(output_path)
            log(f"  ✓ Saved → {new_name}")
            return new_path
    except Exception as e:
        log(f"  ⚠ Could not rename by Rack: {e}")

    log(f"  ✓ Saved → {os.path.basename(output_path)}")
    return output_path


# ── Mismatch group detection / bold-border grouping (v2) ─────────────────────

def _thick_side():
    """Medium (bold) side for group delimiters."""
    return Side(style='medium', color='000000')


def _group_border(is_first, is_last):
    """Return a border that puts a thick line on the top of the first row
    of a mismatch group and on the bottom of the last row of the group.
    All other sides (and non-group rows) use thin borders. This replaces
    the previous orange/yellow fill pairing with clear visual grouping via
    bold grid lines. Supports groups of any size (including 3+)."""
    top    = _thick_side() if is_first else Side(style='thin', color='000000')
    bottom = _thick_side() if is_last  else Side(style='thin', color='000000')
    s = Side(style='thin', color='000000')
    return Border(left=s, right=s, top=top, bottom=bottom)


def highlight_mismatch_pairs(wb, log=lambda *_: None):
    """Mismatches tab: detect swap cycles of ANY size using Union-Find,
    group matched rows together with no fill, and draw thick border lines
    above the first and below the last row of each group.

    Uses the Act.-prefixed A-side columns (from the v2 split/rename) matched
    against the Possible-* columns. Any connected component of size >1 is
    treated as a group (pairs, triples, larger cycles all supported).
    Rows within each group keep their relative original order; groups are
    ordered by the first appearance of their lowest original index.
    Singletons (unconnected) are placed after all groups.
    """
    if 'Mismatches' not in wb.sheetnames:
        return
    ws = wb['Mismatches']
    if ws.max_row < 3:
        return

    header = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]

    def idxs(names):
        return [header.index(n) + 1 for n in names if n in header]

    a_cols = idxs(['Act. Hostname', 'Act. Interface', 'Act. L/R', 'Act. Rack', 'Act. Elevation',
                   'Source_port', 'Destination_port'])
    p_cols = idxs(['Possible Hostname', 'Possible Interface', 'Possible L/R',
                   'Possible Rack', 'Possible Elevation',
                   'Possible Source_port', 'Possible Destination_port'])
    if not a_cols or not p_cols:
        return

    ncol, nrow = ws.max_column, ws.max_row
    rows = [[ws.cell(row=r, column=c).value for c in range(1, ncol + 1)]
            for r in range(2, nrow + 1)]

    def key(rv, cols):
        return tuple(str(rv[c - 1] if rv[c - 1] is not None else '').strip()
                     for c in cols)

    a_keys = [key(rv, a_cols) for rv in rows]
    p_keys = [key(rv, p_cols) for rv in rows]

    from collections import defaultdict
    p_key_map = defaultdict(list)
    for i, pk in enumerate(p_keys):
        if any(pk):
            p_key_map[pk].append(i)

    # Union-Find to connect any A that matches any P (transitive → 3+ groups)
    parent = list(range(len(rows)))
    def find(x):
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x
    def union(x, y):
        parent[find(x)] = find(y)

    for i, ak in enumerate(a_keys):
        if not any(ak):
            continue
        for j in p_key_map.get(ak, []):
            if i != j:
                union(i, j)

    comp = defaultdict(list)
    for i in range(len(rows)):
        comp[find(i)].append(i)

    grouped = [sorted(m) for m in comp.values() if len(m) > 1]
    singletons = [m[0] for m in comp.values() if len(m) == 1]

    order = []
    group_bounds = {}
    for members in sorted(grouped, key=lambda m: m[0]):
        for rank, src_i in enumerate(members):
            off = len(order)
            group_bounds[off] = (rank == 0, rank == len(members) - 1)
            order.append(src_i)
    for src_i in singletons:
        order.append(src_i)

    log(f"  · Found {len(grouped)} mismatch group(s) covering "
        f"{sum(len(m) for m in grouped)} rows "
        f"(sizes: {sorted([len(m) for m in grouped], reverse=True)})")

    pink_names = {'Possible Hostname', 'Possible Interface', 'Possible L/R',
                  'Possible Rack', 'Possible Elevation', 'Possible Source_port',
                  'Possible DMARC1', 'Possible DMARC2', 'Possible Destination_port',
                  'Z Hostname', 'Z Interface', 'Z L/R', 'Z Rack', 'Z Elevation'}
    pink_idx  = {header.index(nm) + 1 for nm in pink_names if nm in header}
    pink_fill = PatternFill('solid', start_color=PINK)
    no_fill   = PatternFill(fill_type=None)

    for out_off, src_i in enumerate(order):
        r  = out_off + 2
        rv = rows[src_i]
        bounds = group_bounds.get(out_off)
        for c in range(1, ncol + 1):
            cell = ws.cell(row=r, column=c, value=rv[c - 1])
            cell.font      = Font(name='Arial', size=10)
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.fill      = pink_fill if c in pink_idx else no_fill
            cell.border    = (_group_border(*bounds) if bounds is not None
                              else thin_border())


def build_downlink_mismatch_tab(wb, log=lambda *_: None):
    """Create a 'Mismatch↔Downlink' tab listing mismatch rows whose
    Expected Hostname + Exp. Interface appears in the Downlinks tab.

    This surfaces cases where a mismatch observed on one side of the link
    corresponds to a pure "interface down" (downlink) on the other side.
    """
    if 'Mismatches' not in wb.sheetnames or 'Downlinks' not in wb.sheetnames:
        return

    ws_m = wb['Mismatches']
    ws_d = wb['Downlinks']

    m_header = [ws_m.cell(row=1, column=c).value for c in range(1, ws_m.max_column + 1)]
    d_header = [ws_d.cell(row=1, column=c).value for c in range(1, ws_d.max_column + 1)]

    dl_host_col = next((i+1 for i, h in enumerate(d_header)
                        if h in ('Act. Hostname', 'Hostname')), None)
    dl_int_col  = next((i+1 for i, h in enumerate(d_header)
                        if h in ('Act. Interface', 'Interface')), None)
    if dl_host_col is None or dl_int_col is None:
        log("  ⚠ Downlinks tab missing Hostname/Interface — skipping Mismatch↔Downlink tab")
        return

    dl_pairs = set()
    for r in range(2, ws_d.max_row + 1):
        h = str(ws_d.cell(row=r, column=dl_host_col).value or '').strip()
        i = str(ws_d.cell(row=r, column=dl_int_col).value  or '').strip()
        if h:
            dl_pairs.add((h, i))

    exp_h_col = next((i+1 for i, h in enumerate(m_header) if h == 'Expected Hostname'), None)
    exp_i_col = next((i+1 for i, h in enumerate(m_header) if h == 'Exp. Interface'), None)
    if exp_h_col is None or exp_i_col is None:
        log("  ⚠ Mismatches tab missing Expected Hostname/Exp. Interface — skipping Mismatch↔Downlink tab")
        return

    matched_rows = []
    for r in range(2, ws_m.max_row + 1):
        eh = str(ws_m.cell(row=r, column=exp_h_col).value or '').strip()
        ei = str(ws_m.cell(row=r, column=exp_i_col).value or '').strip()
        if (eh, ei) in dl_pairs:
            matched_rows.append([ws_m.cell(row=r, column=c).value
                                  for c in range(1, ws_m.max_column + 1)])

    log(f"  · Mismatch↔Downlink: {len(matched_rows)} row(s) matched")
    if not matched_rows:
        return

    if 'Mismatch\u2194Downlink' in wb.sheetnames:
        del wb['Mismatch\u2194Downlink']
    ws_x = wb.create_sheet('Mismatch\u2194Downlink')

    bd          = thin_border()
    yellow_fill = PatternFill('solid', start_color=YELLOW)
    pink_fill   = PatternFill('solid', start_color=PINK)
    pink_names  = {'Possible Hostname', 'Possible Interface', 'Possible L/R',
                   'Possible Rack', 'Possible Elevation', 'Possible Source_port',
                   'Possible DMARC1', 'Possible DMARC2', 'Possible Destination_port',
                   'Z Hostname', 'Z Interface', 'Z L/R', 'Z Rack', 'Z Elevation'}
    pink_idx    = {m_header.index(nm) + 1 for nm in pink_names if nm in m_header}

    for c, h in enumerate(m_header, 1):
        cell = ws_x.cell(row=1, column=c, value=h)
        cell.font      = Font(bold=True, name='Arial', size=10)
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.fill      = yellow_fill
        cell.border    = bd

    for r_off, rv in enumerate(matched_rows):
        r = r_off + 2
        for c, val in enumerate(rv, 1):
            cell = ws_x.cell(row=r, column=c, value=val)
            cell.font      = Font(name='Arial', size=10)
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.fill      = pink_fill if c in pink_idx else PatternFill(fill_type=None)
            cell.border    = bd

    for c, h in enumerate(m_header, 1):
        vals = [str(h or '')] + [str(rv[c-1] or '') for rv in matched_rows]
        ws_x.column_dimensions[get_column_letter(c)].width = min(max(len(v) for v in vals) + 4, 80)

    ws_x.freeze_panes = 'A2'
