"""
HSG17 T0-to-Host clean processor (6-stage architecture).

Stages (as designed):
1. Ingest
2. Normalize (block derivation is here)
3. Enrich (PP join + extra columns)
4. Analyze (light clustering / grouping for rich output)
5. Format (professional 5-tab workbook + Summary)
6. Log (extract counts per DH block → central logger)

This module is intentionally self-contained and testable.
"""

from __future__ import annotations

import io
import re
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path
from typing import Tuple, Optional, Dict, List, Any

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter

from utils.hsg17_models import normalize_block, PG_TO_DH, RACK_TO_BLOCK


# ================== Styles (matching the "pretty" vibe the user likes) ==================
HEADER_FILL = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
HEADER_FONT = Font(name="Arial", bold=True, color="FFFFFF", size=11)
TITLE_FONT = Font(name="Arial", bold=True, size=14, color="1F4E79")
BODY_FONT = Font(name="Arial", size=10)
THIN_BORDER = Border(
    left=Side(style="thin", color="B4B4B4"),
    right=Side(style="thin", color="B4B4B4"),
    top=Side(style="thin", color="B4B4B4"),
    bottom=Side(style="thin", color="B4B4B4"),
)
ORANGE_FILL = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")
YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
LIGHT_BLUE_FILL = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
LIGHT_GREEN_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")


ERROR_SHEETS = [
    "LLDP Mismatch + Link Down",
    "Optic Errors",
    "FEC_BER Errors",
    "Interface Down Errors",
]


def _safe_str(x: Any) -> str:
    return "" if pd.isna(x) else str(x).strip()


def _derive_block_from_pp(pp_text: str) -> Optional[str]:
    """Priority #1 rule from DESIGN: parse DH from Patch Panel Matrix / Full Label."""
    if not pp_text:
        return None
    # Look for DH2, DH-4, DH.7, PP.HSG17.4.DH4. etc.
    m = re.search(r"(?:PP\.[^.]+\.)?DH[\.-]?(\d+)", str(pp_text), re.IGNORECASE)
    if m:
        num = m.group(1)
        if num == "102":
            return "DH-102 (Spines)"
        if len(num) == 1:
            num = f"00{num}"
        elif len(num) == 2 and num[0] != "1":
            num = f"0{num}"
        candidate = f"DH-{num}"
        return candidate
    return None


def _derive_block_from_rack(rack: str) -> Optional[str]:
    """Secondary rule using rack number ranges (handles hsg17:5708:3 and bare numbers)."""
    if not rack:
        return None
    r = str(rack).strip()
    m = re.search(r"(\d{3,4})", r)
    if m:
        rack_num = m.group(1)
        prefix2 = rack_num[:2]
        if prefix2 in RACK_TO_BLOCK:
            return RACK_TO_BLOCK[prefix2]
        # 08xx / 09xx small racks etc. can be extended here later
    return None


def derive_block(row: pd.Series, pp_col_candidates: List[str], rack_col_candidates: List[str]) -> Optional[str]:
    """
    Full priority strategy from DESIGN.md:
    1. Patch Panel DH (strongest signal)
    2. hsg17: rack / numeric rack range
    3. Device name patterns (t1 → spines, etc.)
    4. Fallback None
    """
    # 1. PP columns (Patch Panel Matrix, Full Label, EasyMark...)
    for col in pp_col_candidates:
        if col in row and pd.notna(row[col]):
            blk = _derive_block_from_pp(_safe_str(row[col]))
            if blk:
                return blk

    # 2. Rack columns
    for col in rack_col_candidates:
        if col in row and pd.notna(row[col]):
            blk = _derive_block_from_rack(_safe_str(row[col]))
            if blk:
                return blk

    # 3. Device name heuristic (T1 spines + IPR)
    for col in ["Device A Name", "Device Name", "Source Device Name", "device_a", "device_b", "Device B Name"]:
        if col in row and pd.notna(row[col]):
            name = _safe_str(row[col]).lower()
            if "t1" in name or "-t1-" in name or "t1-r" in name:
                return "DH-102 (Spines)"
            if "ip-r" in name or "q2-ip" in name:
                return "IPR (PG-201)"

    return None


# ================== Stage 1: Ingest ==================
def ingest(allconnections_bytes: bytes, cutsheet_bytes: bytes) -> Dict[str, pd.DataFrame]:
    """Load both inputs. Returns dict of useful frames."""
    # All connections (large — we only keep columns we need for enrichment)
    allc = pd.read_excel(io.BytesIO(allconnections_bytes), sheet_name="Sheet1")

    # Keep the columns that matter for PP enrichment and joining
    keep_allc = [c for c in allc.columns if any(k in str(c).lower() for k in
                 ["devicea", "deviceb", "rack", "port", "full label", "easymark", "patch", "source_port", "destination_port", "t0 switch"])]
    allc = allc[[c for c in keep_allc if c in allc.columns]].copy()

    # Cutsheet — all 4 error sheets + Summary
    cuts = {}
    with io.BytesIO(cutsheet_bytes) as bio:
        xl = pd.ExcelFile(bio)
        for sheet in ERROR_SHEETS:
            if sheet in xl.sheet_names:
                cuts[sheet] = pd.read_excel(xl, sheet_name=sheet)
        if "Summary" in xl.sheet_names:
            cuts["Summary"] = pd.read_excel(xl, sheet_name="Summary")

    return {"allconnections": allc, "cutsheet_sheets": cuts}


# ================== Stage 2: Normalize (block derivation lives here) ==================
def normalize(ingested: Dict[str, Any]) -> Dict[str, pd.DataFrame]:
    """Add canonical 'Block' column to every error sheet. Return enriched frames."""
    cuts = ingested["cutsheet_sheets"]
    allc = ingested["allconnections"]

    pp_candidates = ["Patch Panel Matrix", "Full Label", "EasyMark+ --- Patch Panels", "PP Matrix"]
    rack_candidates = ["Device A Rack", "Device B Rack", "Rack", "device_a_rack", "hsg17:"]

    normalized = {}

    for sheet_name, df in cuts.items():
        if sheet_name == "Summary":
            normalized[sheet_name] = df
            continue

        df = df.copy()
        df["Block"] = df.apply(lambda r: derive_block(r, pp_candidates, rack_candidates), axis=1)

        # Fallback: try a few more obvious columns that often contain DH in this data
        mask = df["Block"].isna()
        if mask.any():
            for col in ["Patch Panel Matrix", "Source Device Location", "Device A Rack"]:
                if col in df.columns:
                    df.loc[mask, "Block"] = df.loc[mask, col].apply(_derive_block_from_pp)
                    mask = df["Block"].isna()
                    if not mask.any():
                        break

        # Final safety: mark unknowns (they will still be logged under a sensible bucket later if needed)
        df["Block"] = df["Block"].fillna("UNKNOWN")

        normalized[sheet_name] = df

    normalized["allconnections"] = allc
    return normalized


# ================== Union-Find for Connected Components ==================
class UnionFind:
    def __init__(self):
        self.parent: Dict[str, str] = {}

    def find(self, x: str) -> str:
        if x not in self.parent:
            self.parent[x] = x
        if self.parent[x] != x:
            self.parent[x] = self.find(self.parent[x])
        return self.parent[x]

    def union(self, x: str, y: str):
        px, py = self.find(x), self.find(y)
        if px != py:
            self.parent[px] = py


def cluster_mismatches(df: pd.DataFrame) -> pd.DataFrame:
    """
    Rich connected-component clustering for LLDP Mismatch + Link Down.
    Nodes are 'device:port'. We union actual + expected mismatches so
    swapped pairs and multi-port issues form visible clusters.
    Returns df with new 'Cluster' column (integer group id).
    """
    if df.empty:
        df = df.copy()
        df["Cluster"] = -1
        return df

    uf = UnionFind()

    for _, row in df.iterrows():
        a = f"{_safe_str(row.get('Device A Name'))}:{_safe_str(row.get('Device A Port'))}"
        b = f"{_safe_str(row.get('Device B Name'))}:{_safe_str(row.get('Device B Port'))}"
        exp_b = f"{_safe_str(row.get('Expected Device B Name'))}:{_safe_str(row.get('Expected Device B Port'))}"

        uf.union(a, b)
        if exp_b and exp_b != b:
            uf.union(a, exp_b)

    # Assign dense cluster ids
    component_to_id: Dict[str, int] = {}
    next_id = 0
    cluster_col = []
    for _, row in df.iterrows():
        a = f"{_safe_str(row.get('Device A Name'))}:{_safe_str(row.get('Device A Port'))}"
        root = uf.find(a)
        if root not in component_to_id:
            component_to_id[root] = next_id
            next_id += 1
        cluster_col.append(component_to_id[root])

    df = df.copy()
    df["Cluster"] = cluster_col
    return df


# ================== Stage 3+4: Enrich + Analyze ==================
def enrich_and_analyze(normalized: Dict[str, pd.DataFrame]) -> Dict[str, pd.DataFrame]:
    """Enrichment + real mismatch clustering + PP joins."""
    out = {}
    allc = normalized.get("allconnections")

    # Build fast lookup from allconnections for PP enrichment
    pp_lookup: Dict[tuple, Dict] = {}
    if allc is not None and not allc.empty:
        # Try common column names present in the user's allconnections file
        dev_a = next((c for c in allc.columns if "DeviceA Name" in str(c) or c == "DeviceA Name"), None)
        port_a = next((c for c in allc.columns if "DeviceA Port" in str(c) or c == "DeviceA Port"), None)
        full_label_col = next((c for c in allc.columns if "Full Label" in str(c)), None)
        easymark_col = next((c for c in allc.columns if "EasyMark" in str(c)), None)

        if dev_a and port_a:
            for _, r in allc.iterrows():
                key = (_safe_str(r.get(dev_a)), _safe_str(r.get(port_a)))
                pp_lookup[key] = {
                    "Full_Label": _safe_str(r.get(full_label_col)) if full_label_col else "",
                    "EasyMark": _safe_str(r.get(easymark_col)) if easymark_col else "",
                }

    for name, df in normalized.items():
        if name in ("allconnections", "Summary"):
            out[name] = df
            continue

        df = df.copy()

        # === PP Enrichment join (from allconnections) ===
        if pp_lookup and not df.empty:
            dev_col = next((c for c in ["Device A Name", "Source Device Name", "Device Name"] if c in df.columns), None)
            port_col = next((c for c in ["Device A Port", "Source Device Port", "Device A Port"] if c in df.columns), None)

            if dev_col and port_col:
                def _lookup_pp(r):
                    key = (_safe_str(r.get(dev_col)), _safe_str(r.get(port_col)))
                    info = pp_lookup.get(key, {})
                    return info.get("Full_Label", "") or info.get("EasyMark", "")

                df["PP_Enriched"] = df.apply(_lookup_pp, axis=1)

        # Promote Block (and Cluster if present) to front
        front = ["Block"]
        if "Cluster" in df.columns:
            front.append("Cluster")
        if "PP_Enriched" in df.columns:
            front.append("PP_Enriched")
        other = [c for c in df.columns if c not in front]
        df = df[front + other]

        # === Real connected-component clustering on LLDP sheet ===
        if name == "LLDP Mismatch + Link Down":
            df = cluster_mismatches(df)

        out[name] = df

    out["allconnections"] = allc
    return out


# ================== Stage 5: Format ==================
def build_workbook(enriched: Dict[str, pd.DataFrame], source_basename: str) -> Tuple[bytes, str]:
    """Create the professional 5-tab Excel the user expects."""
    wb = Workbook()

    # Remove default sheet
    default = wb.active
    wb.remove(default)

    # --- Summary (counts per Block per category) ---
    summary_ws = wb.create_sheet("Summary", 0)
    summary_ws["A1"] = f"HSG17 T0-to-Host Validation Summary — {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    summary_ws["A1"].font = TITLE_FONT
    summary_ws.merge_cells("A1:F1")

    # Build summary data from the 4 error sheets
    summary_rows: List[List[Any]] = [["Block", "LLDP Mismatch + Link Down", "Optic Errors", "FEC_BER Errors", "Interface Down Errors", "Total"]]

    block_counts: Dict[str, Counter] = defaultdict(Counter)

    for sheet_name in ERROR_SHEETS:
        if sheet_name in enriched:
            df = enriched[sheet_name]
            if "Block" in df.columns:
                for blk, cnt in df["Block"].value_counts().items():
                    block_counts[blk][sheet_name] = cnt

    # Sort blocks nicely (DH-001 first, then spines, then specials)
    def block_sort_key(b: str) -> tuple:
        if b.startswith("DH-00"):
            return (0, int(b.split("-")[1]))
        if b.startswith("DH-10"):
            return (1, int(b.split("-")[1][:3]))
        if "Spines" in b:
            return (2, 102)
        if "IPR" in b:
            return (3, 0)
        if "AUX" in b:
            return (4, 0)
        return (5, 0)

    all_blocks = sorted(block_counts.keys(), key=block_sort_key)

    for blk in all_blocks:
        c = block_counts[blk]
        row = [
            blk,
            c.get("LLDP Mismatch + Link Down", 0),
            c.get("Optic Errors", 0),
            c.get("FEC_BER Errors", 0),
            c.get("Interface Down Errors", 0),
        ]
        row.append(sum(row[1:]))
        summary_rows.append(row)

    # Total row
    totals = [sum(r[i] for r in summary_rows[1:]) for i in range(1, 6)]
    summary_rows.append(["TOTAL"] + totals)

    for r_idx, row in enumerate(summary_rows, start=3):
        for c_idx, val in enumerate(row, start=1):
            cell = summary_ws.cell(row=r_idx, column=c_idx, value=val)
            cell.font = BODY_FONT
            cell.border = THIN_BORDER
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if r_idx == 3:  # header
                cell.fill = HEADER_FILL
                cell.font = HEADER_FONT
            if r_idx == len(summary_rows):  # total
                cell.font = Font(name="Arial", bold=True)

    # Column widths
    for col in range(1, 7):
        summary_ws.column_dimensions[get_column_letter(col)].width = 28 if col == 1 else 18

    # --- The 4 detailed error sheets (with Block + Cluster promoted) ---
    for sheet_name in ERROR_SHEETS:
        if sheet_name not in enriched:
            continue
        df = enriched[sheet_name].copy()

        # Drop any remaining internal columns
        for drop in ["_pair_group", "raw_row"]:
            if drop in df.columns:
                df = df.drop(columns=[drop], errors="ignore")

        # Reorder: Block, Cluster (if present), then rest
        front = ["Block"]
        if "Cluster" in df.columns:
            front.append("Cluster")
        other = [c for c in df.columns if c not in front]
        df = df[front + other]

        ws = wb.create_sheet(sheet_name)

        # Title
        ws["A1"] = f"HSG17 — {sheet_name}"
        ws["A1"].font = TITLE_FONT
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=min(10, len(df.columns)))

        is_lldp = sheet_name == "LLDP Mismatch + Link Down"

        # Write data with rich coloring for LLDP clusters
        for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), start=3):
            cluster_val = None
            if is_lldp and "Cluster" in df.columns:
                # Find cluster value in the written row
                try:
                    cluster_idx_in_df = list(df.columns).index("Cluster")
                    cluster_val = row[cluster_idx_in_df] if cluster_idx_in_df < len(row) else None
                except:
                    cluster_val = None

            for c_idx, val in enumerate(row, start=1):
                cell = ws.cell(row=r_idx, column=c_idx, value=val if val is not None else "")
                cell.font = BODY_FONT
                cell.border = THIN_BORDER

                if r_idx == 3:  # header
                    cell.fill = HEADER_FILL
                    cell.font = HEADER_FONT
                else:
                    # Highlight Block (always)
                    if df.columns[c_idx-1] == "Block":
                        cell.fill = LIGHT_BLUE_FILL
                        cell.font = Font(name="Arial", bold=True, size=10)

                    # Rich orange/yellow cluster highlighting on LLDP sheet only
                    # (Block column keeps its blue; other cells get cluster color for strong pair visibility)
                    if is_lldp and cluster_val is not None and df.columns[c_idx-1] != "Block":
                        try:
                            cid = int(cluster_val)
                            if cid % 2 == 0:
                                cell.fill = ORANGE_FILL
                            else:
                                cell.fill = YELLOW_FILL
                            cell.font = Font(name="Arial", size=10, bold=True)
                        except (ValueError, TypeError):
                            pass

        # Auto width
        for c_idx, col in enumerate(df.columns, start=1):
            width = min(50, max(10, len(str(col)) + 3))
            ws.column_dimensions[get_column_letter(c_idx)].width = width

    # Filename
    ts = datetime.now().strftime("%Y%m%d_%H%M")
    filename = f"HSG17_T0_Host_{source_basename}_{ts}.xlsx"

    # Write to bytes
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue(), filename


# ================== Public entry point ==================
def process_hsg17_t0_host(
    allconnections_bytes: bytes,
    cutsheet_bytes: bytes,
    source_filename: str = "HSG17",
) -> Tuple[bytes, str]:
    """
    Full clean pipeline. Returns (excel_bytes, suggested_filename).
    This is what the Streamlit page will call.
    """
    base = Path(source_filename).stem.replace(" ", "_")[:40]

    # 1. Ingest
    ingested = ingest(allconnections_bytes, cutsheet_bytes)

    # 2. Normalize (block derivation)
    normalized = normalize(ingested)

    # 3+4. Enrich + Analyze (light v1)
    enriched = enrich_and_analyze(normalized)

    # 5. Format
    xlsx_bytes, filename = build_workbook(enriched, base)

    return xlsx_bytes, filename


# Convenience for the page: extract the counts we will log (after we have the workbook)
def extract_counts_for_logging(wb_bytes: bytes) -> List[Dict[str, Any]]:
    """Read the Summary tab we just created and return list of log rows."""
    from openpyxl import load_workbook

    wb = load_workbook(io.BytesIO(wb_bytes))
    if "Summary" not in wb.sheetnames:
        return []

    ws = wb["Summary"]
    rows = list(ws.iter_rows(min_row=4, values_only=True))  # skip title + header + blank

    results = []
    for row in rows:
        if not row or not row[0]:
            continue
        blk = str(row[0]).strip()
        if blk.upper() == "TOTAL":
            continue
        results.append({
            "block": blk,
            "LLDP Mismatch + Link Down": row[1] or 0,
            "Optic Errors": row[2] or 0,
            "FEC_BER Errors": row[3] or 0,
            "Interface Down Errors": row[4] or 0,
        })
    return results
