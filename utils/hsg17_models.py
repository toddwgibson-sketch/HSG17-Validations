"""
HSG17 clean data models.

Lightweight (no heavy Pydantic dependency for the hot path on 30k+ rows).
Use these for clarity in the stage functions; the heavy lifting stays in pandas.
"""

from dataclasses import dataclass, field
from typing import Optional, List, Dict, Any


@dataclass
class HSG17Connection:
    """Normalized connection from the allconnections file."""
    source_device: str
    source_port: str
    dest_device: str
    dest_port: str
    t0_switch_port: Optional[str] = None
    block: Optional[str] = None          # derived DH-xxx or special
    rack_a: Optional[str] = None
    rack_b: Optional[str] = None
    patch_panel_path: Optional[str] = None
    full_label: Optional[str] = None
    raw_row: Dict[str, Any] = field(default_factory=dict)


@dataclass
class HSG17Error:
    """One categorized error row (from the pre-classified cutsheet + enrichment)."""
    error_type: str                      # "LLDP Mismatch + Link Down", "Optic Errors", "FEC_BER Errors", "Interface Down Errors"
    device_a: str
    port_a: str
    device_b: Optional[str] = None
    port_b: Optional[str] = None
    block: Optional[str] = None          # the canonical DH-xxx we log against
    details: Dict[str, Any] = field(default_factory=dict)
    patch_panel_matrix: Optional[str] = None
    rack_a: Optional[str] = None
    rack_b: Optional[str] = None
    source_sheet: Optional[str] = None
    raw_row: Dict[str, Any] = field(default_factory=dict)


# Canonical DH blocks (from Bootstrap Sequence authoritative source)
CANONICAL_BLOCKS: List[str] = [
    "DH-001", "DH-002", "DH-003", "DH-004", "DH-005",
    "DH-101", "DH-102 (Spines)", "DH-103", "DH-104", "DH-105",
    "IPR (PG-201)", "AUX Racks",
]

# Known rack ranges → block (from the Confluence table)
RACK_TO_BLOCK: Dict[str, str] = {
    # T1 Spines (PG 151-154)
    "57": "DH-102 (Spines)",
    "58": "DH-102 (Spines)",
    "59": "DH-102 (Spines)",
    "60": "DH-102 (Spines)",
    # 08xx/09xx racks in PG 3-6 (from Bootstrap Sequence)
    "08": "DH-002",
    "09": "DH-002",
    # Additional common prefixes observed in HSG17 data (fall back to PP when possible;
    # extend here for pure rack-based derivation without PP labels)
    "25": "DH-101",
    "28": "DH-103",
    "29": "DH-103",
    "30": "DH-104",
}

# Placement Group to DH (condensed authoritative mapping)
PG_TO_DH: Dict[int, str] = {
    1: "DH-001", 2: "DH-001",
    3: "DH-002", 4: "DH-002", 5: "DH-002", 6: "DH-002",
    7: "DH-003", 8: "DH-003", 9: "DH-003", 10: "DH-003",
    11: "DH-004", 12: "DH-004", 13: "DH-004", 14: "DH-004",
    15: "DH-005", 16: "DH-005", 17: "DH-005", 18: "DH-005",
    19: "DH-101", 20: "DH-101", 21: "DH-101", 22: "DH-101",
    23: "DH-103", 24: "DH-103", 25: "DH-103", 26: "DH-103",
    27: "DH-104", 28: "DH-104", 29: "DH-104", 30: "DH-104",
    31: "DH-105", 32: "DH-105",
    201: "IPR (PG-201)",
}


def normalize_block(raw: Optional[str]) -> Optional[str]:
    """Normalize various DH representations to canonical form."""
    if not raw:
        return None
    s = str(raw).upper().strip()

    # Direct matches
    if "DH-102" in s or "DH102" in s:
        return "DH-102 (Spines)"
    if s.startswith("DH-"):
        # Already good (DH-001 etc.)
        if s in CANONICAL_BLOCKS:
            return s
        # DH-001 style
        return s

    # DH2 / DH4 / DH7 etc from PP labels
    import re
    m = re.search(r"DH[\.-]?(\d+)", s)
    if m:
        num = m.group(1)
        if num == "102":
            return "DH-102 (Spines)"
        # zero-pad to 3 digits where appropriate
        if len(num) == 1:
            num = "00" + num
        elif len(num) == 2 and not num.startswith("10"):
            num = "0" + num
        candidate = f"DH-{num}"
        if candidate in CANONICAL_BLOCKS:
            return candidate
        return f"DH-{num}"

    if "IPR" in s or "201" in s:
        return "IPR (PG-201)"
    if "AUX" in s:
        return "AUX Racks"

    return None


# Placement Group mapping for HSG17 (from Bootstrap Sequence Confluence)
# Racks (4-digit strings like "3110", "0309") -> PG number as string "14"
# Use derive_placement_group(rack) to get "PG14"
RACK_TO_PLACEMENT_GROUP: Dict[str, str] = {
    "0309": "1",
    "0312": "2",
    "0408": "1",
    "0411": "2",
    "0502": "3",
    "0503": "3",
    "0504": "3",
    "0505": "4",
    "0506": "4",
    "0507": "4",
    "0508": "4",
    "0602": "3",
    "0603": "3",
    "0604": "3",
    "0605": "3",
    "0606": "4",
    "0607": "4",
    "0608": "4",
    "0807": "3",
    "0812": "4",
    "0907": "3",
    "0912": "4",
    "1010": "5",
    "1014": "6",
    "1024": "64",
    "1110": "5",
    "1114": "6",
    "1302": "5",
    "1303": "5",
    "1304": "5",
    "1305": "6",
    "1306": "6",
    "1307": "6",
    "1308": "6",
    "1402": "5",
    "1403": "5",
    "1404": "5",
    "1405": "5",
    "1406": "6",
    "1407": "6",
    "1408": "6",
    "1502": "7",
    "1503": "7",
    "1504": "7",
    "1505": "8",
    "1506": "8",
    "1507": "8",
    "1508": "8",
    "1602": "7",
    "1603": "7",
    "1604": "7",
    "1605": "7",
    "1606": "8",
    "1607": "8",
    "1608": "8",
    "1807": "7",
    "1812": "8",
    "1907": "7",
    "1912": "8",
    "2010": "9",
    "2014": "10",
    "2016": "126",
    "2110": "9",
    "2114": "10",
    "2302": "9",
    "2303": "9",
    "2304": "9",
    "2305": "10",
    "2306": "10",
    "2307": "10",
    "2308": "10",
    "2402": "9",
    "2403": "9",
    "2404": "9",
    "2405": "9",
    "2406": "10",
    "2407": "10",
    "2408": "10",
    "2502": "12",
    "2503": "12",
    "2504": "12",
    "2505": "12",
    "2506": "11",
    "2507": "11",
    "2508": "11",
    "2602": "12",
    "2603": "12",
    "2604": "12",
    "2605": "11",
    "2606": "11",
    "2607": "11",
    "2608": "11",
    "2809": "12",
    "2814": "11",
    "2909": "12",
    "2914": "11",
    "3010": "14",
    "3014": "13",
    "3110": "14",
    "3114": "13",
    "3302": "14",
    "3303": "14",
    "3304": "14",
    "3305": "14",
    "3306": "13",
    "3307": "9",
    "3308": "9",
    "3402": "14",
    "3403": "14",
    "3404": "14",
    "3405": "9",
    "3406": "9",
    "3407": "9",
    "3408": "9",
    "3502": "16",
    "3503": "16",
    "3504": "16",
    "3505": "16",
    "3506": "15",
    "3507": "15",
    "3508": "15",
    "3602": "16",
    "3603": "16",
    "3604": "16",
    "3605": "15",
    "3606": "15",
    "3607": "15",
    "3608": "15",
    "3809": "16",
    "3814": "15",
    "3909": "16",
    "3914": "15",
    "4010": "18",
    "4014": "17",
    "4110": "18",
    "4114": "17",
    "4302": "18",
    "4303": "18",
    "4304": "18",
    "4305": "18",
    "4306": "17",
    "4307": "17",
    "4308": "17",
    "4402": "18",
    "4403": "18",
    "4404": "18",
    "4405": "17",
    "4406": "17",
    "4407": "17",
    "4408": "17",
    "4502": "19",
    "4503": "19",
    "4504": "19",
    "4505": "20",
    "4506": "20",
    "4507": "20",
    "4508": "20",
    "4602": "19",
    "4603": "19",
    "4604": "19",
    "4605": "19",
    "4606": "20",
    "4607": "20",
    "4608": "20",
    "4807": "19",
    "4812": "20",
    "4907": "19",
    "4912": "20",
    "5010": "21",
    "5014": "22",
    "5110": "21",
    "5114": "22",
    "5302": "21",
    "5303": "21",
    "5304": "21",
    "5305": "22",
    "5306": "22",
    "5307": "22",
    "5308": "22",
    "5402": "21",
    "5403": "21",
    "5404": "9",
    "5405": "9",
    "5406": "22",
    "5407": "22",
    "5408": "22",
    "5708": "151",
    "5709": "151",
    "5710": "151",
    "5711": "151",
    "5712": "151",
    "5713": "151",
    "5714": "151",
    "5715": "151",
    "5807": "152",
    "5808": "152",
    "5809": "152",
    "5810": "152",
    "5811": "152",
    "5812": "152",
    "5813": "152",
    "5814": "152",
    "5907": "153",
    "5908": "153",
    "5909": "153",
    "5910": "153",
    "5911": "2",
    "5912": "153",
    "5913": "153",
    "5914": "153",
    "5915": "153",
    "5920": "32",
    "6002": "32",
    "6007": "154",
    "6008": "154",
    "6009": "154",
    "6010": "154",
    "6011": "154",
    "6012": "154",
    "6013": "154",
    "6014": "154",
    "6702": "23",
    "6703": "23",
    "6704": "23",
    "6705": "24",
    "6706": "24",
    "6707": "24",
    "6708": "24",
    "6802": "23",
    "6803": "23",
    "6804": "23",
    "6805": "23",
    "6806": "24",
    "6807": "24",
    "6808": "24",
    "7007": "23",
    "7012": "24",
    "7107": "23",
    "7112": "24",
    "7210": "25",
    "7214": "26",
    "7310": "25",
    "7314": "26",
    "7502": "25",
    "7503": "25",
    "7504": "25",
    "7505": "26",
    "7506": "26",
    "7507": "26",
    "7508": "26",
    "7602": "25",
    "7603": "25",
    "7604": "25",
    "7605": "25",
    "7606": "26",
    "7607": "26",
    "7608": "26",
    "7702": "28",
    "7703": "28",
    "7704": "28",
    "7705": "28",
    "7706": "27",
    "7707": "27",
    "7708": "27",
    "7802": "28",
    "7803": "28",
    "7804": "28",
    "7805": "27",
    "7806": "27",
    "7807": "27",
    "7808": "27",
    "8009": "28",
    "8014": "27",
    "8109": "28",
    "8114": "27",
    "8192": "64",
    "8210": "30",
    "8214": "29",
    "8310": "30",
    "8314": "29",
    "8502": "30",
    "8503": "30",
    "8504": "30",
    "8505": "30",
    "8506": "29",
    "8507": "29",
    "8508": "29",
    "8602": "30",
    "8603": "30",
    "8604": "30",
    "8605": "29",
    "8606": "29",
    "8607": "29",
    "8608": "29",
    "8702": "32",
    "8703": "32",
    "8704": "32",
    "8705": "32",
    "8706": "31",
    "8707": "31",
    "8708": "31",
    "8802": "32",
    "8803": "32",
    "8804": "32",
    "8805": "31",
    "8806": "31",
    "8807": "31",
    "8808": "31",
    "9013": "32",
    "9018": "31",
    "9113": "32",
    "9118": "31",
}


# GPU racks: those with Rack SKU "GPU_GB300_NVL72_R.03" from the authoritative
# "HSG17 - Bootstrap Sequence - Cluster Networking - ALM Confluence.txt"
# (user-provided source). Only these racks appear in the dashboard
# "Errors by Category × GPU Rack" section.
GPU_RACKS: set = {
    "0101", "0102", "0103", "0104", "0105", "0106", "0107", "0201", "0202", "0203", "0204", "0205", "0206", "0207",
    "0309", "0312", "0408", "0411",
    "0502", "0503", "0504", "0505", "0506", "0507", "0508", "0602", "0603", "0604", "0605", "0606", "0607", "0608",
    "0807", "0812", "0907", "0912",
    "1010", "1014", "1110", "1114",
    "1302", "1303", "1304", "1305", "1306", "1307", "1308", "1402", "1403", "1404", "1405", "1406", "1407", "1408",
    "1502", "1503", "1504", "1505", "1506", "1507", "1508", "1602", "1603", "1604", "1605", "1606", "1607", "1608",
    "1807", "1812", "1907", "1912",
    "2010", "2014", "2110", "2114",
    "2302", "2303", "2304", "2305", "2306", "2307", "2308", "2402", "2403", "2404", "2405", "2406", "2407", "2408",
    "2502", "2503", "2504", "2505", "2506", "2507", "2508", "2602", "2603", "2604", "2605", "2606", "2607", "2608",
    "2809", "2814", "2909", "2914",
    "3010", "3014", "3110", "3114",
    "3302", "3303", "3304", "3305", "3306", "3307", "3308", "3402", "3403", "3404", "3405", "3406", "3407", "3408",
    "3502", "3503", "3504", "3505", "3506", "3507", "3508", "3602", "3603", "3604", "3605", "3606", "3607", "3608",
    "3809", "3814", "3909", "3914",
    "4010", "4014", "4110", "4114",
    "4302", "4303", "4304", "4305", "4306", "4307", "4308", "4402", "4403", "4404", "4405", "4406", "4407", "4408",
    "4502", "4503", "4504", "4505", "4506", "4507", "4508", "4602", "4603", "4604", "4605", "4606", "4607", "4608",
    "4807", "4812", "4907", "4912",
    "5010", "5014", "5110", "5114",
    "5302", "5303", "5304", "5305", "5306", "5307", "5308", "5402", "5403", "5404", "5405", "5406", "5407", "5408",
    "6702", "6703", "6704", "6705", "6706", "6707", "6708", "6802", "6803", "6804", "6805", "6806", "6807", "6808",
    "7007", "7012", "7107", "7112",
    "7210", "7214", "7310", "7314",
    "7502", "7503", "7504", "7505", "7506", "7507", "7508", "7602", "7603", "7604", "7605", "7606", "7607", "7608",
    "7702", "7703", "7704", "7705", "7706", "7707", "7708", "7802", "7803", "7804", "7805", "7806", "7807", "7808",
    "8009", "8014", "8109", "8114",
    "8210", "8214", "8310", "8314",
    "8502", "8503", "8504", "8505", "8506", "8507", "8508", "8602", "8603", "8604", "8605", "8606", "8607", "8608",
    "8702", "8703", "8704", "8705", "8706", "8707", "8708", "8802", "8803", "8804", "8805", "8806", "8807", "8808",
    "9013", "9018", "9113", "9118"
}


def is_gpu_rack(rack: Optional[str]) -> bool:
    """Return True if the rack is a GPU rack (SKU GPU_GB300_NVL72_R.03)."""
    if not rack:
        return False
    r = str(rack).strip()
    if r.upper().startswith("RACK "):
        parts = r.split(None, 2)
        if len(parts) > 1:
            r = parts[1]
    r = r.zfill(4)
    if r in GPU_RACKS:
        return True
    r2 = r.lstrip("0") or "0"
    return r2.zfill(4) in GPU_RACKS


def derive_placement_group(rack: Optional[str]) -> str:
    """Given a rack like '3110', 'Rack 3110 U1', '0309' return 'PG14' etc.
    Falls back to the rack itself if unknown.
    """
    if not rack:
        return "Unknown"
    r = str(rack).strip()
    if r.upper().startswith("RACK "):
        parts = r.split(None, 2)
        if len(parts) > 1:
            r = parts[1]
    # normalize to 4 digit string as in the map (0309, 3110)
    r = r.zfill(4)
    pg = RACK_TO_PLACEMENT_GROUP.get(r)
    if pg:
        return f"PG{pg}"
    # try without leading zeros
    r2 = r.lstrip("0") or "0"
    pg = RACK_TO_PLACEMENT_GROUP.get(r2.zfill(4))
    if pg:
        return f"PG{pg}"
    return f"Rack-{r}"


def derive_placement_and_rack_from_files(file_paths: list) -> tuple[str, str]:
    """Scan one or more Excel files for rack numbers (in device, source, host, rack columns etc).
    Returns (placement e.g. 'PG14', representative_rack e.g. '3110') based on most common.
    Reusable by both 01 (LV Portal) and 02 (Slack) tools so the Dashboard sees unified PG tracking.
    Falls back to PG14 / 3110 if nothing useful is found.
    """
    import pandas as pd
    import re
    from collections import Counter
    from pathlib import Path

    rack_nums = []
    for p in file_paths or []:
        try:
            p = str(p)
            if not Path(p).exists():
                continue
            xl = pd.ExcelFile(p)
            for sheet_name in xl.sheet_names:
                try:
                    df = pd.read_excel(p, sheet_name=sheet_name)
                    for col in df.columns:
                        col_l = str(col).lower()
                        if any(k in col_l for k in ("device", "source", "rack", "host", "remote", "hostname")):
                            for val in df[col].dropna().astype(str):
                                m = re.search(r"r(\d{3,4})", val.lower())
                                if m:
                                    rack_nums.append(m.group(1).zfill(4))
                                else:
                                    m2 = re.search(r"\b(\d{4})\b", val)
                                    if m2 and 1000 < int(m2.group(1)) < 9999:
                                        rack_nums.append(m2.group(1))
                except Exception:
                    continue
        except Exception:
            continue

    if not rack_nums:
        return "PG14", "3110"

    pgs = [derive_placement_group(r) for r in rack_nums]
    valid_pgs = [pg for pg in pgs if isinstance(pg, str) and pg.startswith("PG")]
    most_pg = Counter(valid_pgs).most_common(1)
    placement = most_pg[0][0] if most_pg else "PG14"

    most_rack = Counter(rack_nums).most_common(1)[0][0]
    return placement, most_rack
