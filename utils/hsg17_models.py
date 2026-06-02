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
