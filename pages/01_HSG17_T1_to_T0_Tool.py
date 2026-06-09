#!/usr/bin/env python3
"""
HSG17 T1-to-T0 Formatter — Streamlit
(UI polished to match the previous formatter. Core logic unchanged.)
"""

import sys
from pathlib import Path

# Ensure project root is on sys.path
_here = Path(__file__).resolve()
_root = _here.parent.parent if _here.parent.name == "pages" else _here.parent
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

import streamlit as st
import tempfile
from pathlib import Path

from utils.data_logger import log_errors
from utils.t1_to_t0_formatter import format_report
from utils.hsg17_models import derive_placement_group

st.set_page_config(page_title="HSG17 T1-to-T0 Formatter", page_icon="🖥️", layout="wide")

st.title("🖥️ HSG17 T1-to-T0 Validation Formatter")
st.caption("")

st.markdown("""
**How to use:**
1. Upload your **LV Portal Validation Export** (.xlsx / .xlsm)
2. Upload the corresponding **Master Cutsheet / Allconnections** file(s)
3. Click **Generate Formatted Report**

The formatted report will be available for immediate download.
""")

# ── Uploaders ────────────────────────────────────────────────────────────────
col1, col2 = st.columns(2)

with col1:
    lv_file = st.file_uploader(
        "LV Portal Validation Export (.xlsx / .xlsm)",
        type=["xlsx", "xlsm"],
        accept_multiple_files=False,
        help="The export containing the error sheets (Optics, FEC, Interface Down, ...)"
    )

with col2:
    cutsheet_files = st.file_uploader(
        "Master Cutsheet(s) / Allconnections",
        type=["xlsx", "xlsm"],
        accept_multiple_files=True,
        help="One or more cutsheet files for enrichment"
    )

run_btn = st.button(
    "🚀 Generate Formatted Report",
    type="primary",
    use_container_width=True,
    disabled=not (lv_file and cutsheet_files)
)

# ── Processing ───────────────────────────────────────────────────────────────
if run_btn and lv_file and cutsheet_files:
    with st.spinner("Processing report with original logic..."):
        tmpdir = Path(tempfile.mkdtemp(prefix="hsg17_t1t0_"))
        try:
            # Save uploads to temp files (core logic expects real paths)
            lv_tmp = tmpdir / lv_file.name
            lv_tmp.write_bytes(lv_file.getvalue())

            cuts_tmp_paths = []
            for f in cutsheet_files:
                p = tmpdir / f.name
                p.write_bytes(f.getvalue())
                cuts_tmp_paths.append(str(p))

            # Call the exact original implementation (interactive=False)
            out_path, counts = format_report(str(lv_tmp), cuts_tmp_paths, interactive=False)

            # Read result for download
            out_bytes = out_path.read_bytes()

            st.success(f"✅ Report ready: {out_path.name}")

            st.download_button(
                label=f"📥 Download {out_path.name}",
                data=out_bytes,
                file_name=out_path.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

            # ====================== SILENT CENTRAL LOGGING (unchanged) ======================
            try:
                source_name = lv_file.name
                placement = "PG14"
                rack = "3110"
                try:
                    import pandas as pd
                    import re
                    from collections import Counter
                    if lv_tmp.exists():
                        lv_df_dict = pd.read_excel(lv_tmp, sheet_name=None)
                        rack_nums = []
                        for sheet_name, sheet_df in lv_df_dict.items():
                            for col in sheet_df.columns:
                                col_l = str(col).lower()
                                if 'device' in col_l or 'source' in col_l:
                                    for val in sheet_df[col].dropna().astype(str):
                                        m = re.search(r'r(\d{3,4})', val.lower())
                                        if m:
                                            rack_nums.append(m.group(1).zfill(4))
                                        else:
                                            m2 = re.search(r'\b(\d{4})\b', val)
                                            if m2 and 1000 < int(m2.group(1)) < 9999:
                                                rack_nums.append(m2.group(1))
                        if rack_nums:
                            pgs = [derive_placement_group(r) for r in rack_nums]
                            most_common = Counter(pgs).most_common(1)[0][0]
                            if most_common and most_common.startswith('PG'):
                                placement = most_common
                            most_common_rack = Counter(rack_nums).most_common(1)[0][0]
                            if most_common_rack:
                                rack = most_common_rack
                except Exception:
                    pass

                for cat_key, cnt in counts.items():
                    if cnt > 0:
                        cat_map = {
                            "mispatches": "LLDP Mismatch + Link Down",
                            "downlinks": "Interface Down Errors",
                            "optics": "Optic Errors",
                            "fec": "FEC_BER Errors"
                        }
                        cat_name = cat_map.get(cat_key, cat_key.title())
                        log_errors(
                            hall="HSG17",
                            rack_type="T1-T0",
                            building=placement,
                            rack=rack,
                            error_category=cat_name,
                            count=int(cnt),
                            source_file=source_name,
                            processed_by="HSG17_T1toT0_Gold"
                        )
            except Exception:
                pass

        finally:
            # Cleanup
            try:
                for p in tmpdir.glob("*"):
                    p.unlink(missing_ok=True)
                tmpdir.rmdir()
            except Exception:
                pass
