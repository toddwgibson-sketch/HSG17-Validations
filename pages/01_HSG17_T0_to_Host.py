#!/usr/bin/env python3

import streamlit as st
import tempfile
from pathlib import Path

from utils.auth import require_login
from utils.data_logger import log_errors
from utils.t1_to_t0_formatter import format_report
from utils.hsg17_models import derive_placement_group

require_login()

st.set_page_config(page_title="HSG17 T0-to-Host", page_icon="🖥️", layout="wide")

with st.sidebar:
    lv_file = st.file_uploader(
        "",
        type=["xlsx", "xlsm"],
        accept_multiple_files=False,
        label_visibility="collapsed"
    )
    cutsheet_files = st.file_uploader(
        "",
        type=["xlsx", "xlsm"],
        accept_multiple_files=True,
        label_visibility="collapsed"
    )
    run_btn = st.button("", type="primary", disabled=not (lv_file and cutsheet_files))

if run_btn and lv_file and cutsheet_files:
    with st.spinner(""):
        tmpdir = Path(tempfile.mkdtemp(prefix="hsg17_t1t0_"))
        try:
            # Write uploads to temp files (the formatter expects real paths + load_workbook)
            lv_tmp = tmpdir / lv_file.name
            lv_tmp.write_bytes(lv_file.getvalue())

            cuts_tmp_paths = []
            for f in cutsheet_files:
                p = tmpdir / f.name
                p.write_bytes(f.getvalue())
                cuts_tmp_paths.append(str(p))

            # Call the exact gold implementation (interactive=False to skip any GUI)
            out_path, counts = format_report(str(lv_tmp), cuts_tmp_paths, interactive=False)

            # Read result for download
            out_bytes = out_path.read_bytes()

            st.download_button(
                label="",
                data=out_bytes,
                file_name=out_path.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

            try:
                source_name = lv_file.name if lv_file else "unknown"
                placement = "PG14"
                try:
                    import pandas as pd
                    import re
                    from collections import Counter
                    from utils.hsg17_models import derive_placement_group
                    if 'lv_tmp' in locals() and lv_tmp.exists():
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
                except Exception as e:
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
                            error_category=cat_name,
                            count=int(cnt),
                            source_file=source_name,
                            processed_by="HSG17_T1toT0_Gold"
                        )
            except Exception as log_exc:
                pass

        finally:
            try:
                for p in tmpdir.glob("*"):
                    p.unlink(missing_ok=True)
                tmpdir.rmdir()
            except Exception:
                pass

elif not (lv_file and cutsheet_files):
    pass