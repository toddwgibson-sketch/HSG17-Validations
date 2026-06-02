#!/usr/bin/env python3
"""
HSG17 T0-to-Host Formatting Tool

"""

import streamlit as st
import tempfile
from pathlib import Path

from utils.auth import require_login
from utils.data_logger import log_errors
from utils.t1_to_t0_formatter import format_report

require_login()

st.set_page_config(page_title="HSG17 T0-to-Host", page_icon="🖥️", layout="wide")
st.title("HSG17 T0-to-Host Validator")
st.caption("")

st.markdown("""
**Inputs:**
- **LV Portal Validation Export**
- **Master Cutsheet(s) / Allconnections**: your T1toT0 allconnect (or master cutsheet).


All processing is local.
""")

with st.sidebar:
    st.header("Inputs")
    lv_file = st.file_uploader(
        "LV Portal Validation Export (.xlsx)",
        type=["xlsx", "xlsm"],
        accept_multiple_files=False,
        help="The export containing the error sheets (Optic, FEC, Interface Down, ...)"
    )
    cutsheet_files = st.file_uploader(
        "Master Cutsheet(s) / Allconnections (hold Ctrl or Cmd for multiple)",
        type=["xlsx", "xlsm"],
        accept_multiple_files=True,
        help="The T1toT0 allconnections or master cutsheet for enrichment."
    )
    run_btn = st.button("🚀 Generate Formatted Report", type="primary", disabled=not (lv_file and cutsheet_files))

if run_btn and lv_file and cutsheet_files:
    with st.spinner("Formatting using the exact gold logic..."):
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

            st.success("✅ Report generated with the exact reference logic.")
            st.write("**Counts:**", counts)

            st.download_button(
                label=f"📥 Download {out_path.name}",
                data=out_bytes,
                file_name=out_path.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

            st.info("")

            # ====================== SILENT CENTRAL LOGGING (to retain Dashboard features) ======================
            # Log the counts so the HSG17 Dashboard shows current state + deltas.
            # Using placeholder building since this flow uses racks from allc rather than DH blocks.
            try:
                source_name = lv_file.name if lv_file else "unknown"
                for cat_key, cnt in counts.items():
                    if cnt > 0:
                        # Map keys to nice category names
                        cat_map = {
                            "mispatches": "Mispatches",
                            "downlinks": "Downlinks",
                            "optics": "Optics",
                            "fec": "FEC Errors"
                        }
                        cat_name = cat_map.get(cat_key, cat_key.title())
                        log_errors(
                            hall="HSG17",
                            rack_type="T1-T0",
                            building="T1toT0",  # placeholder; Dashboard will group under this
                            error_category=cat_name,
                            count=int(cnt),
                            source_file=source_name,
                            processed_by="HSG17_T1toT0_Gold"
                        )
            except Exception as log_exc:
                st.warning(f"Central logging encountered an issue (non-fatal): {log_exc}")

        finally:
            # Best effort cleanup of temps
            try:
                for p in tmpdir.glob("*"):
                    p.unlink(missing_ok=True)
                tmpdir.rmdir()
            except Exception:
                pass

elif not (lv_file and cutsheet_files):
    st.info("Upload the LV Portal export and at least one cutsheet / allconnections file, then click Generate.")

st.markdown("---")
st.caption("")
