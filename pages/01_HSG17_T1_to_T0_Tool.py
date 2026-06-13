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
from utils.hsg17_models import derive_placement_and_rack_from_files

st.set_page_config(page_title="HSG17 T1-to-T0 Formatter", page_icon="🖥️", layout="wide")

st.title("🖥️ HSG17 T1-to-T0 LVV Portal Upload")
st.caption("")

st.markdown("""
**How to use:**
1. Upload your **Master Cutsheet(s) / Allconnections**
2. Upload the **LV Portal Validation Export** (.xlsx / .xlsm)
3. Click **Generate Formatted Report**

The formatted report will be available for immediate download.
""")

# ── Uploaders (stacked vertically, cutsheet first for uniformity) ────────────
# Larger labels to match page 02
st.markdown(
    "<div style='font-size: 1.25rem; font-weight: 600; margin-bottom: 0.15rem;'>"
    "Upload Cutsheet"
    "</div>",
    unsafe_allow_html=True,
)
cutsheet_files = st.file_uploader(
    "Upload Cutsheet",
    type=["xlsx", "xlsm"],
    accept_multiple_files=True,
    label_visibility="collapsed",
    help="One or more cutsheet files for enrichment"
)

st.markdown(
    "<div style='font-size: 1.25rem; font-weight: 600; margin-bottom: 0.15rem;'>"
    "Upload LV Portal Validation"
    "</div>",
    unsafe_allow_html=True,
)
lv_file = st.file_uploader(
    "Upload LV Portal Validation",
    type=["xlsx", "xlsm"],
    accept_multiple_files=False,
    label_visibility="collapsed",
    help="The export containing the error sheets (Optics, FEC, Interface Down, ...)"
)

# Blue primary button (default professional look for the main LV Portal tool)
st.markdown("""
<style>
div[data-testid="stButton"] button[kind="primary"] {
    background-color: #1565c0 !important;
    border-color: #1565c0 !important;
    color: white !important;
}
div[data-testid="stButton"] button[kind="primary"]:hover {
    background-color: #0d47a1 !important;
    border-color: #0d47a1 !important;
}
</style>
""", unsafe_allow_html=True)

run_btn = st.button(
    "🚀 Generate Formatted Report",
    type="primary",
    use_container_width=True,
    disabled=not (lv_file and cutsheet_files)
)

# ── Processing (core logic unchanged) ───────────────────────────────────────
if run_btn and lv_file and cutsheet_files:
    with st.spinner("Processing report with original logic..."):
        tmpdir = Path(tempfile.mkdtemp(prefix="hsg17_t1t0_"))
        try:
            # Save uploads to temp files
            lv_tmp = tmpdir / lv_file.name
            lv_tmp.write_bytes(lv_file.getvalue())

            cuts_tmp_paths = []
            for f in cutsheet_files:
                p = tmpdir / f.name
                p.write_bytes(f.getvalue())
                cuts_tmp_paths.append(str(p))

            # Call the exact original implementation
            out_path, _ = format_report(str(lv_tmp), cuts_tmp_paths, interactive=False)

            # Read result
            out_bytes = out_path.read_bytes()

            st.success(f"✅ Report ready: {out_path.name}")

            st.download_button(
                label=f"📥 Download {out_path.name}",
                data=out_bytes,
                file_name=out_path.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

            # ====================== SILENT CENTRAL LOGGING ======================
            # Read the *filtered* counts directly from the Summary tab of the produced report.
            # This matches exactly the numbers visible in the report's Summary (excludes
            # greyed-out rows, CT-off, -40 optics, etc.).
            try:
                source_name = lv_file.name
                placement, rack = derive_placement_and_rack_from_files([str(lv_tmp)] + cuts_tmp_paths)

                from utils.hsg17_models import extract_filtered_counts_from_summary
                counts = extract_filtered_counts_from_summary(str(out_path))

                for cat_name, cnt in counts.items():
                    if cnt > 0:
                        success = log_errors(
                            hall="HSG17",
                            rack_type="T1-T0",
                            building=placement,
                            rack=rack,
                            error_category=cat_name,
                            count=int(cnt),
                            source_file=source_name,
                            processed_by="HSG17_T1toT0_Gold"
                        )
                        if not success:
                            st.warning("Failed to write some log entries to the central file (see terminal).")
            except Exception as log_err:
                st.warning(f"Logging error: {log_err}")

        finally:
            # Cleanup
            try:
                for p in tmpdir.glob("*"):
                    p.unlink(missing_ok=True)
                tmpdir.rmdir()
            except Exception:
                pass
