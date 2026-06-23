#!/usr/bin/env python3
r'''
HSG17 T1-to-T0 Slack Formatter - Streamlit page

Core formatter logic lives in:
    utils/hsg17_slack_formatter.py
(exactly the updated v2 logic from the reference desktop script)

This page is ONLY the UI wrapper + HSG17-specific silent logging to the central
validation_error_log (Placement Group derivation + counts for the Dashboard).
The actual Excel transformation must be identical to the v2 reference.

UI, button behavior, download (incl. multi-file + ZIP), and all dashboard logging
are intentionally unchanged. New v2 features (3+-way mismatch grouping via
Union-Find + bold grid-line borders instead of orange/yellow fills, plus the
Mismatch↔Downlink cross tab) are delivered inside the formatted .xlsx only.
'''

import sys
import tempfile
import zipfile
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st

from utils.data_logger import log_errors
from utils.hsg17_models import derive_placement_and_rack_from_files
from utils.hsg17_slack_formatter import (
    TAB_ALIASES,
    find_tab,
    load_cutsheet,
    process_file,
)





# --- Streamlit UI -----------------------------------------------------------
st.set_page_config(page_title="HSG17 T1-to-T0 Slack Formatter", page_icon="🖥️", layout="wide")

st.title("🖥️ HSG17 T1-to-T0 Slack Upload")
st.caption("")

st.markdown("""
**How to use:**
1. Upload your **Cutsheet** (Installation Sheet)
2. Upload one or more **Slack Report Excel files**
3. Click **Generate Formatted Report**

The formatted report(s) will be available for immediate download.
The counts are silently logged for the HSG17 Dashboard (same as page 01).
""")

# Red primary button styling (to differentiate from the LV Portal tool)
st.markdown("""
<style>
div[data-testid="stButton"] button[kind="primary"] {
    background-color: #c62828 !important;
    border-color: #c62828 !important;
    color: white !important;
}
div[data-testid="stButton"] button[kind="primary"]:hover {
    background-color: #b71c1c !important;
    border-color: #b71c1c !important;
}
</style>
""", unsafe_allow_html=True)

st.markdown(
    "<div style='font-size: 1.25rem; font-weight: 600; margin-bottom: 0.15rem;'>"
    "Upload Cutsheet"
    "</div>",
    unsafe_allow_html=True,
)
cutsheet_uploader = st.file_uploader(
    "Upload Cutsheet",
    type=["xlsx", "xls"],
    accept_multiple_files=False,
    label_visibility="collapsed",
    help="The Installation Sheet from the master cutsheet"
)

st.markdown(
    "<div style='font-size: 1.25rem; font-weight: 600; margin-bottom: 0.15rem;'>"
    "Upload Slack Report"
    "</div>",
    unsafe_allow_html=True,
)
input_uploaders = st.file_uploader(
    "Upload Slack Report",
    type=["xlsx", "xls"],
    accept_multiple_files=True,
    label_visibility="collapsed",
    help="One or more Slack-style validation report files"
)

run_btn = st.button(
    "🚀 Generate Formatted Report",
    type="primary",
    use_container_width=True,
    disabled=not (cutsheet_uploader and input_uploaders)
)

if run_btn and cutsheet_uploader and input_uploaders:
    with st.spinner("Processing Slack report(s)..."):
        tmpdir = Path(tempfile.mkdtemp(prefix="hsg17_slack_"))
        try:
            cut_tmp = tmpdir / cutsheet_uploader.name
            cut_tmp.write_bytes(cutsheet_uploader.getvalue())

            slack_tmp_paths = []
            for f in input_uploaders:
                p = tmpdir / f.name
                p.write_bytes(f.getvalue())
                slack_tmp_paths.append(str(p))

            # Derive PG + representative rack so the Dashboard sees consistent tracking
            all_files_for_derive = [str(cut_tmp)] + slack_tmp_paths
            placement, rack = derive_placement_and_rack_from_files(all_files_for_derive)

            # --- Actual formatting using the reference logic (via utils) ---
            cut_df = load_cutsheet(str(cut_tmp))
            output_paths = []
            for in_path_str in slack_tmp_paths:
                in_p = Path(in_path_str)
                out_name = in_p.stem + "_FORMATTED.xlsx"
                out_p = tmpdir / out_name
                try:
                    produced = process_file(str(in_p), str(out_p), cut_df, log=lambda *a: None)
                    final_p = Path(produced) if produced else out_p
                    if final_p.exists():
                        output_paths.append(final_p)
                except Exception as proc_err:
                    st.warning(f"Could not fully process {in_p.name}: {proc_err}")

            # ====================== SILENT CENTRAL LOGGING (unified with 01 and 03) ======================
            # We read the *filtered* counts directly from the Summary tab(s) of the
            # produced report(s). This matches exactly the numbers visible in the
            # report's own Summary (excludes greyed-out / filtered rows).
            try:
                source_name = ", ".join([f.name for f in input_uploaders])

                from utils.hsg17_models import extract_filtered_counts_from_summary

                for out_p in output_paths:
                    counts = extract_filtered_counts_from_summary(str(out_p))
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
                                processed_by="HSG17_T1toT0_Slack",
                            )
                            if not success:
                                st.warning("Failed to write some log entries to the central file (see terminal).")
            except Exception as log_err:
                st.warning(f"Logging error: {log_err}")

            if output_paths:
                st.success(f"✅ {len(output_paths)} formatted report(s) ready for download.")
                if len(output_paths) == 1:
                    data = output_paths[0].read_bytes()
                    st.download_button(
                        label=f"📥 Download {output_paths[0].name}",
                        data=data,
                        file_name=output_paths[0].name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                else:
                    for p in output_paths:
                        b = p.read_bytes()
                        st.download_button(
                            label=f"📥 Download {p.name}",
                            data=b,
                            file_name=p.name,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                    zip_buf = BytesIO()
                    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                        for p in output_paths:
                            zf.write(p, arcname=p.name)
                    zip_buf.seek(0)
                    st.download_button(
                        label="📥 Download All as ZIP",
                        data=zip_buf.getvalue(),
                        file_name="HSG17_Slack_Reports_Formatted.zip",
                        mime="application/zip",
                        use_container_width=True
                    )
            else:
                st.info("Processing completed but no output files were produced (check inputs).")

        finally:
            try:
                for p in tmpdir.glob("*"):
                    p.unlink(missing_ok=True)
                tmpdir.rmdir()
            except Exception:
                pass
