#!/usr/bin/env python3
"""
HSG17 T0-to-Host / T1-to-T0 Validator — replacement using gold reference

This replaces the previous broken T0-to-Host implementation.
Uses the exact working T1-to-T0 formatter (lv_portal_formatter_T1toT0.v2 gold).

Retains central logging so the HSG17 Dashboard continues to work with all its features.
"""

import streamlit as st
import tempfile
from pathlib import Path

from utils.auth import require_login
from utils.data_logger import log_errors
from utils.t1_to_t0_formatter import format_report
from utils.hsg17_models import derive_placement_group

require_login()

st.set_page_config(page_title="HSG17 T0-to-Host", page_icon="🖥️", layout="wide")
st.title("HSG17 T0-to-Host Validator (T1-to-T0 Gold)")
st.caption("Exact gold logic • Tracks by Placement Group (per Bootstrap Sequence) • Feeds central Dashboard")

st.markdown("""
**Inputs:**
- **LV Portal Validation Export** (.xlsx): the export with sheets like Optic Errors, FEC_BER Errors, Interface Down Errors (and optional LLDP/mismatch).
- **Master Cutsheet(s) / Allconnections**: your T1toT0 allc (or master cutsheet). The lookup supports the columns in QFABT1toT0_..._allconnections.xlsx (DeviceA/DeviceB combined host+port, RackA/RackB, Source_port, DMARC*, Destination_port, EasyMark+, Physical Ports etc.).

Issues are tracked by **Placement Group** (see HSG17 Bootstrap Sequence for rack->PG mapping; e.g. rack 3110 = PG14).

For your untouched original Dashboard compatibility the error categories are logged under the original names (LLDP Mismatch + Link Down etc.).

Upload, generate, download the `_formatted.xlsx` with the 5 perfect tabs (Summary navy, Mispatches red, Downlinks orange, Optics brown, FEC Errors purple).
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

            st.info("The output has the 5 tabs, exact column orders, pair borders, styling, and enrichment from the gold formatter.")

            # ====================== SILENT CENTRAL LOGGING (to retain Dashboard features) ======================
            # Log the counts so the HSG17 Dashboard shows current state + deltas.
            # Use placement group (from rack in allc, per HSG17 Bootstrap Sequence)
            try:
                source_name = lv_file.name if lv_file else "unknown"
                placement = "PG14"  # fallback for rack 3110 case
                try:
                    import pandas as pd
                    if cuts_tmp_paths:
                        allc_df = pd.read_excel(cuts_tmp_paths[0], sheet_name=0)
                        rack_cols = [c for c in allc_df.columns if "rack" in str(c).lower()]
                        if rack_cols:
                            racks = allc_df[rack_cols[0]].dropna().astype(str).str.replace(r"Rack\s*", "", regex=True).str.strip().str.split().str[0]
                            racks = racks[racks.str.len() > 0]
                            if len(racks) > 0:
                                pgs = racks.apply(derive_placement_group)
                                most_common = pgs.value_counts().index[0]
                                if most_common and most_common.startswith("PG"):
                                    placement = most_common
                except Exception:
                    pass
                for cat_key, cnt in counts.items():
                    if cnt > 0:
                        # Map to legacy category names so the untouched original Dashboard cards continue to work and update
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
                            building=placement,  # placement group e.g. "PG14"
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
st.caption("HSG17 • Gold T1-to-T0 formatter • Issues tracked by Placement Group (PG14 for rack 3110) per Bootstrap Sequence • Categories mapped for your untouched original Dashboard compatibility • All processing is local.")