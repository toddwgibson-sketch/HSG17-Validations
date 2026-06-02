# HSG17 Validations

Clean, purpose-built validation tools for the HSG17 site (Batam).

**Status**: Active (integrated gold T1-to-T0 formatter + Placement Group tracking, June 2026)

## Current Capabilities

- **T0-to-Host Validator** (`pages/01_HSG17_T0_to_Host.py`)
  - Powered by the exact gold T1-to-T0 formatter (the `lv_portal_formatter_T1toT0.v2` reference implementation that "works perfectly")
  - Inputs: LV Portal Validation Export (.xlsx containing Optic Errors, FEC_BER Errors, Interface Down Errors, and optional LLDP/mismatch sheets) + Master Cutsheet(s)/Allconnections (the simpler T1toT0 style; supports combined "DeviceA host+port", RackA, Source_port, DMARC1/2, Destination_port, DeviceB, RackB, EasyMark+, Physical Port columns, etc.)
  - Produces the professional 5-tab report exactly as specified:
    - Summary (navy)
    - Mispatches (red)
    - Downlinks (orange)
    - Optics (brown)
    - FEC Errors (purple)
  - All the gold details preserved: precise column orders per tab, tab colors, thick pair borders for physical cables (s1/s2 etc.), "Also Downlink" grey rows, cutsheet-miss yellow, union-find swap clustering + thick group borders in Mispatches, L&R calculation, logical-pair and reverse-lookup fallbacks, Patch Panel Matrix parsing, per-channel Rx Power and RawBer parsing, BER severity classification, clean strings, freeze_panes, tuned widths, etc.
  - Issues are tracked by **Placement Group** (PG1–PG32 + T1 spines 151–154) per the HSG17 Bootstrap Sequence. The page derives the PG directly from the racks of errored devices/ports in the LV export (e.g. rack 3110 = PG14). The full rack→PG mapping from the Bootstrap Sequence is in `utils/hsg17_models.py`.
  - Automatic silent logging to central `data/validation_error_log.xlsx` (building column = PGxx). Categories are mapped to the legacy names for full compatibility with the original Dashboard.

- **Dashboard** (`pages/10_HSG17_Dashboard.py`)
  - Executive "get to zero" view with widget cards (5 per row)
  - Current state only + deltas vs previous run per (Placement Group + category)
  - Category × Placement Group pivot + totals
  - Live log download button (storage location details removed per request)

## Running

```powershell
cd C:\Users\toddy\Documents\GitHub\HSG17-Validations
streamlit run app.py
```

Login: `admin` / `admin` (case-insensitive; see `utils/auth.py`).

The root `app.py` content is hidden — after login it immediately switches to the main T0-to-Host page. Sidebar navigation between the two tools is still available.

## Data & Logging

- Input files are user-uploaded (never committed):
  - LV Portal validation export (the source of the categorized errors)
  - T1toT0 allconnections / master cutsheet (e.g. QFABT1toT0_17.1to17.4_allconnections.xlsx or equivalent)
- Central log: `data/validation_error_log.xlsx` (gitignored). Uses Placement Group in the `building` column. Same schema used by JPB15 so one dashboard style can cover multiple halls later.
- Re-uploading the same Placement Group + category overwrites previous counts (dashboard always shows latest).

## Architecture Notes

- The gold formatter lives in `utils/t1_to_t0_formatter.py` (self-contained, no tkinter/GUI/persistence, pure stdlib + openpyxl). Reusable from code or tests via:
  ```python
  from utils.t1_to_t0_formatter import format_report
  out_path, counts = format_report(lv_export_path, [allc_or_cutsheet_paths], interactive=False)
  ```
- `utils/hsg17_models.py` holds the authoritative `RACK_TO_PLACEMENT_GROUP` mapping (sourced from the Bootstrap Sequence you provided) + the `derive_placement_group(rack)` helper. (Legacy DH block material is present for reference but not used for HSG17 tracking.)
- The old `utils/hsg17_t0_host.py` + related tests are legacy (kept for reference; the active page uses the gold formatter instead).
- Central logging (`utils/data_logger.py`) is shared so the Dashboard works for both old and new data.

## Recent Work

- Integrated the gold T1-to-T0 formatter as the replacement for the previous T0-to-Host implementation.
- Switched HSG17 tracking from DH Blocks to Placement Groups (full mapping from the provided Bootstrap Sequence txt integrated and used at runtime).
- Updated the T0-to-Host page to derive the correct PG from the actual errored racks in the LV export (not the entire allc).
- Cleared old/wrong data from the central log.
- Simplified the Dashboard UI (just the download button; no more storage location paths/size text).
- Login screen now shows credentials hint (U:Admin P:Admin).
- Root `app.py` hidden (immediate switch to main page after login).
- Categories logged under legacy names for full backward compatibility with the original Dashboard cards/logic.

## Next Ideas

- Direct support for additional input formats (e.g. pre-classified cutsheet without LV Portal wrapper).
- Per-rack or multi-PG breakdown inside a single report.
- Dashboard enhancements (date range, PG search, export of cards, etc.).
- Bring the synthetic tests up to date for the gold path + PG derivation.
- Optional: clean up legacy T0-to-Host processor files once fully migrated.

---

**Clean isolated repo** (deliberately separate from JPB15-Validations). Same logging contract for future unified views.

All processing is local. No data leaves your machine.