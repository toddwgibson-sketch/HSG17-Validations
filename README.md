# HSG17 Validations

Clean, purpose-built validation tools for the HSG17 site (Batam).

**Status**: Active (resumed from parked v1 snapshot, June 2026)

## Current Capabilities

- **T0-to-Host Validator** (`pages/01_HSG17_T0_to_Host.py`)
  - 6-stage clean pipeline: Ingest → Normalize (DH block derivation from Bootstrap Sequence) → Enrich (PP joins) → Analyze (connected-component clustering) → Format → Log
  - Professional 5-tab Excel output (Summary + LLDP Mismatch + Link Down, Optic Errors, FEC_BER Errors, Interface Down Errors)
  - Rich orange/yellow row highlighting by Cluster on the LLDP tab
  - Strong PP enrichment from the large allconnections file (A-side + B-side matching)
  - Prominent "Block" (DH-001 … DH-105 (Spines), IPR, AUX) column using authoritative rack/PG/PP rules
  - Automatic silent logging to central `data/validation_error_log.xlsx`

- **Dashboard** (`pages/10_HSG17_Dashboard.py`)
  - Executive "get to zero" view with widget cards (5 per row)
  - Current state only + deltas vs previous run per (Block + category)
  - Category × Block pivot + totals
  - Storage location helper + live log download

## Running

```powershell
cd C:\Users\toddy\Documents\GitHub\HSG17-Validations
streamlit run app.py
```

Login: `admin` / `admin` (see `utils/auth.py`).

Sidebar navigation to the two tools.

## Data & Logging

- Input files are user-uploaded (never committed):
  - `QFABT0toHOST_allconnections.xlsx`
  - `rack_validation_merged_*.xlsx` (the pre-classified cutsheet with the 4 error sheets)
- Central log: `data/validation_error_log.xlsx` (gitignored). Same schema used by JPB15 so one dashboard style can cover multiple halls later.
- Re-uploading the same Block+category overwrites previous counts (dashboard always shows latest).

## Tests (synthetic, no real data needed)

```powershell
python -m tests.test_hsg17_processor
```

Covers:
- Block derivation (PP regex, rack prefix maps, device name heuristics, spines/IPR)
- Connected component clustering for LLDP swaps
- Full roundtrip through process_hsg17_t0_host + extract
- PP enrichment (A/B side)
- Error cases (missing cutsheet etc.)

## Architecture Notes

See `DESIGN.md` for the original block strategy, data model, and 6-stage breakdown (still accurate).

`utils/hsg17_t0_host.py` is the self-contained processor (easy to unit test or call from other scripts).

`utils/hsg17_models.py` holds the canonical DH list + PG/RACK lookup tables.

## Recent Work (resumption)

- Required both uploads in the T0 tool (cutsheet was "strongly recommended" in UI but mandatory for the pipeline).
- Robust first-sheet read for allconnections + kept all columns for enrichment.
- Added comprehensive synthetic test suite.
- Extended rack→block map per the Confluence Bootstrap Sequence.
- Significantly strengthened PP enrichment (dual-end indexing, B-side fallback, richer combined labels + T0 hints).

## Next Ideas (from DESIGN + resume)

- Even stronger / multi-key PP joins or fuzzy matching
- Additional tools (T1-to-T0, GPU tray validations, etc.)
- Dashboard: date filters, block search, per-rack sub-views, export buttons
- Optional: move common dashboard bits into reusable components if feeding multiple halls

---

**Clean isolated repo** (deliberately separate from JPB15-Validations). Same logging contract for future unified views.

All processing is local. No data leaves your machine.