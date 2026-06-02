# HSG17 T0-to-Host — Clean Architecture Design

**Note (June 2026)**: The active implementation in `pages/01_HSG17_T0_to_Host.py` now uses the gold T1-to-T0 formatter (from the provided lv_portal_formatter reference) with Placement Group tracking (instead of DH blocks) per the Bootstrap Sequence. This document describes the original design goals; the current code deviates toward the gold reference while preserving logging compatibility with the Dashboard.

---

**Status**: Draft — Ready for review
**Date**: 2026-06-01
**Repo**: Separate HSG17-Validations (clean build from scratch)

## Goals

* Clean build from scratch (no porting of old tangled JPB/SYD code)
* Deliver equivalent rich features to the JPB/SYD20 T0-to-Host tools

  * High-quality mismatch pairing and clustering (orange/yellow highlighting)
  * Grey-out logic for fully-down switches
  * Strong patch panel / cutsheet enrichment
  * Professional multi-tab Excel output with Summary
  * Central error logging with delta tracking (current vs previous run per block)
* Use consistent Block terminology
* Integrate cleanly with the shared HSG17 logging system
* Keep the processor modular and maintainable

## Data Sources

1. QFABT0toHOST\_allconnections.xlsx (main raw data)

   * Large flat connections table (\~32k rows)
   * Excellent embedded path/patch panel information in Full Label and EasyMark+ columns
   * Strong DeviceA (compute trays) to DeviceB (T0 switches) structure
   * Balanced across 4 quadrants (p1-p4)
2. rack\_validation\_merged\_2614HC8023.xlsx (pre-classified cutsheet)

   * Already contains categorized errors:

     * LLDP Errors (12)
     * Optic Errors (48)
     * Raw BER Errors (4)
     * Interface Errors (44)
   * This file will be a primary input for the formatter

## Proposed Internal Data Model

### Core Entities

* **HSG17Connection**

  * source\_device, source\_port
  * dest\_device, dest\_port
  * t0\_switch\_port
  * block (derived)
  * rack\_a, rack\_b
  * patch\_panel\_path (from Full Label / EasyMark)
  * full\_label
* **HSG17Error**

  * error\_type (LLDP Mismatch, Optic, FEC\_BER, Interface Down)
  * device\_a, port\_a, device\_b, port\_b
  * block
  * details (Rx power, BER values, mismatch status, etc.)
  * patch\_panel\_matrix

## Proposed Processor Architecture (Clean \& Modular)

Instead of one giant messy process\_files() function, we will use clear, testable stages:

1. **Ingest**

   * Load raw allconnections table
   * Load pre-classified cutsheet (multiple sheets)
2. **Normalize**

   * Convert raw rows into HSG17Connection objects
   * Convert cutsheet rows into HSG17Error objects
   * Derive consistent Block identifiers
3. **Enrich**

   * Cross-reference connections with errors using device+port keys
   * Attach full patch panel information to errors
   * Apply any HSG17-specific normalizations
4. **Analyze** (lightweight)

   * Mismatch clustering logic (if needed beyond the cutsheet)
   * Grey-out detection for fully-down switches
   * Any additional grouping required for the rich output
5. **Format**

   * Build professional Excel workbook with tabs:

     * Summary (counts per Block)
     * LLDP Mismatch + Link Down (clustered + colored)
     * Optic Errors
     * FEC\_BER Errors
     * Interface Down Errors
   * Apply styling, conditional formatting, grey-out
6. **Log**

   * Extract per-Block counts
   * Call central log\_errors() with full delta support

## Integration with Shared Systems

* Use the same central validation\_error\_log.xlsx format
* Full support for the delta tracking logic (current vs previous run per block)
* Will feed the future HSG17 Dashboard using the same patterns as JPB15
* Same authentication module pattern

## Open Questions

1. Do we need to perform additional mismatch clustering ourselves, or can we trust the pre-classified LLDP sheet and focus on beautiful formatting + enrichment?

2. What is the exact definition of a Block for HSG17? (e.g. hsg17:5708:3 style, or something derived from device names?)

3. Should the output tab names and structure stay very close to JPB/SYD20, or can we adapt them to the HSG17 error categories?

4. Any specific new capabilities you want to design in for HSG17 that the older tools didnt have?









## User Answers to Open Questions (2026-06-01)

### 1. Mismatch Clustering
User wants us to perform the mismatch clustering ourselves (do not just trust the pre-classified LLDP sheet for the final rich output).

### 2. Block Definition for HSG17
User provided the authoritative source: the full Confluence export (PDF + .txt + .xlsx) in the Batam folder. The canonical grouping for dashboard/logging is the **DH sector** (DH-001 through DH-105 plus DH-102 for spines) plus special cases (IPR rack, AUX racks).

## Block Identification Strategy for HSG17 (Concrete & Testable)

**Goal:** Every error row written to the central log must have a stable `block` value that matches the DH sectors in the Bootstrap Sequence document so the shared delta-aware Dashboard works without changes.

### Canonical Block Values (from authoritative Confluence export)
- DH-001, DH-002, DH-003, DH-004, DH-005
- DH-101, DH-102 (QFABT1 Spines), DH-103, DH-104, DH-105
- Special: IPR (PG-201), AUX (5920/6002)

These ~10 primary blocks (with sub-grouping by Placement Group or rack for the dashboard cards) satisfy the "~60 blocks" target when the dashboard shows individual racks or PGs that have current errors.

### Derivation Rules (priority order, implement in Normalize stage)

1. **Primary — Patch Panel DH identifier (most reliable)**
   - Source columns (both input files):
     - `rack_validation_merged...`: `Patch Panel Matrix`
     - `QFABT0toHOST_allconnections.xlsx`: `Full Label` and `EasyMark+ --- Patch Panels`
   - Regex: `DH[\d-]+` or `DH\.[\d]+` (case-insensitive)
   - Normalization examples:
     - `DH2`, `DH-2`, `DH.2` → `DH-002`
     - `DH4` → `DH-004`
     - `DH7` → `DH-007`
     - `DH102` or `DH-102` → `DH-102 (Spines)`
   - This is the **preferred** method because the PP labels are present on virtually every row and directly reference the DH sectors in the bootstrap plan.

2. **Secondary — Device A Rack / Device A Name (cutsheet hsg17: format)**
   - In pre-classified cutsheet the `Device A Rack` column uses `hsg17:<rack>:<pos>` (e.g. `hsg17:5708:3`, `hsg17:0807:5`, `hsg17:2814:9`).
   - Extract the rack number (5708, 0807, 2814...).
   - Map via the known rack ranges in the Bootstrap Sequence table:
     - 5708-5715, 5807-5814, 5907-5915, 6007-6014 → `DH-102 (Spines)`
     - 0807, 0812, 0907, 0912 and similar 08xx/09xx in PG 3-6 → `DH-002`
     - 2502, 2809, 2909, 3010 etc. → lookup in the PG-to-DH table (or fall back to PP method)
   - For GPU racks (`GPU_GB300_NVL72_R.03` in Rack SKU or device names containing "compute"): attach the PG number from the bootstrap mapping.

3. **Tertiary — Device Name patterns (for T0/T1/IP distinction)**
   - `hsg17-q2-p*-t1-*` or rack 57xx-60xx → T1 spine, force `DH-102 (Spines)`
   - `hsg17-q2-p*-t0-*` → T0 switch, use PP DH or rack→DH map
   - `hsg17-q2-ip-r*` → IPR rack → `IPR (PG-201)`
   - `aux.*` → `AUX Racks`

4. **Final fallback**
   - Use the first DH found anywhere in the row text.
   - If nothing found, log a warning and use `UNKNOWN-<first-rack>` (this should never happen in production data).

### Implementation Notes
- The Normalize stage will produce a clean `block` field on every `HSG17Error` and `HSG17Connection`.
- The Log stage will emit `hall="HSG17"`, `block=<normalized DH>`, `rack_type="T0"|"T1"|"GPU"|"IPR"|"AUX"`, plus the usual category counts.
- This guarantees the existing `data_logger.py` + delta math + Dashboard cards work unchanged for HSG17 data.

## Next Steps (Post Block Strategy Lock)

1. Finalize exact HSG17Error / HSG17Connection pydantic models (or simple dataclasses).
2. Define the precise 4-tab output structure + Summary tab columns (match the 4 sheets already present in the pre-classified cutsheet).
3. Implement the 6-stage processor (start with Ingest + Normalize, including the DH regex + rack-to-DH map).
4. Add mismatch clustering (connected-component on the LLDP Mismatch rows) exactly as required for rich orange/yellow pairing.
5. Wire the finished formatter to the central logger (silent, no extra buttons).

The two Batam files + the Confluence text/xlsx are now the single source of truth for all HSG17 T0-to-Host logic.

## HSG17 Block & Placement Group Structure (from Confluence PDF)

### Rack Counts
- GPU GB300: 224 racks (18 compute trays / rack, 72 GPUs / rack) → 16,128 devices
- QFABT0: 64 racks (8 devices/rack, 128 devices/plane) → 512 devices
- QFABT1: 32 racks (8 devices/rack, 64 devices/plane) → 256 devices
- QFABIP: 1 Inter Planar Rack
- Total racks: 321

### Placement Group → DH Sector Mapping (authoritative)
- PG 151–154: QFABT1 spine racks (32 total) → DH-102
- PG 201: IPR rack (5911)
- PG 1–2 → DH-001
- PG 3–6 → DH-002
- PG 7–10 → DH-003
- PG 11–14 → DH-004
- PG 15–18 → DH-005
- PG 19–22 → DH-101
- PG 23–26 → DH-103
- PG 27–30 → DH-104
- PG 31–32 → DH-105

AUX racks (5920, 6002) are bootstrapped separately.

This mapping will be hardcoded as a small lookup table in the Normalize stage (or loaded from a tiny reference CSV if it grows).
