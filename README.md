# HSG17 Validations

Clean, purpose-built validation tools for the HSG17 site (Batam).

## Structure
- pp.py - Main entry point
- pages/ - Individual tools (T0-to-Host, etc.)
- utils/ - Shared utilities (auth, central logging with delta tracking)
- data/ - Local error log (gitignored)

## Philosophy
This is a clean build from scratch, not a port of old scripts.

## Running Locally
`powershell
cd C:\Users\toddy\Documents\GitHub\HSG17-Validations
streamlit run app.py
`

## Central Logging
All tools log errors to data/validation_error_log.xlsx using the same format as JPB15-Validations.
The Dashboard (when built) will show current state + deltas vs previous run.
