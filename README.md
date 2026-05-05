# Tally Hub

Tally Hub is a desktop utility for converting monthly Sales and Purchase Excel reports into Tally-compatible XML bundles.

## Current scope

- Sales report import
- Purchase report import
- Automatic grouping of rows into vouchers
- Missing party or supplier ledger export
- Missing stock master export
- ZIP bundle output for import handoff
- GUI mode and CLI mode

## Main files

- `main.py`: simple local entry point
- `src/tally_hub/tally_hub_v2.py`: combined Sales/Purchase desktop app and CLI
- `src/tally_hub/monthly_client_tally_converter.py`: shared Sales conversion backend
- `src/tally_hub/defaults.py`: centralized default app settings and ledger names
- `docs/PURCHASE_FORMAT_HR_FEEDS.md`: purchase report mapping notes
- `packaging/TallyHubV2.spec`: PyInstaller build spec
- `packaging/TallyHubV2_PurchaseFix.spec`: later packaging variant kept for reference

## How the app works

### Sales flow

1. Load the monthly Sales Excel report.
2. Optionally load a current Tally master XML.
3. The app groups rows into vouchers by date, voucher number, and party.
4. It generates:
   - Sales vouchers XML
   - New party ledger masters XML
   - Missing stock masters XML when needed
   - A summary JSON
   - A ZIP bundle containing the generated files

### Purchase flow

1. Load the monthly Purchase Excel report.
2. Optionally load a current Tally master XML.
3. The app groups rows into vouchers by date, voucher number, supplier, and state context.
4. It chooses local vs interstate purchase ledger based on supplier state code.
5. It generates:
   - Purchase vouchers XML
   - Supplier ledger masters XML
   - Missing stock masters XML when needed
   - A summary JSON
   - A ZIP bundle containing the generated files

## Run locally

```powershell
python -m pip install -r requirements.txt
python main.py
```

For CLI usage:

```powershell
python main.py --help
```

## Repository cleanup rules used here

This repo intentionally excludes:

- generated XML files
- ZIP bundles
- state JSON files
- `build/` and `dist/` artifacts
- scratch outputs and local caches

## Notes

- The current GUI still includes a company-specific default profile name in `src/tally_hub/defaults.py`.
- The Sales backend is shared by both the standalone Sales converter logic and the combined Tally Hub app.
- The April 19, 2026 `TallyHubV2_PurchaseFix.exe` appears to be the latest packaged build found in the original workspace.
