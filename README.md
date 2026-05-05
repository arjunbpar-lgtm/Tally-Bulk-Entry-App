# Tally Hub

Cleaned source repository for the Tally Hub desktop tool.

## What this repo contains

- `src/tally_hub/tally_hub_v2.py`: main Sales/Purchase app
- `src/tally_hub/monthly_client_tally_converter.py`: shared sales conversion backend
- `docs/PURCHASE_FORMAT_HR_FEEDS.md`: purchase report mapping notes
- `packaging/*.spec`: PyInstaller build specs
- `main.py`: simple local entry point

## What was intentionally left out

- generated XML files
- ZIP bundles
- state JSON files
- `build/` and `dist/` artifacts
- old test output folders and scratch files

## Run locally

```powershell
python -m pip install -r requirements.txt
python main.py
```

You can also run the CLI:

```powershell
python main.py --help
```

## Suggested next cleanup

- replace the hardcoded company profile in `src/tally_hub/tally_hub_v2.py`
- add a sample `requirements-dev.txt` or packaging script
- separate shared XML helpers into a dedicated module
