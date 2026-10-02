# AI Assistant Instructions

Small Flask web service deployed to Google Cloud Run (project `consumption-web`, region `europe-north1`).
See README.md for complete description.

## Core Components

- **Calculation Engine** (`calculator.py`): `process_excel(input, output)` 
  - Fetches NPS Latvia prices from Elering API (`dashboard.elering.ee`) with 15-min intervals
  - ✅ **FIX (2026-10-01)**: Prices shifted by +15 minutes to align with consumption timing
    - Rule: Consumption recorded at HH:MM uses price from HH:MM+15min interval
    - Why: Elering publishes prices at 15-min marks; consumption is recorded at 15-min intervals too
  - ✅ **FIX (2026-10-02)**: Date parsing fixed (DD.MM.YYYY with dayfirst=True, not MM.DD.YYYY)
  - ✅ **FIX (2026-10-02)**: Month-end filtering excludes next-month data from TOTAL calculations

- **Web Interface & Storage** (`app.py`):
  - Flask routes: upload, history, download, delete
  - ByHours builder (shifts data by BYHOURS_SHIFT_HOURS hours)
  - Google Cloud Storage only (requires `GCS_BUCKET` env var)
  - History persisted in `history.json`

- **Runtime**: Python 3.11 (see Dockerfile)
  - Keep all dependencies pinned in requirements.txt
  - Input sheet `Dati` must have exact Latvian column names (UTF-8)

## Development Rules

- Never commit client Excel files or `.env` files
- Never deploy without explicit user request
- All tests must pass before commits

## Local Testing

```bash
python calculator.py samples/input.xlsx samples/output.xlsx
```

Result: `samples/output.xlsx` with sheets: Dati, Total_Calculation, ByDays

## Environment Variables

Both old and new variable names supported for backward compatibility:
- `BYHOURS_SHIFT_HOURS` (preferred, new name)
- `METER_SHIFT_HOURS` (legacy alternative)
