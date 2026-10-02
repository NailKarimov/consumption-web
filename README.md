# Consumption Web 🔌

Web service for calculating electricity costs from meter data (NPS Latvia).

---

## 🎯 How It Works

1. **Upload Data**: User uploads Excel file (`.xlsx` / `.xls`) with sheet **`Dati`**
   - Required columns: `Datums`, `Patērētā elektroenerģija (A+) (kWh)`, `Tīklā nodotā enerģija (A-) (kWh)`, `Objekta adrese`
   - Date format: DD.MM.YYYY HH:MM

2. **Cost Calculation** (`calculator.py`):
   - Fetches NPS Latvia prices from Elering API (`dashboard.elering.ee`)
   - Aligns prices with meter intervals (15-min / 1-hour)
   - ✅ **Fix (2026-10-01)**: Prices shifted by +15 min to match consumption timing
   - ✅ **Fix (2026-10-02)**: Date parsing corrected (dayfirst=True for DD.MM.YYYY)
   - ✅ **Fix (2026-10-02)**: Month-end filtering excludes next-month data
   - Calculation: `kWh × EUR/MWh ÷ 1000`

3. **Output Sheets** (`app.py`):
   - **ByDays**: Aggregated hourly data (00:00–23:00) per day
     - Columns: each day + daily sum (Sum column)
     - TOTAL row: monthly sum
   - **ByHours** (optional): Hourly data with BYHOURS_SHIFT_HOURS offset
   - **Total_Calculation**: Full calculations per 15-min interval
   - **Dati**: Original meter data

4. **Storage**: Results saved to Google Cloud Storage (GCS)
   - Run history: `history.json`
   - Structure: `runs/<job_id>/{source/, output/, log.txt}`

---

## 📁 Project Structure

| File | Purpose |
|---|---|
| `calculator.py` | **Calculation engine**: main Excel processing logic (standalone CLI) |
| `app.py` | **Flask web interface**: upload, history, download, delete, ByHours builder |
| `main.py` | Gunicorn entry point (`main:app`) |
| `requirements.txt` | Python dependencies (pinned versions) |
| `Dockerfile` | Docker build for Cloud Run (Python 3.11) |
| `samples/` | Test files (input.xlsx, output.xlsx) |
| `templates/` | Jinja HTML templates |
| `CLAUDE.md` | AI assistant instructions |

---

## ⚙️ Environment Variables

| Variable | Default | Description |
|---|---|---|
| `GCS_BUCKET` | *(required)* | GCS bucket for history & files. Prod: `consumption-web-968627726218-runs` |
| `DATA_ROOT` | `/tmp/data` | Temp folder for processing |
| `MAX_HISTORY` | `200` | Max runs in history |
| `BYHOURS_SHIFT_HOURS` | `1` | ByHours sheet shift (hours earlier) |
| `METER_SHIFT_HOURS` | *(optional)* | Legacy alternative to BYHOURS_SHIFT_HOURS (backward compatibility) |

---

## 🚀 Local Setup

### Prerequisites
- Python 3.11+
- Git
- Google Cloud CLI (`gcloud`)
- VS Code (recommended)

### Installation Steps

1. **Clone & open in VS Code**
   ```bash
   git clone <repo>
   code consumption-web
   ```

2. **Install recommended extensions**
   ```
   Ctrl+Shift+P → Tasks: Run Task → Setup: create venv + install requirements
   ```

3. **Select Python interpreter**
   ```
   Ctrl+Shift+P → Python: Select Interpreter → .venv
   ```

4. **Authenticate with Google Cloud**
   ```
   Ctrl+Shift+P → Tasks: Run Task → GCP: login
   ```
   (opens browser for OAuth)

5. **Create test GCS bucket** (one time)
   ```bash
   gcloud storage buckets create gs://consumption-web-968627726218-dev \
     --project consumption-web \
     --location europe-north1 \
     --uniform-bucket-level-access
   ```

6. **Copy `.env.example` → `.env`** and fill:
   ```
   GCS_BUCKET=gs://consumption-web-968627726218-dev
   DATA_ROOT=/tmp/data
   ```

7. **Run locally**
   ```
   F5 → Flask: run web app locally → http://localhost:8080
   ```

### Test Calculation Only (No Web UI)

```bash
python calculator.py samples/input.xlsx samples/output.xlsx
```

**Result**: `samples/output.xlsx` with sheets `Dati`, `Total_Calculation`, `ByDays`

> 💡 `samples/` folder is in `.gitignore` (not committed)

---

## 📊 Input Example

**Sheet `Dati`**:

| Datums | Patērētā elektroenerģija (A+) (kWh) | Tīklā nodotā enerģija (A-) (kWh) | Objekta adrese |
|---|---|---|---|
| 01.08.2026 00:15 | 2 | 0 | Krāslava, Pakalni |
| 01.08.2026 00:30 | 1 | 0 | Krāslava, Pakalni |
| ... | ... | ... | ... |
| 31.08.2026 23:45 | 5 | 60 | Krāslava, Pakalni |

---

## 📈 Output Example

**ByDays Sheet** (hourly aggregates):

| Time | 01.08.2026 | 02.08.2026 | ... | 31.08.2026 | Sum |
|---|---|---|---|---|---|
| 00:00 | 27.61 | 60.39 | ... | 37.21 | *NaN* |
| 01:00 | 55.43 | 90.38 | ... | 29.62 | *NaN* |
| ... | ... | ... | ... | ... | ... |
| TOTAL | 7220.23 | 710.15 | ... | 6156.54 | **33728** |

- **Sum column**: Empty for hours, **33728** in TOTAL (total consumption)
- **Money TOTAL**: 6156.54 EUR

---

## 🔧 Development

### Run Tests
```bash
python -m pytest tests/ -v
```

### Code Quality
Recent refactoring:
- ✅ Removed 14 temporary test files
- ✅ Eliminated code duplication (valid_records calculation)
- ✅ Renamed `ProgramCalcMain.py` → `calculator.py`
- ✅ Added fix documentation

---

## 📦 Deployment

### Deploy to Production
```bash
Ctrl+Shift+P → Tasks: Run Task → Cloud Run: deploy to PRODUCTION
```

Or via PowerShell:
```bash
.\deploy.ps1
```

### Rollback
- Cloud Console → Cloud Run → consumption-web → Revisions
- Switch traffic to previous revision

---

## 🐛 Known Fixes

| Date | Issue | Solution |
|---|---|---|
| 2026-10-01 | Prices offset -15 min from consumption | Added +15 min shift |
| 2026-10-02 | DD.MM.YYYY parsed as MM.DD.YYYY | Added dayfirst=True |
| 2026-10-02 | TOTAL included 01.09 (next month) | Added end_of_month filter |
| 2026-10-02 | Duplicate valid_records calculation | Moved to single variable |

---

## 📞 Support

- **Developer docs**: See `CLAUDE.md`
- **Logs**: All operations logged to GCS `runs/<job_id>/log.txt`
- **History**: Available on web interface (max 200 runs)

---

## 📄 License

Internal project. Not for public distribution.

---

**Last Updated**: 2026-10-02  
**Version**: 1.1 (refactored, internationalized)  
**Status**: ✅ Production-ready
