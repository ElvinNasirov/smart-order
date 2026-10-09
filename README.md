# Smart Order

Hackathon demo: normalizes store Excel sales reports with a local LLM (Ollama/qwen2.5),
stores data in SQLite, and forecasts tomorrow's order quantities with LightGBM.

## Quick start

1. Run `install.bat` once (creates venv, installs deps, pulls model, adds desktop shortcut)
2. Double-click `start.bat` or the desktop shortcut to launch
3. Open the Streamlit UI, upload Excel reports, then build the forecast

## Files

- `data/` — store Excel/CSV sales files
- `normalize/normalize.py` — LLM-assisted normalization to `output/sales.db`
- `ml/forecast.py` — LightGBM forecast, writes `output/forecast_tomorrow.csv`
- `app/app.py` — Streamlit UI
