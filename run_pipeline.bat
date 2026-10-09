@echo off
chcp 65001 >nul
set PYTHONUTF8=1
cd /d "%~dp0"
if exist .venv\Scripts\activate.bat call .venv\Scripts\activate.bat

echo [0/3] Checking Ollama...
curl -s http://localhost:11434/api/tags >nul || (echo Ollama is not running. Start it with: ollama serve & pause & exit /b 1)

echo [1/3] Normalizing data from data\
python normalize\normalize.py || (echo ERROR in normalize & pause & exit /b 1)

echo [2/3] Forecasting
python ml\forecast.py || (echo ERROR in forecast & pause & exit /b 1)

echo [3/3] Starting demo
streamlit run app\app.py
