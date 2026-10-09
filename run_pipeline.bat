@echo off
chcp 65001 >nul
set PYTHONUTF8=1
cd /d "%~dp0"
if exist .venv\Scripts\activate.bat call .venv\Scripts\activate.bat

echo [0/3] Proverka Ollama...
curl -s http://localhost:11434/api/tags >nul || (echo Ollama ne zapushena. Zapusti: ollama serve & pause & exit /b 1)

echo [1/3] Normalizaciya dannyh iz data\
python normalize\normalize.py || (echo OSHIBKA v normalize & pause & exit /b 1)

echo [2/3] Prognoz
python ml\forecast.py || (echo OSHIBKA v forecast & pause & exit /b 1)

echo [3/3] Zapusk demo
streamlit run app\app.py
