@echo off
chcp 65001 >nul
set PYTHONUTF8=1
cd /d "%~dp0"

if not exist .venv (
    echo Creating virtual environment...
    python -m venv .venv
)
call .venv\Scripts\activate.bat

echo Installing dependencies...
pip install -r requirements.txt

echo Pulling Ollama model...
ollama pull qwen2.5:7b-instruct

echo Creating desktop shortcut...
set "STARTBAT=%~dp0start.bat"
set "WORKDIR=%~dp0"
powershell -NoProfile -Command "$ws = New-Object -ComObject WScript.Shell; $lnk = $ws.CreateShortcut([IO.Path]::Combine([Environment]::GetFolderPath('Desktop'), 'Smart Order.lnk')); $lnk.TargetPath = '%STARTBAT%'; $lnk.WorkingDirectory = '%WORKDIR%'; $lnk.Save()"

echo.
echo Install complete. Desktop shortcut 'Smart Order' created.
pause
