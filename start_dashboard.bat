@echo off
setlocal EnableExtensions

set "ROOT=%~dp0"
set "HOST=127.0.0.1"
set "PORT=8111"
set "PYTHONUTF8=1"

cd /d "%ROOT%"

if not exist "%ROOT%logs" mkdir "%ROOT%logs"

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$port=%PORT%; $listener=Get-NetTCPConnection -LocalPort $port -State Listen -ErrorAction SilentlyContinue | Select-Object -First 1; if ($listener) { exit 2 }"

if "%ERRORLEVEL%"=="2" (
  echo Dashboard is already running at http://%HOST%:%PORT%
  start "" "http://%HOST%:%PORT%"
  exit /b 0
)

set "PYTHON_EXE=python"
if exist "%ROOT%.venv\Scripts\python.exe" set "PYTHON_EXE=%ROOT%.venv\Scripts\python.exe"

echo Starting DingTalk to Feishu sync dashboard...
echo URL: http://%HOST%:%PORT%
echo Logs: %ROOT%logs\dashboard_server_%PORT%.out.log

start "DingTalk Feishu Sync Dashboard" /MIN /D "%ROOT%" cmd /c ""%PYTHON_EXE%" -X utf8 "%ROOT%dashboard_server.py" --host %HOST% --port %PORT% >> "%ROOT%logs\dashboard_server_%PORT%.out.log" 2>> "%ROOT%logs\dashboard_server_%PORT%.err.log""

timeout /t 2 /nobreak >nul
start "" "http://%HOST%:%PORT%"

exit /b 0
