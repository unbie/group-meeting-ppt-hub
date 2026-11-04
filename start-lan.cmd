@echo off
setlocal
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  echo Python environment not found. Follow README.md to install dependencies.
  pause
  exit /b 1
)
if not exist "config.json" (
  if not defined PPT_HUB_SECRET_KEY goto missing_config
  if not defined PPT_HUB_ADMIN_PASSWORD_HASH goto missing_config
)
set "PPT_HUB_HOST=0.0.0.0"
echo Starting Group Meeting PPT Hub for LAN access...
echo Keep this window open. Minimize it if needed. Press Ctrl+C to stop.
".venv\Scripts\python.exe" app.py
if errorlevel 1 (
  echo Startup failed. If port 1410 is in use, stop the previous instance first.
  pause
  exit /b 1
)
exit /b 0
:missing_config
echo Administrator configuration is missing. Run .venv\Scripts\python.exe setup.py first.
pause
exit /b 1
