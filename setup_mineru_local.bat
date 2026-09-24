@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
  echo [ERROR] .venv not found. Run install_mineru.bat first.
  pause
  exit /b 1
)

call ".venv\Scripts\activate.bat"

rem Change this to your preferred model root if D: is not available.
if not defined MINERU_MODEL_BASE_DIR set "MINERU_MODEL_BASE_DIR=D:\mineru\models"
if not exist "%MINERU_MODEL_BASE_DIR%" mkdir "%MINERU_MODEL_BASE_DIR%"
set "MINERU_MODEL_SOURCE=modelscope"

echo Model directory: %MINERU_MODEL_BASE_DIR%
echo.

echo ============================================================
echo 1/4 Download MinerU standard-tier models from ModelScope
echo ============================================================
mineru-kit models download --tier standard --source modelscope -v
if errorlevel 1 goto :err

echo.
echo ============================================================
echo 2/4 Verify downloaded models
echo ============================================================
mineru-kit models verify --tier standard
if errorlevel 1 goto :err

echo.
echo ============================================================
echo 3/4 Enable local managed parse server
echo ============================================================
mineru config set parse_server.local.mode managed
if errorlevel 1 goto :err

echo.
echo ============================================================
echo 4/4 Restart MinerU server and show status
echo ============================================================
mineru server stop
mineru server start
mineru server status

echo.
echo If local Parse Server is healthy and tier=standard, run:
echo   python convert_to_md.py --pdf-engine auto --mineru-tier standard --overwrite
pause
exit /b 0

:err
echo.
echo [ERROR] MinerU local model setup failed. Check the error above.
echo Model source: %MINERU_MODEL_SOURCE%
pause
exit /b 1
