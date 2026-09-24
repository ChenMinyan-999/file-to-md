@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

set "PYTHON="
where py >nul 2>nul
if not errorlevel 1 set "PYTHON=py -3"
if not defined PYTHON (
  where python >nul 2>nul
  if not errorlevel 1 set "PYTHON=python"
)
if not defined PYTHON (
  echo [ERROR] Python 3.10+ not found.
  echo Install it from https://www.python.org/downloads/windows/ and check "Add Python to PATH".
  pause
  exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
  echo [1/3] Creating virtual environment .venv ...
  %PYTHON% -m venv .venv
  if errorlevel 1 goto :err
)

call ".venv\Scripts\activate.bat"

echo [2/3] Installing/updating base dependencies ...
python -m pip install --upgrade pip
pip install -r requirements.txt
if errorlevel 1 goto :err

echo [3/3] Converting ...
python convert_to_md.py --source "%~dp0标准2" --output "%~dp0标准2_md" --pdf-engine auto --strict
set "RC=%ERRORLEVEL%"
echo.
if "%RC%"=="0" echo [OK] All converted.
if "%RC%"=="2" echo [REVIEW] Some files need manual review. See "标准2_md\_qa\review_index.html".
if "%RC%"=="1" echo [FAILED] Some files failed. See "标准2_md\_qa\qa_report.csv" and log.
pause
exit /b %RC%

:err
echo [ERROR] Command failed with code %ERRORLEVEL%.
pause
exit /b 1
