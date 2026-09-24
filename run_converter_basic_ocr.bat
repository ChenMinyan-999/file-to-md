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
  pause
  exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
  echo Creating virtual environment .venv ...
  %PYTHON% -m venv .venv
  if errorlevel 1 goto :err
)
call ".venv\Scripts\activate.bat"

echo Installing basic OCR dependencies ...
python -m pip install --upgrade pip
pip install -r requirements_basic_ocr.txt
if errorlevel 1 goto :err

echo Converting with basic OCR fallback (RapidOCR). OCR output is REVIEW-only.
python convert_to_md.py --source "%~dp0标准2" --output "%~dp0标准2_md_basic_ocr" --pdf-engine auto --allow-basic-ocr
set "RC=%ERRORLEVEL%"
echo.
echo Done. Exit code: %RC%
echo See "标准2_md_basic_ocr\_qa\review_index.html" for manual review queue.
pause
exit /b %RC%

:err
echo [ERROR] Command failed with code %ERRORLEVEL%.
pause
exit /b 1
