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

echo Installing MinerU. This may download several GB (PyTorch + models).
echo NVIDIA GPU users can get CUDA acceleration, but CPU is supported.
echo.
pip install -U -r requirements_mineru.txt
if errorlevel 1 (
  echo [ERROR] MinerU installation failed. Please check the error above.
  echo Official docs: https://github.com/opendatalab/MinerU
  pause
  exit /b 1
)

echo.
echo MinerU installed. Check:
mineru --help
pause
exit /b 0

:err
echo [ERROR] Failed to prepare the Python virtual environment.
pause
exit /b 1
