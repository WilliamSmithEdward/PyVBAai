@echo off
REM ============================================================================
REM  PyVBAai build script
REM  Produces:  dist\PyVBAai.exe  (single standalone .exe, no Python required)
REM ============================================================================

setlocal

echo [1/5] Checking Python...
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found. Install Python 3.11+ and add it to PATH.
    pause & exit /b 1
)

echo [2/5] Setting up virtual environment...
if not exist .venv (
    python -m venv .venv
)

echo [3/5] Installing / upgrading dependencies...
.venv\Scripts\python.exe -m pip install --upgrade pip --quiet
.venv\Scripts\pip install -r requirements.txt --quiet

echo [4/5] Building executable with PyInstaller...
.venv\Scripts\python.exe -m PyInstaller PyVBAai.spec --clean --noconfirm

echo [5/5] Done!
if exist dist\PyVBAai.exe (
    echo.
    echo  SUCCESS: dist\PyVBAai.exe created.
    echo.
    echo  Before running, make sure OPENAI_API_KEY is set as a user
    echo  environment variable in Windows Settings.
) else (
    echo  Build may have failed - check output above.
)

pause
