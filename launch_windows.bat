@echo off
REM ============================================================================
REM Karyotype Report Generator - Windows Launch Script
REM Anderson Diagnostics | Peripheral Blood Karyotyping
REM ============================================================================

setlocal enabledelayedexpansion

echo.
echo ========================================================
echo   Karyotype Report Generator
echo   Anderson Diagnostics
echo ========================================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python is not installed or not in PATH.
    echo Please install Python 3.8+ from https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during installation.
    echo.
    pause
    exit /b 1
)

echo [INFO] Python found:
python --version
echo.

REM Check if virtual environment exists
if not exist "venv\Scripts\activate.bat" (
    echo [INFO] Virtual environment not found. Creating one...
    python -m venv venv
    if errorlevel 1 (
        echo [ERROR] Failed to create virtual environment.
        pause
        exit /b 1
    )
    echo [SUCCESS] Virtual environment created.
    echo.
)

REM Activate virtual environment
echo [INFO] Activating virtual environment...
call venv\Scripts\activate.bat
if errorlevel 1 (
    echo [ERROR] Failed to activate virtual environment.
    pause
    exit /b 1
)

REM Check if requirements are installed
echo [INFO] Checking dependencies...
python -c "import PyQt6" >nul 2>&1
if errorlevel 1 (
    echo [INFO] Dependencies not found. Installing from requirements.txt...
    if not exist "requirements.txt" (
        echo [ERROR] requirements.txt not found.
        pause
        exit /b 1
    )
    python -m pip install --upgrade pip
    pip install -r requirements.txt
    if errorlevel 1 (
        echo [ERROR] Failed to install dependencies.
        pause
        exit /b 1
    )
    echo [SUCCESS] Dependencies installed.
    echo.
) else (
    echo [SUCCESS] Dependencies already installed.
    echo.
)

REM Launch the application
echo [INFO] Launching Karyotype Report Generator...
echo.
python karyotype_report_generator.py

REM Deactivate virtual environment on exit
deactivate

echo.
echo [INFO] Application closed.
pause
