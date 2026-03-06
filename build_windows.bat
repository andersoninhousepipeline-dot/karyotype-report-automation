@echo off
REM Build Karyotype Report Generator for Windows
REM Usage: Double-click or run from Command Prompt

echo === Karyotype Report Generator - Windows Build ===
cd /d "%~dp0"

REM 1. Install dependencies
echo Installing Python dependencies...
pip install -r requirements.txt
pip install pyinstaller

REM 2. Clean previous build
if exist build   rmdir /s /q build
if exist dist    rmdir /s /q dist

REM 3. Build with PyInstaller
echo Building with PyInstaller...
pyinstaller KaryotypeReport.spec

echo.
echo === Build complete ===
echo Executable: dist\KaryotypeReportGen\KaryotypeReportGen.exe
echo.
pause
