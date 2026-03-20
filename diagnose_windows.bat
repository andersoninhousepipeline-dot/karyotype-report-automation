@echo off
REM ============================================================================
REM Windows Diagnostic Script - Karyotype Report Generator
REM ============================================================================
REM This script collects diagnostic information to identify why corrections
REM are not being applied on Windows systems.
REM ============================================================================

setlocal enabledelayedexpansion

echo.
echo ========================================================================
echo   Karyotype Report Generator - Windows Diagnostic Tool
echo ========================================================================
echo.
echo This script will collect diagnostic information to help troubleshoot
echo why PDF corrections are not being applied on Windows.
echo.
pause

REM Create output file
set OUTPUT=diagnostic_report.txt
echo Diagnostic Report - Generated on %date% %time% > %OUTPUT%
echo ======================================================================== >> %OUTPUT%
echo. >> %OUTPUT%

echo [1/8] Checking Git status...
echo. >> %OUTPUT%
echo ======== GIT STATUS ======== >> %OUTPUT%
git --version >> %OUTPUT% 2>&1
echo. >> %OUTPUT%
git log -1 --oneline >> %OUTPUT% 2>&1
echo. >> %OUTPUT%
echo Current branch: >> %OUTPUT%
git branch --show-current >> %OUTPUT% 2>&1
echo. >> %OUTPUT%

echo [2/8] Checking Python version...
echo. >> %OUTPUT%
echo ======== PYTHON VERSION ======== >> %OUTPUT%
python --version >> %OUTPUT% 2>&1
echo. >> %OUTPUT%

echo [3/8] Checking if virtual environment exists...
echo. >> %OUTPUT%
echo ======== VIRTUAL ENVIRONMENT ======== >> %OUTPUT%
if exist "venv\Scripts\python.exe" (
    echo Virtual environment found >> %OUTPUT%
    echo. >> %OUTPUT%
    call venv\Scripts\activate
    echo Activated virtual environment >> %OUTPUT%
) else (
    echo WARNING: Virtual environment NOT found! >> %OUTPUT%
    echo This may cause issues. >> %OUTPUT%
)
echo. >> %OUTPUT%

echo [4/8] Checking installed packages...
echo. >> %OUTPUT%
echo ======== INSTALLED PACKAGES ======== >> %OUTPUT%
pip list >> %OUTPUT% 2>&1
echo. >> %OUTPUT%

echo [5/8] Checking critical files...
echo. >> %OUTPUT%
echo ======== FILE INFORMATION ======== >> %OUTPUT%
echo. >> %OUTPUT%
echo karyotype_template.py: >> %OUTPUT%
dir karyotype_template.py >> %OUTPUT% 2>&1
echo. >> %OUTPUT%
echo karyotype_report_generator.py: >> %OUTPUT%
dir karyotype_report_generator.py >> %OUTPUT% 2>&1
echo. >> %OUTPUT%

echo [6/8] Checking for critical methods in code...
echo. >> %OUTPUT%
echo ======== CODE VERIFICATION ======== >> %OUTPUT%
echo. >> %OUTPUT%
echo Checking for _iscn_color method: >> %OUTPUT%
findstr /C:"def _iscn_color" karyotype_template.py >> %OUTPUT% 2>&1
if errorlevel 1 (
    echo WARNING: _iscn_color method NOT FOUND! >> %OUTPUT%
    echo This means the code is OUTDATED. >> %OUTPUT%
) else (
    echo Found: _iscn_color method exists. >> %OUTPUT%
)
echo. >> %OUTPUT%

echo Checking for _image_has_border method: >> %OUTPUT%
findstr /C:"def _image_has_border" karyotype_template.py >> %OUTPUT% 2>&1
if errorlevel 1 (
    echo WARNING: _image_has_border method NOT FOUND! >> %OUTPUT%
    echo This means the code is OUTDATED. >> %OUTPUT%
) else (
    echo Found: _image_has_border method exists. >> %OUTPUT%
)
echo. >> %OUTPUT%

echo Checking for GREEN color definition: >> %OUTPUT%
findstr /C:"GREEN" karyotype_template.py >> %OUTPUT% 2>&1
echo. >> %OUTPUT%

echo [7/8] Running automated tests...
echo. >> %OUTPUT%
echo ======== AUTOMATED TEST RESULTS ======== >> %OUTPUT%
echo. >> %OUTPUT%

echo Running test_generation_normal.py... >> %OUTPUT%
python test_generation_normal.py >> %OUTPUT% 2>&1
if errorlevel 1 (
    echo FAILED: Normal report test failed! >> %OUTPUT%
) else (
    echo PASSED: Normal report test succeeded. >> %OUTPUT%
)
echo. >> %OUTPUT%

echo Running test_generation_abnormal.py... >> %OUTPUT%
python test_generation_abnormal.py >> %OUTPUT% 2>&1
if errorlevel 1 (
    echo FAILED: Abnormal report test failed! >> %OUTPUT%
) else (
    echo PASSED: Abnormal report test succeeded. >> %OUTPUT%
)
echo. >> %OUTPUT%

echo Running test_generation_variant.py... >> %OUTPUT%
python test_generation_variant.py >> %OUTPUT% 2>&1
if errorlevel 1 (
    echo FAILED: Variant report test failed! >> %OUTPUT%
) else (
    echo PASSED: Variant report test succeeded. >> %OUTPUT%
)
echo. >> %OUTPUT%

echo [8/8] Collecting system information...
echo. >> %OUTPUT%
echo ======== SYSTEM INFORMATION ======== >> %OUTPUT%
systeminfo | findstr /C:"OS Name" /C:"OS Version" /C:"System Type" >> %OUTPUT% 2>&1
echo. >> %OUTPUT%

echo. >> %OUTPUT%
echo ======== DIAGNOSTIC COMPLETE ======== >> %OUTPUT%
echo Report saved to: %OUTPUT% >> %OUTPUT%
echo. >> %OUTPUT%

REM Display summary on screen
echo.
echo ========================================================================
echo   Diagnostic Complete
echo ========================================================================
echo.
echo A detailed report has been saved to: %OUTPUT%
echo.
echo Please review the report and look for:
echo   1. Git commit hash (should be: 67cff75)
echo   2. _iscn_color method (should be found)
echo   3. _image_has_border method (should be found)
echo   4. Test results (should all pass)
echo.
echo If any tests FAILED or methods are NOT FOUND:
echo   ^> Your Windows installation has OUTDATED code
echo   ^> Run: git pull
echo   ^> Or manually copy files from Linux
echo.
pause

REM Open the report in notepad
notepad %OUTPUT%
