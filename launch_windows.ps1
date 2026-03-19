# ============================================================================
# Karyotype Report Generator - Windows PowerShell Launch Script
# Anderson Diagnostics | Peripheral Blood Karyotyping
# ============================================================================

Write-Host ""
Write-Host "========================================================"
Write-Host "  Karyotype Report Generator"
Write-Host "  Anderson Diagnostics"
Write-Host "========================================================"
Write-Host ""

# Check if Python is installed
try {
    $pythonVersion = python --version 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw "Python not found"
    }
    Write-Host "[INFO] Python found: $pythonVersion" -ForegroundColor Green
    Write-Host ""
}
catch {
    Write-Host "[ERROR] Python is not installed or not in PATH." -ForegroundColor Red
    Write-Host "Please install Python 3.8+ from https://www.python.org/downloads/" -ForegroundColor Yellow
    Write-Host "Make sure to check 'Add Python to PATH' during installation." -ForegroundColor Yellow
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit 1
}

# Check if virtual environment exists
if (-Not (Test-Path "venv\Scripts\Activate.ps1")) {
    Write-Host "[INFO] Virtual environment not found. Creating one..." -ForegroundColor Yellow
    python -m venv venv
    if ($LASTEXITCODE -ne 0) {
        Write-Host "[ERROR] Failed to create virtual environment." -ForegroundColor Red
        Read-Host "Press Enter to exit"
        exit 1
    }
    Write-Host "[SUCCESS] Virtual environment created." -ForegroundColor Green
    Write-Host ""
}

# Activate virtual environment
Write-Host "[INFO] Activating virtual environment..." -ForegroundColor Cyan
try {
    & "venv\Scripts\Activate.ps1"
}
catch {
    Write-Host "[ERROR] Failed to activate virtual environment." -ForegroundColor Red
    Write-Host "You may need to enable script execution:" -ForegroundColor Yellow
    Write-Host "  Run PowerShell as Administrator and execute:" -ForegroundColor Yellow
    Write-Host "  Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser" -ForegroundColor Yellow
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit 1
}

# Check if requirements are installed
Write-Host "[INFO] Checking dependencies..." -ForegroundColor Cyan
python -c "import PyQt6" 2>&1 | Out-Null
if ($LASTEXITCODE -ne 0) {
    Write-Host "[INFO] Dependencies not found. Installing from requirements.txt..." -ForegroundColor Yellow
    if (-Not (Test-Path "requirements.txt")) {
        Write-Host "[ERROR] requirements.txt not found." -ForegroundColor Red
        Read-Host "Press Enter to exit"
        exit 1
    }
    python -m pip install --upgrade pip
    pip install -r requirements.txt
    if ($LASTEXITCODE -ne 0) {
        Write-Host "[ERROR] Failed to install dependencies." -ForegroundColor Red
        Read-Host "Press Enter to exit"
        exit 1
    }
    Write-Host "[SUCCESS] Dependencies installed." -ForegroundColor Green
    Write-Host ""
}
else {
    Write-Host "[SUCCESS] Dependencies already installed." -ForegroundColor Green
    Write-Host ""
}

# Launch the application
Write-Host "[INFO] Launching Karyotype Report Generator..." -ForegroundColor Cyan
Write-Host ""
python karyotype_report_generator.py

# Deactivate virtual environment on exit
deactivate

Write-Host ""
Write-Host "[INFO] Application closed." -ForegroundColor Green
Read-Host "Press Enter to exit"
