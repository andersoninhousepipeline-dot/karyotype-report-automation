#!/usr/bin/env bash
# Build Karyotype Report Generator for Linux
# Usage: bash build_linux.sh

set -e
cd "$(dirname "$0")"

echo "=== Karyotype Report Generator — Linux Build ==="

# 1. Install dependencies
echo "Installing Python dependencies..."
pip install -r requirements.txt
pip install pyinstaller

# 2. Clean previous build
rm -rf build dist

# 3. Build with PyInstaller
echo "Building with PyInstaller..."
pyinstaller KaryotypeReport.spec

echo ""
echo "=== Build complete ==="
echo "Executable: dist/KaryotypeReportGen/KaryotypeReportGen"
echo ""
echo "To run:  ./dist/KaryotypeReportGen/KaryotypeReportGen"
