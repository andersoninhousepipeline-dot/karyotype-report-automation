#!/bin/bash
cd "$(dirname "$0")"

# Use venv if available, otherwise fall back to miniconda / system Python
if [ -f "venv/bin/python3" ]; then
    PYTHON="venv/bin/python3"
elif [ -f "$HOME/miniconda3/bin/python3" ]; then
    PYTHON="$HOME/miniconda3/bin/python3"
elif [ -f "$HOME/anaconda3/bin/python3" ]; then
    PYTHON="$HOME/anaconda3/bin/python3"
elif command -v python3 &>/dev/null; then
    PYTHON="python3"
else
    echo "[ERROR] Python3 not found. Run install.sh first."
    read -p "Press Enter to exit..."
    exit 1
fi

$PYTHON karyotype_report_generator.py

if [ $? -ne 0 ]; then
    echo
    echo "The application exited with an error."
    echo "If this is your first run, execute: bash install.sh"
    read -p "Press Enter to exit..."
fi
