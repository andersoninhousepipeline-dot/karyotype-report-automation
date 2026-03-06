#!/bin/bash
echo "=============================================="
echo " Karyotype Report Generator - First-Time Setup"
echo "=============================================="
echo

# Check Python3
if ! command -v python3 &>/dev/null; then
    echo "[ERROR] Python3 not found."
    echo
    echo "Install it with:"
    echo "  Ubuntu/Debian : sudo apt install python3 python3-pip python3-venv"
    echo "  Fedora/RHEL   : sudo dnf install python3"
    echo "  Arch          : sudo pacman -S python"
    exit 1
fi

# Check version >= 3.10
python3 -c "import sys; exit(0 if sys.version_info >= (3,10) else 1)"
if [ $? -ne 0 ]; then
    echo "[ERROR] Python 3.10 or higher is required."
    echo "Your version: $(python3 --version)"
    exit 1
fi

echo "[OK] $(python3 --version)"
echo

# Create venv
if [ ! -d "venv" ]; then
    echo "Creating virtual environment..."
    python3 -m venv venv
    if [ $? -ne 0 ]; then
        echo "[ERROR] Failed to create virtual environment."
        echo "Try: sudo apt install python3-venv"
        exit 1
    fi
    echo "[OK] Virtual environment created."
else
    echo "[OK] Virtual environment already exists."
fi

echo
echo "Installing dependencies (this may take a few minutes)..."
venv/bin/pip install --upgrade pip --quiet
venv/bin/pip install -r requirements.txt

if [ $? -ne 0 ]; then
    echo
    echo "[ERROR] Failed to install one or more dependencies."
    echo "Check your internet connection and try again."
    exit 1
fi

# Make launcher executable
chmod +x launch_karyotype.sh

echo
echo "=============================================="
echo " Setup complete!"
echo " Run: bash launch_karyotype.sh"
echo "=============================================="
echo
