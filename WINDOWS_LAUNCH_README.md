# Windows Launch Instructions

## Karyotype Report Generator - Windows Setup & Launch Guide

This guide helps Windows users launch the Karyotype Report Generator application.

---

## Prerequisites

### 1. Install Python
- Download Python 3.8 or higher from [python.org/downloads](https://www.python.org/downloads/)
- **IMPORTANT**: During installation, check the box "Add Python to PATH"
- Verify installation by opening Command Prompt and typing: `python --version`

---

## Launch Methods

### Method 1: Using Batch Script (Recommended for Beginners)

1. **Double-click** `launch_windows.bat` in the project folder
2. The script will automatically:
   - Check if Python is installed
   - Create a virtual environment (first run only)
   - Install all required dependencies (first run only)
   - Launch the application

**First-time launch** may take 2-3 minutes to install dependencies.

---

### Method 2: Using PowerShell Script

1. **Right-click** `launch_windows.ps1`
2. Select **"Run with PowerShell"**

#### If you get an execution policy error:
1. Open PowerShell as Administrator
2. Run this command:
   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```
3. Type `Y` and press Enter
4. Try running the script again

---

### Method 3: Manual Launch (Command Line)

1. Open **Command Prompt** in the project folder
2. Run these commands:

```cmd
# Create virtual environment (first time only)
python -m venv venv

# Activate virtual environment
venv\Scripts\activate

# Install dependencies (first time only)
pip install -r requirements.txt

# Launch application
python karyotype_report_generator.py
```

---

## Troubleshooting

### "Python is not recognized"
- Python is not installed or not in PATH
- Reinstall Python with "Add Python to PATH" checked

### "Failed to create virtual environment"
- Make sure you have write permissions in the folder
- Try running Command Prompt as Administrator

### "Failed to install dependencies"
- Check your internet connection
- Try manually: `pip install -r requirements.txt`

### Application window doesn't appear
- Check if your antivirus is blocking it
- Look for error messages in the terminal window

### Image Auto-Discovery Not Working

If the application cannot find karyogram images by sample number:

**1. Test Image Discovery**
```cmd
python test_image_discovery.py "C:\Path\To\Images" 260154818
```
Replace with your actual image folder path and sample number.

**2. Check Image Filenames**
Images must be named exactly:
- `260154818.jpg` (exact match)
- `260154818 1.jpg` (with space and number)
- `260154818 2.jpg` (additional images)

**3. Supported File Extensions**
- `.jpg`, `.jpeg`, `.png` (case-insensitive)

**4. Common Issues**
- **Extra spaces**: `260154818  1.jpg` (two spaces) won't match
- **Underscores**: `260154818_1.jpg` won't match (use space)
- **Wrong extension**: `.tif`, `.bmp` not supported
- **Subdirectories**: Images must be in the selected folder (not subfolders)
- **File permissions**: Ensure the application can read the image folder

**5. Manual Workaround**
If auto-discovery fails, use the **"Edit / Add Images..."** button to manually select images for each patient.

---

## Quick Start After Installation

After the first successful launch:
- Simply **double-click** `launch_windows.bat` to start the application
- All dependencies are already installed
- Launch time will be much faster (5-10 seconds)

---

## Application Features

- **Manual Entry**: Enter patient data with live PDF preview
- **Bulk Upload**: Upload Excel file for batch processing
- **Image Auto-Discovery**: Automatically finds karyogram images by sample number
- **Draft Save/Load**: Save work in progress as JSON
- **PDF Generation**: High-quality reports with Anderson Diagnostics branding

---

## Support

For issues or questions about the application, contact your IT support or refer to the main documentation.
