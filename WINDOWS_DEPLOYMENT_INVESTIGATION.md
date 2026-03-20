# Windows Deployment Investigation Plan
## Karyotype Report Generator - Linux vs Windows Discrepancy Analysis

---

## Problem Summary

**Symptom**: PDF reports generated on Linux have all corrections applied correctly, but the same code on Windows produces PDFs with the following issues:

1. ❌ **Missing borders** on some files
2. ❌ **Comment box still present** in all files (should be hidden for normal reports)
3. ❌ **Alignment issues** (misalignment of elements)
4. ❌ **Font colors not changed**:
   - Green for Normal reports
   - Red for Abnormal reports
   - Black for Variant reports
5. ❌ **Image sizes not uniform** for abnormal reports

---

## Corrections Applied (Working on Linux, NOT on Windows)

Based on git history (commits: a0f2452 → 67cff75), the following fixes were implemented:

### 1. Font Color Logic (`_iscn_color()` method)
**File**: [karyotype_template.py:305-314](karyotype_template.py#L305-L314)
```python
def _iscn_color(self):
    """Return ISCN text color based on autosome/sex chromosome values.
    Both Normal → GREEN; either Abnormal → RED; either Variant → BLACK."""
    auto = self._get("AUTOSOME").lower()
    sex  = self._get("SEX CHROMOSOME", "SEX CHROMOSOME ").lower()
    if "abnormal" in auto or "abnormal" in sex:
        return RED
    if "variant" in auto or "variant" in sex:
        return BLACK
    return GREEN  # both normal
```

### 2. Border Detection (`_image_has_border()` method)
**File**: [karyotype_template.py:670-699](karyotype_template.py#L670-L699)
```python
@staticmethod
def _image_has_border(path: str) -> bool:
    """Return True if the image already has a built-in dark frame.
    Scans insets 0-15px from each edge and takes the minimum average
    brightness found. Built-in borders appear at different inset depths
    across images, so scanning a range catches them all reliably.
    Threshold 230: no-border images stay at 255; framed images drop to ~130-193."""
```

### 3. Smart Border Drawing in `_place_image()`
**File**: [karyotype_template.py:735-739](karyotype_template.py#L735-L739)
```python
# Draw border only if the image doesn't already have its own frame
if not self._image_has_border(path):
    c.setStrokeColor(BLACK)
    c.setLineWidth(1.0)
    c.rect(cx, cy, dw, dh, fill=0, stroke=1)
```

### 4. Alignment Fixes (Spacing Adjustments)
Multiple spacing changes from 20pt gaps to 12pt gaps:
- [karyotype_template.py:474-478](karyotype_template.py#L474-L478) - Page 2 Normal layout
- [karyotype_template.py:508-510](karyotype_template.py#L508-L510) - Page 2 Abnormal layout
- [karyotype_template.py:816-824](karyotype_template.py#L816-L824) - Recommendations block

### 5. Comments Box Logic (3-page vs 2-page layout)
**File**: [karyotype_template.py:301-303](karyotype_template.py#L301-L303)
```python
# Determine layout variant
has_comments = bool(self._get("COMMENTS"))
self.three_page = has_comments   # True → 3-page layout
```
- If `COMMENTS` field is **empty** → 2-page layout (no comments section rendered)
- If `COMMENTS` field has **content** → 3-page layout (comments section on page 2)

### 6. Translocation Image Alignment
**File**: [karyotype_template.py:637-668](karyotype_template.py#L637-L668)
- Scatter image top-aligned
- Zoom image positioned just below scatter with 4pt gap
- Karyogram top-aligned with scatter

---

## Root Cause Analysis: Why Windows Doesn't Get Updates

### Hypothesis 1: **Code Not Deployed to Windows** ⭐ MOST LIKELY
**Symptoms match**: All corrections missing
**Verification Steps**:
1. Check if `karyotype_template.py` on Windows has the latest changes
2. Compare file modification dates between Linux and Windows
3. Verify git commit hash on Windows machine

**Commands to run on Windows**:
```cmd
# Check current git commit
git log -1 --oneline

# Check if file has _iscn_color method
findstr "_iscn_color" karyotype_template.py

# Check if file has _image_has_border method
findstr "_image_has_border" karyotype_template.py

# Show file modification date
dir karyotype_template.py
```

**Expected on Linux** (current):
```bash
git log -1 --oneline
# Output: 67cff75 updated alignment
```

**Expected on Windows** (if outdated):
```bash
git log -1 --oneline
# Output: 97781e7 or earlier commit
```

---

### Hypothesis 2: **Virtual Environment Not Rebuilt**
**Symptoms match**: Partial - may cause import issues
**Verification Steps**:
1. Check if Windows `venv/` folder has the latest dependencies
2. Verify PIL/Pillow version (needed for `_image_has_border()`)

**Commands to run on Windows**:
```cmd
venv\Scripts\activate
pip list | findstr -i "pillow reportlab"
```

**Expected**:
- Pillow >= 9.0.0
- reportlab >= 3.6.0

---

### Hypothesis 3: **PyInstaller Frozen Executable (Outdated Build)**
**Symptoms match**: All corrections missing
**Verification Steps**:
1. Check if Windows users are running a `.exe` built with PyInstaller
2. Check build date of the `.exe`

**Commands to run on Windows**:
```cmd
# If using built executable
dir dist\KaryotypeReportGen\KaryotypeReportGen.exe

# Check what's inside _MEIPASS (temp extraction folder when running)
# This is harder - need to run the exe and check sys._MEIPASS at runtime
```

**Fix if this is the issue**:
Rebuild the Windows executable:
```cmd
build_windows.bat
```

---

### Hypothesis 4: **Python Path Issues (Running Old Code)**
**Symptoms match**: All corrections missing
**Verification Steps**:
1. Check which `karyotype_template.py` is being imported
2. Verify `PYTHONPATH` doesn't point to an old copy

**Commands to run on Windows**:
```cmd
python -c "import karyotype_template; print(karyotype_template.__file__)"
```

**Expected**:
Should point to the current project directory, not some cached/installed location.

---

### Hypothesis 5: **Cache/Bytecode Issues (.pyc files)**
**Symptoms match**: Partial - unlikely but possible
**Verification Steps**:
1. Check for stale `.pyc` files
2. Clear Python cache

**Commands to run on Windows**:
```cmd
# Delete all .pyc files
del /s /q __pycache__
del /s /q *.pyc

# Restart application
python karyotype_report_generator.py
```

---

## Step-by-Step Investigation Plan

### Phase 1: Verify Code Deployment ⭐ START HERE

**Step 1A: Check Git Status on Windows**
```cmd
cd "C:\Path\To\Karyotyping-Report"
git status
git log -1 --oneline
git diff HEAD karyotype_template.py
```

**Expected Outcome**:
- Should show commit `67cff75 updated alignment`
- `git diff` should show **no changes** (meaning file is up-to-date)

**If git shows old commit or differences**:
```cmd
# Pull latest changes
git pull origin main

# Or if no git remote configured
git fetch
git checkout main
git pull
```

---

**Step 1B: Verify Code Files on Windows**
```cmd
# Check if _iscn_color method exists
findstr /C:"def _iscn_color" karyotype_template.py

# Check if _image_has_border method exists
findstr /C:"def _image_has_border" karyotype_template.py

# Check if GREEN color is defined
findstr /C:"GREEN" karyotype_template.py
```

**Expected Output**:
```
def _iscn_color(self):
def _image_has_border(path: str) -> bool:
GREEN      = HexColor('#00B050')
```

**If methods are MISSING**:
→ **Code is NOT deployed** - Need to copy/pull latest files

---

### Phase 2: Verify Runtime Environment

**Step 2A: Check Python Packages**
```cmd
venv\Scripts\activate
pip list
```

**Required packages**:
- PyQt6
- reportlab >= 3.6.0
- Pillow >= 9.0.0
- pandas
- openpyxl
- pypdfium2 (optional, for preview)

**If packages are outdated**:
```cmd
pip install --upgrade -r requirements.txt
```

---

**Step 2B: Test Image Border Detection**
Create a test script on Windows: `test_border_detection.py`
```python
from karyotype_template import KaryotypeReportGenerator
import os

# Point to an actual image file
img_path = r"C:\Path\To\Images\260154818.jpg"

if os.path.exists(img_path):
    has_border = KaryotypeReportGenerator._image_has_border(img_path)
    print(f"Image: {img_path}")
    print(f"Has built-in border: {has_border}")
else:
    print(f"Image not found: {img_path}")
```

**Run**:
```cmd
python test_border_detection.py
```

**Expected**: Should print `True` or `False` based on image analysis.
**If error**: PIL/Pillow not installed or method doesn't exist.

---

### Phase 3: Test PDF Generation

**Step 3A: Generate Test Report (Normal)**
```cmd
python test_generation_normal.py
```

Create `test_generation_normal.py`:
```python
from karyotype_template import KaryotypeReportGenerator
import os

data = {
    "NAME": "Test Patient",
    "GENDER": "Male",
    "AGE": "25 Years",
    "SPECIMEN": "Peripheral blood",
    "PIN": "TEST001",
    "SAMPLE NUMBER": "TEST123",
    "RESULT": "46,XY",
    "AUTOSOME": "Normal",
    "SEX CHROMOSOME": "Normal",
    "INTERPRETATION": "Karyotype shows an apparently normal male.",
    "COMMENTS": "",  # Empty → should be 2-page
    "RECOMMENDATIONS": "Genetic counseling recommended."
}

output_dir = "test_output"
os.makedirs(output_dir, exist_ok=True)

gen = KaryotypeReportGenerator(data, [], output_dir)
pdf_path = gen.generate()

print(f"Generated: {pdf_path}")
print(f"Layout: {'3-page' if gen.three_page else '2-page'} (expected: 2-page)")

# Check ISCN color
color = gen._iscn_color()
print(f"ISCN Color: {color} (expected: GREEN)")
```

**Expected**:
- 2-page PDF generated
- ISCN text in **GREEN** (#00B050)
- No comments section

**If ISCN is RED**: Color logic not working → code not updated

---

**Step 3B: Generate Test Report (Abnormal)**
Create `test_generation_abnormal.py`:
```python
from karyotype_template import KaryotypeReportGenerator
import os

data = {
    "NAME": "Test Patient Abnormal",
    "GENDER": "Male",
    "AGE": "30 Years",
    "SPECIMEN": "Peripheral blood",
    "PIN": "TEST002",
    "SAMPLE NUMBER": "TEST456",
    "RESULT": "47,XY,+21",
    "AUTOSOME": "Abnormal",
    "SEX CHROMOSOME": "Normal",
    "INTERPRETATION": "Karyotype shows Trisomy 21 (Down Syndrome).",
    "COMMENTS": "Trisomy 21 is associated with developmental delays.",  # Has content → 3-page
    "RECOMMENDATIONS": "Genetic counseling advised."
}

output_dir = "test_output"
os.makedirs(output_dir, exist_ok=True)

gen = KaryotypeReportGenerator(data, [], output_dir)
pdf_path = gen.generate()

print(f"Generated: {pdf_path}")
print(f"Layout: {'3-page' if gen.three_page else '2-page'} (expected: 3-page)")

color = gen._iscn_color()
print(f"ISCN Color: {color} (expected: RED)")
```

**Expected**:
- 3-page PDF generated
- ISCN text in **RED**
- Comments section visible on page 2

**If Comments still hidden**: `three_page` logic not working

---

### Phase 4: Check for Frozen Executable Issues

**Step 4A: Determine How Application is Launched**
Ask the Windows user:
- Are you running `python karyotype_report_generator.py`? (source)
- Or `KaryotypeReportGen.exe`? (frozen executable)
- Or `launch_windows.bat`? (batch script → source)

**If using `.exe`**:
```cmd
# Check exe build date
dir dist\KaryotypeReportGen\KaryotypeReportGen.exe

# Rebuild with latest code
build_windows.bat
```

---

### Phase 5: Compare Linux vs Windows Output

**Step 5A: Generate Same Report on Both Systems**
Use identical test data (from Step 3A) on both Linux and Windows.

**Compare**:
1. File sizes (should be similar)
2. Page count (2-page vs 3-page)
3. Open both PDFs side-by-side and visually inspect:
   - ISCN text color
   - Comments section presence
   - Image borders
   - Alignment/spacing

**Tools**:
- [pdfplumber](https://github.com/jsvine/pdfplumber) to extract text/layout data
- [diffpdf](https://gitlab.com/eang/diffpdf) to compare PDFs visually

---

## Diagnostic Checklist for Windows Client

Send this to your Windows client to run and report back:

### 1️⃣ **Check Code Version**
```cmd
cd C:\Path\To\Karyotyping-Report
git log -1 --oneline
```
📋 **Report**: What commit hash is shown?

---

### 2️⃣ **Check if Methods Exist**
```cmd
findstr /C:"def _iscn_color" karyotype_template.py
findstr /C:"def _image_has_border" karyotype_template.py
```
📋 **Report**: Do both lines appear?

---

### 3️⃣ **Check Python Environment**
```cmd
venv\Scripts\activate
python --version
pip list | findstr -i "reportlab pillow"
```
📋 **Report**: Python version and package versions

---

### 4️⃣ **Run Test Generation**
```cmd
# Copy test_generation_normal.py from Phase 3A above
python test_generation_normal.py
```
📋 **Report**:
- Does it run without errors?
- What layout is generated (2-page or 3-page)?
- What color is the ISCN text in the PDF?

---

### 5️⃣ **Check File Modification Dates**
```cmd
dir karyotype_template.py
dir karyotype_report_generator.py
```
📋 **Report**: When were these files last modified?

---

### 6️⃣ **Check How App is Launched**
📋 **Report**:
- Are you double-clicking `launch_windows.bat`?
- Or running a `.exe` file?
- Or manually typing `python karyotype_report_generator.py`?

---

## Quick Fixes (Ranked by Likelihood)

### Fix #1: Update Code on Windows ⭐⭐⭐⭐⭐
```cmd
# Navigate to project folder
cd C:\Path\To\Karyotyping-Report

# Pull latest changes
git pull

# OR if no git, copy files manually from Linux
# Copy: karyotype_template.py, karyotype_report_generator.py
```

---

### Fix #2: Reinstall Dependencies
```cmd
venv\Scripts\activate
pip install --upgrade -r requirements.txt
```

---

### Fix #3: Clear Python Cache
```cmd
del /s /q __pycache__
del /s /q *.pyc
python karyotype_report_generator.py
```

---

### Fix #4: Rebuild Executable (if using .exe)
```cmd
build_windows.bat
```

---

## Expected Behavior After Fix

✅ **Normal Reports (46,XY or 46,XX)**:
- ISCN text: **GREEN** (#00B050)
- Layout: **2-page** (no comments section)
- Images: Single black border (if image has no built-in border)

✅ **Abnormal Reports (Trisomy, Translocation, etc.)**:
- ISCN text: **RED**
- Layout: **3-page** (with comments section on page 2)
- Images: Proper alignment, no double borders

✅ **Variant Reports**:
- ISCN text: **BLACK**
- Layout: **3-page**

---

## Files to Verify on Windows

Essential files that MUST match Linux version:

1. ✅ `karyotype_template.py` (contains all PDF generation logic)
2. ✅ `karyotype_report_generator.py` (GUI application)
3. ✅ `karyotype_assets.py` (base64-encoded images)
4. ✅ `requirements.txt` (dependency list)
5. ✅ `launch_windows.bat` (launch script)

**Compare checksums** (if available):
```bash
# On Linux
sha256sum karyotype_template.py

# On Windows (PowerShell)
Get-FileHash karyotype_template.py -Algorithm SHA256
```

---

## Next Steps

1. **Send diagnostic checklist** (Section above) to Windows client
2. **Await their responses** to determine exact issue
3. **Apply appropriate fix** based on findings
4. **Test with sample data** on Windows
5. **Generate corrected PDFs** and verify all issues resolved

---

## Contact & Support

If issues persist after applying fixes, collect the following for further analysis:
- Git commit hash on Windows: `git log -1`
- Python version: `python --version`
- Package versions: `pip list`
- Sample PDF output from Windows (attach file)
- Screenshot of any error messages

---

## Summary

**Most Likely Issue**: 🎯 **Code not deployed to Windows**

**Evidence**:
- ALL corrections missing (borders, colors, alignment, comments logic)
- Linux has latest code (commit 67cff75)
- Windows behavior matches older codebase

**Primary Fix**: Pull/copy latest `karyotype_template.py` to Windows machine

**Verification**: Run test generation scripts (Phase 3) and confirm GREEN/RED colors work

---

**Document Version**: 1.0
**Last Updated**: 2026-03-20
**Author**: Claude Code Investigation
