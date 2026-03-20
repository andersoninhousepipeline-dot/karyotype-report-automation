# Windows Fix Instructions - Quick Guide
## Karyotype Report Generator

---

## 🔴 Problem

PDFs generated on Windows have **missing corrections**:
- ❌ Wrong font colors (should be Green/Red/Black based on report type)
- ❌ Comment box still visible in normal reports
- ❌ Alignment issues
- ❌ Border issues on images

**Root Cause**: Windows installation has **outdated code** (not synced with Linux)

---

## ✅ Solution: 3-Step Fix

### Step 1️⃣: Run Diagnostic (5 minutes)

Open Command Prompt in the project folder and run:

```cmd
diagnose_windows.bat
```

This will:
- Check your git commit version
- Verify if critical methods exist
- Run automated tests
- Generate a `diagnostic_report.txt`

**What to look for**:
- Git commit should be: `67cff75 updated alignment`
- Methods should be found: `_iscn_color`, `_image_has_border`
- Tests should **PASS**

**If any FAILED** → Continue to Step 2

---

### Step 2️⃣: Update Code (2 methods)

#### Method A: Using Git (Recommended)

```cmd
cd C:\Path\To\Karyotyping-Report
git pull origin main
```

**If git pull fails** (no remote configured):
```cmd
git status
git diff HEAD karyotype_template.py
```

If you see differences, the code is outdated. Manually copy files from Linux (Method B).

---

#### Method B: Manual File Copy

Copy these files **from Linux to Windows** (overwrite existing):

1. ✅ `karyotype_template.py` ← **MOST IMPORTANT**
2. ✅ `karyotype_report_generator.py`
3. ✅ `karyotype_assets.py`
4. ✅ `requirements.txt`

**How to transfer**:
- USB drive
- Network share
- Email/Cloud (zip the files)
- SCP/SFTP if SSH is available

---

### Step 3️⃣: Reinstall Dependencies & Test

```cmd
cd C:\Path\To\Karyotyping-Report

REM Clear Python cache
del /s /q __pycache__
del /s /q *.pyc

REM Activate virtual environment
call venv\Scripts\activate

REM Update dependencies
pip install --upgrade -r requirements.txt

REM Run test to verify fix
python test_generation_normal.py
python test_generation_abnormal.py
```

**Expected output**:
```
✓ Method _iscn_color() exists (correction applied)
✓ Method _image_has_border() exists (correction applied)
✓ Layout: 2-page (correct for normal report)
✓ ISCN color is GREEN (correct for Normal/Normal)
✓ ALL TESTS PASSED
```

---

## 🎯 Verification Checklist

After applying the fix, verify these in generated PDFs:

### ✅ Normal Reports (46,XY or 46,XX)
- [ ] ISCN text is **GREEN** (#00B050)
- [ ] Only **2 pages** (no page 3)
- [ ] **No "Comments" section** visible
- [ ] Proper alignment and spacing

### ✅ Abnormal Reports (Trisomy, etc.)
- [ ] ISCN text is **RED**
- [ ] **3 pages** total
- [ ] **"Comments" section visible** on page 2
- [ ] No double borders on images
- [ ] Images properly aligned

### ✅ Variant Reports
- [ ] ISCN text is **BLACK**
- [ ] **3 pages** total
- [ ] Comments section visible

---

## 🚨 If Fix Doesn't Work

### Issue: Git pull says "Already up to date" but tests still fail

**Cause**: Local changes conflicting with remote

**Fix**:
```cmd
git status
git diff karyotype_template.py
```

If you see local modifications:
```cmd
# Backup your changes
copy karyotype_template.py karyotype_template.py.backup

# Force pull
git fetch origin
git reset --hard origin/main
```

---

### Issue: Tests fail with "ImportError: No module named..."

**Cause**: Dependencies not installed

**Fix**:
```cmd
pip install --upgrade reportlab Pillow PyQt6 pandas openpyxl pypdfium2
```

---

### Issue: "_iscn_color() method NOT FOUND"

**Cause**: Code file not updated

**Fix**:
1. Verify you updated `karyotype_template.py` (not `karyotype_report_generator.py`)
2. Check file size: Should be ~31-32 KB
3. Open `karyotype_template.py` and search for `def _iscn_color`
4. If not found → file was not copied correctly, try again

---

### Issue: Using .exe file (PyInstaller frozen executable)

**Cause**: Executable built with old code

**Fix**: Rebuild the executable
```cmd
build_windows.bat
```

This will create a new `.exe` in `dist/` folder with the latest code.

---

## 📊 Quick File Verification

To verify files match between Linux and Windows:

**On Linux**:
```bash
sha256sum karyotype_template.py
```

**On Windows (PowerShell)**:
```powershell
Get-FileHash karyotype_template.py -Algorithm SHA256
```

**Expected hash** (for commit 67cff75):
```
SHA256: <hash should match between systems>
```

If hashes don't match → files are different → copy from Linux

---

## 📞 Still Having Issues?

Collect this information and send for further analysis:

1. **Diagnostic report**: `diagnostic_report.txt` (from Step 1)
2. **Git info**:
   ```cmd
   git log -1 --oneline > git_info.txt
   git diff HEAD karyotype_template.py >> git_info.txt
   ```
3. **Sample PDF**: Generate a test report and attach the PDF
4. **Screenshot**: Of any error messages

---

## 🎓 Understanding the Corrections

### Correction 1: Font Colors
- **Before**: All ISCN text was RED
- **After**:
  - Green: Both Autosome AND Sex Chromosome = Normal
  - Red: Either Autosome OR Sex Chromosome = Abnormal
  - Black: Variant Observed

### Correction 2: Comment Box Logic
- **Before**: Comments section always rendered (even if empty)
- **After**:
  - Empty comments → 2-page layout (no comments section)
  - Has comments → 3-page layout (with comments section)

### Correction 3: Image Borders
- **Before**: Always added gray border to all images
- **After**: Detects if image has built-in border, only adds border if missing

### Correction 4: Alignment/Spacing
- **Before**: Large gaps (20pt) between sections
- **After**: Tighter spacing (12pt) for better fit

---

## 📝 Summary

**Problem**: Windows has old code
**Solution**: Update files from Linux
**Verification**: Run test scripts
**Expected**: All tests pass, colors correct, layout correct

**Time required**: 10-15 minutes

---

**Last Updated**: 2026-03-20
**Diagnostic Tool**: `diagnose_windows.bat`
**Test Scripts**: `test_generation_*.py`
