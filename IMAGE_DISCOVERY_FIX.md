# Image Auto-Discovery Fix for Windows

## Problem
Image auto-discovery was working on Linux but failing on Windows systems.

## Root Causes Identified

### 1. **glob.glob() Path Separator Issues**
The original code used `glob.glob(os.path.join(search_dir, ext))` which can behave inconsistently on Windows with mixed path separators.

### 2. **Case-Sensitivity Handling**
Windows file systems are case-insensitive, but the code was listing multiple extension patterns (`*.jpg`, `*.JPG`) which is redundant and can cause issues.

### 3. **Path Normalization**
Paths from QFileDialog on Windows may use forward slashes (`/`) while the file system expects backslashes (`\`), or vice versa.

### 4. **No Error Handling**
The original code didn't catch permission errors or invalid paths, causing silent failures.

---

## Solutions Implemented

### 1. **Replaced glob with Path.iterdir() (Line 234-285)**
**File**: [karyotype_report_generator.py:234-285](karyotype_report_generator.py#L234-L285)

**Changes**:
- Replaced `glob.glob()` with `Path.iterdir()` for consistent cross-platform behavior
- Used `Path` objects throughout for better Windows/Linux compatibility
- Added case-insensitive extension matching: `file_path.suffix.lower()`
- Added proper error handling for `OSError` and `PermissionError`
- Returns absolute paths using `file_path.absolute()`

**Before**:
```python
for ext in ("*.jpg", "*.jpeg", "*.JPG", "*.JPEG", "*.png", "*.PNG"):
    for fpath in glob.glob(os.path.join(search_dir, ext)):
        fname = os.path.splitext(os.path.basename(fpath))[0]
        if fname == sample_no or re.match(rf"^{re.escape(sample_no)}\s+\d+$", fname):
            results.append(fpath)
```

**After**:
```python
search_path = Path(search_dir)
for file_path in search_path.iterdir():
    if not file_path.is_file():
        continue
    if file_path.suffix.lower() not in [ext.lower() for ext in extensions]:
        continue
    fname = file_path.stem
    if fname == sample_no or re.match(rf"^{re.escape(sample_no)}\s+\d+$", fname):
        results.append(str(file_path.absolute()))
```

### 2. **Path Normalization on Save**
**Files Modified**:
- [karyotype_report_generator.py:805](karyotype_report_generator.py#L805) - Manual image directory
- [karyotype_report_generator.py:1351](karyotype_report_generator.py#L1351) - Bulk image directory

**Change**:
```python
self._image_search_dir = os.path.normpath(d)
```

This ensures paths are normalized to use the correct separator for the current OS.

### 3. **Path Normalization on Load**
**File Modified**: [karyotype_report_generator.py:1574](karyotype_report_generator.py#L1574)

**Change**:
```python
self._image_search_dir = os.path.normpath(img_dir)
```

Ensures paths loaded from QSettings are normalized for the current OS.

---

## Testing Tools Created

### 1. **Diagnostic Script: test_image_discovery.py**
**File**: [test_image_discovery.py](test_image_discovery.py)

**Usage**:
```cmd
# Windows
python test_image_discovery.py "C:\Users\Lab\Images" 260154818

# Linux
python test_image_discovery.py /home/lab/images 260154818
```

**Features**:
- Tests image discovery function independently
- Shows all files in the directory
- Displays filename stems and extensions
- Provides detailed troubleshooting output
- Works on both Windows and Linux

### 2. **Updated Documentation**
**File**: [WINDOWS_LAUNCH_README.md](WINDOWS_LAUNCH_README.md)

Added comprehensive troubleshooting section for image auto-discovery issues including:
- How to test image discovery
- Correct filename formats
- Common issues and solutions
- Manual workaround instructions

---

## How to Test the Fix

### On Windows Client:

1. **Launch the application**:
   ```cmd
   python karyotype_report_generator.py
   ```

2. **Set the image folder**:
   - Click "Browse Image Folder"
   - Select the folder containing karyogram images

3. **Enter a sample number**:
   - Type a sample number (e.g., `260154818`)
   - Images should auto-discover immediately

4. **If issues persist, run diagnostics**:
   ```cmd
   python test_image_discovery.py "C:\Path\To\Images" 260154818
   ```

### Expected Image Filenames:
- `260154818.jpg` ✓
- `260154818 1.jpg` ✓
- `260154818 2.jpg` ✓
- `260154818_1.jpg` ✗ (underscores not supported)
- `260154818  1.jpg` ✗ (double spaces not supported)

---

## Technical Details

### Why Path is Better Than glob:

1. **Consistent Behavior**: `Path.iterdir()` works identically on Windows and Linux
2. **No Pattern Matching Issues**: Direct file iteration avoids glob expansion quirks
3. **Better Error Handling**: Can catch specific exceptions
4. **Type Safety**: Path objects are more robust than strings
5. **Cross-Platform**: Automatically handles path separators

### Why Normalization is Important:

On Windows, paths can be:
- `C:\Users\Lab\Images` (native)
- `C:/Users/Lab/Images` (from Qt dialogs)
- `C:\Users\Lab\Images\` (with trailing slash)

`os.path.normpath()` ensures all these become: `C:\Users\Lab\Images`

---

## Verification Checklist

Send this to your Windows client:

- [ ] Can launch application using `launch_windows.bat`
- [ ] Can select image folder using Browse button
- [ ] Images auto-discover when typing sample number
- [ ] Both manual and bulk modes find images
- [ ] Image paths display correctly in the UI
- [ ] PDF generation works with discovered images

If any issues remain, run:
```cmd
python test_image_discovery.py "<IMAGE_FOLDER>" <SAMPLE_NUMBER>
```

And send the output for further diagnosis.
