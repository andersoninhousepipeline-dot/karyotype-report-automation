"""
Karyotype Report Generator - Desktop Application
=================================================
Anderson Diagnostics | Peripheral Blood Karyotyping

Features:
  - Manual patient data entry + live PDF preview
  - Bulk Excel upload with per-row inline editing
  - Report type selector with auto-fill templates
  - Image auto-discovery by sample number (editable)
  - Draft save / load (JSON)
  - QSettings persistence for output/image folders
"""

import sys
import os
import re
import json
import glob
import subprocess
import tempfile
from datetime import datetime
from pathlib import Path


def _resource_path(relative: str) -> str:
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, relative)


import pandas as pd

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QTabWidget, QLabel, QLineEdit, QPushButton, QFileDialog,
    QTableWidget, QTableWidgetItem, QMessageBox, QProgressBar,
    QGroupBox, QFormLayout, QScrollArea, QComboBox,
    QStyle, QSplitter, QTextBrowser, QDialog, QDialogButtonBox,
    QHeaderView, QSizePolicy, QTextEdit, QAbstractItemView,
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QSettings, QTimer, QItemSelectionModel
from PyQt6.QtGui import QPixmap, QFont, QColor, QIcon

from karyotype_template import KaryotypeReportGenerator

try:
    import pypdfium2 as _pdfium
    PYPDFIUM_OK = True
except ImportError:
    PYPDFIUM_OK = False


# ─── Report type templates ──────────────────────────────────────────────────────
REPORT_TEMPLATES = {
    "Normal Male": {
        "INTERPRETATION": "Karyotype shows an apparently normal male.",
        "COMMENTS": "",
        "RECOMMENDATIONS": (
            "\u2022 Genetic counseling is recommended to discuss the implications of the result.\n"
            "\u2022 Additional genetic testing may be warranted based on the specific phenotypic indication."
        ),
        "AUTOSOME": "Normal",
        "SEX CHROMOSOME": "Normal",
    },
    "Normal Female": {
        "INTERPRETATION": "Karyotype shows an apparently normal female.",
        "COMMENTS": "",
        "RECOMMENDATIONS": (
            "\u2022 Genetic counseling is recommended to discuss the implications of the result.\n"
            "\u2022 Additional genetic testing may be warranted based on the specific phenotypic indication."
        ),
        "AUTOSOME": "Normal",
        "SEX CHROMOSOME": "Normal",
    },
    "Trisomy 21 (Down Syndrome)": {
        "INTERPRETATION": (
            "The constitutional karyotype shows a [male/female] with three copies of chromosome 21, "
            "indicating Down (Trisomy 21) syndrome."
        ),
        "COMMENTS": (
            "Trisomy 21 is a genetic syndrome associated with impairment of cognitive ability and "
            "physical growth as well as a particular set of facial characteristics."
        ),
        "RECOMMENDATIONS": "Advised genetic counseling for the parents.",
        "AUTOSOME": "Abnormal",
        "SEX CHROMOSOME": "Normal",
    },
    "Translocation": {
        "INTERPRETATION": (
            "Karyotype shows a [male/female] with an apparently balanced translocation involving "
            "chromosome [X] and [Y] with the breakpoints [Xp/q] and [Yp/q]."
        ),
        "COMMENTS": (
            "Usually this arrangement has no effect on development or general health because no genes "
            "have been lost or gained. However, the carrier of a balanced translocation has an increased "
            "risk of infertility, recurrent abortions and live-born offspring with chromosome imbalances "
            "like Emanuel syndrome."
        ),
        "RECOMMENDATIONS": (
            "Genetic counseling advised (to include review of partners\u2019 karyotype). "
            "Prenatal diagnosis of all subsequent pregnancies is strongly recommended."
        ),
        "AUTOSOME": "Abnormal",
        "SEX CHROMOSOME": "Normal",
    },
    "Mosaic": {
        "INTERPRETATION": (
            "Karyotype analysis showed the presence of two cell lines: [X]% cells with [description] "
            "and [Y]% of the cells with [description]. This indicates a mosaic karyotype."
        ),
        "COMMENTS": (
            "Reports of individuals with this mosaic karyotype have indicated that they can present "
            "a wide spectrum of phenotypes. Clinical correlation is advised."
        ),
        "RECOMMENDATIONS": "Clinical correlation and genetic counselling.",
        "AUTOSOME": "Normal",
        "SEX CHROMOSOME": "Abnormal",
    },
    "Klinefelter's Syndrome (47,XXY)": {
        "INTERPRETATION": (
            "Karyotype analysis shows two X chromosomes and one Y chromosome, "
            "indicating Klinefelter\u2019s syndrome."
        ),
        "COMMENTS": (
            "Klinefelter\u2019s syndrome is characterized by eunuchoid body proportions, taller than "
            "average, gynecomastia, elevated luteinizing hormone (LH) and follicle stimulating hormone "
            "(FSH) levels."
        ),
        "RECOMMENDATIONS": "Advised genetic counseling.",
        "AUTOSOME": "Normal",
        "SEX CHROMOSOME": "Abnormal",
    },
    "Turner Syndrome (45,X)": {
        "INTERPRETATION": (
            "Karyotype analysis shows a single X chromosome (monosomy X), indicating Turner syndrome."
        ),
        "COMMENTS": (
            "Turner syndrome is characterized by short stature, gonadal dysgenesis, and various "
            "clinical features that may include cardiac defects and infertility."
        ),
        "RECOMMENDATIONS": "Advised genetic counseling and specialist evaluation.",
        "AUTOSOME": "Normal",
        "SEX CHROMOSOME": "Abnormal",
    },
    "Chromosomal Variant": {
        "INTERPRETATION": "",
        "COMMENTS": "",
        "RECOMMENDATIONS": "Genetic counseling advised.",
        "AUTOSOME": "Variant Observed",
        "SEX CHROMOSOME": "Normal",
    },
    "Other (Custom)": {
        "INTERPRETATION": "",
        "COMMENTS": "",
        "RECOMMENDATIONS": "",
        "AUTOSOME": "Normal",
        "SEX CHROMOSOME": "Normal",
    },
}

REPORT_TYPE_OPTIONS = list(REPORT_TEMPLATES.keys())


# ─── Field definitions ──────────────────────────────────────────────────────────
# (display_label, excel_col_key, widget_type, placeholder_or_options)
FIELD_DEFS = [
    ("Patient Name",              "NAME",                      "line",  ""),
    ("Gender",                    "GENDER",                    "combo", ["Male", "Female"]),
    ("Age",                       "AGE",                       "line",  "e.g. 25 Years"),
    ("Specimen",                  "SPECIMEN",                  "line",  "Peripheral blood"),
    ("PIN",                       "PIN",                       "line",  ""),
    ("Sample Number",             "SAMPLE NUMBER",             "line",  ""),
    ("Sample Collection Date",    "SAMPLE COLLECTION DATE",    "line",  "DD-MM-YYYY"),
    ("Sample Receipt Date",       "SAMPLE RECEIPT DATE",       "line",  "DD-MM-YYYY"),
    ("Report Date",               "REPORT DATE",               "line",  "DD-MM-YYYY"),
    ("Referring Clinician",       "REFERRING CLINICIAN",       "line",  ""),
    ("Hospital / Clinic",         "HOSPITAL/CLINIC",           "line",  ""),
    ("Test Indication",           "TEST INDICATION",           "line",  "To rule out gross chromosomal abnormality"),
    ("ISCN Result",               "RESULT",                    "line",  "e.g. 46,XX"),
    ("Metaphase Analysed",        "METAPHASE ANALYSED",        "line",  "25"),
    ("Estimated Band Resolution", "ESTIMATED BAND RESOLUTION", "line",  "475"),
    ("Autosome",                  "AUTOSOME",                  "combo", ["Normal", "Abnormal", "Variant Observed"]),
    ("Sex Chromosome",            "SEX CHROMOSOME",            "combo", ["Normal", "Abnormal"]),
    ("Interpretation",            "INTERPRETATION",            "text",  ""),
    ("Comments",                  "COMMENTS",                  "text",  ""),
    ("Recommendations",           "RECOMMENDATIONS",           "text",  ""),
]

BULK_DISPLAY_COLS = ["S. No.", "Patient Name"]


# ─── Helpers ───────────────────────────────────────────────────────────────────
def _clean(v) -> str:
    s = str(v).strip()
    return "" if s in ("nan", "NaT", "None", "NaN", "") else s


def _fmt_date(v) -> str:
    s = _clean(str(v))
    if not s:
        return ""
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(s.split(" ")[0], fmt.split(" ")[0]).strftime("%d-%m-%Y")
        except Exception:
            pass
    return s


def _detect_report_type(iscn: str) -> str:
    s = (iscn or "").strip()
    if not s:
        return "Other (Custom)"
    sl = s.lower()
    if sl.startswith("mos"):
        return "Mosaic"
    if "xxy" in sl:
        return "Klinefelter's Syndrome (47,XXY)"
    if re.search(r"45\s*,\s*x\b", sl):
        return "Turner Syndrome (45,X)"
    if "t(" in sl or "rob(" in sl:
        return "Translocation"
    if "+21" in s:
        return "Trisomy 21 (Down Syndrome)"
    if re.match(r"^46\s*,\s*xy$", s, re.IGNORECASE):
        return "Normal Male"
    if re.match(r"^46\s*,\s*xx$", s, re.IGNORECASE):
        return "Normal Female"
    if re.search(r"del\(|dup\(|inv\(|ins\(|add\(", sl):
        return "Chromosomal Variant"
    return "Other (Custom)"


def _find_images_for_sample(sample_no: str, search_dir: str) -> list:
    """
    Find images matching sample number in search_dir.
    Handles Windows/Linux path differences and case-insensitive matching.

    Matches:
      - Exact: 260154818.jpg
      - Numbered: 260161295 1.jpg, 260161295 2.jpg
    """
    if not sample_no or not search_dir or not os.path.isdir(search_dir):
        return []

    sample_no = str(sample_no).strip()
    if not sample_no:
        return []

    results = []

    # Use Path for better cross-platform compatibility
    search_path = Path(search_dir)

    # Search for common image extensions (case-insensitive on Windows, explicit on Linux)
    extensions = ['.jpg', '.jpeg', '.png', '.JPG', '.JPEG', '.PNG']

    try:
        # List all files in directory
        for file_path in search_path.iterdir():
            if not file_path.is_file():
                continue

            # Check if extension matches (case-insensitive comparison)
            if file_path.suffix.lower() not in [ext.lower() for ext in extensions]:
                continue

            # Extract filename without extension
            fname = file_path.stem

            # Match patterns:
            # 1. Exact match: filename == sample_no
            # 2. Numbered: sample_no followed by space and digit(s)
            if fname == sample_no or re.match(rf"^{re.escape(sample_no)}\s+\d+$", fname):
                results.append(os.path.normpath(str(file_path.absolute())))

        # Sort results for consistent ordering
        results.sort(key=lambda x: Path(x).name)

    except (OSError, PermissionError) as e:
        # Handle permission errors or invalid paths gracefully
        print(f"Warning: Could not access directory {search_dir}: {e}")
        return []

    return results


def _open_folder(path: str):
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception:
        pass


# ─── Workers ───────────────────────────────────────────────────────────────────
class PreviewWorker(QThread):
    finished = pyqtSignal(str)
    error    = pyqtSignal(str)

    def __init__(self, data_row: dict, image_paths: list, tmp_pdf: str,
                 include_logo: bool = True):
        super().__init__()
        self.data_row    = data_row
        self.image_paths = image_paths
        self.tmp_pdf     = tmp_pdf
        self.include_logo = include_logo

    def run(self):
        try:
            tmp_dir = os.path.dirname(self.tmp_pdf)
            gen = KaryotypeReportGenerator(
                self.data_row, self.image_paths, tmp_dir,
                include_logo=self.include_logo)
            gen.filepath = self.tmp_pdf
            gen.filename = os.path.basename(self.tmp_pdf)
            gen.generate()
            self.finished.emit(self.tmp_pdf)
        except Exception:
            import traceback
            self.error.emit(traceback.format_exc())


class BulkWorker(QThread):
    progress = pyqtSignal(int, str)
    finished = pyqtSignal(int, list)

    def __init__(self, jobs: list, output_dir: str, include_logo: bool = True):
        super().__init__()
        self.jobs        = jobs
        self.output_dir  = output_dir
        self.include_logo = include_logo

    def run(self):
        ok, errs = 0, []
        total = len(self.jobs)
        for i, (row, imgs) in enumerate(self.jobs, 1):
            name = _clean(row.get("NAME", f"Row {i}")) or f"Row {i}"
            try:
                self.progress.emit(int((i - 1) / total * 100),
                                   f"Generating {i}/{total}: {name}…")
                KaryotypeReportGenerator(
                    row, imgs, self.output_dir,
                    include_logo=self.include_logo).generate()
                ok += 1
            except Exception as e:
                import traceback
                errs.append(f"{name}: {e}\n{traceback.format_exc()}")
        self.progress.emit(100, "Complete")
        self.finished.emit(ok, errs)


# ─── Image Editor Dialog ────────────────────────────────────────────────────────
class ImageEditorDialog(QDialog):

    def __init__(self, paths: list, sample_no: str, search_dir: str, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"Karyogram Images — Sample {sample_no}")
        self.resize(620, 300)

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(
            "Image paths for this patient. Use Add to browse, Remove to delete."))

        self.table = QTableWidget(0, 2)
        self.table.setHorizontalHeaderLabels(["#", "Image Path"])
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        layout.addWidget(self.table)

        btn_row = QHBoxLayout()
        self.btn_add    = QPushButton("Add Image…")
        self.btn_remove = QPushButton("Remove Selected")
        self.btn_scan   = QPushButton("Re-scan Folder")
        btn_row.addWidget(self.btn_add)
        btn_row.addWidget(self.btn_remove)
        btn_row.addWidget(self.btn_scan)
        btn_row.addStretch()
        layout.addLayout(btn_row)

        btns = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        layout.addWidget(btns)

        self._search_dir = search_dir
        self._sample_no  = sample_no
        self._populate(paths)

        self.btn_add.clicked.connect(self._add)
        self.btn_remove.clicked.connect(self._remove)
        self.btn_scan.clicked.connect(self._rescan)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)

    def _populate(self, paths):
        self.table.setRowCount(0)
        for i, p in enumerate(paths):
            self.table.insertRow(i)
            self.table.setItem(i, 0, QTableWidgetItem(str(i + 1)))
            self.table.setItem(i, 1, QTableWidgetItem(p))

    def _add(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "Select Karyogram Image(s)", self._search_dir,
            "Images (*.jpg *.jpeg *.JPG *.JPEG *.png *.PNG)",
            options=QFileDialog.Option.DontUseNativeDialog)
        for f in files:
            f = os.path.normpath(f)   # Fix: normalize slashes for Windows
            r = self.table.rowCount()
            self.table.insertRow(r)
            self.table.setItem(r, 0, QTableWidgetItem(str(r + 1)))
            self.table.setItem(r, 1, QTableWidgetItem(f))

    def _remove(self):
        rows = sorted({idx.row() for idx in self.table.selectedIndexes()}, reverse=True)
        for r in rows:
            self.table.removeRow(r)
        for i in range(self.table.rowCount()):
            self.table.setItem(i, 0, QTableWidgetItem(str(i + 1)))

    def _rescan(self):
        self._populate(_find_images_for_sample(self._sample_no, self._search_dir))

    def get_paths(self) -> list:
        paths = []
        for r in range(self.table.rowCount()):
            item = self.table.item(r, 1)
            if item and item.text().strip():
                paths.append(item.text().strip())
        return paths


# ─── Main Application Window ────────────────────────────────────────────────────
class KaryotypeReportApp(QMainWindow):

    def __init__(self):
        super().__init__()
        self.settings        = QSettings("AndersonDiagnostics", "KaryotypeReportGen")
        self._manual_inputs  = {}   # key → (widget, wtype)
        self._image_paths    = []
        self._preview_worker = None
        self._gen_worker     = None
        self._tmp_pdf = os.path.join(tempfile.gettempdir(), "_karyo_preview.pdf")

        self._preview_timer = QTimer()
        self._preview_timer.setSingleShot(True)
        self._preview_timer.setInterval(900)
        self._preview_timer.timeout.connect(self._run_preview)

        # Bulk state
        self.bulk_rows:   list = []
        self._bulk_images: list = []
        self._bulk_current_row = -1
        self._bulk_show_cols   = []
        self._bulk_preview_worker = None
        self._bulk_tmp_pdf = os.path.join(tempfile.gettempdir(), "_karyo_bulk_preview.pdf")
        self._bulk_preview_timer = QTimer()
        self._bulk_preview_timer.setSingleShot(True)
        self._bulk_preview_timer.setInterval(900)
        self._bulk_preview_timer.timeout.connect(self._bulk_run_preview)

        self._image_search_dir = self.settings.value(
            "image_search_dir", str(Path.home()))

        self._init_ui()
        self._load_settings()

    # ═══════════════════════════════════════════════════════════════════════════
    # UI bootstrap
    # ═══════════════════════════════════════════════════════════════════════════
    def _init_ui(self):
        self.setWindowTitle("Karyotype Report Generator — Anderson Diagnostics")
        ico = _resource_path("assets/karyotype_icon.png")
        if os.path.isfile(ico):
            self.setWindowIcon(QIcon(ico))
        self.setMinimumSize(1300, 800)
        self.resize(1450, 880)

        central = QWidget()
        self.setCentralWidget(central)
        vbox = QVBoxLayout(central)

        # App title
        title_row = QHBoxLayout()
        lbl = QLabel("Karyotype Report Generator")
        lbl.setStyleSheet(
            "font-size:20px;font-weight:bold;padding:6px 4px;color:#1F497D;")
        sub = QLabel("Anderson Diagnostics — Peripheral Blood Karyotyping")
        sub.setStyleSheet("color:gray;font-size:10px;padding:2px 6px 6px 6px;")
        title_col = QVBoxLayout()
        title_col.setSpacing(0)
        title_col.addWidget(lbl)
        title_col.addWidget(sub)
        title_row.addLayout(title_col)
        title_row.addStretch()
        vbox.addLayout(title_row)

        self.tabs = QTabWidget()
        vbox.addWidget(self.tabs)

        self.tabs.addTab(self._create_manual_tab(), "Manual Entry")
        self.tabs.setTabIcon(
            0, self.style().standardIcon(QStyle.StandardPixmap.SP_FileDialogDetailedView))
        self.tabs.addTab(self._create_bulk_tab(), "Bulk Upload")
        self.tabs.setTabIcon(
            1, self.style().standardIcon(QStyle.StandardPixmap.SP_FileDialogListView))
        self.tabs.addTab(self._create_guide_tab(), "User Guide")
        self.tabs.setTabIcon(
            2, self.style().standardIcon(QStyle.StandardPixmap.SP_MessageBoxInformation))

        self.statusBar().showMessage("Ready")

    # ═══════════════════════════════════════════════════════════════════════════
    # TAB 1 — Manual Entry
    # ═══════════════════════════════════════════════════════════════════════════
    def _create_manual_tab(self) -> QWidget:
        tab   = QWidget()
        outer = QHBoxLayout(tab)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        outer.addWidget(splitter)

        # ── Left: input form ──────────────────────────────────────────────────
        left_w   = QWidget()
        left_vbox = QVBoxLayout(left_w)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        inner  = QWidget()
        form   = QFormLayout(inner)
        form.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.ExpandingFieldsGrow)
        form.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        scroll.setWidget(inner)
        left_vbox.addWidget(scroll)

        # Report Type row (special — not in FIELD_DEFS)
        rt_row = QHBoxLayout()
        self._manual_rt_combo = QComboBox()
        self._manual_rt_combo.addItems(REPORT_TYPE_OPTIONS)
        self._manual_rt_combo.setToolTip(
            "Select the report type to auto-fill Interpretation, "
            "Comments and Recommendations.")
        self._btn_apply_tpl = QPushButton("Apply Template")
        self._btn_apply_tpl.setStyleSheet(
            "background-color:#1F497D;color:white;font-weight:bold;padding:5px 10px;")
        self._btn_apply_tpl.setToolTip(
            "Fills Interpretation / Comments / Recommendations from the selected template")
        self._btn_apply_tpl.clicked.connect(self._apply_manual_template)
        rt_row.addWidget(self._manual_rt_combo, 1)
        rt_row.addWidget(self._btn_apply_tpl)
        form.addRow("Report Type:", rt_row)

        # All patient fields
        for display_lbl, key, wtype, opt in FIELD_DEFS:
            w = self._make_widget(wtype, opt, self._schedule_preview)
            if wtype == "text":
                w.setFixedHeight(65)
            form.addRow(f"{display_lbl}:", w)
            self._manual_inputs[key] = (w, wtype)

        # Draft / clear buttons
        btn_row = QHBoxLayout()
        for text, slot in [("Save Draft", self._manual_save_draft),
                            ("Load Draft", self._manual_load_draft),
                            ("Clear Form", self._manual_clear)]:
            b = QPushButton(text)
            b.clicked.connect(slot)
            btn_row.addWidget(b)
        btn_row.addStretch()
        left_vbox.addLayout(btn_row)

        # Images section
        img_grp = QGroupBox("Karyogram Images  (auto-discovered by Sample Number)")
        img_lay = QVBoxLayout(img_grp)
        self._lbl_images = QLabel("No images")
        self._lbl_images.setWordWrap(True)
        self._lbl_images.setStyleSheet("padding:4px;border:1px solid #ccc;background:white;")
        img_btn_row = QHBoxLayout()
        self._btn_edit_imgs = QPushButton("Edit / Add Images…")
        self._btn_set_dir   = QPushButton("Set Image Folder…")
        img_btn_row.addWidget(self._btn_edit_imgs)
        img_btn_row.addWidget(self._btn_set_dir)
        img_btn_row.addStretch()
        img_lay.addWidget(self._lbl_images)
        img_lay.addLayout(img_btn_row)
        left_vbox.addWidget(img_grp)

        # Generate section
        gen_grp    = QGroupBox("Generate Report")
        gen_layout = QVBoxLayout(gen_grp)

        out_row = QHBoxLayout()
        self._manual_out_lbl = QLabel("No folder selected")
        self._manual_out_lbl.setStyleSheet(
            "padding:4px;border:1px solid #ccc;background:white;")
        btn_out = QPushButton("Select Output Folder")
        btn_out.setIcon(
            self.style().standardIcon(QStyle.StandardPixmap.SP_DirOpenIcon))
        btn_out.clicked.connect(self._manual_browse_output)
        out_row.addWidget(self._manual_out_lbl, 1)
        out_row.addWidget(btn_out)
        # Logo row
        logo_row = QHBoxLayout()
        logo_lbl = QLabel("Logo:")
        self._manual_logo_combo = QComboBox()
        self._manual_logo_combo.addItems(["With Logo", "Without Logo"])
        saved_logo = self.settings.value("logo_mode", "With Logo")
        idx_logo = self._manual_logo_combo.findText(saved_logo)
        if idx_logo >= 0:
            self._manual_logo_combo.setCurrentIndex(idx_logo)
        self._manual_logo_combo.currentTextChanged.connect(
            lambda txt: self.settings.setValue("logo_mode", txt))
        logo_row.addWidget(logo_lbl)
        logo_row.addWidget(self._manual_logo_combo)
        logo_row.addStretch()
        gen_layout.addLayout(logo_row)

        self._manual_gen_btn = QPushButton("Generate Report")
        self._manual_gen_btn.setStyleSheet(
            "background-color:#1F497D;color:white;font-weight:bold;padding:8px;")
        self._manual_gen_btn.setIcon(
            self.style().standardIcon(QStyle.StandardPixmap.SP_MediaPlay))
        self._manual_gen_btn.clicked.connect(self._manual_generate)
        gen_layout.addWidget(self._manual_gen_btn)

        left_vbox.addWidget(gen_grp)
        splitter.addWidget(left_w)

        # ── Right: live preview ───────────────────────────────────────────────
        right_grp  = QGroupBox("Live PDF Preview")
        right_vbox = QVBoxLayout(right_grp)

        prev_toolbar = QHBoxLayout()
        btn_refresh = QPushButton("Refresh Preview")
        btn_refresh.setIcon(
            self.style().standardIcon(QStyle.StandardPixmap.SP_BrowserReload))
        btn_refresh.clicked.connect(self._run_preview)
        self._preview_status = QLabel("Fill in patient details and click Refresh Preview")
        self._preview_status.setStyleSheet("color:gray;font-style:italic;")
        self._preview_status.setWordWrap(True)
        prev_toolbar.addWidget(btn_refresh)
        prev_toolbar.addWidget(self._preview_status, 1)
        right_vbox.addLayout(prev_toolbar)

        prev_scroll = QScrollArea()
        prev_scroll.setWidgetResizable(True)
        self._preview_inner = QWidget()
        self._preview_vbox  = QVBoxLayout(self._preview_inner)
        self._preview_vbox.setAlignment(
            Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignHCenter)
        prev_scroll.setWidget(self._preview_inner)
        right_vbox.addWidget(prev_scroll)

        splitter.addWidget(right_grp)
        splitter.setSizes([440, 800])

        # Wire image buttons and sample number auto-discovery
        self._btn_edit_imgs.clicked.connect(self._manual_edit_images)
        self._btn_set_dir.clicked.connect(self._manual_set_image_dir)
        sn_w, _ = self._manual_inputs["SAMPLE NUMBER"]
        sn_w.textChanged.connect(self._auto_discover_images)

        return tab

    # ── Manual helpers ──────────────────────────────────────────────────────────
    @staticmethod
    def _make_widget(wtype: str, opt, changed_fn):
        if wtype == "combo":
            w = QComboBox()
            w.addItems(opt if isinstance(opt, list) else [opt])
            w.currentTextChanged.connect(changed_fn)
        elif wtype == "text":
            w = QTextEdit()
            w.textChanged.connect(changed_fn)
        else:
            w = QLineEdit()
            if opt:
                w.setPlaceholderText(str(opt))
            w.textChanged.connect(changed_fn)
        return w

    def _apply_manual_template(self):
        tpl = REPORT_TEMPLATES.get(self._manual_rt_combo.currentText(), {})
        self._apply_template_to(tpl, self._manual_inputs)
        self._schedule_preview()

    def _apply_template_to(self, tpl: dict, inputs: dict):
        for key in ("INTERPRETATION", "COMMENTS", "RECOMMENDATIONS"):
            if key in inputs and key in tpl:
                w, wtype = inputs[key]
                if isinstance(w, QTextEdit):
                    w.blockSignals(True)
                    w.setPlainText(tpl[key])
                    w.blockSignals(False)
        for key in ("AUTOSOME", "SEX CHROMOSOME"):
            if key in inputs and key in tpl:
                w, wtype = inputs[key]
                if isinstance(w, QComboBox):
                    idx = w.findText(tpl[key])
                    if idx >= 0:
                        w.blockSignals(True)
                        w.setCurrentIndex(idx)
                        w.blockSignals(False)

    def _get_manual_data(self) -> dict:
        d = {}
        for key, (w, wtype) in self._manual_inputs.items():
            if isinstance(w, QTextEdit):
                d[key] = w.toPlainText().strip()
            elif isinstance(w, QComboBox):
                d[key] = w.currentText()
            else:
                d[key] = w.text().strip()
        return d

    def _set_manual_data(self, data: dict):
        for key, (w, wtype) in self._manual_inputs.items():
            val = _clean(data.get(key, ""))
            w.blockSignals(True)
            if isinstance(w, QTextEdit):
                w.setPlainText(val)
            elif isinstance(w, QComboBox):
                idx = w.findText(val, Qt.MatchFlag.MatchFixedString)
                if idx < 0:
                    idx = w.findText(val, Qt.MatchFlag.MatchContains)
                w.setCurrentIndex(max(idx, 0))
            else:
                w.setText(val)
            w.blockSignals(False)
        # Restore report type combo
        rtype = data.get("REPORT_TYPE", "")
        if not rtype:
            rtype = _detect_report_type(data.get("RESULT", ""))
        idx = self._manual_rt_combo.findText(rtype)
        if idx >= 0:
            self._manual_rt_combo.setCurrentIndex(idx)

    def _schedule_preview(self):
        self._preview_timer.start()

    def _run_preview(self):
        if self._preview_worker and self._preview_worker.isRunning():
            return
        data = self._get_manual_data()
        if not data.get("NAME"):
            self._preview_status.setText("Enter a Patient Name to enable preview.")
            return
        include_logo = (self._manual_logo_combo.currentText() == "With Logo")
        self._preview_status.setText("Generating preview…")
        self._preview_worker = PreviewWorker(
            data, self._image_paths, self._tmp_pdf,
            include_logo=include_logo)
        self._preview_worker.finished.connect(self._on_preview_ready)
        self._preview_worker.error.connect(
            lambda e: self._preview_status.setText(
                f"Preview error: {e.splitlines()[-1]}"))
        self._preview_worker.start()

    def _on_preview_ready(self, pdf_path: str):
        if not PYPDFIUM_OK:
            self._preview_status.setText(
                "PDF preview unavailable — pypdfium2 not installed.")
            return
        self._render_preview_pages(
            pdf_path, self._preview_vbox, self._preview_inner,
            self._preview_status)

    def _render_preview_pages(self, pdf_path, vbox, inner_w, status_lbl):
        try:
            from io import BytesIO
            while vbox.count():
                item = vbox.takeAt(0)
                if item.widget():
                    item.widget().deleteLater()
            doc = _pdfium.PdfDocument(pdf_path)
            target_w = max(inner_w.width() - 24, 560)
            for i in range(len(doc)):
                bm  = doc[i].render(scale=2.5)
                pil = bm.to_pil()
                buf = BytesIO()
                pil.save(buf, format="PNG")
                buf.seek(0)
                px = QPixmap()
                px.loadFromData(buf.read())
                px = px.scaledToWidth(target_w,
                    Qt.TransformationMode.SmoothTransformation)
                lbl = QLabel()
                lbl.setPixmap(px)
                lbl.setAlignment(Qt.AlignmentFlag.AlignHCenter)
                lbl.setStyleSheet("border:1px solid #ccc;margin:4px 0;background:white;")
                vbox.addWidget(lbl)
            doc.close()
            status_lbl.setText(
                f"Preview updated  ({datetime.now().strftime('%H:%M:%S')})")
            status_lbl.setStyleSheet("color:gray;font-style:italic;")
        except Exception as e:
            status_lbl.setText(f"Render error: {e}")

    def _auto_discover_images(self, sample_no: str):
        found = _find_images_for_sample(sample_no.strip(), self._image_search_dir)
        if found:
            self._image_paths = found
            self._refresh_manual_image_label()
            self._schedule_preview()

    def _refresh_manual_image_label(self):
        n = len(self._image_paths)
        if n == 0:
            self._lbl_images.setText("No images — use 'Edit / Add Images…' to add manually")
        else:
            names = [os.path.basename(p) for p in self._image_paths]
            self._lbl_images.setText(f"{n} image(s): " + ",  ".join(names))

    def _manual_edit_images(self):
        sn_w, _ = self._manual_inputs["SAMPLE NUMBER"]
        sample = sn_w.text().strip() if isinstance(sn_w, QLineEdit) else ""
        dlg = ImageEditorDialog(
            self._image_paths, sample, self._image_search_dir, parent=self)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            self._image_paths = dlg.get_paths()
            self._refresh_manual_image_label()
            self._schedule_preview()

    def _manual_set_image_dir(self):
        d = QFileDialog.getExistingDirectory(
            self, "Select Karyogram Search Folder", self._image_search_dir,
            options=QFileDialog.Option.DontUseNativeDialog)
        if d:
            # Normalize path for cross-platform compatibility
            self._image_search_dir = os.path.normpath(d)
            self.settings.setValue("image_search_dir", self._image_search_dir)
            sn_w, _ = self._manual_inputs["SAMPLE NUMBER"]
            if isinstance(sn_w, QLineEdit) and sn_w.text().strip():
                self._auto_discover_images(sn_w.text().strip())

    def _manual_browse_output(self):
        folder = QFileDialog.getExistingDirectory(
            self, "Select Output Folder",
            options=QFileDialog.Option.DontUseNativeDialog)
        if folder:
            self._manual_out_lbl.setText(folder)
            self._manual_out_lbl.setStyleSheet("padding:4px;color:black;")
            self.settings.setValue("manual_outdir", folder)

    def _manual_generate(self):
        out_dir = self._manual_out_lbl.text()
        if not out_dir or out_dir == "No folder selected":
            QMessageBox.warning(self, "No Folder",
                                "Please select an output folder first.")
            return
        data = self._get_manual_data()
        if not data.get("NAME"):
            QMessageBox.warning(self, "Missing Data", "Patient Name is required.")
            return
        include_logo = (self._manual_logo_combo.currentText() == "With Logo")
        try:
            gen  = KaryotypeReportGenerator(
                data, self._image_paths, out_dir,
                include_logo=include_logo)
            path = gen.generate()
            box  = QMessageBox(self)
            box.setWindowTitle("Report Generated")
            box.setIcon(QMessageBox.Icon.Information)
            box.setText(
                f"Report generated successfully.\n\n"
                f"File: {os.path.basename(path)}\n"
                f"Folder: {out_dir}")
            btn_open = box.addButton("Open Folder", QMessageBox.ButtonRole.ActionRole)
            box.addButton(QMessageBox.StandardButton.Ok)
            box.exec()
            if box.clickedButton() == btn_open:
                _open_folder(out_dir)
            self.statusBar().showMessage(f"Saved: {path}")
        except Exception as e:
            import traceback
            QMessageBox.critical(self, "Error", traceback.format_exc())

    def _manual_save_draft(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Save Draft", "karyotype_draft.json", "JSON (*.json)",
            options=QFileDialog.Option.DontUseNativeDialog)
        if not path:
            return
        draft = {
            "data":        self._get_manual_data(),
            "report_type": self._manual_rt_combo.currentText(),
            "images":      self._image_paths,
        }
        with open(path, "w") as f:
            json.dump(draft, f, indent=2)
        self.statusBar().showMessage(f"Draft saved: {path}")

    def _manual_load_draft(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Load Draft", "", "JSON (*.json)",
            options=QFileDialog.Option.DontUseNativeDialog)
        if not path:
            return
        try:
            with open(path) as f:
                draft = json.load(f)
            data = draft.get("data", {})
            data["REPORT_TYPE"] = draft.get("report_type", "")
            self._set_manual_data(data)
            # Fix: normalize paths loaded from JSON (Windows compatibility)
            self._image_paths = [os.path.normpath(p) for p in draft.get("images", [])]
            self._refresh_manual_image_label()
            self._schedule_preview()
            self.statusBar().showMessage(f"Draft loaded: {path}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Could not load draft:\n{e}")

    def _manual_clear(self):
        for key, (w, wtype) in self._manual_inputs.items():
            w.blockSignals(True)
            if isinstance(w, QTextEdit):
                w.clear()
            elif isinstance(w, QLineEdit):
                w.clear()
            elif isinstance(w, QComboBox):
                w.setCurrentIndex(0)
            w.blockSignals(False)
        self._manual_rt_combo.setCurrentIndex(0)
        self._image_paths = []
        self._refresh_manual_image_label()

    # ═══════════════════════════════════════════════════════════════════════════
    # TAB 2 — Bulk Upload
    # ═══════════════════════════════════════════════════════════════════════════
    def _create_bulk_tab(self) -> QWidget:
        tab  = QWidget()
        vbox = QVBoxLayout(tab)

        # ── 1. Excel file ──────────────────────────────────────────────────────
        file_grp = QGroupBox("1. Load Excel File")
        file_row = QHBoxLayout(file_grp)
        self._bulk_file_lbl = QLabel("No file loaded")
        self._bulk_file_lbl.setStyleSheet("color:gray;font-style:italic;padding:2px;")
        btn_browse = QPushButton("Browse…")
        btn_browse.setIcon(
            self.style().standardIcon(QStyle.StandardPixmap.SP_FileDialogStart))
        btn_browse.clicked.connect(self._bulk_load_excel)
        file_row.addWidget(self._bulk_file_lbl, 1)
        file_row.addWidget(btn_browse)
        vbox.addWidget(file_grp)

        # ── 2. Image Folder ────────────────────────────────────────────────────
        imgf_grp = QGroupBox("2. Image Folder  (auto-discovers karyogram images by Sample Number)")
        imgf_row = QHBoxLayout(imgf_grp)
        self._bulk_img_dir_lbl = QLabel(self._image_search_dir)
        self._bulk_img_dir_lbl.setStyleSheet("padding:2px;")
        btn_imgf = QPushButton("Browse…")
        btn_imgf.setIcon(
            self.style().standardIcon(QStyle.StandardPixmap.SP_DirOpenIcon))
        btn_imgf.clicked.connect(self._bulk_browse_img_dir)
        imgf_row.addWidget(self._bulk_img_dir_lbl, 1)
        imgf_row.addWidget(btn_imgf)
        vbox.addWidget(imgf_grp)

        # ── 3. Output Folder ───────────────────────────────────────────────────
        out_grp = QGroupBox("3. Output Folder")
        out_row = QHBoxLayout(out_grp)
        self._bulk_out_lbl = QLabel("No folder selected")
        self._bulk_out_lbl.setStyleSheet("color:gray;font-style:italic;padding:2px;")
        btn_out = QPushButton("Browse…")
        btn_out.setIcon(
            self.style().standardIcon(QStyle.StandardPixmap.SP_DirOpenIcon))
        btn_out.clicked.connect(self._bulk_browse_output)
        out_row.addWidget(self._bulk_out_lbl, 1)
        out_row.addWidget(btn_out)
        vbox.addWidget(out_grp)

        # ── 4. Review & Edit Patients ──────────────────────────────────────────
        data_grp    = QGroupBox("4. Review & Edit Patients")
        data_layout = QVBoxLayout(data_grp)

        # Toolbar
        toolbar = QHBoxLayout()
        for text, slot in [("Select All",   self._bulk_select_all),
                            ("Deselect All", self._bulk_deselect_all)]:
            b = QPushButton(text)
            b.clicked.connect(slot)
            toolbar.addWidget(b)
        toolbar.addStretch()
        data_layout.addLayout(toolbar)

        main_splitter = QSplitter(Qt.Orientation.Horizontal)

        # LEFT — patient table
        self._bulk_table = QTableWidget()
        self._bulk_table.setSelectionBehavior(
            QTableWidget.SelectionBehavior.SelectRows)
        self._bulk_table.setEditTriggers(
            QTableWidget.EditTrigger.NoEditTriggers)
        self._bulk_table.verticalHeader().setVisible(False)
        self._bulk_table.horizontalHeader().setStretchLastSection(True)
        self._bulk_table.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.ResizeToContents)
        self._bulk_table.setAlternatingRowColors(True)
        self._bulk_table.setSelectionMode(
            QTableWidget.SelectionMode.ExtendedSelection)
        self._bulk_table.itemSelectionChanged.connect(self._bulk_on_row_selected)
        main_splitter.addWidget(self._bulk_table)

        # MIDDLE — inline editor
        editor_grp  = QGroupBox("Patient Editor")
        editor_vbox = QVBoxLayout(editor_grp)

        self._bulk_editor_placeholder = QLabel("Click a row in the table to edit")
        self._bulk_editor_placeholder.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._bulk_editor_placeholder.setStyleSheet(
            "color:gray;font-style:italic;padding:20px;")
        editor_vbox.addWidget(self._bulk_editor_placeholder)

        editor_scroll = QScrollArea()
        editor_scroll.setWidgetResizable(True)
        editor_inner = QWidget()
        editor_form  = QFormLayout(editor_inner)
        editor_form.setFieldGrowthPolicy(
            QFormLayout.FieldGrowthPolicy.ExpandingFieldsGrow)
        editor_form.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        editor_scroll.setWidget(editor_inner)
        editor_vbox.addWidget(editor_scroll, 1)
        self._bulk_editor_scroll = editor_scroll

        # Report type row in editor
        self._bulk_rt_combo = QComboBox()
        self._bulk_rt_combo.addItems(REPORT_TYPE_OPTIONS)
        self._bulk_apply_btn = QPushButton("Apply Template")
        self._bulk_apply_btn.setStyleSheet(
            "background-color:#1F497D;color:white;font-weight:bold;padding:4px 8px;")
        self._bulk_apply_btn.clicked.connect(self._bulk_apply_template)
        rt_editor_row = QHBoxLayout()
        rt_editor_row.addWidget(self._bulk_rt_combo, 1)
        rt_editor_row.addWidget(self._bulk_apply_btn)
        editor_form.addRow("Report Type:", rt_editor_row)

        # All patient fields in editor
        self._bulk_editor_inputs = {}
        for display_lbl, key, wtype, opt in FIELD_DEFS:
            w = self._make_widget(wtype, opt, self._bulk_schedule_preview)
            if wtype == "text":
                w.setFixedHeight(60)
            editor_form.addRow(f"{display_lbl}:", w)
            self._bulk_editor_inputs[key] = (w, wtype)

        editor_scroll.setVisible(False)

        # Images row in editor
        self._bulk_img_lbl = QLabel("No images")
        self._bulk_img_lbl.setWordWrap(True)
        self._bulk_img_lbl.setStyleSheet(
            "padding:4px;border:1px solid #ccc;background:white;")
        self._bulk_edit_imgs_btn = QPushButton("Edit Images…")
        self._bulk_edit_imgs_btn.clicked.connect(self._bulk_edit_images)
        self._bulk_edit_imgs_btn.setVisible(False)
        img_lbl_row = QHBoxLayout()
        img_lbl_row.addWidget(self._bulk_img_lbl, 1)
        img_lbl_row.addWidget(self._bulk_edit_imgs_btn)
        editor_vbox.addLayout(img_lbl_row)

        # Editor action buttons
        editor_btn_row = QHBoxLayout()
        self._bulk_save_row_btn = QPushButton("Save Changes")
        self._bulk_save_row_btn.setStyleSheet(
            "background-color:#1F497D;color:white;font-weight:bold;padding:6px;")
        self._bulk_save_row_btn.clicked.connect(self._bulk_save_current_row)
        self._bulk_save_row_btn.setVisible(False)
        editor_btn_row.addWidget(self._bulk_save_row_btn)
        editor_vbox.addLayout(editor_btn_row)

        main_splitter.addWidget(editor_grp)

        # RIGHT — live preview
        prev_grp  = QGroupBox("Live Preview")
        prev_vbox = QVBoxLayout(prev_grp)

        prev_top = QHBoxLayout()
        self._bulk_preview_status = QLabel("Select a row to preview")
        self._bulk_preview_status.setStyleSheet("color:gray;font-style:italic;")
        self._bulk_preview_status.setWordWrap(True)
        prev_top.addWidget(self._bulk_preview_status, 1)
        btn_bulk_refresh = QPushButton("Refresh")
        btn_bulk_refresh.setIcon(
            self.style().standardIcon(QStyle.StandardPixmap.SP_BrowserReload))
        btn_bulk_refresh.clicked.connect(self._bulk_run_preview)
        prev_top.addWidget(btn_bulk_refresh)
        prev_vbox.addLayout(prev_top)

        prev_scroll = QScrollArea()
        prev_scroll.setWidgetResizable(True)
        self._bulk_preview_inner = QWidget()
        self._bulk_preview_vbox  = QVBoxLayout(self._bulk_preview_inner)
        self._bulk_preview_vbox.setAlignment(
            Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignTop)
        prev_scroll.setWidget(self._bulk_preview_inner)
        prev_vbox.addWidget(prev_scroll, 1)

        main_splitter.addWidget(prev_grp)
        main_splitter.setSizes([280, 380, 600])

        data_layout.addWidget(main_splitter, 1)
        vbox.addWidget(data_grp, 1)

        # ── 5. Generate ────────────────────────────────────────────────────────
        gen_grp    = QGroupBox("5. Generate Reports")
        gen_layout = QVBoxLayout(gen_grp)

        act_row = QHBoxLayout()
        self._bulk_gen_sel_btn = QPushButton("Generate Selected")
        self._bulk_gen_sel_btn.setStyleSheet(
            "background-color:#1F497D;color:white;font-weight:bold;padding:8px;")
        self._bulk_gen_sel_btn.setEnabled(False)
        self._bulk_gen_sel_btn.clicked.connect(self._bulk_generate_selected)

        self._bulk_gen_all_btn = QPushButton("Generate All")
        self._bulk_gen_all_btn.setStyleSheet(
            "background-color:#27AE60;color:white;font-weight:bold;padding:8px;")
        self._bulk_gen_all_btn.setEnabled(False)
        self._bulk_gen_all_btn.clicked.connect(self._bulk_generate_all)

        act_row.addWidget(self._bulk_gen_sel_btn)
        act_row.addWidget(self._bulk_gen_all_btn)
        act_row.addStretch()
        gen_layout.addLayout(act_row)

        # Bulk logo row
        bulk_logo_row = QHBoxLayout()
        bulk_logo_lbl = QLabel("Logo:")
        self._bulk_logo_combo = QComboBox()
        self._bulk_logo_combo.addItems(["With Logo", "Without Logo"])
        saved_bulk_logo = self.settings.value("bulk_logo_mode", "With Logo")
        idx_bulk_logo = self._bulk_logo_combo.findText(saved_bulk_logo)
        if idx_bulk_logo >= 0:
            self._bulk_logo_combo.setCurrentIndex(idx_bulk_logo)
        self._bulk_logo_combo.currentTextChanged.connect(
            lambda txt: self.settings.setValue("bulk_logo_mode", txt))
        bulk_logo_row.addWidget(bulk_logo_lbl)
        bulk_logo_row.addWidget(self._bulk_logo_combo)
        bulk_logo_row.addStretch()
        gen_layout.addLayout(bulk_logo_row)

        self._bulk_prog_lbl  = QLabel("")
        self._bulk_prog_lbl.setVisible(False)
        self._bulk_progress  = QProgressBar()
        self._bulk_progress.setVisible(False)
        gen_layout.addWidget(self._bulk_prog_lbl)
        gen_layout.addWidget(self._bulk_progress)
        vbox.addWidget(gen_grp)

        return tab

    # ── Bulk helpers ────────────────────────────────────────────────────────────
    def _bulk_load_excel(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Open Excel File", "",
            "Excel Files (*.xls *.xlsx *.xlsm)",
            options=QFileDialog.Option.DontUseNativeDialog)
        if not path:
            return
        try:
            # Auto-detect header row
            raw = pd.read_excel(path, header=None, dtype=str, nrows=8)
            header_row = 0
            for i, row in raw.iterrows():
                vals = [str(v).strip().upper() for v in row.values if str(v).strip()]
                if any(k in vals for k in ("NAME", "S.NO.", "S. NO.", "S.NO")):
                    header_row = int(i)
                    break

            df = pd.read_excel(path, header=header_row, dtype=str)
            df.columns = [str(c).strip().upper() for c in df.columns]
            df = df.dropna(how="all")

            self.bulk_rows   = []
            self._bulk_images = []
            for _, ser in df.iterrows():
                row = {k: _clean(v) for k, v in ser.items()}
                if not _clean(row.get("NAME", "")):
                    continue
                for dc in ("SAMPLE COLLECTION DATE",
                           "SAMPLE RECEIPT DATE", "REPORT DATE"):
                    for variant in (dc, dc + " ", " " + dc):
                        if variant in row:
                            row[dc] = _fmt_date(row[variant])
                for k in list(row.keys()):
                    canonical = k.strip()
                    if canonical != k and canonical not in row:
                        row[canonical] = row[k]
                row["REPORT_TYPE"] = _detect_report_type(row.get("RESULT", ""))
                self.bulk_rows.append(row)
                self._bulk_images.append(
                    _find_images_for_sample(
                        row.get("SAMPLE NUMBER", ""), self._image_search_dir))

            self._bulk_file_lbl.setText(os.path.basename(path))
            self._bulk_file_lbl.setStyleSheet("color:black;padding:2px;")
            self._populate_bulk_table()
            self._bulk_gen_sel_btn.setEnabled(True)
            self._bulk_gen_all_btn.setEnabled(True)
            self.statusBar().showMessage(
                f"Loaded {len(self.bulk_rows)} patients from {os.path.basename(path)}")
        except Exception as e:
            import traceback
            QMessageBox.critical(self, "Load Error", traceback.format_exc())

    def _populate_bulk_table(self):
        cols = ["S. No.", "Patient Name", "Sample Number", "ISCN Result",
                "Report Type", "Images"]
        self._bulk_show_cols = cols
        self._bulk_table.setColumnCount(len(cols))
        self._bulk_table.setHorizontalHeaderLabels(cols)
        self._bulk_table.setRowCount(len(self.bulk_rows))

        for r, row in enumerate(self.bulk_rows):
            imgs = self._bulk_images[r]
            cells = [
                str(r + 1),
                row.get("NAME", ""),
                row.get("SAMPLE NUMBER", ""),
                row.get("RESULT", ""),
                row.get("REPORT_TYPE", ""),
                f"{len(imgs)} image(s)" if imgs else "No images",
            ]
            for c, val in enumerate(cells):
                item = QTableWidgetItem(val)
                item.setFlags(
                    Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled)
                if c == 5 and not imgs:
                    item.setForeground(QColor("red"))
                self._bulk_table.setItem(r, c, item)

        self._bulk_current_row = -1
        self._bulk_table.selectRow(0)

    def _bulk_select_all(self):
        model = self._bulk_table.selectionModel()
        model.clearSelection()
        for row in range(self._bulk_table.rowCount()):
            if not self._bulk_table.isRowHidden(row):
                idx = self._bulk_table.model().index(row, 0)
                model.select(
                    idx,
                    QItemSelectionModel.SelectionFlag.Select |
                    QItemSelectionModel.SelectionFlag.Rows)

    def _bulk_deselect_all(self):
        self._bulk_table.clearSelection()

    def _bulk_on_row_selected(self):
        row_idx = self._bulk_table.currentRow()
        if row_idx < 0 or row_idx >= len(self.bulk_rows):
            return
        self._bulk_populate_editor(row_idx)
        self._bulk_schedule_preview()

    def _bulk_populate_editor(self, row_idx: int):
        self._bulk_current_row = row_idx
        data = self.bulk_rows[row_idx]

        self._bulk_editor_placeholder.setVisible(False)
        self._bulk_editor_scroll.setVisible(True)
        self._bulk_save_row_btn.setVisible(True)
        self._bulk_edit_imgs_btn.setVisible(True)

        # Set report type combo
        rtype = data.get("REPORT_TYPE", "")
        idx = self._bulk_rt_combo.findText(rtype)
        if idx >= 0:
            self._bulk_rt_combo.blockSignals(True)
            self._bulk_rt_combo.setCurrentIndex(idx)
            self._bulk_rt_combo.blockSignals(False)

        # Fill all fields
        for key, (w, wtype) in self._bulk_editor_inputs.items():
            val = _clean(data.get(key, ""))
            w.blockSignals(True)
            if isinstance(w, QTextEdit):
                w.setPlainText(val)
            elif isinstance(w, QComboBox):
                idx2 = w.findText(val, Qt.MatchFlag.MatchFixedString)
                if idx2 < 0:
                    idx2 = w.findText(val, Qt.MatchFlag.MatchContains)
                w.setCurrentIndex(max(idx2, 0))
            else:
                w.setText(val)
            w.blockSignals(False)

        # Refresh image label
        imgs = self._bulk_images[row_idx]
        if imgs:
            names = [os.path.basename(p) for p in imgs]
            self._bulk_img_lbl.setText(
                f"{len(imgs)} image(s): " + ",  ".join(names))
        else:
            self._bulk_img_lbl.setText("No images found for this sample")

    def _bulk_apply_template(self):
        tpl = REPORT_TEMPLATES.get(self._bulk_rt_combo.currentText(), {})
        self._apply_template_to(tpl, self._bulk_editor_inputs)
        self._bulk_schedule_preview()

    def _bulk_save_current_row(self):
        if self._bulk_current_row < 0 or \
                self._bulk_current_row >= len(self.bulk_rows):
            return
        idx = self._bulk_current_row
        d   = dict(self.bulk_rows[idx])
        for key, (w, wtype) in self._bulk_editor_inputs.items():
            if isinstance(w, QTextEdit):
                d[key] = w.toPlainText().strip()
            elif isinstance(w, QComboBox):
                d[key] = w.currentText()
            else:
                d[key] = w.text().strip()
        d["REPORT_TYPE"] = self._bulk_rt_combo.currentText()
        self.bulk_rows[idx] = d
        # Refresh table row
        imgs = self._bulk_images[idx]
        cells = [
            str(idx + 1),
            d.get("NAME", ""),
            d.get("SAMPLE NUMBER", ""),
            d.get("RESULT", ""),
            d.get("REPORT_TYPE", ""),
            f"{len(imgs)} image(s)" if imgs else "No images",
        ]
        for c, val in enumerate(cells):
            self._bulk_table.setItem(idx, c, QTableWidgetItem(val))
        self.statusBar().showMessage(f"Row {idx + 1} saved")
        self._bulk_run_preview()

    def _bulk_edit_images(self):
        if self._bulk_current_row < 0:
            return
        row    = self.bulk_rows[self._bulk_current_row]
        sample = row.get("SAMPLE NUMBER", "")
        dlg    = ImageEditorDialog(
            self._bulk_images[self._bulk_current_row],
            sample, self._image_search_dir, parent=self)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            self._bulk_images[self._bulk_current_row] = dlg.get_paths()
            imgs  = self._bulk_images[self._bulk_current_row]
            names = [os.path.basename(p) for p in imgs]
            self._bulk_img_lbl.setText(
                f"{len(imgs)} image(s): " + ",  ".join(names) if imgs
                else "No images")
            # Update table
            self._bulk_table.setItem(
                self._bulk_current_row, 5,
                QTableWidgetItem(f"{len(imgs)} image(s)" if imgs else "No images"))

    def _bulk_schedule_preview(self):
        if self._bulk_current_row >= 0:
            self._bulk_preview_timer.start()

    def _bulk_run_preview(self):
        if self._bulk_current_row < 0 or \
                self._bulk_current_row >= len(self.bulk_rows):
            return
        if self._bulk_preview_worker and self._bulk_preview_worker.isRunning():
            return
        # Use live editor values
        d = dict(self.bulk_rows[self._bulk_current_row])
        for key, (w, wtype) in self._bulk_editor_inputs.items():
            if isinstance(w, QTextEdit):
                d[key] = w.toPlainText().strip()
            elif isinstance(w, QComboBox):
                d[key] = w.currentText()
            else:
                d[key] = w.text().strip()
        if not _clean(d.get("NAME", "")):
            self._bulk_preview_status.setText(
                "Enter a Patient Name to enable preview.")
            return
        self._bulk_preview_status.setText("Generating preview…")
        imgs = self._bulk_images[self._bulk_current_row]
        include_logo = (self._bulk_logo_combo.currentText() == "With Logo")
        self._bulk_preview_worker = PreviewWorker(
            d, imgs, self._bulk_tmp_pdf,
            include_logo=include_logo)
        self._bulk_preview_worker.finished.connect(self._bulk_on_preview_ready)
        self._bulk_preview_worker.error.connect(
            lambda e: self._bulk_preview_status.setText(
                f"Preview error: {e.splitlines()[-1]}"))
        self._bulk_preview_worker.start()

    def _bulk_on_preview_ready(self, pdf_path: str):
        if not PYPDFIUM_OK:
            self._bulk_preview_status.setText(
                "PDF preview unavailable — pypdfium2 not installed.")
            return
        self._render_preview_pages(
            pdf_path,
            self._bulk_preview_vbox,
            self._bulk_preview_inner,
            self._bulk_preview_status)

    def _bulk_browse_img_dir(self):
        d = QFileDialog.getExistingDirectory(
            self, "Select Image Folder", self._image_search_dir,
            options=QFileDialog.Option.DontUseNativeDialog)
        if d:
            # Normalize path for cross-platform compatibility
            self._image_search_dir = os.path.normpath(d)
            self.settings.setValue("image_search_dir", self._image_search_dir)
            self._bulk_img_dir_lbl.setText(self._image_search_dir)
            # Re-scan all images
            if self.bulk_rows:
                for i, row in enumerate(self.bulk_rows):
                    self._bulk_images[i] = _find_images_for_sample(
                        row.get("SAMPLE NUMBER", ""), self._image_search_dir)
                self._populate_bulk_table()

    def _bulk_browse_output(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        if folder:
            self._bulk_out_lbl.setText(folder)
            self._bulk_out_lbl.setStyleSheet("color:black;padding:2px;")
            self.settings.setValue("bulk_outdir", folder)

    def _bulk_generate_selected(self):
        ranges = self._bulk_table.selectedRanges()
        if not ranges:
            QMessageBox.warning(self, "No Selection", "Select rows to generate.")
            return
        jobs = []
        for rng in ranges:
            for r in range(rng.topRow(), rng.bottomRow() + 1):
                if r < len(self.bulk_rows) and \
                        not self._bulk_table.isRowHidden(r):
                    jobs.append((self.bulk_rows[r], self._bulk_images[r]))
        self._start_bulk_gen(jobs)

    def _bulk_generate_all(self):
        if not self.bulk_rows:
            QMessageBox.warning(self, "No Data", "Load an Excel file first.")
            return
        self._start_bulk_gen(list(zip(self.bulk_rows, self._bulk_images)))

    def _start_bulk_gen(self, jobs: list):
        out_dir = self._bulk_out_lbl.text()
        if not out_dir or out_dir == "No folder selected":
            out_dir = QFileDialog.getExistingDirectory(
                self, "Select Output Folder",
                options=QFileDialog.Option.DontUseNativeDialog)
            if not out_dir:
                return
            self._bulk_out_lbl.setText(out_dir)
            self._bulk_out_lbl.setStyleSheet("color:black;padding:2px;")

        os.makedirs(out_dir, exist_ok=True)
        self._bulk_progress.setValue(0)
        self._bulk_progress.setVisible(True)
        self._bulk_prog_lbl.setVisible(True)
        self._bulk_gen_sel_btn.setEnabled(False)
        self._bulk_gen_all_btn.setEnabled(False)

        self._gen_worker = BulkWorker(
            jobs, out_dir,
            include_logo=(self._bulk_logo_combo.currentText() == "With Logo"))
        self._gen_worker.progress.connect(self._on_bulk_progress)
        self._gen_worker.finished.connect(self._on_bulk_finished)
        self._gen_worker.start()

    def _on_bulk_progress(self, pct: int, msg: str):
        self._bulk_progress.setValue(pct)
        self._bulk_prog_lbl.setText(msg)

    def _on_bulk_finished(self, successes: int, errors: list):
        self._bulk_progress.setVisible(False)
        self._bulk_prog_lbl.setVisible(False)
        self._bulk_gen_sel_btn.setEnabled(True)
        self._bulk_gen_all_btn.setEnabled(True)
        out_dir = self._bulk_out_lbl.text()

        if errors:
            box = QMessageBox(self)
            box.setWindowTitle("Generation Complete")
            box.setIcon(QMessageBox.Icon.Warning)
            box.setText(f"{successes} report(s) generated, {len(errors)} failed.")
            box.setDetailedText("\n\n".join(errors))
            box.exec()
        else:
            box = QMessageBox(self)
            box.setWindowTitle("Generation Complete")
            box.setIcon(QMessageBox.Icon.Information)
            box.setText(f"All {successes} report(s) generated successfully.")
            btn_open = box.addButton("Open Folder", QMessageBox.ButtonRole.ActionRole)
            box.addButton(QMessageBox.StandardButton.Ok)
            box.exec()
            if box.clickedButton() == btn_open:
                _open_folder(out_dir)

        self.statusBar().showMessage(
            f"Generated {successes} report(s)"
            + (f", {len(errors)} error(s)" if errors else ""))

    # ═══════════════════════════════════════════════════════════════════════════
    # TAB 3 — User Guide
    # ═══════════════════════════════════════════════════════════════════════════
    def _create_guide_tab(self) -> QWidget:
        tab = QWidget()
        lay = QVBoxLayout(tab)
        browser = QTextBrowser()
        browser.setOpenExternalLinks(True)
        browser.setHtml("""
<html><head><style>
body  { font-family:'Segoe UI',Arial,sans-serif; padding:16px; color:#222; }
h1   { color:#1F497D; border-bottom:2px solid #1F497D; padding-bottom:6px; }
h2   { color:#1F497D; margin-top:18px; }
code { background:#eef2f7; padding:2px 5px; border-radius:3px;
       font-family:Consolas,monospace; }
table { border-collapse:collapse; width:100%; margin:8px 0; }
th    { background:#1F497D; color:white; padding:7px 10px; text-align:left; }
td    { border:1px solid #ddd; padding:5px 10px; }
tr:nth-child(even) { background:#f7f9fc; }
.tip  { background:#fffbe6; border-left:4px solid #d4a000; padding:8px 12px;
        margin:8px 0; }
</style></head><body>

<h1>Karyotype Report Generator — User Guide</h1>
<p>Anderson Diagnostics &nbsp;|&nbsp; Peripheral Blood Karyotyping (PBK)</p>

<h2>Quick Start</h2>
<ol>
  <li><b>Manual Entry</b> — Select a <b>Report Type</b>, click <i>Apply Template</i>,
      fill in patient details, then click <b>Generate Report</b>.</li>
  <li><b>Bulk Upload</b> — Load an Excel file, set the Image Folder and Output Folder,
      review / edit each patient in the inline editor, then click <b>Generate All</b>.</li>
</ol>

<h2>Report Types &amp; Templates</h2>
<table>
<tr><th>Report Type</th><th>Typical ISCN</th><th>Layout</th></tr>
<tr><td>Normal Male</td><td>46,XY</td><td>2-page</td></tr>
<tr><td>Normal Female</td><td>46,XX</td><td>2-page</td></tr>
<tr><td>Trisomy 21 (Down Syndrome)</td><td>47,XY,+21 / 47,XX,+21</td><td>3-page</td></tr>
<tr><td>Translocation</td><td>46,XX,t(11;22)…</td><td>3-page</td></tr>
<tr><td>Mosaic</td><td>mos 46,X,del(Y)…[43]/45,X[7]</td><td>3-page</td></tr>
<tr><td>Klinefelter's Syndrome (47,XXY)</td><td>47,XXY</td><td>3-page</td></tr>
<tr><td>Turner Syndrome (45,X)</td><td>45,X</td><td>3-page</td></tr>
<tr><td>Chromosomal Variant</td><td>46,XY,del(…)</td><td>3-page</td></tr>
<tr><td>Other (Custom)</td><td>Any</td><td>Auto (by Comments field)</td></tr>
</table>
<div class="tip"><b>Auto-detection:</b> When loading from Excel the Report Type is
automatically guessed from the ISCN result. You can change it in the editor and
click <i>Apply Template</i> to update Interpretation / Comments / Recommendations.</div>

<h2>Layout Rules</h2>
<ul>
  <li><b>2-page</b> — Comments field is empty (Normal Male / Female).</li>
  <li><b>3-page</b> — Comments field has content (all abnormal / variant reports).</li>
</ul>

<h2>Image Auto-Discovery</h2>
<p>Point the <b>Image Folder</b> to the directory containing karyogram images.
Files are matched by Sample Number:</p>
<ul>
  <li><code>260154818.jpg</code> — single image (exact match)</li>
  <li><code>260161295 1.jpg</code>, <code>260161295 2.jpg</code> — multiple images (numbered suffix)</li>
</ul>
<p>Use <i>Edit / Add Images…</i> to manually adjust if auto-discovery misses files.</p>

<h2>Excel Column Reference</h2>
<table>
<tr><th>Column</th><th>Description</th></tr>
<tr><td>NAME</td><td>Patient full name</td></tr>
<tr><td>GENDER</td><td>Male / Female</td></tr>
<tr><td>AGE</td><td>Age (e.g. 25 Years)</td></tr>
<tr><td>SPECIMEN</td><td>e.g. Peripheral blood</td></tr>
<tr><td>PIN</td><td>Patient ID number</td></tr>
<tr><td>SAMPLE NUMBER</td><td>Lab sample number (used for image matching)</td></tr>
<tr><td>SAMPLE COLLECTION DATE</td><td>DD-MM-YYYY</td></tr>
<tr><td>SAMPLE RECEIPT DATE</td><td>DD-MM-YYYY</td></tr>
<tr><td>REPORT DATE</td><td>DD-MM-YYYY</td></tr>
<tr><td>REFERRING CLINICIAN</td><td>Doctor / clinician name</td></tr>
<tr><td>HOSPITAL/CLINIC</td><td>Referring hospital or clinic</td></tr>
<tr><td>TEST INDICATION</td><td>Clinical reason for test</td></tr>
<tr><td>RESULT</td><td>ISCN notation (e.g. 46,XX)</td></tr>
<tr><td>METAPHASE ANALYSED</td><td>Number of metaphases counted</td></tr>
<tr><td>ESTIMATED BAND RESOLUTION</td><td>e.g. 475</td></tr>
<tr><td>AUTOSOME</td><td>Normal / Abnormal / Variant Observed</td></tr>
<tr><td>SEX CHROMOSOME</td><td>Normal / Abnormal</td></tr>
<tr><td>INTERPRETATION</td><td>Clinical interpretation paragraph</td></tr>
<tr><td>COMMENTS</td><td>Additional clinical comments (blank → 2-page layout)</td></tr>
<tr><td>RECOMMENDATIONS</td><td>Clinical recommendations</td></tr>
</table>

<h2>Output Filename</h2>
<p><code>{Patient_Name}_{Sample_Number}_Karyotype.pdf</code></p>

<h2>Fixed Signatories</h2>
<ul>
  <li>Dr. R. Deepika — Consultant Cytogeneticist</li>
  <li>Dr. Teena Koshy — Consultant Cytogeneticist</li>
  <li>Dr. Suriya kumar G — Senior Consultant &amp; HOD</li>
</ul>

<h2>Building a Standalone Executable</h2>
<pre style="background:#eef2f7;padding:10px;border-radius:4px;font-family:Consolas,monospace;">
# Linux
bash build_linux.sh

# Windows
build_windows.bat
</pre>
<p>Output: <code>dist/KaryotypeReportGen/KaryotypeReportGen</code> (Linux) or
<code>.exe</code> (Windows)</p>

</body></html>
""")
        lay.addWidget(browser)
        return tab

    # ═══════════════════════════════════════════════════════════════════════════
    # Settings persistence
    # ═══════════════════════════════════════════════════════════════════════════
    def _load_settings(self):
        out = self.settings.value("manual_outdir", "")
        if out:
            self._manual_out_lbl.setText(out)
            self._manual_out_lbl.setStyleSheet("padding:4px;color:black;")
        bulk_out = self.settings.value("bulk_outdir", "")
        if bulk_out:
            self._bulk_out_lbl.setText(bulk_out)
            self._bulk_out_lbl.setStyleSheet("color:black;padding:2px;")
        img_dir = self.settings.value("image_search_dir", "")
        if img_dir:
            # Normalize path for cross-platform compatibility
            self._image_search_dir = os.path.normpath(img_dir)
            self._bulk_img_dir_lbl.setText(self._image_search_dir)


# ─── Entry point ───────────────────────────────────────────────────────────────
def main():
    app = QApplication(sys.argv)
    app.setApplicationName("Karyotype Report Generator")
    app.setOrganizationName("Anderson Diagnostics")
    app.setStyle("Fusion")
    font = QFont("Segoe UI" if sys.platform.startswith("win") else "Ubuntu", 10)
    app.setFont(font)

    win = KaryotypeReportApp()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
