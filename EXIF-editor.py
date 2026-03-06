#!/usr/bin/env python3
from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QFont
from PyQt6.QtWidgets import (
    QApplication,
    QCheckBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QPlainTextEdit,
    QProgressBar,
    QScrollArea,
    QSplitter,
    QStatusBar,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

SUPPORTED_EXTS = {
    ".jpg", ".jpeg", ".tif", ".tiff", ".png", ".heic", ".webp",
    ".nef", ".cr2", ".cr3", ".arw", ".raf", ".orf", ".rw2", ".dng", ".pef", ".srw", ".3fr", ".rwl",
}

RAW_EXTS = {
    ".nef", ".cr2", ".cr3", ".arw", ".raf", ".orf",
    ".rw2", ".dng", ".pef", ".srw", ".3fr", ".rwl",
}

DT_RE = re.compile(r"^\d{4}:\d{2}:\d{2} \d{2}:\d{2}:\d{2}$")
FRACTION_RE = re.compile(r"^\d+/\d+$")
EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")
URL_RE = re.compile(r"^https?://", re.IGNORECASE)

MIXED_TEXT = "(mixed values)"


# ---------- Styling ----------
def apply_professional_theme(app: QApplication) -> None:
    app.setStyle("Fusion")
    app.setStyleSheet(
        """
        QWidget {
            background-color: #113b2f;
            color: #ffffff;
            font-size: 12px;
        }
        QMainWindow {
            background-color: #113b2f;
        }
        QLabel#titleLabel,
        QGroupBox::title,
        QTabBar::tab:selected,
        QTabBar::tab:hover {
            color: #ff9f1c;
        }
        QLabel#hintLabel {
            color: #d7f3df;
        }
        QLabel#statusInfoLabel {
            color: #e8f8ee;
        }
        QGroupBox {
            font-weight: 700;
            border: 1px solid #2d705b;
            border-radius: 12px;
            margin-top: 12px;
            padding: 12px;
            background-color: #164736;
        }
        QGroupBox::title {
            subcontrol-origin: margin;
            left: 12px;
            padding: 0 8px;
        }
        QLineEdit, QPlainTextEdit, QListWidget {
            background-color: #0d2e24;
            border: 1px solid #2d705b;
            border-radius: 10px;
            padding: 8px;
            color: #ffffff;
            selection-background-color: #ff9f1c;
            selection-color: #1b1b1b;
        }
        QListWidget::item {
            padding: 6px;
            border-radius: 6px;
        }
        QListWidget::item:selected {
            background: #255f4d;
            color: #ffffff;
        }
        QTabWidget::pane {
            border: 1px solid #2d705b;
            border-radius: 12px;
            top: -1px;
            background: #164736;
        }
        QTabBar::tab {
            background: #225746;
            border: 1px solid #2d705b;
            padding: 9px 16px;
            margin-right: 6px;
            border-top-left-radius: 10px;
            border-top-right-radius: 10px;
            color: #f4fff7;
            font-weight: 600;
        }
        QPushButton {
            background-color: #ff9f1c;
            color: #162219;
            border: none;
            border-radius: 12px;
            padding: 10px 16px;
            font-weight: 700;
            min-height: 18px;
        }
        QPushButton:hover {
            background-color: #ffb347;
        }
        QPushButton:pressed {
            background-color: #e98e12;
        }
        QPushButton:disabled {
            background-color: #8c7048;
            color: #2a2a2a;
        }
        QPushButton#secondaryButton {
            background-color: #2a6b57;
            color: #ffffff;
            border: 1px solid #41856f;
        }
        QPushButton#secondaryButton:hover {
            background-color: #318066;
        }
        QPushButton#dangerButton {
            background-color: #944d3b;
            color: #ffffff;
        }
        QPushButton#dangerButton:hover {
            background-color: #a85a46;
        }
        QCheckBox {
            spacing: 6px;
        }
        QCheckBox::indicator {
            width: 16px;
            height: 16px;
        }
        QCheckBox::indicator:unchecked {
            border: 1px solid #6db39a;
            background: #0d2e24;
            border-radius: 4px;
        }
        QCheckBox::indicator:checked {
            border: 1px solid #ff9f1c;
            background: #ff9f1c;
            border-radius: 4px;
        }
        QProgressBar {
            background: #0d2e24;
            border: 1px solid #2d705b;
            border-radius: 10px;
            text-align: center;
            min-height: 20px;
            font-weight: 700;
        }
        QProgressBar::chunk {
            background-color: #ff9f1c;
            border-radius: 10px;
        }
        QStatusBar {
            background: #0d2e24;
            color: #ffffff;
            border-top: 1px solid #2d705b;
        }
        QMessageBox {
            background-color: #113b2f;
        }
        """
    )


def make_title_label(text: str) -> QLabel:
    label = QLabel(text)
    label.setObjectName("titleLabel")
    font = QFont()
    font.setPointSize(12)
    font.setBold(True)
    label.setFont(font)
    return label


def wrap_in_scroll_area(widget: QWidget) -> QScrollArea:
    scroll = QScrollArea()
    scroll.setWidgetResizable(True)
    scroll.setFrameShape(QScrollArea.Shape.NoFrame)
    scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
    scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
    scroll.setWidget(widget)
    return scroll


# ---------- Helpers ----------
def is_int(text: str) -> bool:
    try:
        int(text)
        return True
    except Exception:
        return False



def is_float(text: str) -> bool:
    try:
        float(text)
        return True
    except Exception:
        return False



def normalize_decimal(text: str) -> str:
    return (text or "").strip().replace(",", ".")



def stringify_meta_value(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, list):
        return ", ".join(str(v) for v in value)
    return str(value)



def split_keywords(text: str) -> List[str]:
    parts = [p.strip() for p in text.split(",")]
    seen: set[str] = set()
    out: List[str] = []
    for p in parts:
        if p and p.lower() not in seen:
            out.append(p)
            seen.add(p.lower())
    return out



def validate_tag_value(tag: str, value: str) -> Optional[str]:
    v = normalize_decimal(value)

    if tag == "EXIF:ISO":
        if not is_int(v) or int(v) <= 0:
            return "ISO must be a positive integer (e.g. 100, 400)."

    if tag == "EXIF:FNumber":
        if not is_float(v) or float(v) <= 0:
            return "F-Number must be a positive number (e.g. 2.8, 8)."

    if tag == "EXIF:ExposureTime":
        if not (FRACTION_RE.match(v) or is_float(v)):
            return "Exposure Time must be like 1/125 or a decimal like 0.008."
        if FRACTION_RE.match(v):
            num, den = v.split("/")
            try:
                if int(num) <= 0 or int(den) <= 0:
                    return "Exposure Time fraction must be positive (e.g. 1/125)."
            except Exception:
                return "Exposure Time must be like 1/125."
        if is_float(v) and float(v) <= 0:
            return "Exposure Time must be greater than 0."

    if tag == "EXIF:FocalLength":
        if not is_float(v) or float(v) <= 0:
            return "Focal Length must be a positive number (e.g. 24, 35, 70)."

    if tag in ("EXIF:ExposureCompensation", "EXIF:ExposureBiasValue"):
        if not is_float(v):
            return "Exposure Compensation must be a number (e.g. -0.3, 0, 1.0)."

    if tag == "EXIF:DateTimeOriginal":
        if not DT_RE.match(v):
            return "Date Taken must be YYYY:MM:DD HH:MM:SS (e.g. 2026:03:01 14:25:00)."

    if tag == "XMP:CreatorWorkEmail" and v and not EMAIL_RE.match(v):
        return "Email looks invalid (example: agent@company.com)."

    if tag == "XMP:CreatorWorkURL" and v and not URL_RE.match(v):
        return "Website should start with http:// or https://"

    if tag == "XMP:CreatorWorkTelephone" and v and not re.fullmatch(r"[0-9+\-\s().]{5,}", v):
        return "Phone contains unusual characters. Use digits, +, spaces, - or ()."

    return None


# ---------- ExifTool ----------
def app_base_dir() -> Path:
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)  # type: ignore[attr-defined]
    return Path(__file__).resolve().parent



def find_exiftool() -> str:
    candidates: List[Path] = []

    # PyInstaller one-file/one-dir temporary extraction folder
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        candidates.append(Path(sys._MEIPASS) / "tools" / "exiftool.exe")  # type: ignore[attr-defined]

    base = app_base_dir()
    candidates.append(base / "tools" / "exiftool.exe")

    # Next to the built executable (helpful for one-dir builds)
    if getattr(sys, "frozen", False):
        candidates.append(Path(sys.executable).resolve().parent / "tools" / "exiftool.exe")

    for candidate in candidates:
        if candidate.exists():
            return str(candidate)

    path = shutil.which("exiftool") or shutil.which("exiftool.exe")
    return path or ""


EXIFTOOL = find_exiftool()



def run_exiftool(args: List[str]) -> Tuple[int, str, str]:
    if not EXIFTOOL:
        return 1, "", "ExifTool not found. Put exiftool.exe in tools/ or install it to PATH."

    cmd = [EXIFTOOL] + args

    startupinfo = None
    creationflags = 0

    if sys.platform == "win32":
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        creationflags = subprocess.CREATE_NO_WINDOW

    p = subprocess.run(
        cmd,
        capture_output=True,
        text=True,
        startupinfo=startupinfo,
        creationflags=creationflags,
    )
    return p.returncode, p.stdout, p.stderr



def scan_files(folder: Path, include_subfolders: bool) -> List[Path]:
    paths = folder.rglob("*") if include_subfolders else folder.glob("*")
    files = [p for p in paths if p.is_file() and p.suffix.lower() in SUPPORTED_EXTS]
    files.sort(key=lambda p: str(p).lower())
    return files


# ---------- Preset storage ----------
def presets_path() -> Path:
    appdata = Path(os.environ.get("APPDATA", str(Path.home())))
    base = appdata / "MetadataEditor"
    base.mkdir(parents=True, exist_ok=True)
    return base / "presets.json"



def load_presets() -> List[Dict[str, str]]:
    path = presets_path()
    if not path.exists():
        return []
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        if isinstance(data, list):
            out: List[Dict[str, str]] = []
            for item in data:
                if isinstance(item, dict) and item.get("name"):
                    out.append(
                        {
                            "name": str(item.get("name", "")).strip(),
                            "creator": str(item.get("creator", "")).strip(),
                            "email": str(item.get("email", "")).strip(),
                            "phone": str(item.get("phone", "")).strip(),
                            "website": str(item.get("website", "")).strip(),
                            "copyright": str(item.get("copyright", "")).strip(),
                        }
                    )
            return out
    except Exception:
        return []
    return []



def save_presets(presets: List[Dict[str, str]]) -> None:
    presets_path().write_text(json.dumps(presets, indent=2, ensure_ascii=False), encoding="utf-8")


# ---------- Threads ----------
class ScanThread(QThread):
    finished_scan = pyqtSignal(list, str)

    def __init__(self, folder: Path, include_subfolders: bool) -> None:
        super().__init__()
        self.folder = folder
        self.include_subfolders = include_subfolders

    def run(self) -> None:
        try:
            self.finished_scan.emit(scan_files(self.folder, self.include_subfolders), "")
        except Exception as exc:
            self.finished_scan.emit([], str(exc))


class ReadMetaThread(QThread):
    finished_read = pyqtSignal(list, str)

    def __init__(self, file_paths: List[Path]) -> None:
        super().__init__()
        self.file_paths = file_paths

    def run(self) -> None:
        try:
            rc, out, err = run_exiftool(["-j", "-G", "-a", "-s"] + [str(p) for p in self.file_paths])
            if rc != 0:
                self.finished_read.emit([], err.strip() or "ExifTool read error")
                return
            data = json.loads(out)
            self.finished_read.emit(data if isinstance(data, list) else [], "")
        except Exception as exc:
            self.finished_read.emit([], str(exc))


class WriteMetaThread(QThread):
    progress = pyqtSignal(int, int, str)
    finished_write = pyqtSignal(bool, str)

    def __init__(
        self,
        files: List[Path],
        tag_updates: Dict[str, str],
        overwrite_original: bool,
        create_undo_backup: bool,
        backup_root: Optional[Path],
        base_folder: Optional[Path],
    ) -> None:
        super().__init__()
        self.files = files
        self.tag_updates = tag_updates
        self.overwrite_original = overwrite_original
        self.create_undo_backup = create_undo_backup
        self.backup_root = backup_root
        self.base_folder = base_folder

    def _backup_one(self, src: Path) -> None:
        if not self.create_undo_backup or not self.backup_root:
            return
        rel = src.name
        try:
            if self.base_folder:
                rel = str(src.relative_to(self.base_folder))
        except Exception:
            rel = src.name
        dst = self.backup_root / rel
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, dst)

    def _build_args(self, file_path: Path) -> List[str]:
        args: List[str] = ["-P", "-m"]
        if self.overwrite_original:
            args.append("-overwrite_original")

        for tag, value in self.tag_updates.items():
            if tag == "XMP:Subject":
                keywords = split_keywords(value)
                args.append("-XMP:Subject=")
                args.append("-IPTC:Keywords=")
                if keywords:
                    for kw in keywords:
                        args.append(f"-XMP:Subject+={kw}")
                        args.append(f"-IPTC:Keywords+={kw}")
                continue
            args.append(f"-{tag}={value}")

        args.append(str(file_path))
        return args

    def _write_one(self, file_path: Path) -> Tuple[bool, str]:
        rc, out, err = run_exiftool(self._build_args(file_path))
        if rc != 0:
            return False, err.strip() or out.strip() or "Write failed"
        return True, out.strip() or "OK"

    def run(self) -> None:
        try:
            if not self.files:
                self.finished_write.emit(False, "No files selected.")
                return
            if not self.tag_updates:
                self.finished_write.emit(False, "No fields selected to apply.")
                return

            total_steps = len(self.files) * (2 if self.create_undo_backup and self.backup_root else 1)
            done = 0

            if self.create_undo_backup and self.backup_root:
                for file_path in self.files:
                    self._backup_one(file_path)
                    done += 1
                    self.progress.emit(done, total_steps, f"Backup: {file_path.name}")

            for file_path in self.files:
                ok, msg = self._write_one(file_path)
                done += 1
                self.progress.emit(done, total_steps, f"Write: {file_path.name}")
                if not ok:
                    self.finished_write.emit(False, f"Failed on {file_path.name}:\n{msg}")
                    return

            self.finished_write.emit(True, "Changes applied successfully.")
        except Exception as exc:
            self.finished_write.emit(False, str(exc))


class RestoreBackupThread(QThread):
    progress = pyqtSignal(int, int, str)
    finished_restore = pyqtSignal(bool, str)

    def __init__(self, backup_root: Path, target_root: Path) -> None:
        super().__init__()
        self.backup_root = backup_root
        self.target_root = target_root

    def run(self) -> None:
        try:
            backup_files = [p for p in self.backup_root.rglob("*") if p.is_file()]
            if not backup_files:
                self.finished_restore.emit(False, "The selected backup folder is empty.")
                return

            total = len(backup_files)
            for i, src in enumerate(backup_files, start=1):
                rel = src.relative_to(self.backup_root)
                dst = self.target_root / rel
                dst.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(src, dst)
                self.progress.emit(i, total, f"Restore: {rel}")

            self.finished_restore.emit(True, f"Restored {len(backup_files)} file(s) from backup.")
        except Exception as exc:
            self.finished_restore.emit(False, str(exc))


# ---------- Preset dialog ----------
class PresetEditorDialog(QDialog):
    def __init__(self, parent: QWidget, title: str, preset: Optional[Dict[str, str]] = None) -> None:
        super().__init__(parent)
        self.setWindowTitle(title)
        self.resize(520, 260)

        self.name = QLineEdit()
        self.creator = QLineEdit()
        self.email = QLineEdit()
        self.phone = QLineEdit()
        self.website = QLineEdit()
        self.copyright = QLineEdit()

        if preset:
            self.name.setText(preset.get("name", ""))
            self.creator.setText(preset.get("creator", ""))
            self.email.setText(preset.get("email", ""))
            self.phone.setText(preset.get("phone", ""))
            self.website.setText(preset.get("website", ""))
            self.copyright.setText(preset.get("copyright", ""))

        form = QFormLayout()
        form.addRow("Preset name*", self.name)
        form.addRow("Creator / Author", self.creator)
        form.addRow("Email", self.email)
        form.addRow("Phone", self.phone)
        form.addRow("Website", self.website)
        form.addRow("Copyright", self.copyright)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)

        layout = QVBoxLayout()
        layout.addLayout(form)
        layout.addWidget(buttons)
        self.setLayout(layout)

    def get_value(self) -> Optional[Dict[str, str]]:
        name = (self.name.text() or "").strip()
        if not name:
            QMessageBox.warning(self, "Missing name", "Preset name is required.")
            return None

        email = (self.email.text() or "").strip()
        if email and not EMAIL_RE.match(email):
            QMessageBox.warning(self, "Invalid email", "Email looks invalid.")
            return None

        website = (self.website.text() or "").strip()
        if website and not URL_RE.match(website):
            QMessageBox.warning(self, "Invalid website", "Website should start with http:// or https://")
            return None

        phone = (self.phone.text() or "").strip()
        if phone and not re.fullmatch(r"[0-9+\-\s().]{5,}", phone):
            QMessageBox.warning(self, "Invalid phone", "Phone contains unusual characters.")
            return None

        return {
            "name": name,
            "creator": (self.creator.text() or "").strip(),
            "email": email,
            "phone": phone,
            "website": website,
            "copyright": (self.copyright.text() or "").strip(),
        }


# ---------- UI ----------
@dataclass
class FieldDef:
    label: str
    tag: str
    placeholder: str = "—"


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()

        if not EXIFTOOL:
            QMessageBox.critical(
                self,
                "ExifTool missing",
                "ExifTool was not found.\n\n"
                "Fix:\n"
                f"1) Put exiftool.exe into:\n   {app_base_dir() / 'tools'}\n"
                "2) Or install ExifTool and add it to PATH.\n",
            )
            raise SystemExit(1)

        self.setWindowTitle("Metadata Editor Pro")
        self.resize(1260, 820)

        self.current_folder: Optional[Path] = None
        self.files: List[Path] = []
        self.current_meta: Dict[str, Any] = {}
        self.presets: List[Dict[str, str]] = load_presets()
        self.last_backup_root: Optional[Path] = None

        self._status = QStatusBar()
        self.setStatusBar(self._status)
        self._status.showMessage("Ready")

        # left panel
        self.folder_title = make_title_label("Project folder")
        self.folder_label = QLabel("No folder selected")
        self.folder_label.setObjectName("statusInfoLabel")
        self.folder_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)

        self.btn_choose_folder = QPushButton("Choose Folder…")
        self.btn_choose_folder.setObjectName("secondaryButton")
        self.btn_scan = QPushButton("Scan")
        self.cb_subfolders = QCheckBox("Include subfolders")
        self.cb_subfolders.setChecked(True)

        self.files_hint = QLabel("Click a file to edit, or select several to edit at the same time.")
        self.files_hint.setObjectName("hintLabel")
        self.files_hint.setWordWrap(True)

        self.file_list = QListWidget()
        self.file_list.setSelectionMode(QListWidget.SelectionMode.ExtendedSelection)

        left_top = QHBoxLayout()
        left_top.addWidget(self.btn_choose_folder)
        left_top.addWidget(self.btn_scan)

        left_layout = QVBoxLayout()
        left_layout.setSpacing(10)
        left_layout.addWidget(self.folder_title)
        left_layout.addLayout(left_top)
        left_layout.addWidget(self.cb_subfolders)
        left_layout.addWidget(self.folder_label)
        left_layout.addWidget(self.files_hint)
        left_layout.addWidget(self.file_list, 1)

        left_widget = QWidget()
        left_widget.setLayout(left_layout)
        left_scroll = wrap_in_scroll_area(left_widget)

        # right panel
        self.selection_label = make_title_label("Select a file to view metadata")
        self.selection_label.setWordWrap(True)

        self.tabs = QTabWidget()

        self.common_fields: List[tuple[str, QLineEdit, QCheckBox]] = []
        self.exif_fields: List[tuple[str, QLineEdit, QCheckBox]] = []
        self.common_tab = QWidget()
        self.exif_tab = QWidget()
        self.presets_tab = QWidget()
        self.raw_tab = QWidget()

        self.common_group = QGroupBox("Common fields")
        self.exif_group = QGroupBox("Camera and EXIF")

        common_form = QFormLayout()
        exif_form = QFormLayout()

        def add_fields(group_form: QFormLayout, store: List[tuple[str, QLineEdit, QCheckBox]], field: FieldDef) -> None:
            editor = QLineEdit()
            editor.setPlaceholderText(field.placeholder)
            apply_cb = QCheckBox("Apply")
            row = QHBoxLayout()
            row.addWidget(editor, 1)
            row.addWidget(apply_cb)
            wrapper = QWidget()
            wrapper.setLayout(row)
            group_form.addRow(field.label, wrapper)
            store.append((field.tag, editor, apply_cb))

        for field in [
            FieldDef("Creator / Author", "XMP:Creator"),
            FieldDef("Description", "XMP:Description"),
            FieldDef("Title", "XMP:Title"),
            FieldDef("Keywords (comma-separated)", "XMP:Subject", "kitchen, living room, exterior"),
            FieldDef("Copyright", "XMP:Rights"),
            FieldDef("Contact Email", "XMP:CreatorWorkEmail", "agent@company.com"),
            FieldDef("Contact Phone", "XMP:CreatorWorkTelephone", "+48 123 456 789"),
            FieldDef("Website", "XMP:CreatorWorkURL", "https://example.com"),
            FieldDef("Date Taken (YYYY:MM:DD HH:MM:SS)", "EXIF:DateTimeOriginal", "2026:03:01 14:25:00"),
        ]:
            add_fields(common_form, self.common_fields, field)

        for field in [
            FieldDef("Camera Make", "EXIF:Make", "Canon / Nikon / Sony"),
            FieldDef("Camera Model", "EXIF:Model", "EOS R5 / Z7 / A7IV"),
            FieldDef("Lens Model", "EXIF:LensModel", "15-35mm / 24-70mm"),
            FieldDef("Serial Number", "EXIF:SerialNumber", "(optional)"),
            FieldDef("ISO", "EXIF:ISO", "100"),
            FieldDef("F-Number", "EXIF:FNumber", "2.8"),
            FieldDef("Exposure Time", "EXIF:ExposureTime", "1/125"),
            FieldDef("Focal Length (mm)", "EXIF:FocalLength", "24"),
            FieldDef("Exposure Compensation", "EXIF:ExposureCompensation", "0"),
        ]:
            add_fields(exif_form, self.exif_fields, field)

        self.common_group.setLayout(common_form)
        common_inner = QWidget()
        common_layout = QVBoxLayout()
        common_layout.setContentsMargins(8, 8, 8, 8)
        common_layout.addWidget(self.common_group)
        common_layout.addStretch(1)
        common_inner.setLayout(common_layout)
        common_outer = QVBoxLayout()
        common_outer.setContentsMargins(0, 0, 0, 0)
        common_outer.addWidget(wrap_in_scroll_area(common_inner))
        self.common_tab.setLayout(common_outer)

        self.exif_group.setLayout(exif_form)
        exif_inner = QWidget()
        exif_layout = QVBoxLayout()
        exif_layout.setContentsMargins(8, 8, 8, 8)
        exif_layout.addWidget(self.exif_group)
        exif_layout.addStretch(1)
        exif_inner.setLayout(exif_layout)
        exif_outer = QVBoxLayout()
        exif_outer.setContentsMargins(0, 0, 0, 0)
        exif_outer.addWidget(wrap_in_scroll_area(exif_inner))
        self.exif_tab.setLayout(exif_outer)

        self.presets_list = QListWidget()
        self.btn_preset_new = QPushButton("New…")
        self.btn_preset_new.setObjectName("secondaryButton")
        self.btn_preset_edit = QPushButton("Edit…")
        self.btn_preset_edit.setObjectName("secondaryButton")
        self.btn_preset_delete = QPushButton("Delete")
        self.btn_preset_delete.setObjectName("dangerButton")
        self.btn_preset_apply = QPushButton("Use preset to fill fields")
        self.cb_preset_tick_apply = QCheckBox("Auto-tick Apply for filled fields")
        self.cb_preset_tick_apply.setChecked(True)

        self.preset_info = QLabel(
            "Create presets once, then apply them to selected photos.\n"
            "This fills Creator, Email, Phone, Website and Copyright in the Common tab."
        )
        self.preset_info.setObjectName("hintLabel")
        self.preset_info.setWordWrap(True)

        preset_buttons = QHBoxLayout()
        preset_buttons.addWidget(self.btn_preset_new)
        preset_buttons.addWidget(self.btn_preset_edit)
        preset_buttons.addWidget(self.btn_preset_delete)
        preset_buttons.addStretch(1)

        preset_apply_row = QHBoxLayout()
        preset_apply_row.addWidget(self.btn_preset_apply)
        preset_apply_row.addStretch(1)
        preset_apply_row.addWidget(self.cb_preset_tick_apply)

        presets_inner = QWidget()
        presets_layout = QVBoxLayout()
        presets_layout.setContentsMargins(8, 8, 8, 8)
        presets_layout.addWidget(make_title_label("Presets"))
        presets_layout.addWidget(self.preset_info)
        presets_layout.addLayout(preset_buttons)
        presets_layout.addWidget(self.presets_list, 1)
        presets_layout.addLayout(preset_apply_row)
        presets_inner.setLayout(presets_layout)
        presets_outer = QVBoxLayout()
        presets_outer.setContentsMargins(0, 0, 0, 0)
        presets_outer.addWidget(wrap_in_scroll_area(presets_inner))
        self.presets_tab.setLayout(presets_outer)

        self.raw_meta = QPlainTextEdit()
        self.raw_meta.setReadOnly(True)
        raw_inner = QWidget()
        raw_layout = QVBoxLayout()
        raw_layout.setContentsMargins(8, 8, 8, 8)
        raw_layout.addWidget(make_title_label("All metadata (read-only)"))
        raw_layout.addWidget(self.raw_meta, 1)
        raw_inner.setLayout(raw_layout)
        raw_outer = QVBoxLayout()
        raw_outer.setContentsMargins(0, 0, 0, 0)
        raw_outer.addWidget(wrap_in_scroll_area(raw_inner))
        self.raw_tab.setLayout(raw_outer)

        self.tabs.addTab(self.common_tab, "Common")
        self.tabs.addTab(self.exif_tab, "Camera & EXIF")
        self.tabs.addTab(self.presets_tab, "Presets")
        self.tabs.addTab(self.raw_tab, "All Metadata")

        self.cb_warn_raw = QCheckBox("Warn before editing RAW files")
        self.cb_warn_raw.setChecked(True)
        self.cb_undo_backup = QCheckBox("Create UNDO backup folder automatically")
        self.cb_undo_backup.setChecked(True)
        self.cb_overwrite_original = QCheckBox("Overwrite original files (no _original backups)")
        self.cb_overwrite_original.setChecked(True)

        self.btn_restore_backup = QPushButton("Restore Last Backup")
        self.btn_restore_backup.setObjectName("secondaryButton")
        self.btn_apply = QPushButton("Apply Changes")

        self.progress = QProgressBar()
        self.progress.setVisible(False)
        self.progress.setValue(0)

        self.status_label = QLabel("")
        self.status_label.setObjectName("statusInfoLabel")
        self.status_label.setWordWrap(True)

        options_row = QHBoxLayout()
        options_row.addWidget(self.cb_warn_raw)
        options_row.addSpacing(12)
        options_row.addWidget(self.cb_undo_backup)
        options_row.addSpacing(12)
        options_row.addWidget(self.cb_overwrite_original)
        options_row.addStretch(1)

        buttons_row = QHBoxLayout()
        buttons_row.addWidget(self.btn_restore_backup)
        buttons_row.addStretch(1)
        buttons_row.addWidget(self.btn_apply)

        right_layout = QVBoxLayout()
        right_layout.setSpacing(10)
        right_layout.addWidget(self.selection_label)
        right_layout.addWidget(self.tabs, 1)
        right_layout.addLayout(options_row)
        right_layout.addLayout(buttons_row)
        right_layout.addWidget(self.progress)
        right_layout.addWidget(self.status_label)

        right_widget = QWidget()
        right_widget.setLayout(right_layout)
        right_scroll = wrap_in_scroll_area(right_widget)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.addWidget(left_scroll)
        splitter.addWidget(right_scroll)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 2)

        root = QVBoxLayout()
        root.setContentsMargins(12, 12, 12, 12)
        root.addWidget(splitter)
        central = QWidget()
        central.setLayout(root)
        self.setCentralWidget(central)

        self.btn_choose_folder.clicked.connect(self.choose_folder)
        self.btn_scan.clicked.connect(self.start_scan)
        self.file_list.itemSelectionChanged.connect(self.on_selection_changed)
        self.btn_apply.clicked.connect(self.apply_changes)
        self.btn_restore_backup.clicked.connect(self.restore_last_backup)

        self.btn_preset_new.clicked.connect(self.preset_new)
        self.btn_preset_edit.clicked.connect(self.preset_edit)
        self.btn_preset_delete.clicked.connect(self.preset_delete)
        self.btn_preset_apply.clicked.connect(self.preset_use)

        self._scan_thread: Optional[ScanThread] = None
        self._read_thread: Optional[ReadMetaThread] = None
        self._write_thread: Optional[WriteMetaThread] = None
        self._restore_thread: Optional[RestoreBackupThread] = None

        self.refresh_presets_list()
        self._update_status_bar()

    # ---------- status ----------
    def _update_status_bar(self) -> None:
        total = len(self.files)
        selected = len(self.selected_files())
        folder = str(self.current_folder) if self.current_folder else "—"
        self._status.showMessage(f"Folder: {folder}    |    Files: {total}    |    Selected: {selected}")

    def _set_busy(self, busy: bool, total: int = 0) -> None:
        widgets = [
            self.btn_apply,
            self.btn_restore_backup,
            self.btn_choose_folder,
            self.btn_scan,
            self.file_list,
            self.presets_list,
        ]
        for widget in widgets:
            widget.setEnabled(not busy)

        self.progress.setVisible(busy)
        if busy:
            self.progress.setRange(0, total)
            self.progress.setValue(0)
        else:
            self.progress.setValue(0)

    def _success_popup(self, title: str, message: str) -> None:
        box = QMessageBox(self)
        box.setWindowTitle(title)
        box.setIcon(QMessageBox.Icon.Information)
        box.setText(f"✅ {message}")
        box.exec()

    # ---------- selection ----------
    def selected_files(self) -> List[Path]:
        out: List[Path] = []
        for item in self.file_list.selectedItems():
            path = item.data(Qt.ItemDataRole.UserRole)
            if path:
                out.append(Path(path))
        return out

    # ---------- presets ----------
    def refresh_presets_list(self) -> None:
        self.presets_list.clear()
        for preset in self.presets:
            item = QListWidgetItem(preset.get("name", "Unnamed"))
            item.setData(Qt.ItemDataRole.UserRole, preset.get("name", ""))
            self.presets_list.addItem(item)

    def _selected_preset_index(self) -> Optional[int]:
        items = self.presets_list.selectedItems()
        if not items:
            return None
        name = items[0].data(Qt.ItemDataRole.UserRole)
        for i, preset in enumerate(self.presets):
            if preset.get("name") == name:
                return i
        return None

    def preset_new(self) -> None:
        dialog = PresetEditorDialog(self, "New preset")
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        value = dialog.get_value()
        if not value:
            return
        if value["name"] in {p.get("name") for p in self.presets}:
            QMessageBox.warning(self, "Name exists", "A preset with this name already exists.")
            return
        self.presets.append(value)
        self.presets.sort(key=lambda p: (p.get("name") or "").lower())
        save_presets(self.presets)
        self.refresh_presets_list()

    def preset_edit(self) -> None:
        idx = self._selected_preset_index()
        if idx is None:
            QMessageBox.information(self, "Select preset", "Select a preset first.")
            return
        current = self.presets[idx]
        dialog = PresetEditorDialog(self, "Edit preset", current)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        value = dialog.get_value()
        if not value:
            return
        if value["name"] != current.get("name") and value["name"] in {p.get("name") for p in self.presets}:
            QMessageBox.warning(self, "Name exists", "A preset with this name already exists.")
            return
        self.presets[idx] = value
        self.presets.sort(key=lambda p: (p.get("name") or "").lower())
        save_presets(self.presets)
        self.refresh_presets_list()

    def preset_delete(self) -> None:
        idx = self._selected_preset_index()
        if idx is None:
            QMessageBox.information(self, "Select preset", "Select a preset first.")
            return
        name = self.presets[idx].get("name", "this preset")
        res = QMessageBox.question(
            self,
            "Delete preset?",
            f"Delete preset '{name}'?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if res != QMessageBox.StandardButton.Yes:
            return
        self.presets.pop(idx)
        save_presets(self.presets)
        self.refresh_presets_list()

    def preset_use(self) -> None:
        idx = self._selected_preset_index()
        if idx is None:
            QMessageBox.information(self, "Select preset", "Select a preset first.")
            return
        preset = self.presets[idx]
        mapping = {
            "creator": "XMP:Creator",
            "email": "XMP:CreatorWorkEmail",
            "phone": "XMP:CreatorWorkTelephone",
            "website": "XMP:CreatorWorkURL",
            "copyright": "XMP:Rights",
        }
        tag_to_widgets = {tag: (editor, cb) for tag, editor, cb in self.common_fields}
        for key, tag in mapping.items():
            value = (preset.get(key) or "").strip()
            if tag in tag_to_widgets:
                editor, cb = tag_to_widgets[tag]
                editor.setText(value)
                if self.cb_preset_tick_apply.isChecked() and value:
                    cb.setChecked(True)
        self.tabs.setCurrentWidget(self.common_tab)
        self.status_label.setText(f"Preset '{preset.get('name')}' loaded into Common fields.")

    # ---------- actions ----------
    def choose_folder(self) -> None:
        folder = QFileDialog.getExistingDirectory(self, "Choose Folder")
        if not folder:
            return
        self.current_folder = Path(folder)
        self.folder_label.setText(str(self.current_folder))
        self.start_scan()

    def start_scan(self) -> None:
        if not self.current_folder:
            QMessageBox.information(self, "Choose folder", "Please choose a folder first.")
            return
        self.status_label.setText("Scanning files…")
        self.file_list.clear()
        self.files = []
        self._update_status_bar()
        self._scan_thread = ScanThread(self.current_folder, self.cb_subfolders.isChecked())
        self._scan_thread.finished_scan.connect(self.on_scan_done)
        self._scan_thread.start()

    def on_scan_done(self, files: list, error: str) -> None:
        if error:
            self.status_label.setText(f"Scan error: {error}")
            return
        self.files = [Path(p) for p in files]
        for path in self.files:
            item = QListWidgetItem(path.name)
            item.setToolTip(str(path))
            item.setData(Qt.ItemDataRole.UserRole, str(path))
            if path.suffix.lower() in RAW_EXTS:
                item.setText(f"[RAW] {path.name}")
            self.file_list.addItem(item)
        self.status_label.setText(f"Found {len(self.files)} files.")
        self.selection_label.setText("Select one or more files to view or edit metadata")
        self._update_status_bar()

    def on_selection_changed(self) -> None:
        selection = self.selected_files()
        self._update_status_bar()
        if not selection:
            self.selection_label.setText("Select a file to view metadata")
            self.raw_meta.clear()
            self._clear_apply_fields()
            return
        if len(selection) == 1:
            self.selection_label.setText(f"1 file selected: {selection[0].name}")
        else:
            self.selection_label.setText(f"{len(selection)} files selected: shared values are shown, differences are marked as {MIXED_TEXT}")
        self.read_metadata(selection)

    def _clear_apply_fields(self) -> None:
        for _tag, editor, _cb in self.common_fields + self.exif_fields:
            editor.clear()

    def read_metadata(self, file_paths: List[Path]) -> None:
        self.status_label.setText("Reading metadata…")
        self.raw_meta.setPlainText("Loading…")
        self._read_thread = ReadMetaThread(file_paths)
        self._read_thread.finished_read.connect(self.on_read_done)
        self._read_thread.start()

    def _merged_value(self, metas: List[Dict[str, Any]], keys: List[str]) -> str:
        found: List[str] = []
        for meta in metas:
            value = ""
            for key in keys:
                if key in meta and meta[key] not in (None, ""):
                    value = stringify_meta_value(meta[key])
                    break
            found.append(value)

        non_empty = [v for v in found if v != ""]
        if not non_empty:
            return ""
        if len(set(non_empty)) == 1 and len(non_empty) == len(found):
            return non_empty[0]
        if len(set(non_empty)) == 1 and len(non_empty) < len(found):
            return MIXED_TEXT
        return MIXED_TEXT

    def on_read_done(self, metas: list, error: str) -> None:
        if error:
            self.status_label.setText(f"Read error: {error}")
            self.raw_meta.clear()
            return
        if not metas:
            self.status_label.setText("No metadata returned.")
            self.raw_meta.clear()
            return

        self.current_meta = metas[0]
        self.raw_meta.setPlainText(json.dumps(metas[0], indent=2, ensure_ascii=False))

        common_map = {
            "XMP:Creator": ["XMP:Creator", "XMP-dc:Creator", "EXIF:Artist", "IPTC:By-line"],
            "XMP:Description": ["XMP:Description", "XMP-dc:Description", "IPTC:Caption-Abstract", "EXIF:ImageDescription"],
            "XMP:Title": ["XMP:Title", "XMP-dc:Title", "IPTC:ObjectName"],
            "XMP:Subject": ["XMP:Subject", "IPTC:Keywords"],
            "XMP:Rights": ["XMP:Rights", "XMP-dc:Rights", "EXIF:Copyright", "IPTC:CopyrightNotice"],
            "XMP:CreatorWorkEmail": ["XMP:CreatorWorkEmail"],
            "XMP:CreatorWorkTelephone": ["XMP:CreatorWorkTelephone"],
            "XMP:CreatorWorkURL": ["XMP:CreatorWorkURL"],
            "EXIF:DateTimeOriginal": ["EXIF:DateTimeOriginal", "Composite:DateTimeOriginal"],
        }
        exif_map = {
            "EXIF:Make": ["EXIF:Make"],
            "EXIF:Model": ["EXIF:Model"],
            "EXIF:LensModel": ["EXIF:LensModel", "EXIF:LensID", "Composite:LensID"],
            "EXIF:SerialNumber": ["EXIF:SerialNumber", "EXIF:BodySerialNumber"],
            "EXIF:ISO": ["EXIF:ISO"],
            "EXIF:FNumber": ["EXIF:FNumber"],
            "EXIF:ExposureTime": ["EXIF:ExposureTime"],
            "EXIF:FocalLength": ["EXIF:FocalLength"],
            "EXIF:ExposureCompensation": ["EXIF:ExposureCompensation", "EXIF:ExposureBiasValue"],
        }

        for tag, editor, _cb in self.common_fields:
            editor.setText(self._merged_value(metas, common_map.get(tag, [tag])))
        for tag, editor, _cb in self.exif_fields:
            editor.setText(self._merged_value(metas, exif_map.get(tag, [tag])))

        if len(metas) > 1:
            self.status_label.setText("Shared values are shown. Differing values are marked as (mixed values).")
        else:
            self.status_label.setText("Metadata loaded. Tick Apply next to fields you want to write.")

    def label_for_tag(self, tag: str) -> str:
        mapping = {
            "XMP:Creator": "Creator / Author",
            "XMP:Description": "Description",
            "XMP:Title": "Title",
            "XMP:Subject": "Keywords",
            "XMP:Rights": "Copyright",
            "XMP:CreatorWorkEmail": "Contact Email",
            "XMP:CreatorWorkTelephone": "Contact Phone",
            "XMP:CreatorWorkURL": "Website",
            "EXIF:DateTimeOriginal": "Date Taken",
            "EXIF:Make": "Camera Make",
            "EXIF:Model": "Camera Model",
            "EXIF:LensModel": "Lens Model",
            "EXIF:SerialNumber": "Serial Number",
            "EXIF:ISO": "ISO",
            "EXIF:FNumber": "F-Number",
            "EXIF:ExposureTime": "Exposure Time",
            "EXIF:FocalLength": "Focal Length",
            "EXIF:ExposureCompensation": "Exposure Compensation",
        }
        return mapping.get(tag, tag)

    def apply_changes(self) -> None:
        selection = self.selected_files()
        if not selection:
            QMessageBox.information(self, "No selection", "Select one or more files first.")
            return

        if self.cb_warn_raw.isChecked():
            raw_files = [p for p in selection if p.suffix.lower() in RAW_EXTS]
            if raw_files:
                msg = (
                    "RAW file(s) detected.\n\n"
                    "Editing RAW files can be risky in some workflows.\n"
                    "This app writes metadata directly into the file.\n\n"
                    f"RAW files selected: {len(raw_files)}\n\n"
                    "Proceed at your own risk?"
                )
                res = QMessageBox.warning(
                    self,
                    "RAW File Warning",
                    msg,
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                )
                if res != QMessageBox.StandardButton.Yes:
                    return

        updates: Dict[str, str] = {}
        applied_fields: List[tuple[str, str, str]] = []

        def gather(fields: List[tuple[str, QLineEdit, QCheckBox]]) -> None:
            for tag, editor, cb in fields:
                if cb.isChecked():
                    value = editor.text().strip()
                    if value == MIXED_TEXT:
                        continue
                    updates[tag] = value
                    applied_fields.append((tag, self.label_for_tag(tag), value))

        gather(self.common_fields)
        gather(self.exif_fields)

        if not updates:
            QMessageBox.information(self, "Nothing to apply", "Tick Apply next to at least one field that is not mixed.")
            return

        empties = [(tag, label) for tag, label, value in applied_fields if value == ""]
        if empties:
            names = "\n".join(f"• {label}" for _tag, label in empties)
            res = QMessageBox.question(
                self,
                "Clear fields?",
                "You ticked Apply on empty fields.\n\n"
                "That will clear these tags in the selected files:\n"
                f"{names}\n\n"
                "Do you want to continue?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if res != QMessageBox.StandardButton.Yes:
                return

        errors: List[str] = []
        for tag, label, value in applied_fields:
            if value == "":
                continue
            error = validate_tag_value(tag, value)
            if error:
                errors.append(f"{label}: {error}")
        if errors:
            QMessageBox.warning(self, "Fix these fields", "Some values look invalid:\n\n" + "\n".join(f"• {e}" for e in errors))
            return

        if len(selection) > 1:
            res = QMessageBox.question(
                self,
                "Apply to multiple files?",
                f"You are about to apply changes to {len(selection)} files. Continue?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if res != QMessageBox.StandardButton.Yes:
                return

        backup_root: Optional[Path] = None
        if self.cb_undo_backup.isChecked():
            base = self.current_folder if self.current_folder else selection[0].parent
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_root = base / "_UNDO_METADATA_BACKUP" / stamp
            self.last_backup_root = backup_root

        self._set_busy(True, len(selection) * (2 if backup_root else 1))
        self.status_label.setText("Applying changes…")

        self._write_thread = WriteMetaThread(
            files=selection,
            tag_updates=updates,
            overwrite_original=self.cb_overwrite_original.isChecked(),
            create_undo_backup=self.cb_undo_backup.isChecked(),
            backup_root=backup_root,
            base_folder=self.current_folder,
        )
        self._write_thread.progress.connect(self.on_write_progress)
        self._write_thread.finished_write.connect(self.on_write_done)
        self._write_thread.start()

    def on_write_progress(self, done: int, total: int, step_name: str) -> None:
        self.progress.setRange(0, total)
        self.progress.setValue(done)
        self.status_label.setText(f"{step_name} ({done}/{total})")

    def on_write_done(self, ok: bool, message: str) -> None:
        self._set_busy(False)
        if ok:
            self.status_label.setText(message)
            self._success_popup("Success", "Changes applied successfully.")
            selected = self.selected_files()
            if selected:
                self.read_metadata(selected)
        else:
            self.status_label.setText(f"Write error: {message}")
            QMessageBox.warning(self, "Write error", message)
        self._update_status_bar()

    def restore_last_backup(self) -> None:
        if not self.current_folder:
            QMessageBox.information(self, "Choose folder", "Choose the working folder first.")
            return

        backup_parent = self.current_folder / "_UNDO_METADATA_BACKUP"
        if not backup_parent.exists():
            QMessageBox.information(self, "No backup", "No backup folder was found for the current project folder.")
            return

        backup_candidates = [p for p in backup_parent.iterdir() if p.is_dir()]
        if not backup_candidates:
            QMessageBox.information(self, "No backup", "No saved backup batches were found.")
            return

        latest = max(backup_candidates, key=lambda p: p.name)
        res = QMessageBox.question(
            self,
            "Restore latest backup?",
            f"This will restore files from:\n{latest}\n\nand overwrite current files with the backed-up versions. Continue?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if res != QMessageBox.StandardButton.Yes:
            return

        count = len([p for p in latest.rglob("*") if p.is_file()])
        self._set_busy(True, count)
        self.status_label.setText("Restoring backup…")
        self._restore_thread = RestoreBackupThread(latest, self.current_folder)
        self._restore_thread.progress.connect(self.on_restore_progress)
        self._restore_thread.finished_restore.connect(self.on_restore_done)
        self._restore_thread.start()

    def on_restore_progress(self, done: int, total: int, step_name: str) -> None:
        self.progress.setRange(0, total)
        self.progress.setValue(done)
        self.status_label.setText(f"{step_name} ({done}/{total})")

    def on_restore_done(self, ok: bool, message: str) -> None:
        self._set_busy(False)
        if ok:
            self.status_label.setText(message)
            self._success_popup("Restore complete", message)
            if self.current_folder:
                self.start_scan()
        else:
            self.status_label.setText(f"Restore error: {message}")
            QMessageBox.warning(self, "Restore error", message)
        self._update_status_bar()


def main() -> None:
    app = QApplication(sys.argv)
    apply_professional_theme(app)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
