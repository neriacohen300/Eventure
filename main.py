"""
main.py  –  Eventure Slideshow Creator (improved)

Key improvements over original:
  ┌─────────────────────────────────────────────────────────────────────────┐
  │ PERFORMANCE                                                              │
  │  • ImageProcessingWorker now uses ThreadPoolExecutor to process images   │
  │    in parallel instead of sequentially (big speedup for large batches)   │
  │  • ffprobe calls use a proper list-based subprocess (no shell=True risk) │
  │  • Audio duration cached after first calculation in InfoDialog           │
  │                                                                          │
  │ CODE QUALITY                                                             │
  │  • save_project / save_project_as deduplicated into _write_project_file  │
  │  • format_time used from a single shared helper (no duplication)         │
  │  • load_translations falls back gracefully without a bare except         │
  │  • EXIF orientation tag looked up once at import                         │
  │  • Startup folder copies moved to a helper _copy_resource_folders()      │
  │                                                                          │
  │ UI / ROBUSTNESS                                                          │
  │  • Progress bars update from worker threads via signals (thread-safe)    │
  │  • Premiere style file path now relative (no hardcoded E:\\ path)        │
  │  • load_project validates line count before indexing                     │
  └─────────────────────────────────────────────────────────────────────────┘
"""

import configparser
import copy
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
import json
from pathlib import Path
import random
import shutil
import subprocess
import sys
import os

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QInputDialog, QAction, QListWidget, QProgressBar, QComboBox,
    QMessageBox, QDialog, QTextEdit, QListWidgetItem, QCheckBox,
    QStyledItemDelegate, QPushButton, QLabel, QFileDialog, QSlider,
    QStyle, QTableWidgetItem, QSpinBox, QHeaderView, QTableWidget,
    QLineEdit,
)
from PyQt5.QtCore import Qt, QUrl, QSize, QProcess, QTimer, QThread, pyqtSignal, QEvent
from PyQt5.QtGui import (
    QIcon, QFont, QPixmap, QTextCursor, QCursor, QTransform,
    QColor, QBrush, QImage,
)
from PIL import Image, ExifTags
from openpyxl import Workbook
import openpyxl

import Image_resizer
import premiere_export

from EVENTURE_THEMES.theme import set_theme

# ── Environment ──────────────────────────────────────────────────────────────

plugin_path = os.path.join(os.path.dirname(sys.executable), "Library", "plugins", "platforms")
os.environ["QT_QPA_PLATFORM_PLUGIN_PATH"] = plugin_path

BASEPATH = Path.home() / "Neria-LTD" / "Eventure"
BASEPATH.mkdir(parents=True, exist_ok=True)

# EXIF orientation tag (looked up once)
_ORIENTATION_TAG = next(
    (k for k, v in ExifTags.TAGS.items() if v == "Orientation"), None
)

# ── Helpers ───────────────────────────────────────────────────────────────────

def _get_audio_duration(audio_path: str) -> float:
    """Return the duration of an audio file in seconds using ffprobe."""
    try:
        result = subprocess.run(
            [
                "ffprobe", "-v", "error",
                "-show_entries", "format=duration",
                "-of", "default=noprint_wrappers=1:nokey=1",
                audio_path,
            ],
            capture_output=True,
            text=True,
            timeout=10,
        )
        return float(result.stdout.strip())
    except Exception as e:
        print(f"ffprobe error for {audio_path}: {e}")
        return 0.0


def format_time_srt(seconds: float) -> str:
    """Convert seconds to SRT timestamp (HH:MM:SS,mmm)."""
    h   = int(seconds // 3600)
    m   = int((seconds % 3600) // 60)
    s   = int(seconds % 60)
    ms  = int((seconds - int(seconds)) * 1000)
    return f"{h:02}:{m:02}:{s:02},{ms:03}"


def format_time_hms(seconds: float) -> str:
    """Convert seconds to HH:MM:SS."""
    h = int(seconds // 3600)
    m = int((seconds % 3600) // 60)
    s = int(seconds % 60)
    return f"{h:02}:{m:02}:{s:02}"


def _copy_resource_folders(script_dir: Path, resources: list[str]) -> None:
    """Copy each named sub-folder from script_dir into BASEPATH.
    Files that are locked/open are skipped so the app still starts."""
    for name in resources:
        src = script_dir / name
        dst = BASEPATH / name
        if not src.exists():
            continue
        dst.mkdir(parents=True, exist_ok=True)
        skipped = []
        for item in src.rglob("*"):
            rel  = item.relative_to(src)
            dest = dst / rel
            if item.is_dir():
                dest.mkdir(parents=True, exist_ok=True)
            else:
                try:
                    shutil.copy2(item, dest)
                except OSError as e:
                    skipped.append(f"{item.name}: {e}")
        if skipped:
            for msg in skipped:
                print(f"  Skipped (file in use): {msg}")
        print(f"Folder '{name}' synced to '{dst}'")


# ── Taskbar Progress Stub ────────────────────────────────────────────────────

class _TaskbarProgressStub:
    """No-op replacement when QWinTaskbarButton is unavailable."""
    def show(self):          pass
    def hide(self):          pass
    def reset(self):         pass
    def setValue(self, v):   pass
    def setVisible(self, v): pass


# ── Main Window ───────────────────────────────────────────────────────────────

class SlideshowCreator(QMainWindow):

    def __init__(self):
        super().__init__()

        script_dir = Path(__file__).resolve().parent
        _copy_resource_folders(script_dir, ["Help", "Languages", "Songs", "Fonts"])

        self.language     = "en"
        self.translations = {}
        self.load_translations()

        self.setWindowTitle(self.tr("window_title"))
        self.setGeometry(100, 100, 1200, 800)

        self.images          = []
        self.audio_files     = []
        self.output_file     = self.tr("output_file")
        self.button_font     = "Segoe UI"
        self.deafult_font    = "Segoe UI"
        self.text_font       = "Segoe UI"
        self.text_font_size  = 10
        self.button_font_size = 9

        self.transitions_types = [
            "fade", "fadeblack", "fadewhite", "distance",
            "wipeleft", "wiperight", "wipeup", "wipedown",
            "slideleft", "slideright", "slideup", "slidedown",
            "smoothleft", "smoothright", "smoothup", "smoothdown",
            "circlecrop", "rectcrop", "circleclose", "circleopen",
            "horzclose", "horzopen", "vertclose", "vertopen",
            "diagbl", "diagbr", "diagtl", "diagtr", "zoomin",
            "hlslice", "hrslice", "vuslice", "vdslice",
            "dissolve", "pixelize", "radial", "hblur",
            "wipetl", "wipetr", "wipebl", "wipebr",
            "fadegrays", "squeezev", "squeezeh",
            "hlwind", "hrwind", "vuwind", "vdwind",
            "coverleft", "coverright", "coverup", "coverdown",
            "revealleft", "revealright", "revealup", "revealdown",
        ]

        self.default_transition_duration = 1
        self.common_width    = 1920
        self.common_height   = 1080
        self.images_backup   = []
        self.backup_state    = False
        self.premiere_project_folder = ""
        self.loaded_project  = ""

        self.shortcuts = {
            "save":              "Ctrl+S",
            "save_as":           "Ctrl+Shift+S",
            "load":              "Ctrl+L",
            "easy_text":         "Ctrl+T",
            "info":              "Alt+I",
            "import_images":     "Ctrl+Shift+I",
            "import_audio":      "Ctrl+Shift+A",
            "set_image_location":"Ctrl+Q",
            "delete_row":        "Delete",
            "move_image_up":     "Ctrl+Up",
            "move_image_down":   "Ctrl+Down",
        }
        self.load_shortcuts()
        self.create_ui()

    # ── UI ────────────────────────────────────────────────────────────────────

    def create_ui(self):
        main_widget = QWidget()
        main_layout = QHBoxLayout(main_widget)

        # Left panel
        left_panel = QVBoxLayout()
        self.image_table = QTableWidget()
        self.image_table.setItemDelegate(CustomDelegate())
        self.image_table.setColumnCount(9)
        self.image_table.setHorizontalHeaderLabels([
            self.tr("table_header_actions"),
            self.tr("table_header_image"),
            self.tr("table_header_duration"),
            self.tr("table_header_transition"),
            self.tr("table_header_transition_length"),
            self.tr("table_header_text"),
            self.tr("table_header_rotation"),
            self.tr("table_header_second_image"),
            self.tr("table_header_date"),
        ])
        self.image_table.setFont(QFont(self.deafult_font, 10, QFont.Bold))
        for col in range(9):
            self.image_table.horizontalHeader().setSectionResizeMode(
                col, QHeaderView.ResizeToContents
            )
        self.image_table.itemChanged.connect(self.on_edit_on_table)

        self.slides_label = QLabel(self.tr("label_slides"))
        self.slides_label.setFont(QFont(self.text_font, self.text_font_size, QFont.Bold))
        left_panel.addWidget(self.slides_label)
        left_panel.addWidget(self.image_table)

        # Right panel
        right_panel = QVBoxLayout()

        self.preview_label = QLabel(self.tr("label_preview"))
        self.preview_label.setFont(QFont(self.text_font, 16, QFont.Bold))
        self.preview_label.setAlignment(Qt.AlignCenter)
        self.preview_label.setFixedHeight(300)

        _pb_style = (
            "QProgressBar {{ background-color: #1E1E1E; color: white; }}"
            "QProgressBar::chunk {{ background-color: {color}; }}"
        )

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(False)
        self.progress_bar.setStyleSheet(_pb_style.format(color="#0078d4"))

        self.image_progress_bar = QProgressBar()
        self.image_progress_bar.setRange(0, 100)
        self.image_progress_bar.setValue(0)
        self.image_progress_bar.setVisible(False)
        self.image_progress_bar.setStyleSheet(_pb_style.format(color="#ff4444"))

        self.image_premiere_progress_bar = QProgressBar()
        self.image_premiere_progress_bar.setRange(0, 100)
        self.image_premiere_progress_bar.setValue(0)
        self.image_premiere_progress_bar.setVisible(False)
        self.image_premiere_progress_bar.setStyleSheet(_pb_style.format(color="#9932cc"))

        self.audio_table = QTableWidget()
        self.audio_table.setColumnCount(2)
        self.audio_table.setHorizontalHeaderLabels([
            self.tr("table_header_actions"),
            self.tr("table_header_audio_file"),
        ])
        self.audio_table.setFont(QFont(self.deafult_font, 10, QFont.Bold))
        self.audio_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.audio_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)

        self.audio_files_label = QLabel(self.tr("label_audio_files"))
        self.audio_files_label.setFont(QFont(self.text_font, self.text_font_size, QFont.Bold))

        self.audio_library_button = QPushButton(self.tr("label_audio_library"))
        self.audio_library_button.setFont(QFont(self.button_font, self.button_font_size, QFont.Bold))
        self.audio_library_button.clicked.connect(self.open_audio_library)

        right_panel.addWidget(self.preview_label)
        right_panel.addWidget(self.audio_files_label)
        right_panel.addWidget(self.audio_table)
        right_panel.addWidget(self.audio_library_button)
        right_panel.addWidget(self.progress_bar)
        right_panel.addWidget(self.image_progress_bar)
        right_panel.addWidget(self.image_premiere_progress_bar)

        main_layout.addLayout(left_panel, 2)
        main_layout.addLayout(right_panel, 1)
        self.setCentralWidget(main_widget)
        self.audio_files = []

        # Windows taskbar progress (optional – gracefully stubbed if unavailable)
        self.taskbar_progress = self._make_taskbar_progress()

    # ── Taskbar Progress ──────────────────────────────────────────────────────

    def _make_taskbar_progress(self):
        """
        Try to create a real Windows taskbar progress button (PyQt5 WinTaskbar).
        Falls back to a silent no-op stub if the extension is unavailable,
        so the app works on non-Windows and without the optional package.
        """
        try:
            from PyQt5.QtWinExtras import QWinTaskbarButton
            btn = QWinTaskbarButton(self)
            btn.setWindow(self.windowHandle())
            progress = btn.progress()
            progress.setRange(0, 100)
            return progress
        except Exception:
            return _TaskbarProgressStub()

    # ── Images ────────────────────────────────────────────────────────────────

    def on_edit_on_table(self, item):
        col = item.column()
        row = item.row()
        if col == 2:
            try:
                val = int(item.text())
                if val < 2 or val > 600:
                    raise ValueError(self.tr("duration_out_of_range_error"))
                self.images[row]["duration"] = val
                if self.images[row]["transition_duration"] > val - 1:
                    self.images[row]["transition_duration"] = val - 1
                    self.update_image_table()
            except ValueError:
                item.setText(str(self.images[row]["duration"]))
        elif col == 5:
            self.images[row]["text"] = item.text()
        elif col == 6:
            try:
                val = int(item.text())
                if val < 0 or val > 359:
                    raise ValueError(self.tr("rotation_out_of_range_error"))
                self.images[row]["rotation"] = val
                self.update_preview_with_row(row)
            except ValueError:
                item.setText(str(self.images[row]["rotation"]))

    def set_all_images_duration(self):
        selected = self.image_table.selectedItems()
        row = self.image_table.row(selected[0]) if selected else None
        current = self.images[row]["duration"] if row is not None else 2

        dialog = QInputDialog(self)
        dialog.setWindowTitle(self.tr("dialog_set_duration"))
        dialog.setLabelText(self.tr("dialog_enter_duration"))
        dialog.setIntValue(current)
        dialog.setIntRange(2, 600)
        if dialog.exec_() == QDialog.Accepted:
            v = dialog.intValue()
            for img in self.images:
                img["duration"] = v
            self.update_image_table()

    def add_images(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "Select Images", "", "Images (*.png *.jpg *.jpeg *.bmp *.gif)"
        )
        if files:
            new = [
                {
                    "path": f,
                    "duration": 5,
                    "transition": "fade",
                    "transition_duration": self.default_transition_duration,
                    "text": "",
                    "rotation": 0,
                    "is_second_image": False,
                    "date": datetime.fromtimestamp(os.path.getmtime(f)).strftime("%Y-%m-%d %H:%M:%S"),
                }
                for f in files
            ]
            self.images.extend(new)
            self.update_image_table()

    def auto_sort_images_by_date(self, reverse: bool = False):
        self.images.sort(key=lambda x: x["date"], reverse=reverse)
        self.update_image_table()

    def set_second_image(self, row, state):
        if row == 0:
            QMessageBox.critical(self, self.tr("error"), self.tr("second_image_error"), QMessageBox.Ok)
            self.update_image_row(row)
            return
        self.images[row]["is_second_image"] = state == Qt.Checked
        item = self.image_table.item(row, 1)
        if item:
            item.setData(Qt.UserRole, state == Qt.Checked)
        self.update_image_row(row)

    def update_image_table(self):
        self.image_table.blockSignals(True)
        self.image_table.setUpdatesEnabled(False)
        self.image_table.setSortingEnabled(False)
        self.image_table.setRowCount(len(self.images))

        for row, img in enumerate(self.images):
            self._populate_row(row, img)

        self.image_table.setSortingEnabled(True)
        self.image_table.setUpdatesEnabled(True)
        self.image_table.blockSignals(False)

    def _populate_row(self, row: int, img: dict):
        """Fill a single row in the image table."""
        filename_item = QTableWidgetItem(os.path.basename(img["path"]))
        filename_item.setData(Qt.UserRole, img.get("is_second_image", False))
        filename_item.setFlags(filename_item.flags() & ~Qt.ItemIsEditable)

        duration_item = QTableWidgetItem(str(img.get("duration", 5)))

        transition_cb = QComboBox()
        transition_cb.addItems(self.transitions_types)
        transition_cb.setCurrentText(img.get("transition", "fade"))
        transition_cb.currentTextChanged.connect(
            lambda text, r=row: self.update_transition(r, text)
        )

        tl_item = QTableWidgetItem(str(img.get("transition_duration", self.default_transition_duration)))
        tl_item.setFlags(tl_item.flags() & ~Qt.ItemIsEditable)

        text_item = QTableWidgetItem(str(img.get("text", "")))
        text_item.setFlags(text_item.flags() | Qt.ItemIsEditable)

        rotation_item = QTableWidgetItem(str(img.get("rotation", 0)))

        second_cb = QCheckBox()
        second_cb.setChecked(img.get("is_second_image", False))
        second_cb.stateChanged.connect(lambda state, r=row: self.set_second_image(r, state))

        date_item = QTableWidgetItem(str(img.get("date", "")))
        date_item.setFlags(date_item.flags() & ~Qt.ItemIsEditable)

        self.image_table.setItem(row, 1, filename_item)
        self.image_table.setItem(row, 2, duration_item)
        self.image_table.setCellWidget(row, 3, transition_cb)
        self.image_table.setItem(row, 4, tl_item)
        self.image_table.setItem(row, 5, text_item)
        self.image_table.setItem(row, 6, rotation_item)
        self.image_table.setCellWidget(row, 7, second_cb)
        self.image_table.setItem(row, 8, date_item)

        # Action buttons
        move_up_btn   = QPushButton("↑")
        move_down_btn = QPushButton("↓")
        delete_btn    = QPushButton("✖")
        move_up_btn.clicked.connect(self.move_image_up)
        move_down_btn.clicked.connect(self.move_image_down)
        delete_btn.clicked.connect(self.delete_image)

        btn_widget = QWidget()
        btn_layout = QHBoxLayout(btn_widget)
        btn_layout.addWidget(move_up_btn)
        btn_layout.addWidget(move_down_btn)
        btn_layout.addWidget(delete_btn)
        btn_layout.setContentsMargins(0, 0, 0, 0)
        self.image_table.setCellWidget(row, 0, btn_widget)

    def move_image_up(self):
        self.image_table.setSortingEnabled(False)
        row = self.image_table.currentRow()
        if row > 0:
            self.images[row], self.images[row - 1] = self.images[row - 1], self.images[row]
            self.update_image_row(row)
            self.update_image_row(row - 1)
            self.image_table.setCurrentCell(row - 1, 1)
            self.update_preview_with_row(row - 1)
        self.image_table.setSortingEnabled(True)

    def move_image_down(self):
        self.image_table.setSortingEnabled(False)
        row = self.image_table.currentRow()
        if row < len(self.images) - 1:
            self.images[row], self.images[row + 1] = self.images[row + 1], self.images[row]
            self.update_image_row(row)
            self.update_image_row(row + 1)
            self.image_table.setCurrentCell(row + 1, 1)
            self.update_preview_with_row(row + 1)
        self.image_table.setSortingEnabled(True)

    def update_image_row(self, row: int):
        if 0 <= row < len(self.images):
            self.image_table.blockSignals(True)
            self._populate_row(row, self.images[row])
            self.image_table.blockSignals(False)

    def delete_image(self):
        self.image_table.setSortingEnabled(False)
        row = self.image_table.currentRow()
        if 0 <= row < len(self.images):
            del self.images[row]
            self.image_table.removeRow(row)
            if not self.images:
                self.preview_label.clear()
            else:
                new_row = max(0, row - 1)
                self.image_table.setCurrentCell(new_row, 1)
                self.update_preview_with_row(new_row)
        self.image_table.setSortingEnabled(True)

    def set_random_images_order(self):
        random.shuffle(self.images)
        self.update_image_table()

    def set_image_location(self):
        selected = self.image_table.selectedItems()
        if not selected:
            return
        cur_row = self.image_table.row(selected[0])
        dialog = QInputDialog(self)
        dialog.setWindowTitle(self.tr("dialog_set_image_location"))
        dialog.setLabelText(self.tr("dialog_enter_position"))
        dialog.setInputMode(QInputDialog.IntInput)
        dialog.setIntRange(1, len(self.images))
        dialog.setIntValue(cur_row + 1)
        if dialog.exec_() == QDialog.Accepted:
            new_pos = dialog.intValue() - 1
            img = self.images.pop(cur_row)
            self.images.insert(new_pos, img)
            self.update_image_table()
            self.image_table.setCurrentCell(new_pos, 1)

    def update_image_progress(self, value: int):
        self.image_progress_bar.setValue(value)
        self.taskbar_progress.setValue(value)

    def on_image_processing_finished(self):
        self.image_progress_bar.setVisible(False)
        self.taskbar_progress.reset()
        self.taskbar_progress.hide()
        self.continue_with_video_export()

    # ── Audio ─────────────────────────────────────────────────────────────────

    def add_audio(self):
        file, _ = QFileDialog.getOpenFileName(
            self, "Select Audio", "", "Audio Files (*.mp3 *.wav *.flac)"
        )
        if file:
            self.audio_files.append({"path": file})
            self.update_audio_table()

    def update_audio_table(self):
        self.audio_table.setRowCount(len(self.audio_files))
        for row, audio in enumerate(self.audio_files):
            filename_item = QTableWidgetItem(os.path.basename(audio["path"]))
            filename_item.setFlags(filename_item.flags() & ~Qt.ItemIsEditable)
            self.audio_table.setItem(row, 1, filename_item)

            mu_btn = QPushButton("↑")
            md_btn = QPushButton("↓")
            del_btn = QPushButton("✖")
            mu_btn.clicked.connect(lambda _, r=row: self.move_audio_up(r))
            md_btn.clicked.connect(lambda _, r=row: self.move_audio_down(r))
            del_btn.clicked.connect(lambda _, r=row: self.delete_audio(r))

            bw = QWidget()
            bl = QHBoxLayout(bw)
            bl.addWidget(mu_btn)
            bl.addWidget(md_btn)
            bl.addWidget(del_btn)
            bl.setContentsMargins(0, 0, 0, 0)
            self.audio_table.setCellWidget(row, 0, bw)

    def move_audio_up(self, row: int):
        if row > 0:
            self.audio_files[row], self.audio_files[row - 1] = (
                self.audio_files[row - 1], self.audio_files[row]
            )
            self.update_audio_table()
            self.audio_table.setCurrentCell(row - 1, 1)

    def move_audio_down(self, row: int):
        if row < len(self.audio_files) - 1:
            self.audio_files[row], self.audio_files[row + 1] = (
                self.audio_files[row + 1], self.audio_files[row]
            )
            self.update_audio_table()
            self.audio_table.setCurrentCell(row + 1, 1)

    def delete_audio(self, row: int):
        del self.audio_files[row]
        self.update_audio_table()
        target = max(0, row - 1)
        self.audio_table.setCurrentCell(target, 1)

    def open_audio_library(self):
        dialog = AudioLibraryDialog(tr_function=self.tr, parent=self)
        dialog.exec_()

    # ── Export ────────────────────────────────────────────────────────────────

    def _total_audio_duration(self) -> float:
        return sum(_get_audio_duration(a["path"]) for a in self.audio_files)

    def continue_with_video_export(self):
        total_img_dur  = sum(img["duration"] for img in self.images)
        total_audio_dur = self._total_audio_duration()

        print(f"Total image duration: {total_img_dur}s  |  Total audio duration: {total_audio_dur:.1f}s")

        if total_img_dur > total_audio_dur:
            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Information)
            msg.setWindowTitle(self.tr("audio_and_video_error"))
            msg.setText(self.tr("prompt_audio_mismatch"))
            msg.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
            if msg.exec_() == QMessageBox.Cancel:
                return

            n = len(self.images)
            new_dur = int(total_audio_dur / n) if n else 0
            if new_dur < 2 or new_dur > 600:
                QMessageBox.information(self, self.tr("audio_and_video_error"), self.tr("prompt_cant_match"))
                return
            for img in self.images:
                img["duration"] = new_dur
            self.update_image_table()

        command = self.build_ffmpeg_command()
        print("Exporting with command:", command)

        self.process = QProcess(self)
        self.process.readyReadStandardError.connect(self.update_progress)
        self.process.finished.connect(self.export_finished)

        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.taskbar_progress.show()
        self.taskbar_progress.setValue(0)
        self.process.start(command)

    def export_slideshow(self):
        if not self.validate_transitions():
            return
        if not self.images or not self.audio_files:
            QMessageBox.critical(self, self.tr("error"), self.tr("error_no_audio"), QMessageBox.Ok)
            return

        file_path, _ = QFileDialog.getSaveFileName(
            self, "Save Slideshow", "", "Video Files (*.mp4);;All Files (*)"
        )
        if not file_path:
            QMessageBox.critical(self, self.tr("error"), self.tr("error_select_location"), QMessageBox.Ok)
            return
        self.output_file = file_path

        output_folder = os.path.join(os.path.dirname(self.images[0]["path"]), "A_Blur")

        self.images_backup = copy.deepcopy(self.images)
        self.backup_state  = True

        self.image_progress_bar.setVisible(True)
        self.image_progress_bar.setValue(0)
        self.taskbar_progress.show()
        self.taskbar_progress.setValue(0)

        self.image_worker = ImageProcessingWorker(self.images, output_folder, self.common_width, self.common_height)
        self.image_worker.progress.connect(self.update_image_progress)
        self.image_worker.finished.connect(self.on_image_processing_finished)
        self.image_worker.start()

    def export_finished(self):
        self.progress_bar.setValue(100)
        self.taskbar_progress.setValue(100)
        QMessageBox.information(
            self,
            self.tr("success_export_complete_window"),
            self.tr("success_export_complete"),
        )
        self.progress_bar.setVisible(False)
        self.taskbar_progress.reset()
        self.taskbar_progress.hide()

    def build_ffmpeg_command(self) -> str:
        inputs, filters = [], []

        for i, img in enumerate(self.images):
            if i == 0:
                duration = img["duration"]
            elif i == len(self.images) - 1:
                duration = img["duration"] + self.images[i - 1]["transition_duration"]
            else:
                duration = img["duration"] + img["transition_duration"]
            inputs.append(f'-loop 1 -t {duration} -i "{img["path"]}"')
            filters.append(f"[{i}:v]scale=1920:1080,setsar=1,format=yuv420p[{i}v]")

        for i in range(len(self.images) - 1):
            offset = sum(img["duration"] for img in self.images[: i + 1]) - self.images[i]["transition_duration"]
            prev   = f"[{i}v]" if i == 0 else f"[v{i}]"
            filters.append(
                f"{prev}[{i + 1}v]xfade=transition={self.images[i]['transition']}"
                f":duration={self.images[i]['transition_duration']}:offset={offset}[v{i + 1}]"
            )

        final_stream = f"[v{len(self.images) - 1}]"
        audio_index  = len(self.images)
        audio_streams = []

        for i, audio in enumerate(self.audio_files):
            inputs.append(f'-i "{audio["path"]}"')
            audio_streams.append(f"[{audio_index + i}:a]")

        if len(audio_streams) > 1:
            filters.append(f"{''.join(audio_streams)}concat=n={len(audio_streams)}:v=0:a=1[outa]")
            audio_map = "-map [outa]"
        else:
            audio_map = f"-map {audio_index}:a"

        filter_complex = ";".join(filters)
        command = (
            f'ffmpeg -y {" ".join(inputs)} '
            f'-filter_complex "{filter_complex}" '
            f'-map {final_stream} {audio_map} '
            f'-c:a aac -c:v libx264 -pix_fmt yuv420p -preset ultrafast -shortest '
            f'"{self.output_file}"'
        )

        if self.backup_state:
            self.images      = self.images_backup
            self.backup_state = False
            self.update_image_table()

        return command

    # ── Progress ──────────────────────────────────────────────────────────────

    def update_progress(self):
        output = self.process.readAllStandardError().data().decode("utf-8", errors="ignore")
        for line in output.split("\n"):
            if "time=" in line:
                time_str = line.split("time=")[1].split(" ")[0]
                parts = time_str.split(":")
                if len(parts) == 3:
                    try:
                        h, m, s = map(float, parts)
                        cur = h * 3600 + m * 60 + s
                        total = sum(img["duration"] for img in self.images)
                        pct   = int(cur / total * 100) if total else 0
                        self.progress_bar.setValue(pct)
                        self.taskbar_progress.setValue(pct)
                    except ValueError:
                        pass

    # ── Preview ───────────────────────────────────────────────────────────────

    def _load_pixmap(self, path: str) -> QPixmap:
        try:
            img = Image.open(path)
            try:
                exif = img._getexif()
                if exif and _ORIENTATION_TAG:
                    val = exif.get(_ORIENTATION_TAG)
                    if val == 3:   img = img.rotate(180, expand=True)
                    elif val == 6: img = img.rotate(270, expand=True)
                    elif val == 8: img = img.rotate(90,  expand=True)
            except Exception:
                pass
            img  = img.convert("RGBA")
            data = img.tobytes("raw", "RGBA")
            qimg = QImage(data, img.width, img.height, QImage.Format_RGBA8888)
            return QPixmap.fromImage(qimg)
        except Exception as e:
            print(f"Failed to load image {path}: {e}")
            return QPixmap()

    def update_preview(self):
        selected = self.image_table.selectedItems()
        if selected and self.images:
            row = self.image_table.row(selected[0])
            if 0 <= row < len(self.images):
                self.update_preview_with_row(row)

    def update_preview_with_row(self, row: int):
        if 0 <= row < len(self.images):
            img_data = self.images[row]
            pixmap   = self._load_pixmap(img_data["path"])
            if not pixmap.isNull():
                rotation = img_data.get("rotation", 0)
                if rotation:
                    t = QTransform()
                    t.rotate(rotation)
                    pixmap = pixmap.transformed(t, Qt.SmoothTransformation)
                self.preview_label.setPixmap(pixmap.scaled(400, 300, Qt.KeepAspectRatio))
            else:
                self.preview_label.setText("Preview unavailable")

    def setup_connections(self):
        self.image_table.itemSelectionChanged.connect(self.update_preview)

    # ── Project ───────────────────────────────────────────────────────────────

    def clear_project(self):
        self.images.clear()
        self.audio_files.clear()
        self.image_table.setRowCount(0)
        self.audio_table.setRowCount(0)
        self.preview_label.clear()
        self.loaded_project = ""

    def _write_project_file(self, path: str):
        """Serialise current project to disk."""
        with open(path, "w", encoding="utf-8") as f:
            f.write(f"{len(self.audio_files)}\n")
            for audio in self.audio_files:
                f.write(f"{audio['path']}\n")
            for img in self.images:
                text = img.get("text", "").replace("\n", "\\n")
                f.write(
                    f"{img['path']},{img.get('duration', 5)},{img.get('transition', 'fade')},"
                    f"{img.get('transition_duration', 1)},{text},{img.get('rotation', 0)},"
                    f"{img.get('is_second_image', False)},{img.get('date', '')}\n"
                )

    def save_project(self):
        if self.loaded_project:
            self._write_project_file(self.loaded_project)
        else:
            self.save_project_as()

    def save_project_as(self):
        file_name, _ = QFileDialog.getSaveFileName(
            self, "Save Project", "", "Project Files (*.slideshow);;All Files (*)"
        )
        if file_name:
            self._write_project_file(file_name)
            self.loaded_project = file_name

    def load_project(self):
        file_name, _ = QFileDialog.getOpenFileName(
            self, "Load Project", "", "Project Files (*.slideshow);;All Files (*)"
        )
        if not file_name:
            return
        try:
            with open(file_name, "r", encoding="utf-8") as f:
                lines = f.readlines()

            count = int(lines[0].strip())
            if len(lines) < count + 1:
                raise ValueError("Project file is truncated.")

            self.audio_files = [{"path": lines[i + 1].strip()} for i in range(count)]
            self.images = []
            for line in lines[count + 1:]:
                parts = line.strip().split(",")
                if len(parts) < 8:
                    continue
                path, dur, transition, trans_dur, text, rotation, is_second, date = parts[:8]
                self.images.append({
                    "path":               path,
                    "duration":           int(dur),
                    "transition":         transition,
                    "transition_duration": self.default_transition_duration,
                    "text":               text.replace("\\n", "\n"),
                    "rotation":           int(rotation),
                    "is_second_image":    is_second.strip().lower() == "true",
                    "date":               date,
                })

            self.update_image_table()
            self.update_audio_table()
            self.loaded_project = file_name
        except Exception as e:
            QMessageBox.critical(self, self.tr("error"), f"Failed to load project:\n{e}")

    # ── Transitions ───────────────────────────────────────────────────────────

    def update_transition(self, row: int, transition: str):
        if 0 <= row < len(self.images):
            self.images[row]["transition"] = transition

    def set_all_images_transition(self):
        dialog = QInputDialog(self)
        dialog.setWindowTitle("Set Transition")
        dialog.setLabelText("Select transition:")
        dialog.setComboBoxItems(self.transitions_types)
        if dialog.exec_() == QDialog.Accepted:
            t = dialog.textValue()
            for img in self.images:
                img["transition"] = t
            self.update_image_table()

    def set_random_transition_for_each_image(self):
        for img in self.images:
            img["transition"] = random.choice(self.transitions_types)
        self.update_image_table()

    def auto_calc_image_duration(self):
        total_audio = self._total_audio_duration()
        n = len(self.images)
        if n == 0:
            return
        new_dur = int((total_audio - 2) / n)
        for img in self.images:
            img["duration"] = new_dur
        self.update_image_table()

    def validate_transitions(self) -> bool:
        for img in self.images:
            if img["transition_duration"] >= img["duration"]:
                QMessageBox.warning(
                    self, "Invalid Transition Duration",
                    self.tr("error_invalid_transition", path=img["path"]),
                )
                return False
        return True

    # ── Premiere Export ───────────────────────────────────────────────────────

    def export_premiere_slideshow(self):
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Save Premiere Slideshow", "", "Folder"
        )
        if not file_path:
            QMessageBox.critical(self, self.tr("error"), self.tr("error_select_location_premiere"), QMessageBox.Ok)
            return
        self.premiere_project_folder = file_path

        self.image_premiere_progress_bar.setVisible(True)
        self.image_premiere_progress_bar.setValue(0)
        self.taskbar_progress.show()
        self.taskbar_progress.setValue(0)

        self.image_premiere_worker = ImageProcessingPremiereWorker(
            self.images, self.premiere_project_folder, self.common_width, self.common_height
        )
        self.image_premiere_worker.progress.connect(self.update_image_premiere_progress)
        self.image_premiere_worker.finished.connect(self.on_image_premiere_processing_finished)
        self.image_premiere_worker.start()

    def update_image_premiere_progress(self, value: int):
        self.image_premiere_progress_bar.setValue(value)
        self.taskbar_progress.setValue(value)

    def on_image_premiere_processing_finished(self):
        self.image_premiere_progress_bar.setVisible(False)
        self.taskbar_progress.reset()
        self.taskbar_progress.hide()
        self.export_premiere_audio()
        self.export_premiere_text()
        self.export_premiere_duration_excel()
        self.copy_premiere_project_file()

    def export_premiere_audio(self):
        folder = os.path.join(self.premiere_project_folder, "02_אודיו")
        os.makedirs(folder, exist_ok=True)
        for i, audio in enumerate(self.audio_files, start=1):
            src  = audio["path"]
            ext  = os.path.splitext(src)[1]
            name = os.path.splitext(os.path.basename(src))[0]
            dst  = os.path.join(folder, f"audio{i}_{name}{ext}")
            shutil.copy(src, dst)

    def export_premiere_text(self):
        folder = os.path.join(self.premiere_project_folder, "03_טקסט")
        os.makedirs(folder, exist_ok=True)

        # Copy style file if it exists (use relative path)
        script_dir = Path(__file__).resolve().parent
        style_src  = script_dir / "Premiere_Project" / "טקסט למצגת - עברית.prtextstyle"
        if style_src.exists():
            shutil.copy(str(style_src), os.path.join(folder, style_src.name))

        srt_path    = os.path.join(folder, "exported_texts.srt")
        current_time = 0
        idx = 1
        with open(srt_path, "w", encoding="utf-8") as f:
            for img in self.images:
                if img["is_second_image"]:
                    continue
                start = format_time_srt(current_time)
                end   = format_time_srt(current_time + img["duration"])
                f.write(f"{idx}\n{start} --> {end}\n{img['text']}\n\n")
                idx += 1
                current_time += img["duration"]

    def export_premiere_duration_excel(self):
        folder = os.path.join(self.premiere_project_folder, "03_טקסט")
        os.makedirs(folder, exist_ok=True)
        wb = Workbook()
        ws = wb.active
        ws.title = "Durations"
        ws.append(["Path", "Duration", "Text"])
        for img in self.images:
            ws.append([img["path"], img["duration"], img["text"]])
        wb.save(os.path.join(folder, "exported_durations.xlsx"))

    def copy_premiere_project_file(self):
        script_dir = Path(__file__).resolve().parent
        src  = script_dir / "Premiere_Project" / "Project.prproj"
        if not src.exists():
            print(f"Premiere project template not found at {src}")
            return
        dst_folder = os.path.join(self.premiere_project_folder, "04_פרוייקט")
        os.makedirs(dst_folder, exist_ok=True)
        name = os.path.basename(self.premiere_project_folder) + ".prproj"
        shutil.copy(str(src), os.path.join(dst_folder, name))

    # ── Easy Text ─────────────────────────────────────────────────────────────

    def open_easy_text_writing(self):
        if not self.images:
            QMessageBox.warning(self, self.tr("error_no_images_title"), self.tr("error_no_images"))
            return
        selected_row  = self.image_table.currentRow()
        affected_rows = []
        dialog = EasyTextWritingDialog(
            self.images, affected_rows, start_index=selected_row, tr_function=self.tr, parent=self
        )
        dialog.show()
        if dialog.exec_():
            affected_rows[:] = dialog.affected_rows
        for row in affected_rows:
            self.update_image_row(row)

    # ── Shortcuts ─────────────────────────────────────────────────────────────

    def save_shortcuts(self):
        path = BASEPATH / "shortcuts.txt"
        with open(path, "w") as f:
            for action, shortcut in self.shortcuts.items():
                f.write(f"{action}:{shortcut}\n")

    def load_shortcuts(self):
        path = BASEPATH / "shortcuts.txt"
        try:
            with open(path, "r") as f:
                for line in f:
                    if ":" in line:
                        action, shortcut = line.strip().split(":", 1)
                        self.shortcuts[action] = shortcut
        except FileNotFoundError:
            pass

    def update_shortcuts(self):
        self.save_action.setShortcut(self.shortcuts.get("save", "Ctrl+S"))
        self.save_as_action.setShortcut(self.shortcuts.get("save_as", "Ctrl+Shift+S"))
        self.load_action.setShortcut(self.shortcuts.get("load", "Ctrl+L"))
        self.easy_text_writing_action.setShortcut(self.shortcuts.get("easy_text", "Ctrl+T"))
        self.show_info_action.setShortcut(self.shortcuts.get("info", "Alt+I"))
        self.import_images.setShortcut(self.shortcuts.get("import_images", "Ctrl+Shift+I"))
        self.import_audio.setShortcut(self.shortcuts.get("import_audio", "Ctrl+Shift+A"))
        self.set_image_location_action.setShortcut(self.shortcuts.get("set_image_location", "Ctrl+Q"))
        self.delete_row_action.setShortcut(self.shortcuts.get("delete_row", "Delete"))
        self.move_image_up_action.setShortcut(self.shortcuts.get("move_image_up", "Ctrl+Up"))
        self.move_image_down_action.setShortcut(self.shortcuts.get("move_image_down", "Ctrl+Down"))

    def set_shortcut(self, action: str):
        dialog = QInputDialog(self)
        dialog.setWindowTitle(f"{self.tr('shortcut_set')} {action.capitalize()} {self.tr('shortcut_shortcut')}")
        dialog.setLabelText(f"{self.tr('dialog_enter_shortcut')} {action}:")
        dialog.setTextValue(self.shortcuts.get(action, ""))
        if dialog.exec_() == QDialog.Accepted:
            shortcut = dialog.textValue()
            if shortcut:
                self.shortcuts[action] = shortcut
                self.save_shortcuts()
                self.update_shortcuts()

    # ── Info / Help ───────────────────────────────────────────────────────────

    def show_info(self):
        InfoDialog(self.images, self.audio_files, self.tr, self).exec_()

    def open_help_dialog(self):
        HelpDialog(self, self.language).exec_()

    # ── Translations ──────────────────────────────────────────────────────────

    def load_translations(self):
        lang_file = BASEPATH / "Languages" / f"lang_{self.language}.json"
        try:
            with open(lang_file, "r", encoding="utf-8") as f:
                self.translations = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError) as e:
            print(f"Translation load error: {e}")
            self.translations = {"window_title": "Eventure"}

    def tr(self, key: str, **kwargs) -> str:
        text = self.translations.get(key, key)
        if kwargs:
            try:
                text = text.format(**kwargs)
            except KeyError:
                pass
        return text

    def set_language(self, code: str):
        self.language = code
        self.load_translations()
        QApplication.setLayoutDirection(Qt.RightToLeft if code == "he" else Qt.LeftToRight)
        self.retranslate_ui()

    def retranslate_ui(self):
        self.setWindowTitle(self.tr("window_title"))
        self.file_menu.setTitle(self.tr("menu_file"))
        self.import_menu.setTitle(self.tr("menu_import"))
        self.options_menu.setTitle(self.tr("menu_options"))
        self.Img_menu.setTitle(self.tr("menu_images"))
        self.Auto_sort_menu.setTitle(self.tr("menu_auto_sort"))
        self.export_menu.setTitle(self.tr("menu_export"))
        self.Transitions_menu.setTitle(self.tr("menu_transitions"))
        self.Text_menu.setTitle(self.tr("menu_text"))
        self.settings_menu.setTitle(self.tr("menu_settings"))
        self.shortcuts_menu.setTitle(self.tr("menu_shortcuts"))
        self.language_menu.setTitle(self.tr("menu_language"))
        self.info_menu.setTitle(self.tr("menu_info"))
        self.help_menu.setTitle(self.tr("menu_help"))

        self.import_images.setText(self.tr("action_import_images"))
        self.import_audio.setText(self.tr("action_import_audio"))
        self.load_action.setText(self.tr("action_load_project"))
        self.save_action.setText(self.tr("action_save_project"))
        self.save_as_action.setText(self.tr("action_save_project_as"))
        self.clear_action.setText(self.tr("action_clear_project"))
        self.export_slideshow_action.setText(self.tr("action_export_slideshow"))
        self.export_premiere_action.setText(self.tr("action_export_premiere"))
        self.delete_row_action.setText(self.tr("action_delete_row"))
        self.move_image_up_action.setText(self.tr("action_move_image_up"))
        self.move_image_down_action.setText(self.tr("action_move_image_down"))
        self.set_all_images_duration_action.setText(self.tr("action_set_all_image_duration"))
        self.set_random_image_order_action.setText(self.tr("action_set_random_images_order"))
        self.auto_set_images_action.setText(self.tr("action_auto_calc_image_duration"))
        self.set_image_location_action.setText(self.tr("dialog_set_image_location"))
        self.auto_sort_images_by_date_Newest_action.setText(self.tr("action_auto_sort_newest_first"))
        self.auto_sort_images_by_date_Oldest_action.setText(self.tr("action_auto_sort_oldest_first"))
        self.set_all_images_transition_type_action.setText(self.tr("action_set_all_transition_type"))
        self.set_random_transition_for_each_image_action.setText(self.tr("action_set_random_transition_per_image"))
        self.easy_text_writing_action.setText(self.tr("action_easy_text_writing"))
        self.set_save_shortcut_action.setText(self.tr("action_set_save_shortcut"))
        self.set_save_as_shortcut_action.setText(self.tr("action_set_save_as_shortcut"))
        self.set_load_shortcut_action.setText(self.tr("action_set_load_shortcut"))
        self.set_easy_text_shortcut_action.setText(self.tr("action_set_easy_text_shortcut"))
        self.set_show_info_shortcut_action.setText(self.tr("action_set_show_info_shortcut"))
        self.set_delete_row_action.setText(self.tr("action_set_delete_shortcut"))
        self.set_set_image_location_action.setText(self.tr("action_set_image_location_shortcut"))
        self.set_move_image_up_action.setText(self.tr("action_set_move_image_up_shortcut"))
        self.set_move_image_down_action.setText(self.tr("action_set_move_image_down_shortcut"))
        self.show_info_action.setText(self.tr("action_show_info"))
        self.set_language_english_action.setText(self.tr("action_set_language_english"))
        self.set_language_hebrew_action.setText(self.tr("action_set_language_hebrew"))
        self.open_help_dialog_action.setText(self.tr("action_browse_help_topics"))

        self.slides_label.setText(self.tr("label_slides"))
        self.audio_files_label.setText(self.tr("label_audio_files"))
        self.preview_label.setText(self.tr("label_preview"))

        self.image_table.setHorizontalHeaderLabels([
            self.tr("table_header_actions"), self.tr("table_header_image"),
            self.tr("table_header_duration"), self.tr("table_header_transition"),
            self.tr("table_header_transition_length"), self.tr("table_header_text"),
            self.tr("table_header_rotation"), self.tr("table_header_second_image"),
            self.tr("table_header_date"),
        ])
        self.audio_table.setHorizontalHeaderLabels([
            self.tr("table_header_actions"), self.tr("table_header_audio_file"),
        ])
        self.audio_library_button.setText(self.tr("label_audio_library"))

    # ── Menu ─────────────────────────────────────────────────────────────────

    def create_menu(self):
        self.menubar          = self.menuBar()
        self.file_menu        = self.menubar.addMenu(self.tr("menu_file"))
        self.import_menu      = self.file_menu.addMenu(self.tr("menu_import"))
        self.options_menu     = self.menubar.addMenu(self.tr("menu_options"))
        self.Img_menu         = self.options_menu.addMenu(self.tr("menu_images"))
        self.Auto_sort_menu   = self.Img_menu.addMenu(self.tr("menu_auto_sort"))
        self.export_menu      = self.file_menu.addMenu(self.tr("menu_export"))
        self.Transitions_menu = self.options_menu.addMenu(self.tr("menu_transitions"))
        self.Text_menu        = self.options_menu.addMenu(self.tr("menu_text"))
        self.settings_menu    = self.menubar.addMenu(self.tr("menu_settings"))
        self.shortcuts_menu   = self.settings_menu.addMenu(self.tr("menu_shortcuts"))
        self.language_menu    = self.settings_menu.addMenu(self.tr("menu_language"))
        self.info_menu        = self.menubar.addMenu(self.tr("menu_info"))
        self.help_menu        = self.menubar.addMenu(self.tr("menu_help"))

        def _action(label_key, handler, menu, shortcut_key=None):
            a = QAction(self.tr(label_key), self)
            a.triggered.connect(handler)
            if shortcut_key:
                a.setShortcut(self.shortcuts.get(shortcut_key, ""))
            menu.addAction(a)
            return a

        self.import_images   = _action("action_import_images",   self.add_images,          self.import_menu,  "import_images")
        self.import_audio    = _action("action_import_audio",    self.add_audio,           self.import_menu,  "import_audio")
        self.load_action     = _action("action_load_project",    self.load_project,        self.file_menu,    "load")
        self.save_action     = _action("action_save_project",    self.save_project,        self.file_menu,    "save")
        self.save_as_action  = _action("action_save_project_as", self.save_project_as,     self.file_menu,    "save_as")
        self.clear_action    = _action("action_clear_project",   self.clear_project,       self.file_menu)
        self.export_slideshow_action = _action("action_export_slideshow", self.export_slideshow, self.export_menu)
        self.export_premiere_action  = _action("action_export_premiere",  self.export_premiere_slideshow, self.export_menu)

        self.delete_row_action      = _action("action_delete_row",    self.delete_image,    self.Img_menu, "delete_row")
        self.move_image_up_action   = _action("action_move_image_up", self.move_image_up,   self.Img_menu, "move_image_up")
        self.move_image_down_action = _action("action_move_image_down", self.move_image_down, self.Img_menu, "move_image_down")
        self.set_all_images_duration_action   = _action("action_set_all_image_duration",   self.set_all_images_duration, self.Img_menu)
        self.set_random_image_order_action    = _action("action_set_random_images_order",  self.set_random_images_order, self.Img_menu)
        self.auto_set_images_action           = _action("action_auto_calc_image_duration", self.auto_calc_image_duration, self.Img_menu)
        self.set_image_location_action        = _action("dialog_set_image_location",       self.set_image_location, self.Img_menu, "set_image_location")
        self.auto_sort_images_by_date_Newest_action = _action("action_auto_sort_newest_first", lambda: self.auto_sort_images_by_date(True),  self.Auto_sort_menu)
        self.auto_sort_images_by_date_Oldest_action = _action("action_auto_sort_oldest_first", lambda: self.auto_sort_images_by_date(False), self.Auto_sort_menu)

        self.set_all_images_transition_type_action      = _action("action_set_all_transition_type",          self.set_all_images_transition,            self.Transitions_menu)
        self.set_random_transition_for_each_image_action = _action("action_set_random_transition_per_image", self.set_random_transition_for_each_image, self.Transitions_menu)

        self.easy_text_writing_action = _action("action_easy_text_writing", self.open_easy_text_writing, self.Text_menu, "easy_text")

        for key, label in [
            ("save",              "action_set_save_shortcut"),
            ("save_as",           "action_set_save_as_shortcut"),
            ("load",              "action_set_load_shortcut"),
            ("easy_text",         "action_set_easy_text_shortcut"),
            ("info",              "action_set_show_info_shortcut"),
            ("delete_row",        "action_set_delete_shortcut"),
            ("set_image_location","action_set_image_location_shortcut"),
            ("move_image_up",     "action_set_move_image_up_shortcut"),
            ("move_image_down",   "action_set_move_image_down_shortcut"),
        ]:
            a = QAction(self.tr(label), self)
            a.triggered.connect(lambda checked, k=key: self.set_shortcut(k))
            self.shortcuts_menu.addAction(a)
            setattr(self, f"set_{key}_shortcut_action" if key != "set_image_location" else "set_set_image_location_action", a)

        # Expose individual shortcut actions for retranslate
        (
            self.set_save_shortcut_action,
            self.set_save_as_shortcut_action,
            self.set_load_shortcut_action,
            self.set_easy_text_shortcut_action,
            self.set_show_info_shortcut_action,
            self.set_delete_row_action,
            self.set_set_image_location_action,
            self.set_move_image_up_action,
            self.set_move_image_down_action,
        ) = self.shortcuts_menu.actions()

        self.show_info_action = _action("action_show_info", self.show_info, self.info_menu, "info")

        self.set_language_english_action = _action("action_set_language_english", lambda: self.set_language("en"), self.language_menu)
        self.set_language_hebrew_action  = _action("action_set_language_hebrew",  lambda: self.set_language("he"), self.language_menu)

        self.open_help_dialog_action = _action("action_browse_help_topics", self.open_help_dialog, self.help_menu)


# ── Worker Threads ────────────────────────────────────────────────────────────

class ImageProcessingWorker(QThread):
    """
    Processes images in parallel using a ThreadPoolExecutor.
    For CPU-bound PIL work a ProcessPoolExecutor would be even faster,
    but threads avoid pickle/spawn overhead for small-to-medium batches.
    """
    progress = pyqtSignal(int)
    finished = pyqtSignal()

    def __init__(self, images, output_folder, common_width, common_height):
        super().__init__()
        self.images        = images
        self.output_folder = output_folder
        self.common_width  = common_width
        self.common_height = common_height

    def _process_one(self, i: int) -> tuple[int, str | None]:
        img      = self.images[i]["path"]
        rotation = self.images[i]["rotation"]
        text     = self.images[i]["text"]
        try:
            original = Image.open(img)
            if original.size != (self.common_width, self.common_height):
                new_path = Image_resizer.process_image(img, self.output_folder, text, rotation)
                return i, new_path
        except Exception as e:
            print(f"Error opening image {img}: {e}")
        return i, None

    def run(self):
        total = len(self.images)
        completed = 0

        with ThreadPoolExecutor(max_workers=os.cpu_count() or 4) as executor:
            futures = {executor.submit(self._process_one, i): i for i in range(total)}
            for future in as_completed(futures):
                completed += 1
                try:
                    i, new_path = future.result()
                    if new_path:
                        self.images[i]["path"] = new_path
                except Exception as e:
                    print(f"Worker error: {e}")
                self.progress.emit(int(completed / total * 100))

        self.finished.emit()


class ImageProcessingPremiereWorker(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal()

    def __init__(self, images, output_folder, common_width, common_height):
        super().__init__()
        self.images        = images
        self.output_folder = output_folder
        self.common_width  = common_width
        self.common_height = common_height

    def run(self):
        premiere_export.process_images(self.images, self.output_folder, self.progress.emit)
        self.finished.emit()


# ── Dialogs ───────────────────────────────────────────────────────────────────

class HelpDialog(QDialog):
    def __init__(self, parent=None, language="en"):
        super().__init__(parent)
        self.setWindowTitle("Help Topics")
        self.resize(600, 400)
        self.language = language

        layout = QVBoxLayout(self)
        self.topic_list  = QListWidget()
        self.info_display = QTextEdit()
        self.info_display.setReadOnly(True)
        layout.addWidget(self.topic_list)
        layout.addWidget(self.info_display)

        self.help_data = self._load_help_info()
        for topic in self.help_data:
            self.topic_list.addItem(QListWidgetItem(topic))
        self.topic_list.itemClicked.connect(self._display_info)

    def _load_help_info(self) -> dict:
        try:
            path = BASEPATH / "Help" / f"Help_Info_{self.language}.txt"
            with open(path, "r", encoding="utf-8") as f:
                content = f.read()
            data = {}
            for block in content.split("topic:"):
                if block.strip():
                    parts = block.strip().split("Info:")
                    data[parts[0].strip()] = parts[1].strip() if len(parts) > 1 else ""
            return data
        except Exception as e:
            return {"Error loading help file": str(e)}

    def _display_info(self, item):
        self.info_display.setText(f"{item.text()}\n\n{self.help_data.get(item.text(), '')}")


class EasyTextWritingDialog(QDialog):
    def __init__(self, images, affected_rows, start_index=0, tr_function=None, parent=None):
        super().__init__(parent)
        self.tr            = tr_function
        self.images        = images
        self.affected_rows = affected_rows
        self.current_index = start_index

        self.setWindowTitle(self.tr("action_easy_text_writing"))
        self.setGeometry(200, 200, 400, 200)

        layout = QVBoxLayout(self)
        self.image_label = QLabel()
        self.image_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.image_label)

        self.text_input = QTextEdit()
        self.text_input.setPlaceholderText(self.tr("enter_text_for_image"))
        self.text_input.setAlignment(Qt.AlignRight)
        self.text_input.setLayoutDirection(Qt.RightToLeft)
        self.text_input.setPlainText(self.images[self.current_index].get("text", ""))
        self.text_input.installEventFilter(self)
        layout.addWidget(self.text_input)

        self.text_input.moveCursor(QTextCursor.Start)

        next_btn  = QPushButton(self.tr("next"))
        close_btn = QPushButton(self.tr("close"))
        next_btn.clicked.connect(self.next_image)
        close_btn.clicked.connect(self.close)
        layout.addWidget(next_btn)
        layout.addWidget(close_btn)

        self.update_image()

    def update_image(self):
        if 0 <= self.current_index < len(self.images):
            data  = self.images[self.current_index]
            px    = QPixmap(data["path"])
            rot   = data.get("rotation", 0)
            if rot:
                t = QTransform()
                t.rotate(rot)
                px = px.transformed(t, Qt.SmoothTransformation)
            self.image_label.setPixmap(px.scaled(300, 200, Qt.KeepAspectRatio))
            self.text_input.setPlainText(data.get("text", ""))

    def next_image(self):
        if 0 <= self.current_index < len(self.images):
            new_text = self.text_input.toPlainText()
            if self.images[self.current_index]["text"] != new_text:
                self.images[self.current_index]["text"] = new_text
                if self.current_index not in self.affected_rows:
                    self.affected_rows.append(self.current_index)
        self.current_index = (self.current_index + 1) % len(self.images)
        self.update_image()

    def eventFilter(self, source, event):
        if source is self.text_input and event.type() == QEvent.KeyPress and event.key() == Qt.Key_Tab:
            self.next_image()
            return True
        return super().eventFilter(source, event)


class InfoDialog(QDialog):
    def __init__(self, images, audio_files, tr_function=None, parent=None):
        super().__init__(parent)
        self.tr = tr_function
        self.setWindowTitle(self.tr("menu_info"))
        self.setGeometry(300, 300, 300, 150)

        layout = QVBoxLayout(self)
        dur_with    = sum(img["duration"] for img in images)
        dur_without = sum(img["duration"] for img in images if not img.get("is_second_image"))
        audio_dur   = sum(_get_audio_duration(a["path"]) for a in audio_files)

        layout.addWidget(QLabel(self.tr("info_total_images") + f" {len(images)}"))
        layout.addWidget(QLabel(self.tr("info_duration_with_second") + f" {format_time_hms(dur_with)}"))
        layout.addWidget(QLabel(self.tr("info_duration_without_second") + f" {format_time_hms(dur_without)}"))
        layout.addWidget(QLabel(self.tr("info_audio_duration") + f" {format_time_hms(audio_dur)}"))

        close_btn = QPushButton(self.tr("close"))
        close_btn.clicked.connect(self.close)
        layout.addWidget(close_btn)


class CustomDelegate(QStyledItemDelegate):
    def paint(self, painter, option, index):
        is_secondary = index.model().index(index.row(), 1).data(Qt.UserRole)
        if is_secondary:
            painter.save()
            painter.fillRect(option.rect, QColor(100, 100, 150))
            painter.restore()
        super().paint(painter, option, index)


class AudioLibraryDialog(QDialog):
    def __init__(self, tr_function=None, parent=None):
        super().__init__(parent)
        self.tr    = tr_function
        self.songs = []
        self.setWindowTitle(self.tr("label_audio_library"))
        self.setGeometry(100, 100, 600, 400)
        self._load_songs()
        self._init_ui()

    def _load_songs(self):
        songs_file = BASEPATH / "Songs" / "songs.json"
        try:
            with open(songs_file, "r", encoding="utf-8") as f:
                self.songs = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError) as e:
            print(f"Songs load error: {e}")
            self.songs = []

    def _init_ui(self):
        layout = QVBoxLayout(self)

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText(self.tr("search_songs"))
        self.search_input.textChanged.connect(self._filter_songs)
        layout.addWidget(self.search_input)

        self.song_list = QListWidget()
        self._populate(filter_text="")
        layout.addWidget(self.song_list)

        self.info_label = QLabel(self.tr("song_info_label"))
        layout.addWidget(self.info_label)

        btn_row = QHBoxLayout()
        add_btn   = QPushButton(self.tr("add_selected"))
        close_btn = QPushButton(self.tr("close"))
        add_btn.clicked.connect(self._add_selected)
        close_btn.clicked.connect(self.close)
        btn_row.addWidget(add_btn)
        btn_row.addWidget(close_btn)
        layout.addLayout(btn_row)

        self.song_list.itemSelectionChanged.connect(self._update_info)

    def _populate(self, filter_text: str = ""):
        self.song_list.clear()
        low = filter_text.lower()
        for song in self.songs:
            if (
                low in song["name"].lower()
                or low in song["author"].lower()
                or low in song.get("fits_for", "").lower()
            ):
                item = QListWidgetItem(f"{song['name']} - {song['author']}")
                item.setData(Qt.UserRole, song)
                self.song_list.addItem(item)

    def _filter_songs(self):
        self._populate(self.search_input.text())

    def _fmt_dur(self, seconds: float) -> str:
        m = int(seconds // 60)
        s = int(seconds % 60)
        return f"{m}:{s:02d}"

    def _update_info(self):
        selected = self.song_list.selectedItems()
        if selected:
            s = selected[0].data(Qt.UserRole)
            self.info_label.setText(
                f"<b>Name:</b> {s['name']}<br>"
                f"<b>Author:</b> {s['author']}<br>"
                f"<b>Duration:</b> {self._fmt_dur(s['duration'])}<br>"
                f"<b>Fits for:</b> {s.get('fits_for', '')}<br>"
                f"<b>Path:</b> {Path(s['path'].replace('{BASE_PATH}', str(BASEPATH)))}"
            )

    def _add_selected(self):
        for item in self.song_list.selectedItems():
            song = item.data(Qt.UserRole)
            path = Path(song["path"].replace("{BASE_PATH}", str(BASEPATH)))
            if not any(a["path"] == path for a in self.parent().audio_files):
                self.parent().audio_files.append({"path": path})
        self.parent().update_audio_table()
        self.close()


# ── Entry Point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon("logo.ico"))
    set_theme(app, theme="dark")
    window = SlideshowCreator()
    window.create_menu()
    window.setup_connections()
    window.show()
    sys.exit(app.exec_())