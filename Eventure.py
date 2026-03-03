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

CRITICAL – multiprocessing fix (Windows):
  freeze_support() MUST be the very first executable line in the entry-point
  file.  On Windows, ProcessPoolExecutor spawns fresh Python processes that
  re-import this module; if freeze_support() isn't first, every worker tries
  to boot the full Qt app and crashes immediately.
"""

# ── freeze_support MUST come before every other import ───────────────────────
import multiprocessing
import threading

from pptx_export import extract_pptx_content_to_slideshow_file
multiprocessing.freeze_support()
# ─────────────────────────────────────────────────────────────────────────────

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
from PyQt5.QtCore import pyqtSlot
from PyQt5.QtCore import Qt, QUrl, QSize, QProcess, QTimer, QThread, pyqtSignal, QEvent
from PyQt5.QtGui import (
    QIcon, QFont, QPixmap, QTextCursor, QCursor, QTransform,
    QColor, QBrush, QImage,
)
from PIL import Image, ExifTags
from openpyxl import Workbook
import openpyxl

import premiere_export

from EVENTURE_THEMES.theme import set_theme

APP_VERSION = "1.0.3"


# ── Environment ──────────────────────────────────────────────────────────────

plugin_path = os.path.join(os.path.dirname(sys.executable), "Library", "plugins", "platforms")
os.environ["QT_QPA_PLATFORM_PLUGIN_PATH"] = plugin_path

BASEPATH = Path.home() / "Neria-LTD" / "Eventure"
BASEPATH.mkdir(parents=True, exist_ok=True)

# ── Resolve the folder that contains the running app ─────────────────────────
# Works correctly both when running as a plain .py script AND when bundled
# with PyInstaller (frozen exe).  PyInstaller sets sys.frozen and unpacks
# bundled files next to sys.executable, so ffmpeg.exe will be found there.
if getattr(sys, "frozen", False):
    # Running as PyInstaller bundle — exe lives in this folder
    APP_DIR = Path(sys.executable).resolve().parent
else:
    # Running as a normal Python script
    APP_DIR = Path(__file__).resolve().parent

# EXIF orientation tag (looked up once)
_ORIENTATION_TAG = next(
    (k for k, v in ExifTags.TAGS.items() if v == "Orientation"), None
)


def check_for_updates(parent_window, current_version: str):
    """
    Checks GitHub releases API in a background thread.
    If a newer version exists, shows a non-blocking dialog with a download link.

    Replace GITHUB_USER and GITHUB_REPO with your actual values.
    """
    GITHUB_USER = "neriacohen300"       # ← change this
    GITHUB_REPO = "Eventure"      # ← change this
    API_URL = f"https://api.github.com/repos/{GITHUB_USER}/{GITHUB_REPO}/releases/latest"

    def _fetch():
        try:
            import urllib.request, json
            req = urllib.request.Request(
                API_URL,
                headers={"User-Agent": "Eventure-App"},
            )
            with urllib.request.urlopen(req, timeout=5) as resp:
                data = json.loads(resp.read().decode())

            latest_tag  = data.get("tag_name", "").lstrip("v")   # e.g. "1.2.3"
            release_url = data.get("html_url", "")

            if not latest_tag:
                return

            # Simple tuple comparison: (1, 2, 3) > (1, 0, 0)
            def _ver(s):
                try:    return tuple(int(x) for x in s.split("."))
                except: return (0,)

            if _ver(latest_tag) > _ver(current_version):
                # Must update the UI on the main thread via a Qt signal
                from PyQt5.QtCore import QMetaObject, Qt, Q_ARG
                QMetaObject.invokeMethod(
                    parent_window,
                    "_show_update_dialog",
                    Qt.QueuedConnection,
                    Q_ARG(str, latest_tag),
                    Q_ARG(str, release_url),
                )
        except Exception as e:
            print(f"Update check failed: {e}")   # silent — never block startup

    threading.Thread(target=_fetch, daemon=True).start()


# ── Helpers ───────────────────────────────────────────────────────────────────


def _ffmpeg_exe() -> str:
    """Resolve ffmpeg from PATH, then next to the running app (script or PyInstaller exe)."""
    import shutil as _shutil
    return _shutil.which("ffmpeg") or str(APP_DIR / "ffmpeg.exe")

def _ffprobe_exe() -> str:
    """Resolve ffprobe from PATH, then next to the running app (script or PyInstaller exe)."""
    import shutil as _shutil
    return _shutil.which("ffprobe") or str(APP_DIR / "ffprobe.exe")

def _get_audio_duration(audio_path: str) -> float:
    """Return the duration of an audio file in seconds using ffprobe."""
    try:
        result = subprocess.run(
            [
                _ffprobe_exe(), "-v", "error",
                "-show_entries", "format=duration",
                "-of", "default=noprint_wrappers=1:nokey=1",
                audio_path,
            ],
            capture_output=True,
            text=True,
            timeout=10,
            creationflags=subprocess.CREATE_NO_WINDOW if hasattr(subprocess, "CREATE_NO_WINDOW") else 0
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


# ── Ken Burns pre-renderer ───────────────────────────────────────────────────
#
# Instead of using FFmpeg's notoriously fragile zoompan filter (which produces
# variable-framerate streams that break xfade), we pre-render Ken Burns as a
# real .mp4 clip per image using Pillow + ffmpeg pipe.  The main export then
# uses these clips as normal video inputs — stable timestamps, correct size,
# no surprises.

_KB_FPS = 25

def render_ken_burns_clip(image_path: str, effect: str, duration: float,
                           output_path: str, rotation: int = 0,
                           text: str = "", text_on_kb: bool = True) -> bool:
    """
    Render a smooth Ken Burns clip at 1920x1080 25fps using sub-pixel interpolation.

    Speed improvements vs original:
      - All frames pre-built into a contiguous numpy buffer → single pipe write (no per-frame syscall)
      - CRF 28 (was 18) for intermediate clips — re-encoded in final pass anyway
      - -tune fastdecode + -g 25 for faster ffmpeg encode
    """
    import cv2 as _cv2
    import numpy as _np
    import subprocess as _sp
    from PIL import Image as _Image, ImageDraw as _Draw, ImageFont as _Font
    from PIL import ExifTags as _ExifTags

    W, H, FPS = 1920, 1080, _KB_FPS
    frames = max(1, int(duration * FPS))

    # ── Load & orient ─────────────────────────────────────────────────────────
    try:
        img = _Image.open(image_path)
        img = img.convert("RGBA")
        try:
            exif = img._getexif() if hasattr(img, "_getexif") else None
            if exif and _ORIENTATION_TAG:
                val = exif.get(_ORIENTATION_TAG)
                if val == 3:   img = img.rotate(180, expand=True)
                elif val == 6: img = img.rotate(270, expand=True)
                elif val == 8: img = img.rotate(90,  expand=True)
        except Exception:
            pass
        if rotation:
            img = img.rotate(rotation, expand=True)
        img = img.convert("RGB")
    except Exception as e:
        print(f"KB render error: {e}")
        return False

    # ── Smoothstep easing ─────────────────────────────────────────────────────
    def smooth(raw_t):
        t = max(0.0, min(1.0, raw_t))
        return t * t * (3.0 - 2.0 * t)

    # ── Build high-res source canvas ──────────────────────────────────────────
    ZOOM = 1.10
    src_w, src_h = int(W * ZOOM), int(H * ZOOM)
    img_arr = _np.array(img)
    ih, iw = img_arr.shape[:2]
    scale_cov = max(src_w / iw, src_h / ih)
    new_iw, new_ih = int(iw * scale_cov), int(ih * scale_cov)

    resized = _cv2.resize(img_arr, (new_iw, new_ih), interpolation=_cv2.INTER_LANCZOS4)
    cx, cy = (new_iw - src_w) // 2, (new_ih - src_h) // 2
    src = resized[cy:cy + src_h, cx:cx + src_w].copy()

    # ── Pre-compute float rectangles ─────────────────────────────────────────
    IS_PAN = effect in ("pan_left", "pan_right", "pan_up", "pan_down", "none")
    pad_x = src_w - W
    pad_y = src_h - H
    rects = []
    for f in range(frames):
        t = smooth(f / max(frames - 1, 1))
        sw, sh = float(src_w), float(src_h)
        sx, sy = 0.0, 0.0

        if effect == "zoom_in":
            sw = src_w - (src_w - W) * t
            sh = src_h - (src_h - H) * t
            sx = (src_w - sw) / 2.0
            sy = (src_h - sh) / 2.0
        elif effect == "zoom_out":
            sw = W + (src_w - W) * t
            sh = H + (src_h - H) * t
            sx = (src_w - sw) / 2.0
            sy = (src_h - sh) / 2.0
        elif effect == "pan_left":
            sw, sh = float(W), float(H)
            sx = (src_w - W) * (1.0 - t)
            sy = (src_h - H) / 2.0
        elif effect == "pan_right":
            sw, sh = float(W), float(H)
            sx = (src_w - W) * t
            sy = (src_h - H) / 2.0
        elif effect == "pan_up":
            sw, sh = float(W), float(H)
            sx = float(pad_x) / 2.0
            sy = float(pad_y) * (1.0 - t)
        elif effect == "pan_down":
            sw, sh = float(W), float(H)
            sx = float(pad_x) / 2.0
            sy = float(pad_y) * t
        else:
            sw, sh = float(W), float(H)
            sx = float(pad_x) / 2.0
            sy = float(pad_y) / 2.0

        rects.append((sx, sy, sw, sh))

    # ── Static text overlay (text_on_kb=False) ────────────────────────────────
    static_overlay_bgr = None
    static_overlay_mask = None
    if text and text.strip() and not text_on_kb:
        try:
            from bidi.algorithm import get_display as _bidi
            from pathlib import Path as _Path
            _font_path = _Path.home() / "Neria-LTD" / "Eventure" / "Fonts" / "Birzia-Black.otf"
            try:   _font = _Font.truetype(str(_font_path), 85)
            except Exception: _font = _Font.load_default()
            overlay = _Image.new("RGBA", (W, H), (0, 0, 0, 0))
            draw = _Draw.Draw(overlay)
            htext = _bidi(text)
            bbox  = draw.textbbox((0, 0), htext, font=_font)
            tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
            bg_w, bg_h = tw + 40, th + 20
            bg_x = (W - bg_w) // 2;  bg_y = H - bg_h - 50
            draw.rounded_rectangle((bg_x, bg_y, bg_x+bg_w, bg_y+bg_h), radius=12, fill="white")
            draw.text(((W - tw) // 2, bg_y - 4), htext, font=_font, fill="black")
            ov_arr = _np.array(overlay)
            alpha  = ov_arr[:, :, 3:4].astype(_np.float32) / 255.0
            static_overlay_bgr  = ov_arr[:, :, :3].astype(_np.float32)
            static_overlay_mask = alpha
        except Exception as e:
            print(f"KB text overlay error: {e}")

    # ── Pre-build all frames into a contiguous buffer ─────────────────────────
    # One numpy allocation → one pipe write → eliminates per-frame syscall overhead
    frame_buffer = _np.empty((frames, H, W, 3), dtype=_np.uint8)

    for f in range(frames):
        sx, sy, sw, sh = rects[f]
        scale_x = W / sw
        scale_y = H / sh
        M = _np.array([
            [scale_x,  0.0, -sx * scale_x],
            [0.0,  scale_y, -sy * scale_y],
        ], dtype=_np.float64)
        frame = _cv2.warpAffine(
            src, M, (W, H),
            flags=_cv2.INTER_LINEAR,
            borderMode=_cv2.BORDER_REFLECT_101,
        )
        if static_overlay_bgr is not None:
            frame = (
                frame.astype(_np.float32) * (1.0 - static_overlay_mask)
                + static_overlay_bgr * static_overlay_mask
            ).clip(0, 255).astype(_np.uint8)

        frame_buffer[f] = frame

    # ── Pipe entire buffer into ffmpeg in one write ───────────────────────────
    cmd = [
        _ffmpeg_exe(), "-y",
        "-f", "rawvideo", "-vcodec", "rawvideo",
        "-s", f"{W}x{H}", "-pix_fmt", "rgb24", "-r", str(FPS),
        "-i", "pipe:0",
        "-vcodec", "libx264", "-pix_fmt", "yuv420p",
        "-preset", "ultrafast", "-crf", "28",
        "-tune", "fastdecode",
        "-g", "25",
        "-r", str(FPS), "-movflags", "+faststart",
        output_path,
    ]
    proc = None
    try:
        proc = _sp.Popen(
            cmd,
            stdin=_sp.PIPE,
            stdout=_sp.DEVNULL,
            stderr=_sp.DEVNULL,
            creationflags=_sp.CREATE_NO_WINDOW if hasattr(_sp, "CREATE_NO_WINDOW") else 0,
        )
        proc.stdin.write(frame_buffer.tobytes())  # single syscall instead of N writes
        proc.stdin.close()
        proc.wait()
        return proc.returncode == 0
    except Exception as e:
        print(f"KB render pipe error: {e}")
        if proc:
            try: proc.stdin.close()
            except Exception: pass
        return False


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


    @pyqtSlot(str, str)
    def _show_update_dialog(self, new_version: str, url: str):
        """Called on the main thread when a newer version is available."""
        from PyQt5.QtWidgets import QMessageBox
        msg = QMessageBox(self)
        msg.setWindowTitle(self.tr("update_available"))
        msg.setIcon(QMessageBox.Information)
        msg.setText(
            f"<b>{self.tr("new_version")} v{new_version}</b><br><br>"
            f"{self.tr("cur_version")} v{APP_VERSION}.<br><br>"
            f'<a href="{url}">{self.tr("download")}</a>'
        )
        msg.setTextFormat(Qt.RichText)
        msg.setTextInteractionFlags(Qt.TextBrowserInteraction)
        msg.setStandardButtons(QMessageBox.Ok)
        msg.exec_()

    def __init__(self):
        super().__init__()

        # Resource folders are synced once at startup in __main__ block below.
        # Do NOT call _copy_resource_folders here — worker processes also
        # instantiate classes during import, and we must not do file I/O there.

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

        self._pending_temp_dirs: list[str] = []   # filled by worker, deleted after export

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
        self.image_table.setSortingEnabled(False)
        self.image_table.setColumnCount(10)
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
            self.tr("ken_burns"),
        ])
        self.image_table.setFont(QFont(self.deafult_font, 10, QFont.Bold))
        for col in range(10):
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
                    "ken_burns": "none"
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
        if (state == Qt.Checked and self.images[row-1]["is_second_image"] == True) or (state == Qt.Checked and self.images[row+1]["is_second_image"] == True):
            QMessageBox.warning(self, self.tr("error"), self.tr("subsequent_line"), QMessageBox.Ok)
            self.update_image_row(row)
            return
        
        is_checked = (state == Qt.Checked)
        self.images[row]["is_second_image"] = is_checked
        item = self.image_table.item(row, 1)
        if item:
            item.setData(Qt.UserRole, is_checked)
            
        self.update_image_row(row)

    def update_image_table(self):
        self.image_table.blockSignals(True)
        self.image_table.setUpdatesEnabled(False)
        self.image_table.setSortingEnabled(False)
        self.image_table.setRowCount(len(self.images))

        for row, img in enumerate(self.images):
            self._populate_row(row, img)

        self.image_table.setSortingEnabled(False)
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

        kb_cb = QComboBox()
        kb_cb.addItems(["none", "zoom_in", "zoom_out",
                         "pan_left", "pan_right", "pan_up", "pan_down"])
        kb_cb.setCurrentText(img.get("ken_burns", "none"))
        kb_cb.currentTextChanged.connect(lambda text, r=row: self._update_ken_burns(r, text))
        self.image_table.setCellWidget(row, 9, kb_cb)


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
        self.image_table.setSortingEnabled(False)

    def move_image_down(self):
        self.image_table.setSortingEnabled(False)
        row = self.image_table.currentRow()
        if row < len(self.images) - 1:
            self.images[row], self.images[row + 1] = self.images[row + 1], self.images[row]
            self.update_image_row(row)
            self.update_image_row(row + 1)
            self.image_table.setCurrentCell(row + 1, 1)
            self.update_preview_with_row(row + 1)
        self.image_table.setSortingEnabled(False)

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
        self.image_table.setSortingEnabled(False)

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

    def _warn_corrupted_image(self, path: str):
        name = os.path.basename(path)
        QMessageBox.warning(
            self,
            "Corrupted Image",
            f"The following image appears to be corrupted and will be skipped:\n\n{name}",
            QMessageBox.Ok,
        )

    def _store_temp_dirs(self, dirs: list):
        """Slot: remember temp dirs emitted by the worker so we can delete them later."""
        self._pending_temp_dirs = dirs

    def _cleanup_temp_dirs(self):
        """Delete all temporary folders (A_Blur + kb_clips) after the final export."""
        for d in self._pending_temp_dirs:
            if d and os.path.isdir(d):
                try:
                    shutil.rmtree(d, ignore_errors=True)
                    print(f"Cleaned up temp folder: {d}")
                except Exception as e:
                    print(f"Could not remove temp folder {d}: {e}")
        self._pending_temp_dirs = []

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

        # Locate ffmpeg — QProcess on Windows does NOT inherit the full system
        # PATH, so we resolve the executable to an absolute path ourselves.
        import shutil as _shutil
        import shlex  as _shlex
        # 1) Check system PATH, 2) fall back to script folder
        ffmpeg_exe = _ffmpeg_exe()
        import os as _os
        if not _os.path.isfile(ffmpeg_exe):
            ffmpeg_exe = None
        if not ffmpeg_exe:
            QMessageBox.critical(
                self, "FFmpeg not found",
                "ffmpeg could not be found on your PATH or in the script folder.\n"
                "Please place ffmpeg.exe next to Eventure.py or add it to your PATH.",
                QMessageBox.Ok,
            )
            return
        print("ffmpeg resolved to:", ffmpeg_exe)

        # Split into program + args list; strip quotes shlex leaves on tokens.
        raw_args = _shlex.split(command, posix=False)
        args = [a.strip('"') for a in raw_args[1:]]

        self.process = QProcess(self)

        # Pass the full current environment so ffmpeg can find its libraries.
        from PyQt5.QtCore import QProcessEnvironment
        self.process.setProcessEnvironment(QProcessEnvironment.systemEnvironment())

        self.process.readyReadStandardError.connect(self.update_progress)
        self.process.finished.connect(self.export_finished)

        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.taskbar_progress.show()
        self.taskbar_progress.setValue(0)

        self.process.start(ffmpeg_exe, args)
        print("QProcess state after start:", self.process.state())
        if self.process.state() == 0:
            err = self.process.errorString()
            QMessageBox.critical(self, "Export failed",
                                 f"Failed to launch ffmpeg:\n{err}", QMessageBox.Ok)

    def export_slideshow(self):
        if not self.images or not self.audio_files:
            QMessageBox.critical(self, self.tr("error"), self.tr("error_no_audio"), QMessageBox.Ok)
            return
        QMessageBox.warning(self, self.tr("just_know"), self.tr("no_secondery"), QMessageBox.Ok)


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
        self.image_worker.cleanup_dirs.connect(self._store_temp_dirs)
        self.image_worker.finished.connect(self.on_image_processing_finished)
        self.image_worker.corrupted_image.connect(self._warn_corrupted_image)
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
        # Remove A_Blur and kb_clips now that the final .mp4 is written.
        self._cleanup_temp_dirs()
        # Delete the temporary filter_complex script file if one was created.
        fc = getattr(self, "_fc_script_path", None)
        if fc:
            try:
                import os as _os
                _os.remove(fc)
            except OSError:
                pass
            self._fc_script_path = None

    def build_ffmpeg_command(self) -> str:
        inputs, filters = [], []

        for i, img in enumerate(self.images):
            if i == 0:
                duration = img["duration"]
            elif i == len(self.images) - 1:
                duration = img["duration"] + self.images[i - 1]["transition_duration"]
            else:
                duration = img["duration"] + img["transition_duration"]
            # If a pre-rendered KB clip exists, use it as a video input (not looped still)
            kb_clip = img.get("_kb_clip_path")
            if kb_clip and os.path.exists(kb_clip):
                # Normalize to forward slashes — os.path.join uses backslash on
                # Windows but original paths may use forward slash, causing FFmpeg
                # to silently fail on mixed-separator paths like E:/foo\bar.mp4
                kb_clip_norm = str(kb_clip).replace("\\", "/")
                inputs.append(f'-t {duration} -i "{kb_clip_norm}"')
                filters.append(
                    f"[{i}:v]fps=25,setpts=PTS-STARTPTS,scale=1920:1080,setsar=1,format=yuv420p[{i}v]"
                )
            else:
                img_path_norm = str(img["path"]).replace("\\", "/")
                inputs.append(f'-loop 1 -t {duration} -i "{img_path_norm}"')
                filters.append(
                    f"[{i}:v]fps=25,scale=1920:1080,setsar=1,"
                    f"setpts=PTS-STARTPTS,format=yuv420p[{i}v]"
                )

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
            audio_norm = str(audio["path"]).replace("\\", "/")
            inputs.append(f'-i "{audio_norm}"')
            audio_streams.append(f"[{audio_index + i}:a]")

        # ── Audio fade-out at the end of the video ────────────────────────────
        # Total video duration used to anchor the fade start time.
        total_video_dur = sum(img["duration"] for img in self.images)
        fade_duration   = 3.0   # seconds of fade-out
        fade_start      = max(0.0, total_video_dur - fade_duration)

        if len(audio_streams) > 1:
            # Concatenate all audio tracks, then apply fade-out on the result.
            filters.append(
                f"{''.join(audio_streams)}concat=n={len(audio_streams)}:v=0:a=1[outa_raw]"
            )
            filters.append(
                f"[outa_raw]afade=t=out:st={fade_start:.3f}:d={fade_duration:.3f}[outa]"
            )
            audio_map = "-map [outa]"
        else:
            # Single audio stream — apply fade-out directly in the filter graph.
            filters.append(
                f"[{audio_index}:a]afade=t=out:st={fade_start:.3f}:d={fade_duration:.3f}[outa]"
            )
            audio_map = "-map [outa]"

        filter_complex = ";".join(filters)

        # ── Write filter_complex to a temp script file ────────────────────────
        # Windows has a 32 767-char command-line limit.  With many images the
        # filter_complex alone exceeds this.  ffmpeg's -filter_complex_script
        # reads it from a file instead, keeping the command line short.
        import tempfile as _tempfile
        fc_file = _tempfile.NamedTemporaryFile(
            mode="w", suffix=".txt", delete=False, encoding="utf-8"
        )
        fc_file.write(filter_complex)
        fc_file.close()
        self._fc_script_path = fc_file.name   # remember so we can delete it later

        output_norm = str(self.output_file).replace(chr(92), "/")
        command = (
            f'ffmpeg -y {" ".join(inputs)} '
            f'-filter_complex_script "{fc_file.name}" '
            f'-map {final_stream} {audio_map} '
            f'-c:a aac -c:v libx264 -pix_fmt yuv420p -preset ultrafast -movflags +faststart -shortest '
            f'"{output_norm}"'
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
                    f"{img.get('is_second_image', False)},{img.get('date', '')},"
                    f"{img.get('ken_burns', 'none')},{img.get('text_on_kb', True)}\n"
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
        if file_name:
            self._load_project_from_path(file_name)

    def import_pptx(self):
        file_name, _ = QFileDialog.getOpenFileName(
            self, "Select a PowerPoint file", "", "PowerPoint files (*.pptx;*.pptm);;All Files (*)"
        )
        
        if file_name:
            slideshow_file = extract_pptx_content_to_slideshow_file(file_name)
            if slideshow_file:
                self._load_project_from_path(slideshow_file)

        


    def _load_project_from_path(self, file_name: str):
        """Load a .slideshow file from an explicit path (also called by file association)."""
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
                # parts: path, dur, transition, trans_dur, text, rotation,
                #        second_image_path, second_image_rotation, date, ken_burns
                path         = parts[0]
                dur          = parts[1]
                transition   = parts[2]
                trans_dur    = parts[3]
                text         = parts[4]
                rotation     = parts[5]
                is_second  = parts[6]
                date         = parts[7] if len(parts) > 7 else ""
                ken_burns    = parts[8].strip()  if len(parts) > 8  else "none"

                self.images.append({
                    "path":                path,
                    "duration":            int(dur),
                    "transition":          transition,
                    "transition_duration": self.default_transition_duration,
                    "text":                text.replace("\\n", "\n"),
                    "rotation":            int(rotation),
                    "is_second_image":     is_second.strip().lower() == "true",
                    "date":                date,
                    "ken_burns":           ken_burns
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

    def _update_ken_burns(self, row: int, value: str):
        if 0 <= row < len(self.images):
            self.images[row]["ken_burns"] = value

    def _set_all_ken_burns(self):
        KB_OPTIONS = ["none", "zoom_in", "zoom_out",
                      "pan_left", "pan_right", "pan_up", "pan_down"]
        dialog = QInputDialog(self)
        dialog.setWindowTitle("Set Ken Burns Effect")
        dialog.setLabelText("Select effect for all images:")
        dialog.setComboBoxItems(KB_OPTIONS)
        if dialog.exec_() == QDialog.Accepted:
            effect = dialog.textValue()
            for img in self.images:
                img["ken_burns"] = effect
            self.update_image_table()

    def _set_random_ken_burns_per_image(self):
        """Assign a random Ken Burns effect to every image independently."""
        KB_OPTIONS = ["zoom_in", "zoom_out",
                      "pan_left", "pan_right", "pan_up", "pan_down"]
        for img in self.images:
            img["ken_burns"] = random.choice(KB_OPTIONS)
        self.update_image_table()

    def _set_smart_ken_burns(self):
        """
        Assign Ken Burns effects using a cinematic continuity algorithm.

        Rules:
          1. Motion continuity — the next effect should feel like it picks up
             where the last one left off:
               zoom_in  → zoom_out  (camera reverses, starts at the zoomed-in frame)
               zoom_out → zoom_in   (same logic in reverse)
               pan_left → pan_right (reversal feels natural)
               pan_right→ pan_left
               pan_up   → pan_down
               pan_down → pan_up
          2. Variety — after two reversals in a row, force a category switch
             (zoom ↔ pan) to break the monotony.
          3. Occasional wildcards (≈ 20 %) — insert a pan between two zooms or
             vice-versa to keep the sequence interesting.
          4. Second images (PiP) are skipped.
        """
        # Continuity map: what naturally follows each effect
        REVERSAL = {
            "zoom_in":   "zoom_out",
            "zoom_out":  "zoom_in",
            "pan_left":  "pan_right",
            "pan_right": "pan_left",
            "pan_up":    "pan_down",
            "pan_down":  "pan_up",
        }
        ZOOM_EFFECTS = ["zoom_in",  "zoom_out"]
        PAN_EFFECTS  = ["pan_left", "pan_right", "pan_up", "pan_down"]

        def _opposite_category(effect: str) -> list:
            return PAN_EFFECTS if effect in ZOOM_EFFECTS else ZOOM_EFFECTS

        def _same_category(effect: str) -> list:
            pool = ZOOM_EFFECTS if effect in ZOOM_EFFECTS else PAN_EFFECTS
            return [e for e in pool if e != effect]

        prev_effect   = None
        reversal_streak = 0   # how many consecutive reversals we've done

        for img in self.images:
            if img.get("is_second_image", False):
                continue   # PiP slides don't need a KB assignment

            if prev_effect is None:
                # First image: pick randomly from all effects
                chosen = random.choice(ZOOM_EFFECTS + PAN_EFFECTS)
            else:
                # 20 % wildcard: jump to the opposite category
                if random.random() < 0.20:
                    chosen = random.choice(_opposite_category(prev_effect))
                    reversal_streak = 0
                # After 2+ reversals in a row: force a category switch
                elif reversal_streak >= 2:
                    chosen = random.choice(_opposite_category(prev_effect))
                    reversal_streak = 0
                else:
                    # Normal continuity: use the reversal
                    chosen = REVERSAL[prev_effect]
                    reversal_streak += 1

            img["ken_burns"] = chosen
            prev_effect = chosen

        self.update_image_table()



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
            self.images, self.premiere_project_folder,
            self.common_width, self.common_height,
            audio_files=self.audio_files,          # ← pass audio for XML music track
        )
        self.image_premiere_worker.progress.connect(self.update_image_premiere_progress)
        self.image_premiere_worker.finished.connect(self.on_image_premiere_processing_finished)
        self.image_premiere_worker.xml_ready.connect(self.on_premiere_xml_ready)
        self.image_premiere_worker.start()

    def on_premiere_xml_ready(self, xml_path: str):
        """Called when the Premiere XML timeline has been written to disk."""
        print(f"Premiere XML ready: {xml_path}")
        # Move the XML into the 04_project folder so it lives next to the .prproj
        dst_folder = os.path.join(self.premiere_project_folder, "04_פרוייקט")
        os.makedirs(dst_folder, exist_ok=True)
        dst = os.path.join(dst_folder, "premiere_timeline.xml")
        try:
            shutil.move(xml_path, dst)
            QMessageBox.information(
                self,
                "Premiere XML",
                f"Timeline XML saved:\n{dst}\n\nIn Premiere Pro: File → Import → premiere_timeline.xml",
            )
        except Exception as e:
            print(f"Could not move XML to project folder: {e}")

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
        style_src  = APP_DIR / "Premiere_Project" / "טקסט למצגת - עברית.prtextstyle"
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
        src  = APP_DIR / "Premiere_Project" / "Project.prproj"
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
        self.import_pptx_action.setText(self.tr("action_import_pptx"))
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
        self.set_all_ken_burns_action.setText(self.tr("set_all_ken"))
        self.set_random_ken_burns_action.setText(self.tr("random_ken"))
        self.set_smart_ken_burns_action.setText(self.tr("smart_ken"))
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
            self.tr("table_header_date"), self.tr("ken_burns"),
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
        self.import_pptx_action = _action("action_import_pptx", self.import_pptx,self.import_menu)
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

        self.set_all_ken_burns_action = QAction(self.tr("set_all_ken"), self)
        self.set_all_ken_burns_action.triggered.connect(self._set_all_ken_burns)
        self.Img_menu.addAction(self.set_all_ken_burns_action)

        self.set_random_ken_burns_action = QAction(self.tr("random_ken"), self)
        self.set_random_ken_burns_action.triggered.connect(self._set_random_ken_burns_per_image)
        self.Img_menu.addAction(self.set_random_ken_burns_action)

        self.set_smart_ken_burns_action = QAction(self.tr("smart_ken"), self)
        self.set_smart_ken_burns_action.triggered.connect(self._set_smart_ken_burns)
        self.Img_menu.addAction(self.set_smart_ken_burns_action)

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
    progress        = pyqtSignal(int)
    finished        = pyqtSignal()
    corrupted_image = pyqtSignal(str)   # emits path of any image that fails to open
    # Emitted after all work is done; carries the list of temp dirs to delete.
    cleanup_dirs = pyqtSignal(list)

    def __init__(self, images, output_folder, common_width, common_height):
        super().__init__()
        self.images        = images
        self.output_folder = output_folder
        self.common_width  = common_width
        self.common_height = common_height
        # Temp directories created during this export (populated in run()).
        self._temp_dirs: list[str] = []

    def _resize_one(self, i: int) -> tuple[int, str | None]:
        """Step 1: resize/blur image if needed. Run in parallel (PIL, no ffmpeg).

        When there is no Ken Burns effect the KB renderer never runs, so the
        regardless of what the checkbox says.
        """
        img_path   = self.images[i]["path"]
        rotation   = self.images[i]["rotation"]
        text       = self.images[i]["text"]
        has_kb     = self.images[i].get("ken_burns", "none") != "none"
        text_on_static = not has_kb
        try:
            original = Image.open(img_path)
            original.verify()          # catches truncated / corrupt files
            original = Image.open(img_path)   # re-open after verify() closes it
            if original.size != (self.common_width, self.common_height):
                new_path = Image_resizer.process_image(img_path, self.output_folder, text, rotation, text_on_static)
                return i, new_path
        except Exception as e:
            print(f"Corrupted image {img_path}: {e}")
            self.corrupted_image.emit(img_path)
        return i, None

    def run(self):
        total = len(self.images)

        # ── Phase 1: parallel image resize/blur ───────────────────────────────
        with ThreadPoolExecutor(max_workers=os.cpu_count() or 4) as executor:
            futures = {executor.submit(self._resize_one, i): i for i in range(total)}
            done = 0
            for future in as_completed(futures):
                done += 1
                try:
                    i, new_path = future.result()
                    if new_path:
                        self.images[i]["path"] = new_path
                except Exception as e:
                    print(f"Resize error: {e}")
                # Phase 1 = first 50% of progress bar
                self.progress.emit(int(done / total * 50))

        # ── Phase 2: limited-parallel Ken Burns rendering ─────────────────────
        # 2 workers: enough to overlap disk I/O + encoding; more causes contention.
        KB_WORKERS = min(2, os.cpu_count() or 1)
        kb_images = [i for i in range(total)
                     if self.images[i].get("ken_burns", "none") != "none"]
        kb_dir = os.path.join(self.output_folder, "kb_clips")
        if kb_images:
            os.makedirs(kb_dir, exist_ok=True)

        # Track both temp folders so the main thread can delete them later.
        self._temp_dirs = [self.output_folder, kb_dir]

        def _render_kb_one(i: int):
            img = self.images[i]
            effect = img["ken_burns"]
            if i == 0:
                clip_duration = img.get("duration", 5)
            elif i == len(self.images) - 1:
                clip_duration = img.get("duration", 5) + self.images[i - 1].get("transition_duration", 1)
            else:
                clip_duration = img.get("duration", 5) + img.get("transition_duration", 1)
            kb_out = os.path.join(kb_dir, f"kb_{i}_{effect}.mp4").replace("\\", "/")
            print(f"  KB render start: {effect} for image {i} ({clip_duration}s)")
            has_kb     = self.images[i].get("ken_burns", "none") != "none"
            text_on_static = not has_kb
            success = render_ken_burns_clip(
                img["path"], effect, clip_duration, kb_out,
                text=img.get("text", ""),
                text_on_kb=text_on_static,
            )
            return i, kb_out, success

        done_kb = 0
        with ThreadPoolExecutor(max_workers=KB_WORKERS) as kb_executor:
            kb_futures = {kb_executor.submit(_render_kb_one, i): i for i in kb_images}
            for future in as_completed(kb_futures):
                done_kb += 1
                try:
                    i, kb_out, success = future.result()
                    if success:
                        self.images[i]["_kb_clip_path"] = kb_out
                        print(f"  KB render done {done_kb}/{len(kb_images)}: image {i}")
                    else:
                        print(f"  KB render failed for image {i}, will use still image")
                except Exception as e:
                    print(f"  KB render exception: {e}")
                self.progress.emit(50 + int(done_kb / max(len(kb_images), 1) * 50))

        self.cleanup_dirs.emit(self._temp_dirs)
        self.finished.emit()


class ImageProcessingPremiereWorker(QThread):
    progress    = pyqtSignal(int)
    finished    = pyqtSignal()
    xml_ready   = pyqtSignal(str)   # emits path to the generated XML file

    def __init__(self, images, output_folder, common_width, common_height,
                 audio_files=None):
        super().__init__()
        self.images        = images
        self.output_folder = output_folder
        self.common_width  = common_width
        self.common_height = common_height
        self.audio_files   = audio_files or []

    def run(self):
        # ── Step 1: process images (backgrounds + foregrounds) ────────────────
        premiere_export.process_images(self.images, self.output_folder, self.progress.emit)

        # ── Step 2: build slide list mapping processed files → slide data ─────
        bg_folder  = os.path.join(self.output_folder, "01_images", "backgrounds")
        img_folder = os.path.join(self.output_folder, "01_images", "foregrounds")

        slide_list = []
        for i, img in enumerate(self.images, start=1):
            if img.get("is_second_image"):
                fg = os.path.join(img_folder, f"img{i}_2nd_of_img{i-1}.png")
                slide_list.append({
                    "bg_path":        None,
                    "fg_path":        fg if os.path.exists(fg) else None,
                    "duration":       img.get("duration", 5.0),
                    "text":           img.get("text", ""),
                    "is_second_image": True,
                })
            else:
                bg = os.path.join(bg_folder, f"background_img{i}.jpg")
                fg = os.path.join(img_folder, f"img{i}.png")
                slide_list.append({
                    "bg_path":        bg if os.path.exists(bg) else None,
                    "fg_path":        fg if os.path.exists(fg) else None,
                    "duration":       img.get("duration", 5.0),
                    "text":           img.get("text", ""),
                    "is_second_image": False,
                })

        # ── Step 3: generate the Premiere XML timeline ────────────────────────
        try:
            xml_path = premiere_export.generate_premiere_xml(
                slide_list   = slide_list,
                output_folder= self.output_folder,
                music_paths   = self.audio_files,
            )
            self.xml_ready.emit(xml_path)
        except Exception as e:
            print(f"XML generation error: {e}")

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
    import Image_resizer
    Image_resizer.sync_app_folders()

    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon("logo.ico"))
    set_theme(app, theme="dark")
    window = SlideshowCreator()
    window.create_menu()
    window.setup_connections()
    window.show()

    # ── File association fix: open .slideshow passed by Windows shell ─────────
    # When the user double-clicks a .slideshow file (or "Open with" the app),
    # Windows passes the file path as sys.argv[1].  Load it automatically.
    if len(sys.argv) > 1:
        arg = sys.argv[1]
        if arg.endswith(".slideshow") and os.path.exists(arg):
            window._load_project_from_path(arg)

    check_for_updates(window, APP_VERSION)
    sys.exit(app.exec_())