"""
Eventure.py  –  Redesigned UI
Full redesign: dark cinema aesthetic, card-based layout, modern toolbar,
sidebar preview, inline controls. All original functionality preserved.
"""

import base64
import numpy as np
import multiprocessing
import threading
from pptx_export import extract_pptx_content_to_slideshow_file
multiprocessing.freeze_support()

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
    QLineEdit, QFrame, QScrollArea, QSizePolicy, QToolBar, QStatusBar,
    QSplitter, QGridLayout, QToolButton, QMenu, QRadioButton, QButtonGroup,
)
from PyQt5.QtCore import pyqtSlot
from PyQt5.QtCore import Qt, QUrl, QSize, QProcess, QTimer, QThread, pyqtSignal, QEvent, QPoint, QRect
from PyQt5.QtGui import (
    QIcon, QFont, QPixmap, QTextCursor, QCursor, QTransform,
    QColor, QBrush, QImage, QPalette, QPainter, QLinearGradient,
    QFontDatabase, QPen, QPainterPath,
)
from PIL import Image, ExifTags, ImageDraw, ImageFont
from openpyxl import Workbook
import openpyxl

import premiere_export
import urllib.request
import urllib.error
import urllib.parse
import hashlib
import tempfile

APP_VERSION = "1.0.8"
# ── GitHub Audio Library Config ───────────────────────────────────────────────
GITHUB_AUDIO_USER   = "neriacohen300"
GITHUB_AUDIO_REPO   = "Eventure"
GITHUB_AUDIO_BRANCH = "main"
GITHUB_AUDIO_BASE   = f"https://raw.githubusercontent.com/{GITHUB_AUDIO_USER}/{GITHUB_AUDIO_REPO}/{GITHUB_AUDIO_BRANCH}/Songs"
GITHUB_AUDIO_INDEX  = f"{GITHUB_AUDIO_BASE}/songs.json"

# ── GPU Detection ──────────────────────────────────────────────────────────────

plugin_path = os.path.join(os.path.dirname(sys.executable), "Library", "plugins", "platforms")
os.environ["QT_QPA_PLATFORM_PLUGIN_PATH"] = plugin_path

BASEPATH = Path.home() / "Neria-LTD" / "Eventure"
BASEPATH.mkdir(parents=True, exist_ok=True)

if getattr(sys, "frozen", False):
    APP_DIR = Path(sys.executable).resolve().parent
else:
    APP_DIR = Path(__file__).resolve().parent

_ORIENTATION_TAG = next(
    (k for k, v in ExifTags.TAGS.items() if v == "Orientation"), None
)

# ── Design System ─────────────────────────────────────────────────────────────

COLORS = {
    "bg_deep":       "#1E2228",
    "bg_panel":      "#252B34",
    "bg_card":       "#2C333D",
    "bg_hover":      "#333B47",
    "bg_selected":   "#2B3A50",
    "border":        "#363E4A",
    "border_light":  "#404A58",
    "accent":        "#5B9BFF",
    "accent_dim":    "#2D4E7A",
    "accent_glow":   "rgba(91,155,255,0.15)",
    "success":       "#4ADB9A",
    "warning":       "#F5A623",
    "danger":        "#F76E6E",
    "purple":        "#A67CFF",
    "text_primary":  "#E8EDF5",
    "text_secondary":"#95A0B4",
    "text_muted":    "#647080",
    "header_bg":     "#20262F",
    "toolbar_bg":    "#232930",
    "second_image":  "rgba(166,124,255,0.12)",
}

STYLESHEET = f"""
/* ── Base ── */
QMainWindow, QWidget {{
    background-color: {COLORS['bg_deep']};
    color: {COLORS['text_primary']};
    font-family: 'Segoe UI', 'SF Pro Display', sans-serif;
    font-size: 13px;
}}

/* ── MenuBar ── */
QMenuBar {{
    background-color: {COLORS['header_bg']};
    color: {COLORS['text_secondary']};
    border-bottom: 1px solid {COLORS['border']};
    padding: 2px 8px;
    spacing: 4px;
    font-size: 12px;
    letter-spacing: 0.3px;
}}
QMenuBar::item {{
    background: transparent;
    padding: 5px 12px;
    border-radius: 4px;
}}
QMenuBar::item:selected {{
    background-color: {COLORS['bg_hover']};
    color: {COLORS['text_primary']};
}}
QMenu {{
    background-color: {COLORS['bg_card']};
    border: 1px solid {COLORS['border_light']};
    border-radius: 8px;
    padding: 6px 4px;
    color: {COLORS['text_primary']};
}}
QMenu::item {{
    padding: 7px 28px 7px 16px;
    border-radius: 4px;
    margin: 1px 4px;
}}
QMenu::item:selected {{
    background-color: {COLORS['bg_hover']};
    color: {COLORS['accent']};
}}
QMenu::separator {{
    height: 1px;
    background: {COLORS['border']};
    margin: 4px 8px;
}}

/* ── Toolbar ── */
QToolBar {{
    background-color: {COLORS['toolbar_bg']};
    border-bottom: 1px solid {COLORS['border']};
    padding: 4px 12px;
    spacing: 6px;
}}
QToolButton {{
    background-color: transparent;
    color: {COLORS['text_secondary']};
    border: none;
    border-radius: 6px;
    padding: 6px 14px;
    font-size: 12px;
    font-weight: 500;
}}
QToolButton:hover {{
    background-color: {COLORS['bg_hover']};
    color: {COLORS['text_primary']};
}}
QToolButton:pressed {{
    background-color: {COLORS['bg_selected']};
    color: {COLORS['accent']};
}}
QToolButton[accent="true"] {{
    background-color: {COLORS['accent_dim']};
    color: {COLORS['accent']};
    border: 1px solid {COLORS['accent_dim']};
}}
QToolButton[accent="true"]:hover {{
    background-color: {COLORS['accent']};
    color: #FFFFFF;
}}

/* ── Table ── */
QTableWidget {{
    background-color: {COLORS['bg_panel']};
    border: 1px solid {COLORS['border']};
    border-radius: 10px;
    gridline-color: {COLORS['border']};
    color: {COLORS['text_primary']};
    selection-background-color: {COLORS['bg_selected']};
    selection-color: {COLORS['text_primary']};
    outline: none;
    font-size: 12px;
}}
QTableWidget::item {{
    padding: 6px 10px;
    border: none;
}}
QTableWidget::item:selected {{
    background-color: {COLORS['bg_selected']};
    border-left: 2px solid {COLORS['accent']};
}}
QHeaderView::section {{
    background-color: {COLORS['bg_deep']};
    color: {COLORS['text_muted']};
    border: none;
    border-bottom: 1px solid {COLORS['border']};
    padding: 8px 10px;
    font-size: 11px;
    font-weight: 600;
    letter-spacing: 0.8px;
    text-transform: uppercase;
}}
QHeaderView {{
    background-color: {COLORS['bg_deep']};
    border: none;
}}

/* ── Buttons ── */
QPushButton {{
    background-color: {COLORS['bg_card']};
    color: {COLORS['text_primary']};
    border: 1px solid {COLORS['border_light']};
    border-radius: 6px;
    padding: 5px 12px;
    font-size: 12px;
    font-weight: 500;
    min-height: 26px;
}}
QPushButton:hover {{
    background-color: {COLORS['bg_hover']};
    border-color: {COLORS['accent']};
    color: {COLORS['accent']};
}}
QPushButton:pressed {{
    background-color: {COLORS['bg_selected']};
}}
QPushButton[action="primary"] {{
    background-color: {COLORS['accent']};
    border-color: {COLORS['accent']};
    color: #ffffff;
    font-weight: 600;
}}
QPushButton[action="primary"]:hover {{
    background-color: #6BA0FF;
    color: #ffffff;
}}
QPushButton[action="danger"] {{
    background-color: transparent;
    border-color: {COLORS['danger']};
    color: {COLORS['danger']};
}}
QPushButton[action="danger"]:hover {{
    background-color: {COLORS['danger']};
    color: white;
}}
QPushButton[action="icon"] {{
    background: transparent;
    border: none;
    padding: 3px 6px;
    color: {COLORS['text_muted']};
    font-size: 14px;
    min-height: 20px;
    border-radius: 4px;
}}
QPushButton[action="icon"]:hover {{
    background: {COLORS['bg_hover']};
    color: {COLORS['text_primary']};
}}

/* ── ComboBox ── */
QComboBox {{
    background-color: {COLORS['bg_card']};
    color: {COLORS['text_primary']};
    border: 1px solid {COLORS['border']};
    border-radius: 5px;
    padding: 4px 8px;
    font-size: 12px;
    min-width: 90px;
}}
QComboBox:hover {{
    border-color: {COLORS['accent']};
}}
QComboBox::drop-down {{
    border: none;
    width: 20px;
}}
QComboBox::down-arrow {{
    image: none;
    border-left: 4px solid transparent;
    border-right: 4px solid transparent;
    border-top: 5px solid {COLORS['text_muted']};
    margin-right: 6px;
}}
QComboBox QAbstractItemView {{
    background-color: {COLORS['bg_card']};
    border: 1px solid {COLORS['border_light']};
    border-radius: 6px;
    color: {COLORS['text_primary']};
    selection-background-color: {COLORS['bg_hover']};
    padding: 4px;
}}

/* ── CheckBox ── */
QCheckBox {{
    color: {COLORS['text_secondary']};
    spacing: 6px;
}}
QCheckBox::indicator {{
    width: 16px;
    height: 16px;
    border: 1.5px solid {COLORS['border_light']};
    border-radius: 4px;
    background: {COLORS['bg_card']};
}}
QCheckBox::indicator:checked {{
    background-color: {COLORS['accent']};
    border-color: {COLORS['accent']};
}}

/* ── LineEdit / TextEdit ── */
QLineEdit, QTextEdit {{
    background-color: {COLORS['bg_card']};
    color: {COLORS['text_primary']};
    border: 1px solid {COLORS['border']};
    border-radius: 6px;
    padding: 6px 10px;
    font-size: 13px;
    selection-background-color: {COLORS['accent_dim']};
}}
QLineEdit:focus, QTextEdit:focus {{
    border-color: {COLORS['accent']};
    background-color: {COLORS['bg_hover']};
}}
QLineEdit::placeholder {{
    color: {COLORS['text_muted']};
}}

/* ── SpinBox ── */
QSpinBox {{
    background-color: {COLORS['bg_card']};
    color: {COLORS['text_primary']};
    border: 1px solid {COLORS['border']};
    border-radius: 5px;
    padding: 4px 6px;
}}
QSpinBox:focus {{
    border-color: {COLORS['accent']};
}}

/* ── ProgressBar ── */
QProgressBar {{
    background-color: {COLORS['bg_card']};
    border: none;
    border-radius: 3px;
    height: 4px;
    text-align: center;
    color: transparent;
}}
QProgressBar::chunk {{
    border-radius: 3px;
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
        stop:0 {COLORS['accent']}, stop:1 #7AB8FF);
}}
QProgressBar[type="kb"]::chunk {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
        stop:0 {COLORS['purple']}, stop:1 #C49FFF);
}}
QProgressBar[type="premiere"]::chunk {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
        stop:0 {COLORS['success']}, stop:1 #7FFFD4);
}}

/* ── ScrollBar ── */
QScrollBar:vertical {{
    background: transparent;
    width: 8px;
    border-radius: 4px;
}}
QScrollBar::handle:vertical {{
    background: {COLORS['border_light']};
    border-radius: 4px;
    min-height: 20px;
}}
QScrollBar::handle:vertical:hover {{
    background: {COLORS['text_muted']};
}}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
    height: 0;
}}
QScrollBar:horizontal {{
    background: transparent;
    height: 8px;
    border-radius: 4px;
}}
QScrollBar::handle:horizontal {{
    background: {COLORS['border_light']};
    border-radius: 4px;
    min-width: 20px;
}}
QScrollBar::handle:horizontal:hover {{
    background: {COLORS['text_muted']};
}}
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{
    width: 0;
}}

/* ── Frames & Panels ── */
QFrame[class="panel"] {{
    background-color: {COLORS['bg_panel']};
    border: 1px solid {COLORS['border']};
    border-radius: 10px;
}}
QFrame[class="divider"] {{
    background-color: {COLORS['border']};
    max-height: 1px;
}}

/* ── Labels ── */
QLabel[class="section-title"] {{
    color: {COLORS['text_muted']};
    font-size: 10px;
    font-weight: 700;
    letter-spacing: 1.2px;
    text-transform: uppercase;
}}
QLabel[class="preview-empty"] {{
    color: {COLORS['text_muted']};
    font-size: 13px;
    background-color: {COLORS['bg_card']};
    border: 1px dashed {COLORS['border_light']};
    border-radius: 8px;
}}

/* ── StatusBar ── */
QStatusBar {{
    background-color: {COLORS['header_bg']};
    color: {COLORS['text_muted']};
    border-top: 1px solid {COLORS['border']};
    font-size: 11px;
    padding: 0 12px;
}}

/* ── Dialog ── */
QDialog {{
    background-color: {COLORS['bg_panel']};
    color: {COLORS['text_primary']};
    border: 1px solid {COLORS['border_light']};
    border-radius: 12px;
}}

/* ── Splitter ── */
QSplitter::handle {{
    background-color: {COLORS['border']};
}}
QSplitter::handle:horizontal {{
    width: 1px;
}}

/* ── ListWidget ── */
QListWidget {{
    background-color: {COLORS['bg_card']};
    border: 1px solid {COLORS['border']};
    border-radius: 8px;
    color: {COLORS['text_primary']};
    outline: none;
    padding: 4px;
}}
QListWidget::item {{
    padding: 8px 12px;
    border-radius: 6px;
    margin: 1px 0;
}}
QListWidget::item:selected {{
    background-color: {COLORS['bg_selected']};
    color: {COLORS['accent']};
}}
QListWidget::item:hover {{
    background-color: {COLORS['bg_hover']};
}}
"""


def _make_section_label(text: str) -> QLabel:
    lbl = QLabel(text.upper())
    lbl.setProperty("class", "section-title")
    return lbl


def _make_divider() -> QFrame:
    f = QFrame()
    f.setProperty("class", "divider")
    f.setFrameShape(QFrame.HLine)
    f.setFixedHeight(1)
    return f


def _styled_btn(text: str, action: str = "") -> QPushButton:
    btn = QPushButton(text)
    if action:
        btn.setProperty("action", action)
    return btn


def _make_crop_btn(has_crop: bool) -> QPushButton:
    """Return a styled ✂ Crop button; glows orange when a crop is active."""
    crop_btn = QPushButton("✂ Crop")
    crop_btn.setFixedHeight(22)
    if has_crop:
        crop_btn.setStyleSheet(
            f"QPushButton {{ background: {COLORS['warning']}; color: #1a1a1a; "
            f"border: none; border-radius: 4px; padding: 2px 7px; "
            f"font-size: 11px; font-weight: 700; }}"
            f"QPushButton:hover {{ background: #ffc53d; }}"
        )
        crop_btn.setToolTip("Edit crop  (crop active)")
    else:
        crop_btn.setStyleSheet(
            f"QPushButton {{ background: {COLORS['bg_hover']}; color: {COLORS['text_secondary']}; "
            f"border: 1px solid {COLORS['border_light']}; border-radius: 4px; "
            f"padding: 2px 7px; font-size: 11px; font-weight: 600; }}"
            f"QPushButton:hover {{ background: {COLORS['accent_dim']}; color: {COLORS['accent']}; "
            f"border-color: {COLORS['accent']}; }}"
        )
        crop_btn.setToolTip("Crop image")
    return crop_btn



# ── Update check ─────────────────────────────────────────────────────────────

def check_for_updates(parent_window, current_version: str):
    GITHUB_USER = "neriacohen300"
    GITHUB_REPO = "Eventure"
    API_URL = f"https://api.github.com/repos/{GITHUB_USER}/{GITHUB_REPO}/releases/latest"

    def _fetch():
        try:
            import urllib.request, json
            req = urllib.request.Request(API_URL, headers={"User-Agent": "Eventure-App"})
            with urllib.request.urlopen(req, timeout=5) as resp:
                data = json.loads(resp.read().decode())
            latest_tag  = data.get("tag_name", "").lstrip("v")
            release_url = data.get("html_url", "")
            changelog   = data.get("body", "").strip()
            if not latest_tag:
                return
            def _ver(s):
                try:    return tuple(int(x) for x in s.split("."))
                except: return (0,)
            if _ver(latest_tag) > _ver(current_version):
                from PyQt5.QtCore import QMetaObject, Qt, Q_ARG
                QMetaObject.invokeMethod(
                    parent_window, "_show_update_dialog", Qt.QueuedConnection,
                    Q_ARG(str, latest_tag), Q_ARG(str, release_url), Q_ARG(str, changelog),
                )
        except Exception as e:
            print(f"Update check failed: {e}")

    threading.Thread(target=_fetch, daemon=True).start()


# ── Helpers ───────────────────────────────────────────────────────────────────

def _ffmpeg_exe() -> str:
    import shutil as _shutil
    return _shutil.which("ffmpeg") or str(APP_DIR / "ffmpeg.exe")

def _ffprobe_exe() -> str:
    import shutil as _shutil
    return _shutil.which("ffprobe") or str(APP_DIR / "ffprobe.exe")

def _get_audio_duration(audio_path: str) -> float:
    try:
        result = subprocess.run(
            [_ffprobe_exe(), "-v", "error", "-show_entries", "format=duration",
             "-of", "default=noprint_wrappers=1:nokey=1", audio_path],
            capture_output=True, text=True, timeout=10,
            creationflags=subprocess.CREATE_NO_WINDOW if hasattr(subprocess, "CREATE_NO_WINDOW") else 0
        )
        return float(result.stdout.strip())
    except Exception as e:
        print(f"ffprobe error for {audio_path}: {e}")
        return 0.0


def format_time_srt(seconds: float) -> str:
    h   = int(seconds // 3600)
    m   = int((seconds % 3600) // 60)
    s   = int(seconds % 60)
    ms  = int((seconds - int(seconds)) * 1000)
    return f"{h:02}:{m:02}:{s:02},{ms:03}"


def format_time_hms(seconds: float) -> str:
    h = int(seconds // 3600)
    m = int((seconds % 3600) // 60)
    s = int(seconds % 60)
    return f"{h:02}:{m:02}:{s:02}"


def _copy_resource_folders(script_dir: Path, resources: list) -> None:
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


# ── Ken Burns Renderer ────────────────────────────────────────────────────────

_KB_FPS = 25

def render_ken_burns_clip(image_path, effect, duration, output_path, rotation=0, text="", text_on_kb=True, crop=None):
    import cv2 as _cv2
    import numpy as _np
    import subprocess as _sp
    from PIL import Image as _Image, ImageDraw as _Draw, ImageFont as _Font
    from PIL import ExifTags as _ExifTags

    W, H, FPS = 1920, 1080, _KB_FPS
    frames = max(1, int(duration * FPS))

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

    # ── Algorithm ─────────────────────────────────────────────────────────────
    # 1. Scale image to COVER W x H (no letterbox/pillarbox).
    # 2. Ken Burns = animate a sub-rect of that cover canvas.
    #    zoom_in: rect starts at W*ZOOM x H*ZOOM, shrinks to W x H (zooms in).
    #    zoom_out: rect grows from W x H to W*ZOOM x H*ZOOM (zooms out).
    #    pan_*: shift the W x H window across the canvas.
    # This works identically for landscape, portrait, and cropped images.

    def smooth(raw_t):
        t = max(0.0, min(1.0, raw_t))
        return t * t * (3.0 - 2.0 * t)

    ZOOM = 1.10
    img_arr = _np.array(img)
    ih, iw  = img_arr.shape[:2]

    # Cover-scale to W*ZOOM x H*ZOOM so every animated rect fits without clamping
    scale_cover = max(W * ZOOM / iw, H * ZOOM / ih)
    cov_w = int(round(iw * scale_cover))
    cov_h = int(round(ih * scale_cover))
    canvas = _cv2.resize(img_arr, (cov_w, cov_h), interpolation=_cv2.INTER_LANCZOS4)

    # Pan travel clamped to actual canvas headroom in each axis so the window
    # never gets clipped — the vertical headroom (cov_h-H)/2 is smaller than
    # (cov_w-W)/2, which is why pan_up/down appeared frozen at the start.
    h_travel = min((cov_w - W) / 2.0, min(W, H) * 0.08)   # horizontal headroom
    v_travel = min((cov_h - H) / 2.0, min(W, H) * 0.08)   # vertical headroom

    rects = []
    for f in range(frames):
        t = smooth(f / max(frames - 1, 1))
        sw, sh = float(W), float(H)
        sx = (cov_w - W) / 2.0
        sy = (cov_h - H) / 2.0

        if effect == "zoom_in":
            z = ZOOM - (ZOOM - 1.0) * t        # ZOOM → 1.0
            sw, sh = W * z, H * z
            sx = (cov_w - sw) / 2.0
            sy = (cov_h - sh) / 2.0
        elif effect == "zoom_out":
            z = 1.0 + (ZOOM - 1.0) * t         # 1.0 → ZOOM
            sw, sh = W * z, H * z
            sx = (cov_w - sw) / 2.0
            sy = (cov_h - sh) / 2.0
        elif effect == "pan_left":
            sx = (cov_w - W) / 2.0 + h_travel * (1.0 - t)
            sy = (cov_h - H) / 2.0
        elif effect == "pan_right":
            sx = (cov_w - W) / 2.0 - h_travel * (1.0 - t)
            sy = (cov_h - H) / 2.0
        elif effect == "pan_up":
            sx = (cov_w - W) / 2.0
            sy = (cov_h - H) / 2.0 + v_travel * (1.0 - t)
        elif effect == "pan_down":
            sx = (cov_w - W) / 2.0
            sy = (cov_h - H) / 2.0 - v_travel * (1.0 - t)

        sx = max(0.0, min(float(cov_w) - sw, sx))
        sy = max(0.0, min(float(cov_h) - sh, sy))
        rects.append((sx, sy, sw, sh))

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
            bbox = draw.textbbox((0, 0), htext, font=_font)
            tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
            bg_w, bg_h = tw + 40, th + 20
            bg_x = (W - bg_w) // 2; bg_y = H - bg_h - 50
            draw.rounded_rectangle((bg_x, bg_y, bg_x+bg_w, bg_y+bg_h), radius=12, fill="white")
            draw.text(((W - tw) // 2, bg_y - 4), htext, font=_font, fill="black")
            ov_arr = _np.array(overlay)
            alpha = ov_arr[:, :, 3:4].astype(_np.float32) / 255.0
            static_overlay_bgr = ov_arr[:, :, :3].astype(_np.float32)
            static_overlay_mask = alpha
        except Exception as e:
            print(f"KB text overlay error: {e}")

    frame_buffer = _np.empty((frames, H, W, 3), dtype=_np.uint8)
    for f in range(frames):
        sx, sy, sw, sh = rects[f]
        scale_x = W / sw
        scale_y = H / sh
        M = _np.array([[scale_x, 0.0, -sx * scale_x],
                       [0.0, scale_y, -sy * scale_y]], dtype=_np.float64)
        frame = _cv2.warpAffine(canvas, M, (W, H),
                                flags=_cv2.INTER_LANCZOS4,
                                borderMode=_cv2.BORDER_REFLECT_101)
        if static_overlay_bgr is not None:
            frame = (frame.astype(_np.float32) * (1.0 - static_overlay_mask)
                     + static_overlay_bgr * static_overlay_mask).clip(0, 255).astype(_np.uint8)
        frame_buffer[f] = frame

    vcodec_kb = ["-vcodec", "libx264", "-pix_fmt", "yuv420p", "-preset", "ultrafast", "-crf", "28"]
    extra_kb  = ["-tune", "fastdecode"]

    cmd = [
        _ffmpeg_exe(), "-y",
        "-f", "rawvideo", "-vcodec", "rawvideo",
        "-s", f"{W}x{H}", "-pix_fmt", "rgb24", "-r", str(FPS),
        "-i", "pipe:0",
    ] + vcodec_kb + extra_kb + [
        "-g", "25",
        "-r", str(FPS), "-movflags", "+faststart",
        output_path,
    ]
    proc = None
    try:
        proc = _sp.Popen(
            cmd, stdin=_sp.PIPE, stdout=_sp.DEVNULL, stderr=_sp.DEVNULL,
            creationflags=_sp.CREATE_NO_WINDOW if hasattr(_sp, "CREATE_NO_WINDOW") else 0,
        )
        proc.stdin.write(frame_buffer.tobytes())
        proc.stdin.close()
        proc.wait()
        return proc.returncode == 0
    except Exception as e:
        print(f"KB render pipe error: {e}")
        if proc:
            try: proc.stdin.close()
            except Exception: pass
        return False


# ── Taskbar Progress Stub ─────────────────────────────────────────────────────

class _TaskbarProgressStub:
    def show(self):          pass
    def hide(self):          pass
    def reset(self):         pass
    def setValue(self, v):   pass
    def setVisible(self, v): pass


# ── Custom Delegate ───────────────────────────────────────────────────────────

class CustomDelegate(QStyledItemDelegate):
    """Paints alternating row tints and a purple highlight for second-image rows."""

    _ROW_EVEN = QColor(0x25, 0x2B, 0x34, 255)   # bg_panel
    _ROW_ODD  = QColor(0x28, 0x2F, 0x39, 255)   # slightly lighter
    _SECOND   = QColor(80, 55, 120, 55)           # subtle purple tint

    def paint(self, painter, option, index):
        is_secondary = index.model().index(index.row(), 1).data(Qt.UserRole)
        if is_secondary:
            painter.save()
            painter.fillRect(option.rect, self._SECOND)
            painter.restore()
        elif index.row() % 2 == 1:
            # Odd rows get a very subtle alternate tint for scannability
            painter.save()
            painter.fillRect(option.rect, self._ROW_ODD)
            painter.restore()
        super().paint(painter, option, index)


# ── Preview Panel ─────────────────────────────────────────────────────────────

class PreviewPanel(QFrame):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setProperty("class", "panel")
        self.setMinimumHeight(220)
        self.setMinimumWidth(280)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        title = _make_section_label("Preview")
        layout.addWidget(title)

        self.image_label = QLabel()
        self.image_label.setAlignment(Qt.AlignCenter)
        self.image_label.setProperty("class", "preview-empty")
        self.image_label.setMinimumHeight(170)
        self.image_label.setText("No image selected")
        layout.addWidget(self.image_label)

        self.filename_label = QLabel()
        self.filename_label.setAlignment(Qt.AlignCenter)
        self.filename_label.setStyleSheet(f"color: {COLORS['text_muted']}; font-size: 11px;")
        layout.addWidget(self.filename_label)

    def set_pixmap(self, pixmap: QPixmap, filename: str = ""):
        if pixmap and not pixmap.isNull():
            self.image_label.setPixmap(pixmap.scaled(
                self.image_label.width() - 4,
                self.image_label.height() - 4,
                Qt.KeepAspectRatio, Qt.SmoothTransformation
            ))
            self.image_label.setProperty("class", "")
        else:
            self.image_label.clear()
            self.image_label.setText("No image selected")
            self.image_label.setProperty("class", "preview-empty")
        self.filename_label.setText(filename)
        self.style().unpolish(self.image_label)
        self.style().polish(self.image_label)

    def clear(self):
        self.image_label.clear()
        self.image_label.setText("No image selected")
        self.filename_label.setText("")


# ── Stats Panel ───────────────────────────────────────────────────────────────

class StatsBar(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(16, 4, 16, 4)
        layout.setSpacing(6)

        self._labels = {}
        entries = [
            ("slides",   "🎞",  "0 slides"),
            ("duration", "⏱",  "0:00:00"),
            ("audio",    "🎵",  "No audio"),
        ]
        for i, (key, icon, default) in enumerate(entries):
            if i > 0:
                sep = QLabel("·")
                sep.setStyleSheet(f"color: {COLORS['border_light']}; font-size: 14px;")
                layout.addWidget(sep)
            icon_lbl = QLabel(icon)
            icon_lbl.setStyleSheet(f"color: {COLORS['text_muted']}; font-size: 12px;")
            lbl = QLabel(default)
            lbl.setStyleSheet(f"color: {COLORS['text_muted']}; font-size: 11px;")
            layout.addWidget(icon_lbl)
            layout.addWidget(lbl)
            self._labels[key] = lbl

        layout.addStretch()

    def update_stats(self, n_slides, duration_sec, audio_count):
        self._labels["slides"].setText(f"{n_slides} slides")
        self._labels["duration"].setText(format_time_hms(duration_sec))
        self._labels["audio"].setText(
            f"{audio_count} audio file{'s' if audio_count != 1 else ''}"
        )


# ── Progress Section ──────────────────────────────────────────────────────────

class ProgressSection(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 4, 0, 0)
        layout.setSpacing(6)

        _PCT_STYLE = (
            f"color: {COLORS['text_secondary']}; font-size: 11px; "
            f"font-weight: 600; min-width: 34px; text-align: right;"
        )

        def _bar_row(color_type):
            """Return (container_widget, bar, pct_label) packed in an HBoxLayout."""
            container = QWidget()
            row = QHBoxLayout(container)
            row.setContentsMargins(0, 0, 0, 0)
            row.setSpacing(6)

            bar = QProgressBar()
            bar.setRange(0, 100)
            bar.setValue(0)
            bar.setFixedHeight(4)
            bar.setProperty("type", color_type)

            pct_lbl = QLabel("0%")
            pct_lbl.setStyleSheet(_PCT_STYLE)
            pct_lbl.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
            pct_lbl.setFixedWidth(34)

            row.addWidget(bar)
            row.addWidget(pct_lbl)

            container.setVisible(False)
            return container, bar, pct_lbl

        self._export_container,  self.export_bar,   self._export_pct   = _bar_row("")
        self._image_container,   self.image_bar,    self._image_pct    = _bar_row("kb")
        self._premiere_container, self.premiere_bar, self._premiere_pct = _bar_row("premiere")

        for container in [self._export_container, self._image_container, self._premiere_container]:
            layout.addWidget(container)

        self.status_label = QLabel()
        self.status_label.setStyleSheet(f"color: {COLORS['text_muted']}; font-size: 11px;")
        self.status_label.setVisible(False)
        layout.addWidget(self.status_label)

    def set_export_progress(self, value: int, visible: bool = True):
        self._export_container.setVisible(visible)
        if visible:
            self.export_bar.setValue(value)
            self._export_pct.setText(f"{value}%")

    def set_image_progress(self, value: int, visible: bool = True):
        self._image_container.setVisible(visible)
        if visible:
            self.image_bar.setValue(value)
            self._image_pct.setText(f"{value}%")

    def set_premiere_progress(self, value: int, visible: bool = True):
        self._premiere_container.setVisible(visible)
        if visible:
            self.premiere_bar.setValue(value)
            self._premiere_pct.setText(f"{value}%")

    # ── Visibility helpers (used by external callers) ─────────────────────────
    def show_export(self):
        self._export_container.setVisible(True)

    def hide_export(self):
        self._export_container.setVisible(False)

    def show_image(self):
        self._image_container.setVisible(True)

    def hide_image(self):
        self._image_container.setVisible(False)

    def show_premiere(self):
        self._premiere_container.setVisible(True)

    def hide_premiere(self):
        self._premiere_container.setVisible(False)



# ── Filmstrip Timeline ────────────────────────────────────────────────────────

# ── Background thumbnail loader ───────────────────────────────────────────────

class _ThumbLoader(QThread):
    """Loads one thumbnail off the UI thread and emits it when ready."""
    thumb_ready = pyqtSignal(str, object)   # (path, QPixmap)

    THUMB_W = 200
    THUMB_H = 120

    def __init__(self, path: str, parent=None):
        super().__init__(parent)
        self.path = path

    def run(self):
        try:
            img = Image.open(self.path)
            img.thumbnail((self.THUMB_W, self.THUMB_H), Image.LANCZOS)
            img = img.convert("RGBA")
            data = img.tobytes("raw", "RGBA")
            qi   = QImage(data, img.width, img.height, QImage.Format_RGBA8888)
            px   = QPixmap.fromImage(qi)
        except Exception:
            px = QPixmap()
        self.thumb_ready.emit(self.path, px)


# ── Shared canvas logic (reused by inline strip and full-view dialog) ─────────

class _FilmstripCanvas(QWidget):
    """
    Painted canvas: drag-to-reorder cards, right-click context menu.
    Works both as the inner widget of FilmstripTimeline and inside
    FilmstripFullDialog (larger cards).
    """
    order_changed   = pyqtSignal(int, int)   # (from_idx, to_idx)
    delete_at       = pyqtSignal(int)
    move_to         = pyqtSignal(int, int)   # (current_idx, target_pos_1based)
    card_clicked    = pyqtSignal(int)        # for syncing table selection

    CARD_W  = 110
    CARD_H  = 90
    THUMB_H = 60
    GAP     = 8
    RADIUS  = 8

    def __init__(self, card_w=110, card_h=90, thumb_h=60, gap=8, parent=None):
        super().__init__(parent)
        self.CARD_W  = card_w
        self.CARD_H  = card_h
        self.THUMB_H = thumb_h
        self.GAP     = gap

        self.images: list        = []
        self._thumbs: dict       = {}          # path → QPixmap (None = loading)
        self._loaders: dict      = {}          # path → _ThumbLoader
        self._drag_idx: int      = -1
        self._drag_abs_x: int    = 0
        self._drag_offset: int   = 0
        self._hover_idx: int     = -1
        self._drop_idx: int      = -1
        self._selected_idx: int  = -1

        self.setMouseTracking(True)
        self.setCursor(Qt.OpenHandCursor)
        self.setContextMenuPolicy(Qt.CustomContextMenu)
        self.customContextMenuRequested.connect(self._show_context_menu)

    # ── public ───────────────────────────────────────────────────────────────

    def set_images(self, images: list):
        self.images = images
        # Cancel loaders for paths no longer present
        current_paths = {img["path"] for img in images}
        for path in list(self._loaders.keys()):
            if path not in current_paths:
                self._loaders.pop(path, None)
        # Evict stale cached thumbs
        self._thumbs = {k: v for k, v in self._thumbs.items() if k in current_paths}
        # Kick off loaders for any new paths
        for img in images:
            self._request_thumb(img["path"])
        self._resize_canvas()
        self.update()

    def set_selected(self, idx: int):
        self._selected_idx = idx
        self.update()

    # ── thumb loading ────────────────────────────────────────────────────────

    def _request_thumb(self, path: str):
        if path in self._thumbs or path in self._loaders:
            return
        self._thumbs[path] = None   # sentinel: loading in progress
        loader = _ThumbLoader(path, self)
        loader.thumb_ready.connect(self._on_thumb_ready)
        self._loaders[path] = loader
        loader.start()

    def _on_thumb_ready(self, path: str, px: QPixmap):
        # Scale to card dimensions now that we're back on the UI thread
        if not px.isNull():
            px = px.scaled(self.CARD_W - 12, self.THUMB_H - 4,
                           Qt.KeepAspectRatio, Qt.SmoothTransformation)
        self._thumbs[path] = px
        self._loaders.pop(path, None)
        self.update()

    # ── geometry ─────────────────────────────────────────────────────────────

    def _card_x(self, idx: int) -> int:
        return self.GAP + idx * (self.CARD_W + self.GAP)

    def _resize_canvas(self):
        n = len(self.images)
        w = self.GAP + n * (self.CARD_W + self.GAP)
        h = self.CARD_H + self.GAP * 2
        self.setMinimumSize(max(w, 200), h)
        self.resize(max(w, 200), h)

    def _idx_at(self, x: int) -> int:
        for i in range(len(self.images)):
            cx = self._card_x(i)
            if cx <= x <= cx + self.CARD_W:
                return i
        return -1

    def _drop_pos_at(self, x: int) -> int:
        n = len(self.images)
        for i in range(n):
            mid = self._card_x(i) + self.CARD_W // 2
            if x < mid:
                return i
        return n

    # ── context menu ─────────────────────────────────────────────────────────

    def _show_context_menu(self, pos):
        idx = self._idx_at(pos.x())
        if idx < 0:
            return
        img  = self.images[idx]
        name = os.path.basename(img["path"])

        menu = QMenu(self)
        menu.setStyleSheet(
            f"QMenu {{ background: {COLORS['bg_card']}; border: 1px solid {COLORS['border_light']};"
            f"  border-radius: 8px; padding: 4px; color: {COLORS['text_primary']}; }}"
            f"QMenu::item {{ padding: 7px 24px 7px 14px; border-radius: 4px; margin: 1px 4px; }}"
            f"QMenu::item:selected {{ background: {COLORS['bg_hover']}; color: {COLORS['accent']}; }}"
            f"QMenu::separator {{ height: 1px; background: {COLORS['border']}; margin: 4px 8px; }}"
        )

        # Header (non-interactive title)
        title_action = QAction(f"#{idx + 1}  {name[:28]}", menu)
        title_action.setEnabled(False)
        menu.addAction(title_action)
        menu.addSeparator()

        move_action   = QAction("✦  Set Position…", menu)
        delete_action = QAction("✕  Delete", menu)
        delete_action.setProperty("danger", True)

        menu.addAction(move_action)
        menu.addSeparator()
        menu.addAction(delete_action)

        action = menu.exec_(self.mapToGlobal(pos))

        if action == move_action:
            n = len(self.images)
            new_pos, ok = QInputDialog.getInt(
                self, "Set Position",
                f"Move slide #{idx + 1} to position (1–{n}):",
                idx + 1, 1, n
            )
            if ok and new_pos - 1 != idx:
                self.move_to.emit(idx, new_pos)

        elif action == delete_action:
            self.delete_at.emit(idx)

    # ── painting ─────────────────────────────────────────────────────────────

    def paintEvent(self, _event):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        p.setRenderHint(QPainter.SmoothPixmapTransform)
        p.fillRect(self.rect(), QColor(COLORS["bg_deep"]))

        if not self.images:
            p.setPen(QColor(COLORS["text_muted"]))
            p.setFont(QFont("Segoe UI", 11))
            p.drawText(self.rect(), Qt.AlignCenter,
                       "Add images — they'll appear here for easy drag-to-reorder")
            p.end()
            return

        for i in range(len(self.images)):
            if i == self._drag_idx:
                continue
            self._draw_card(p, i, self._card_x(i), self.GAP, dragging=False)

        # Drop indicator line + dot
        if self._drag_idx >= 0 and self._drop_idx >= 0:
            if self._drop_idx < len(self.images):
                lx = self._card_x(self._drop_idx) - self.GAP // 2 - 1
            else:
                lx = self._card_x(len(self.images) - 1) + self.CARD_W + self.GAP // 2
            p.setPen(QColor(COLORS["accent"]))
            p.setBrush(QBrush(QColor(COLORS["accent"])))
            p.drawLine(lx, self.GAP - 2, lx, self.GAP + self.CARD_H + 2)
            p.drawEllipse(lx - 3, self.GAP - 4, 6, 6)

        # Dragged ghost card on top
        if self._drag_idx >= 0:
            ghost_x = self._drag_abs_x - self._drag_offset
            self._draw_card(p, self._drag_idx, ghost_x, self.GAP - 4, dragging=True)

        p.end()

    def _draw_card(self, p: QPainter, idx: int, x: int, y: int, dragging: bool):
        img       = self.images[idx]
        is_second = img.get("is_second_image", False)
        is_sel    = (idx == self._selected_idx and not dragging)
        is_hover  = (idx == self._hover_idx and not dragging)

        if dragging:
            # Drop shadow
            sr = type(self.rect())(x + 3, y + 5, self.CARD_W, self.CARD_H)
            p.setBrush(QBrush(QColor(0, 0, 0, 60)))
            p.setPen(Qt.NoPen)
            p.drawRoundedRect(sr, self.RADIUS, self.RADIUS)
            bg, bdr = QColor(COLORS["bg_selected"]), QColor(COLORS["accent"])
        elif is_sel:
            bg, bdr = QColor(COLORS["bg_selected"]), QColor(COLORS["accent"])
        elif is_second:
            bg, bdr = QColor(COLORS["bg_card"]), QColor(COLORS["purple"])
        elif is_hover:
            bg, bdr = QColor(COLORS["bg_hover"]), QColor(COLORS["border_light"])
        else:
            bg, bdr = QColor(COLORS["bg_card"]), QColor(COLORS["border"])

        card = type(self.rect())(x, y, self.CARD_W, self.CARD_H)
        p.setBrush(QBrush(bg))
        p.setPen(bdr)
        p.drawRoundedRect(card, self.RADIUS, self.RADIUS)

        # Thumbnail (clipped to rounded top)
        thumb = self._thumbs.get(img["path"])
        thumb_rect = type(self.rect())(x + 4, y + 3, self.CARD_W - 8, self.THUMB_H)
        if thumb and not thumb.isNull():
            tw, th = thumb.width(), thumb.height()
            tx = x + (self.CARD_W - tw) // 2
            ty = y + 3 + (self.THUMB_H - th) // 2
            p.save()
            clip = QPainterPath()
            clip.addRoundedRect(thumb_rect.x(), thumb_rect.y(),
                                thumb_rect.width(), thumb_rect.height(), 5, 5)
            p.setClipPath(clip)
            p.drawPixmap(tx, ty, thumb)
            p.restore()
        else:
            # Loading spinner placeholder
            p.fillRect(thumb_rect, QColor(COLORS["bg_hover"]))
            p.setPen(QColor(COLORS["text_muted"]))
            p.setFont(QFont("Segoe UI", 8))
            label = "…" if img["path"] in self._loaders else "?"
            p.drawText(thumb_rect, Qt.AlignCenter, label)

        # Number badge (top-left)
        badge_bg = QColor(COLORS["accent"]) if not is_second else QColor(COLORS["purple"])
        p.setBrush(QBrush(badge_bg))
        p.setPen(Qt.NoPen)
        badge = type(self.rect())(x + 5, y + 6, 20, 14)
        p.drawRoundedRect(badge, 3, 3)
        p.setPen(QColor("#FFFFFF"))
        p.setFont(QFont("Segoe UI", 7, QFont.Bold))
        p.drawText(badge, Qt.AlignCenter, str(idx + 1))

        # Ken Burns badge (top-right)
        kb = img.get("ken_burns", "none")
        if kb != "none":
            kb_short = {"zoom_in": "Z+", "zoom_out": "Z−", "pan_left": "←",
                        "pan_right": "→", "pan_up": "↑", "pan_down": "↓"}.get(kb, "KB")
            p.setBrush(QBrush(QColor(COLORS["purple"])))
            p.setPen(Qt.NoPen)
            kb_badge = type(self.rect())(x + self.CARD_W - 24, y + 6, 18, 14)
            p.drawRoundedRect(kb_badge, 3, 3)
            p.setPen(QColor("#FFFFFF"))
            p.setFont(QFont("Segoe UI", 7, QFont.Bold))
            p.drawText(kb_badge, Qt.AlignCenter, kb_short)

        # Filename + duration
        name = os.path.basename(img["path"])
        if len(name) > 13:
            name = name[:11] + "…"
        text_y = y + self.THUMB_H + 7
        label_rect = type(self.rect())(x + 5, text_y, self.CARD_W - 10, 14)

        p.setPen(QColor(COLORS["text_primary"] if (is_sel or is_hover) else COLORS["text_secondary"]))
        p.setFont(QFont("Segoe UI", 7))
        p.drawText(label_rect, Qt.AlignLeft | Qt.AlignVCenter, name)

        p.setPen(QColor(COLORS["accent"] if is_sel else COLORS["text_muted"]))
        p.setFont(QFont("Segoe UI", 7, QFont.Bold))
        p.drawText(label_rect, Qt.AlignRight | Qt.AlignVCenter, f"{img.get('duration', 5)}s")

        # Bottom highlight bar (selected)
        if is_sel:
            p.setBrush(QBrush(QColor(COLORS["accent"])))
            p.setPen(Qt.NoPen)
            p.drawRoundedRect(x + 12, y + self.CARD_H - 4, self.CARD_W - 24, 3, 2, 2)

    # ── mouse ────────────────────────────────────────────────────────────────

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            idx = self._idx_at(event.x())
            if idx >= 0:
                self._drag_idx    = idx
                self._drag_abs_x  = event.x()
                self._drag_offset = event.x() - self._card_x(idx)
                self._drop_idx    = idx
                self.setCursor(Qt.ClosedHandCursor)
                self.card_clicked.emit(idx)
                self.update()

    def mouseMoveEvent(self, event):
        if self._drag_idx >= 0:
            self._drag_abs_x = event.x()
            self._drop_idx   = self._drop_pos_at(event.x())
            self.update()
        else:
            new_h = self._idx_at(event.x())
            if new_h != self._hover_idx:
                self._hover_idx = new_h
                self.setCursor(Qt.OpenHandCursor if new_h >= 0 else Qt.ArrowCursor)
                self.update()

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.LeftButton and self._drag_idx >= 0:
            from_idx = self._drag_idx
            to_idx   = self._drop_pos_at(event.x())
            to_idx   = max(0, min(to_idx, len(self.images)))
            if to_idx > from_idx:
                to_idx -= 1

            # Reset drag state BEFORE emitting signal so the canvas
            # repaints clean immediately (no freeze visual)
            self._drag_idx = -1
            self._drop_idx = -1
            self.setCursor(Qt.OpenHandCursor)
            self.update()
            QApplication.processEvents()   # flush repaint NOW, before table rebuild

            if from_idx != to_idx:
                self.order_changed.emit(from_idx, to_idx)

    def leaveEvent(self, _event):
        self._hover_idx = -1
        self.update()


# ── Inline filmstrip (scrollable, fixed-height strip) ─────────────────────────

class FilmstripTimeline(QScrollArea):
    order_changed = pyqtSignal(int, int)
    delete_at     = pyqtSignal(int)
    move_to       = pyqtSignal(int, int)
    card_clicked  = pyqtSignal(int)

    CARD_W = 110
    CARD_H = 90
    GAP    = 8

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.setFrameShape(QFrame.NoFrame)
        self.setFixedHeight(self.CARD_H + self.GAP * 2 + 6)
        self.setStyleSheet(
            f"QScrollArea {{ background: {COLORS['bg_deep']}; border-top: 1px solid {COLORS['border']}; }}"
            f"QScrollBar:horizontal {{ background: transparent; height: 6px; border-radius: 3px; }}"
            f"QScrollBar::handle:horizontal {{ background: {COLORS['border_light']}; border-radius: 3px; min-width: 20px; }}"
            f"QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{ width: 0; }}"
        )
        self._canvas = _FilmstripCanvas(self.CARD_W, self.CARD_H, 60, self.GAP, self)
        self._canvas.order_changed.connect(self.order_changed)
        self._canvas.delete_at.connect(self.delete_at)
        self._canvas.move_to.connect(self.move_to)
        self._canvas.card_clicked.connect(self.card_clicked)
        self.setWidget(self._canvas)
        self.setWidgetResizable(False)

    def set_images(self, images: list):
        self._canvas.set_images(images)

    def highlight_index(self, idx: int):
        self._canvas.set_selected(idx)
        x = self.GAP + idx * (self.CARD_W + self.GAP) - 20
        self.horizontalScrollBar().setValue(max(0, x))


# ── Full-view timeline dialog ─────────────────────────────────────────────────

class _WrappingFilmstripCanvas(QWidget):
    """
    Like _FilmstripCanvas but wraps cards into multiple rows based on the
    widget's current width.  Used by FilmstripFullDialog.
    All signals and interactions are identical to _FilmstripCanvas.
    """
    order_changed = pyqtSignal(int, int)
    delete_at     = pyqtSignal(int)
    move_to       = pyqtSignal(int, int)
    card_clicked  = pyqtSignal(int)

    CARD_W  = 160
    CARD_H  = 130
    THUMB_H = 95
    GAP     = 14
    RADIUS  = 8

    def __init__(self, card_w=160, card_h=130, thumb_h=95, gap=14, parent=None):
        super().__init__(parent)
        self.CARD_W  = card_w
        self.CARD_H  = card_h
        self.THUMB_H = thumb_h
        self.GAP     = gap

        self.images: list       = []
        self._thumbs: dict      = {}
        self._loaders: dict     = {}
        self._drag_idx: int     = -1
        self._drag_abs_x: int   = 0
        self._drag_abs_y: int   = 0
        self._drag_offset_x: int = 0
        self._drag_offset_y: int = 0
        self._hover_idx: int    = -1
        self._drop_idx: int     = -1   # flat insertion index
        self._selected_idx: int = -1

        self.setMouseTracking(True)
        self.setCursor(Qt.OpenHandCursor)
        self.setContextMenuPolicy(Qt.CustomContextMenu)
        self.customContextMenuRequested.connect(self._show_context_menu)

    # ── public ───────────────────────────────────────────────────────────────

    def set_images(self, images: list):
        self.images = images
        current_paths = {img["path"] for img in images}
        for path in list(self._loaders.keys()):
            if path not in current_paths:
                self._loaders.pop(path, None)
        self._thumbs = {k: v for k, v in self._thumbs.items() if k in current_paths}
        for img in images:
            self._request_thumb(img["path"])
        self._relayout()
        self.update()

    def set_selected(self, idx: int):
        self._selected_idx = idx
        self.update()

    # ── thumb loading ─────────────────────────────────────────────────────────

    def _request_thumb(self, path: str):
        if path in self._thumbs or path in self._loaders:
            return
        self._thumbs[path] = None
        loader = _ThumbLoader(path, self)
        loader.thumb_ready.connect(self._on_thumb_ready)
        self._loaders[path] = loader
        loader.start()

    def _on_thumb_ready(self, path: str, px: QPixmap):
        if not px.isNull():
            px = px.scaled(self.CARD_W - 12, self.THUMB_H - 4,
                           Qt.KeepAspectRatio, Qt.SmoothTransformation)
        self._thumbs[path] = px
        self._loaders.pop(path, None)
        self.update()

    # ── layout helpers ────────────────────────────────────────────────────────

    def _cols(self) -> int:
        """How many cards fit per row given current width."""
        avail = self.width() - self.GAP
        cols  = avail // (self.CARD_W + self.GAP)
        return max(1, cols)

    def _rows(self) -> int:
        import math
        n = len(self.images)
        if n == 0:
            return 1
        return math.ceil(n / self._cols())

    def _card_pos(self, idx: int):
        """Return (x, y) top-left for card at flat index idx."""
        cols = self._cols()
        row  = idx // cols
        col  = idx  % cols
        x    = self.GAP + col * (self.CARD_W + self.GAP)
        y    = self.GAP + row * (self.CARD_H + self.GAP)
        return x, y

    def _relayout(self):
        rows  = self._rows()
        total_h = self.GAP + rows * (self.CARD_H + self.GAP)
        self.setMinimumHeight(total_h)
        # Use setFixedHeight instead of resize() to avoid recursive resizeEvent
        # loops when the parent QScrollArea (with widgetResizable=True) resizes
        # us during window maximize / fullscreen transitions.
        if self.height() != total_h:
            self.setFixedHeight(total_h)

    def _idx_at(self, x: int, y: int) -> int:
        for i in range(len(self.images)):
            cx, cy = self._card_pos(i)
            if cx <= x <= cx + self.CARD_W and cy <= y <= cy + self.CARD_H:
                return i
        return -1

    def _drop_pos_at(self, x: int, y: int) -> int:
        """Find nearest insertion gap using distance to card centres."""
        cols = self._cols()
        n    = len(self.images)
        if n == 0:
            return 0
        best_dist = float("inf")
        best_pos  = 0
        # Check gaps: before card 0, between each pair, after last card
        for i in range(n + 1):
            if i < n:
                cx, cy = self._card_pos(i)
                gap_x  = cx          # left edge of card i
                gap_y  = cy + self.CARD_H // 2
            else:
                cx, cy = self._card_pos(n - 1)
                gap_x  = cx + self.CARD_W   # right edge of last card
                gap_y  = cy + self.CARD_H // 2
            d = ((x - gap_x) ** 2 + (y - gap_y) ** 2) ** 0.5
            if d < best_dist:
                best_dist = d
                best_pos  = i
        return best_pos

    # ── context menu ──────────────────────────────────────────────────────────

    def _show_context_menu(self, pos):
        idx = self._idx_at(pos.x(), pos.y())
        if idx < 0:
            return
        img  = self.images[idx]
        name = os.path.basename(img["path"])
        menu = QMenu(self)
        menu.setStyleSheet(
            f"QMenu {{ background: {COLORS['bg_card']}; border: 1px solid {COLORS['border_light']};"
            f"  border-radius: 8px; padding: 4px; color: {COLORS['text_primary']}; }}"
            f"QMenu::item {{ padding: 7px 24px 7px 14px; border-radius: 4px; margin: 1px 4px; }}"
            f"QMenu::item:selected {{ background: {COLORS['bg_hover']}; color: {COLORS['accent']}; }}"
            f"QMenu::separator {{ height: 1px; background: {COLORS['border']}; margin: 4px 8px; }}"
        )
        title_action = QAction(f"#{idx + 1}  {name[:28]}", menu)
        title_action.setEnabled(False)
        menu.addAction(title_action)
        menu.addSeparator()
        move_action   = QAction("✦  Set Position…", menu)
        delete_action = QAction("✕  Delete", menu)
        menu.addAction(move_action)
        menu.addSeparator()
        menu.addAction(delete_action)
        action = menu.exec_(self.mapToGlobal(pos))
        if action == move_action:
            n = len(self.images)
            new_pos, ok = QInputDialog.getInt(
                self, "Set Position",
                f"Move slide #{idx + 1} to position (1–{n}):",
                idx + 1, 1, n)
            if ok and new_pos - 1 != idx:
                self.move_to.emit(idx, new_pos)
        elif action == delete_action:
            self.delete_at.emit(idx)

    # ── painting ──────────────────────────────────────────────────────────────

    def paintEvent(self, _event):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        p.setRenderHint(QPainter.SmoothPixmapTransform)
        p.fillRect(self.rect(), QColor(COLORS["bg_deep"]))

        if not self.images:
            p.setPen(QColor(COLORS["text_muted"]))
            p.setFont(QFont("Segoe UI", 12))
            p.drawText(self.rect(), Qt.AlignCenter,
                       "Add images — they'll appear here for easy drag-to-reorder")
            p.end()
            return

        for i in range(len(self.images)):
            if i == self._drag_idx:
                continue
            cx, cy = self._card_pos(i)
            self._draw_card(p, i, cx, cy, dragging=False)
        # Drop indicator: blue line between cards
        if self._drag_idx >= 0 and 0 <= self._drop_idx <= len(self.images):
            real_di = self._drop_idx
            if real_di >= self._drag_idx:
                real_di = max(0, real_di - 1)
            cols = self._cols()
            if real_di < len(self.images):
                lx, ly = self._card_pos(real_di)
                # draw line to the LEFT of the target card
                line_x = lx - self.GAP // 2
                p.setPen(QColor(COLORS["accent"]))
                p.setBrush(QBrush(QColor(COLORS["accent"])))
                p.drawLine(line_x, ly, line_x, ly + self.CARD_H)
                p.drawEllipse(line_x - 3, ly - 3, 6, 6)
            else:
                # after last card
                lx, ly = self._card_pos(len(self.images) - 1)
                line_x = lx + self.CARD_W + self.GAP // 2
                p.setPen(QColor(COLORS["accent"]))
                p.setBrush(QBrush(QColor(COLORS["accent"])))
                p.drawLine(line_x, ly, line_x, ly + self.CARD_H)
                p.drawEllipse(line_x - 3, ly - 3, 6, 6)

        # Ghost (dragged card)
        if self._drag_idx >= 0:
            gx = self._drag_abs_x - self._drag_offset_x
            gy = self._drag_abs_y - self._drag_offset_y
            self._draw_card(p, self._drag_idx, gx, gy, dragging=True)

        p.end()

    def _draw_card(self, p: QPainter, idx: int, x: int, y: int, dragging: bool):
        img       = self.images[idx]
        is_second = img.get("is_second_image", False)
        is_sel    = (idx == self._selected_idx and not dragging)
        is_hover  = (idx == self._hover_idx and not dragging)

        if dragging:
            sr = type(self.rect())(x + 3, y + 5, self.CARD_W, self.CARD_H)
            p.setBrush(QBrush(QColor(0, 0, 0, 60)))
            p.setPen(Qt.NoPen)
            p.drawRoundedRect(sr, self.RADIUS, self.RADIUS)
            bg, bdr = QColor(COLORS["bg_selected"]), QColor(COLORS["accent"])
        elif is_sel:
            bg, bdr = QColor(COLORS["bg_selected"]), QColor(COLORS["accent"])
        elif is_second:
            bg, bdr = QColor(COLORS["bg_card"]), QColor(COLORS["purple"])
        elif is_hover:
            bg, bdr = QColor(COLORS["bg_hover"]), QColor(COLORS["border_light"])
        else:
            bg, bdr = QColor(COLORS["bg_card"]), QColor(COLORS["border"])

        card = type(self.rect())(x, y, self.CARD_W, self.CARD_H)
        p.setBrush(QBrush(bg))
        p.setPen(bdr)
        p.drawRoundedRect(card, self.RADIUS, self.RADIUS)

        # Thumbnail
        thumb = self._thumbs.get(img["path"])
        thumb_rect = type(self.rect())(x + 4, y + 3, self.CARD_W - 8, self.THUMB_H)
        if thumb and not thumb.isNull():
            tw, th = thumb.width(), thumb.height()
            tx = x + (self.CARD_W - tw) // 2
            ty = y + 3 + (self.THUMB_H - th) // 2
            p.save()
            clip = QPainterPath()
            clip.addRoundedRect(thumb_rect.x(), thumb_rect.y(),
                                thumb_rect.width(), thumb_rect.height(), 5, 5)
            p.setClipPath(clip)
            p.drawPixmap(tx, ty, thumb)
            p.restore()
        else:
            p.fillRect(thumb_rect, QColor(COLORS["bg_hover"]))
            p.setPen(QColor(COLORS["text_muted"]))
            p.setFont(QFont("Segoe UI", 9))
            label = "…" if img["path"] in self._loaders else "?"
            p.drawText(thumb_rect, Qt.AlignCenter, label)

        # Number badge
        badge_bg = QColor(COLORS["accent"]) if not is_second else QColor(COLORS["purple"])
        p.setBrush(QBrush(badge_bg))
        p.setPen(Qt.NoPen)
        badge = type(self.rect())(x + 5, y + 6, 22, 15)
        p.drawRoundedRect(badge, 3, 3)
        p.setPen(QColor("#FFFFFF"))
        p.setFont(QFont("Segoe UI", 8, QFont.Bold))
        p.drawText(badge, Qt.AlignCenter, str(idx + 1))

        # Ken Burns badge
        kb = img.get("ken_burns", "none")
        if kb != "none":
            kb_short = {"zoom_in": "Z+", "zoom_out": "Z−", "pan_left": "←",
                        "pan_right": "→", "pan_up": "↑", "pan_down": "↓"}.get(kb, "KB")
            p.setBrush(QBrush(QColor(COLORS["purple"])))
            p.setPen(Qt.NoPen)
            kb_b = type(self.rect())(x + self.CARD_W - 26, y + 6, 20, 15)
            p.drawRoundedRect(kb_b, 3, 3)
            p.setPen(QColor("#FFFFFF"))
            p.setFont(QFont("Segoe UI", 8, QFont.Bold))
            p.drawText(kb_b, Qt.AlignCenter, kb_short)

        # Filename + duration
        name = os.path.basename(img["path"])
        if len(name) > 18:
            name = name[:16] + "…"
        text_y   = y + self.THUMB_H + 8
        lbl_rect = type(self.rect())(x + 5, text_y, self.CARD_W - 10, 16)
        p.setPen(QColor(COLORS["text_primary"] if (is_sel or is_hover) else COLORS["text_secondary"]))
        p.setFont(QFont("Segoe UI", 8))
        p.drawText(lbl_rect, Qt.AlignLeft | Qt.AlignVCenter, name)
        p.setPen(QColor(COLORS["accent"] if is_sel else COLORS["text_muted"]))
        p.setFont(QFont("Segoe UI", 8, QFont.Bold))
        p.drawText(lbl_rect, Qt.AlignRight | Qt.AlignVCenter, f"{img.get('duration', 5)}s")

        if is_sel:
            p.setBrush(QBrush(QColor(COLORS["accent"])))
            p.setPen(Qt.NoPen)
            p.drawRoundedRect(x + 14, y + self.CARD_H - 5, self.CARD_W - 28, 3, 2, 2)

    # ── resize ────────────────────────────────────────────────────────────────

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._relayout()
        self.update()

    # ── mouse ─────────────────────────────────────────────────────────────────

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            idx = self._idx_at(event.x(), event.y())
            if idx >= 0:
                cx, cy = self._card_pos(idx)
                self._drag_idx      = idx
                self._drag_abs_x    = event.x()
                self._drag_abs_y    = event.y()
                self._drag_offset_x = event.x() - cx
                self._drag_offset_y = event.y() - cy
                self._drop_idx      = idx
                self.setCursor(Qt.ClosedHandCursor)
                self.card_clicked.emit(idx)
                self.update()

    def mouseMoveEvent(self, event):
        if self._drag_idx >= 0:
            self._drag_abs_x = event.x()
            self._drag_abs_y = event.y()
            self._drop_idx   = self._drop_pos_at(event.x(), event.y())
            self.update()
        else:
            new_h = self._idx_at(event.x(), event.y())
            if new_h != self._hover_idx:
                self._hover_idx = new_h
                self.setCursor(Qt.OpenHandCursor if new_h >= 0 else Qt.ArrowCursor)
                self.update()

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.LeftButton and self._drag_idx >= 0:
            from_idx = self._drag_idx
            to_idx   = self._drop_pos_at(event.x(), event.y())
            to_idx   = max(0, min(to_idx, len(self.images)))
            if to_idx > from_idx:
                to_idx -= 1

            self._drag_idx = -1
            self._drop_idx = -1
            self.setCursor(Qt.OpenHandCursor)
            self.update()
            QApplication.processEvents()

            if from_idx != to_idx:
                self.order_changed.emit(from_idx, to_idx)

    def leaveEvent(self, _event):
        self._hover_idx = -1
        self.update()


class FilmstripFullDialog(QDialog):
    """
    Full-view timeline dialog.
    Cards wrap into multiple rows based on the window width —
    resize the window to see more or fewer cards per row.
    """
    order_changed = pyqtSignal(int, int)
    delete_at     = pyqtSignal(int)
    move_to       = pyqtSignal(int, int)

    def __init__(self, images: list, selected_idx: int = 0, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Timeline — Full View")
        self.setMinimumSize(600, 300)
        self.resize(1200, 500)
        self.setWindowFlags(self.windowFlags() | Qt.WindowMaximizeButtonHint)
        self.setStyleSheet(
            f"QDialog {{ background: {COLORS['bg_deep']}; }}"
        )

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # Header bar
        header = QWidget()
        header.setFixedHeight(44)
        header.setStyleSheet(
            f"background: {COLORS['toolbar_bg']}; border-bottom: 1px solid {COLORS['border']};"
        )
        hlay = QHBoxLayout(header)
        hlay.setContentsMargins(16, 0, 16, 0)
        title = QLabel("  Timeline")
        title.setStyleSheet(
            f"color: {COLORS['text_primary']}; font-size: 13px; font-weight: 600;"
        )
        hint = QLabel("Drag cards to reorder  •  Right-click for options  •  Resize window to change cards per row")
        hint.setStyleSheet(f"color: {COLORS['text_muted']}; font-size: 11px;")
        close_btn = _styled_btn("✕  Close", "")
        close_btn.setFixedHeight(28)
        close_btn.clicked.connect(self.close)
        hlay.addWidget(title)
        hlay.addWidget(hint)
        hlay.addStretch()
        hlay.addWidget(close_btn)
        root.addWidget(header)

        # Scroll area — vertical scroll, no horizontal
        scroll = QScrollArea()
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setWidgetResizable(True)   # canvas width tracks scroll area width
        scroll.setStyleSheet(
            f"QScrollArea {{ background: {COLORS['bg_deep']}; }}"
            f"QScrollBar:vertical {{ background: transparent; width: 8px; border-radius: 4px; }}"
            f"QScrollBar::handle:vertical {{ background: {COLORS['border_light']}; border-radius: 4px; min-height: 20px; }}"
            f"QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}"
        )

        self._canvas = _WrappingFilmstripCanvas(160, 130, 95, 14, self)
        self._canvas.set_images(images)
        self._canvas.set_selected(selected_idx)
        self._canvas.order_changed.connect(self._fwd_order)
        self._canvas.delete_at.connect(self._fwd_delete)
        self._canvas.move_to.connect(self._fwd_move)
        scroll.setWidget(self._canvas)
        root.addWidget(scroll, 1)

        # Scroll to show selected card
        QTimer.singleShot(80, lambda: self._scroll_to(scroll, selected_idx))

    def _scroll_to(self, scroll: QScrollArea, idx: int):
        if 0 <= idx < len(self._canvas.images):
            _, cy = self._canvas._card_pos(idx)
            scroll.verticalScrollBar().setValue(max(0, cy - 20))

    def refresh(self, images: list, selected_idx: int = -1):
        self._canvas.set_images(images)
        if selected_idx >= 0:
            self._canvas.set_selected(selected_idx)

    def _fwd_order(self, a, b):  self.order_changed.emit(a, b)
    def _fwd_delete(self, i):    self.delete_at.emit(i)
    def _fwd_move(self, i, p):   self.move_to.emit(i, p)


# ── Main Window ───────────────────────────────────────────────────────────────

class SlideshowCreator(QMainWindow):

    @pyqtSlot(str, str)
    @pyqtSlot(str, str, str)
    def _show_update_dialog(self, new_version: str, url: str, changelog: str = ""):
        dlg = QDialog(self)
        dlg.setWindowTitle("Update Available")
        dlg.setMinimumWidth(480)
        layout = QVBoxLayout(dlg)
        layout.setSpacing(12)
        layout.setContentsMargins(20, 20, 20, 20)

        header = QLabel(
            f"<b style='font-size:15px; color:{COLORS['accent']}'>New Version Available: v{new_version}</b><br>"
            f"<span style='color:{COLORS['text_secondary']}'>You are running v{APP_VERSION}</span>"
        )
        header.setTextFormat(Qt.RichText)
        layout.addWidget(header)

        if changelog:
            notes = QTextEdit()
            notes.setReadOnly(True)
            notes.setPlainText(changelog)
            notes.setFixedHeight(180)
            layout.addWidget(notes)

        link = QLabel(f'<a href="{url}" style="color:{COLORS["accent"]}">Download latest version →</a>')
        link.setTextFormat(Qt.RichText)
        link.setTextInteractionFlags(Qt.TextBrowserInteraction)
        link.setOpenExternalLinks(True)
        layout.addWidget(link)

        btn_row = QHBoxLayout()
        close_btn = _styled_btn("Later", "")
        close_btn.clicked.connect(dlg.accept)
        btn_row.addStretch()
        btn_row.addWidget(close_btn)
        layout.addLayout(btn_row)
        dlg.exec_()

    def __init__(self):
        super().__init__()
        self.language     = "en"
        self.translations = {}
        self.load_translations()

        self.setWindowTitle(self.tr("window_title"))
        self.setGeometry(100, 100, 1400, 820)
        self.setMinimumSize(900, 600)

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
        self._pending_temp_dirs: list = []
        self._full_timeline_dlg = None

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
        self._load_font_settings()
        self.create_ui()

    # ── UI ────────────────────────────────────────────────────────────────────

    def create_ui(self):
        # Central widget + main layout
        central = QWidget()
        self.setCentralWidget(central)
        root_layout = QVBoxLayout(central)
        root_layout.setContentsMargins(0, 0, 0, 0)
        root_layout.setSpacing(0)

        # Toolbar
        self._build_quick_toolbar(root_layout)

        # Main splitter
        splitter = QSplitter(Qt.Horizontal)
        splitter.setHandleWidth(1)
        splitter.setChildrenCollapsible(False)
        root_layout.addWidget(splitter, 1)

        # ── Left: Slides table ─────────────────────────────────────────────
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(12, 12, 6, 8)
        left_layout.setSpacing(8)

        header_row = QHBoxLayout()
        slides_title = _make_section_label("Slides")
        slides_title.setStyleSheet(
            f"color: {COLORS['text_muted']}; font-size: 11px; font-weight: 700; letter-spacing: 1px;"
        )
        header_row.addWidget(slides_title)
        header_row.addStretch()

        self.slide_count_label = QLabel("0")
        self.slide_count_label.setStyleSheet(
            f"background: {COLORS['bg_card']}; color: {COLORS['accent']}; "
            f"border-radius: 10px; padding: 1px 8px; font-size: 11px; font-weight: 600;"
        )
        header_row.addWidget(self.slide_count_label)
        left_layout.addLayout(header_row)

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
        self.image_table.setFont(QFont("Segoe UI", 11))
        self.image_table.setShowGrid(False)
        self.image_table.setAlternatingRowColors(False)
        self.image_table.verticalHeader().setVisible(False)
        self.image_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.image_table.setEditTriggers(QTableWidget.DoubleClicked | QTableWidget.SelectedClicked)

        for col in range(10):
            self.image_table.horizontalHeader().setSectionResizeMode(col, QHeaderView.ResizeToContents)
        self.image_table.horizontalHeader().setStretchLastSection(False)
        self.image_table.setColumnWidth(5, 160)  # text col wider

        self.image_table.itemChanged.connect(self.on_edit_on_table)
        left_layout.addWidget(self.image_table)

        # ── Filmstrip timeline ─────────────────────────────────────────────
        strip_header = QHBoxLayout()
        strip_header.setContentsMargins(0, 4, 0, 0)
        strip_lbl = _make_section_label("Timeline — drag to reorder  •  right-click for options")
        strip_header.addWidget(strip_lbl)
        strip_header.addStretch()

        expand_btn = QPushButton("⤢  Full View")
        expand_btn.setFixedHeight(22)
        expand_btn.setStyleSheet(
            f"QPushButton {{ background: {COLORS['bg_card']}; color: {COLORS['text_secondary']};"
            f"  border: 1px solid {COLORS['border']}; border-radius: 4px; padding: 0 8px; font-size: 11px; }}"
            f"QPushButton:hover {{ color: {COLORS['accent']}; border-color: {COLORS['accent']}; }}"
        )
        expand_btn.clicked.connect(self._open_full_timeline)
        strip_header.addWidget(expand_btn)
        left_layout.addLayout(strip_header)

        self.filmstrip = FilmstripTimeline()
        self.filmstrip.order_changed.connect(self._on_filmstrip_reorder)
        self.filmstrip.delete_at.connect(self._on_filmstrip_delete)
        self.filmstrip.move_to.connect(self._on_filmstrip_move_to)
        self.filmstrip.card_clicked.connect(lambda idx: self.image_table.setCurrentCell(idx, 1))
        left_layout.addWidget(self.filmstrip)

        self._full_timeline_dlg = None   # lazily created

        # Bottom action strip
        action_strip = QHBoxLayout()
        action_strip.setSpacing(6)
        add_img_btn = _styled_btn("＋  Add Images", "primary")
        add_img_btn.clicked.connect(self.add_images)
        add_img_btn.setFixedHeight(30)

        easy_text_btn = _styled_btn("✏  Easy Text", "")
        easy_text_btn.clicked.connect(self.open_easy_text_writing)
        easy_text_btn.setFixedHeight(30)

        action_strip.addWidget(add_img_btn)
        action_strip.addWidget(easy_text_btn)
        action_strip.addStretch()
        left_layout.addLayout(action_strip)

        # ── Right: Sidebar ─────────────────────────────────────────────────
        right_widget = QWidget()
        right_widget.setFixedWidth(300)
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(6, 12, 12, 8)
        right_layout.setSpacing(12)

        # Preview
        self.preview_panel = PreviewPanel()
        right_layout.addWidget(self.preview_panel)

        # Audio section
        audio_header = QHBoxLayout()
        audio_title = _make_section_label("Audio")
        audio_header.addWidget(audio_title)
        audio_header.addStretch()

        self.audio_count_label = QLabel("0")
        self.audio_count_label.setStyleSheet(
            f"background: {COLORS['bg_card']}; color: {COLORS['success']}; "
            f"border-radius: 10px; padding: 1px 8px; font-size: 11px; font-weight: 600;"
        )
        audio_header.addWidget(self.audio_count_label)
        right_layout.addLayout(audio_header)

        self.audio_table = QTableWidget()
        self.audio_table.setColumnCount(2)
        self.audio_table.setHorizontalHeaderLabels([
            self.tr("table_header_actions"),
            self.tr("table_header_audio_file"),
        ])
        self.audio_table.setFont(QFont("Segoe UI", 11))
        self.audio_table.setShowGrid(False)
        self.audio_table.verticalHeader().setVisible(False)
        self.audio_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.audio_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.audio_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.audio_table.setMaximumHeight(140)
        right_layout.addWidget(self.audio_table)

        audio_btn_row = QHBoxLayout()
        audio_btn_row.setSpacing(6)
        add_audio_btn = _styled_btn("＋  Audio", "")
        add_audio_btn.clicked.connect(self.add_audio)
        add_audio_btn.setFixedHeight(28)
        library_btn = _styled_btn("♪  Library", "")
        library_btn.clicked.connect(self.open_audio_library)
        library_btn.setFixedHeight(28)
        audio_btn_row.addWidget(add_audio_btn)
        audio_btn_row.addWidget(library_btn)
        right_layout.addLayout(audio_btn_row)

        # Export buttons
        right_layout.addWidget(_make_divider())
        export_title = _make_section_label("Export")
        right_layout.addWidget(export_title)

        preview_btn = _styled_btn("▶  Preview Slideshow", "primary")
        preview_btn.clicked.connect(self.open_preview_dialog)
        preview_btn.setFixedHeight(36)
        preview_btn.setStyleSheet(
            f"QPushButton {{ background: {COLORS['purple']}; color: #fff; "
            f"border: none; border-radius: 6px; font-weight: 700; font-size: 13px; }}"
            f"QPushButton:hover {{ background: #8a5ef5; }}"
        )
        right_layout.addWidget(preview_btn)

        export_btn = _styled_btn("▶  Export Slideshow", "primary")
        export_btn.clicked.connect(self.export_slideshow)
        export_btn.setFixedHeight(36)
        right_layout.addWidget(export_btn)

        premiere_btn = _styled_btn("⬡  Export to Premiere", "")
        premiere_btn.clicked.connect(self.export_premiere_slideshow)
        premiere_btn.setFixedHeight(32)
        right_layout.addWidget(premiere_btn)

        # Progress bars
        self.progress_section = ProgressSection()
        right_layout.addWidget(self.progress_section)
        self.progress_bar              = self.progress_section.export_bar
        self.image_progress_bar        = self.progress_section.image_bar
        self.image_premiere_progress_bar = self.progress_section.premiere_bar

        right_layout.addStretch()

        splitter.addWidget(left_widget)
        splitter.addWidget(right_widget)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 0)

        # Status bar
        self.stats_bar = StatsBar()
        self.setStatusBar(QStatusBar())
        self.statusBar().addPermanentWidget(self.stats_bar, 1)

        self.taskbar_progress = self._make_taskbar_progress()

    def _build_quick_toolbar(self, parent_layout: QVBoxLayout):
        """Build the top quick-action toolbar."""
        tb_widget = QWidget()
        tb_widget.setObjectName("quickToolbar")
        tb_widget.setStyleSheet(f"#quickToolbar {{ background: {COLORS['toolbar_bg']}; border-bottom: 1px solid {COLORS['border']}; }}")
        tb_widget.setFixedHeight(44)

        tb_layout = QHBoxLayout(tb_widget)
        tb_layout.setContentsMargins(12, 4, 12, 4)
        tb_layout.setSpacing(4)

        def _tb_btn(text, tip=""):
            b = QToolButton()
            b.setText(text)
            if tip:
                b.setToolTip(tip)
            return b

        btn_open   = _tb_btn("  Open",    "Load project (Ctrl+L)")
        btn_save   = _tb_btn("  Save",    "Save project (Ctrl+S)")
        btn_saveas = _tb_btn("  Save As", "Save project as...")

        btn_open.clicked.connect(self.load_project)
        btn_save.clicked.connect(self.save_project)
        btn_saveas.clicked.connect(self.save_project_as)

        sep1 = QFrame(); sep1.setFrameShape(QFrame.VLine)
        sep1.setStyleSheet(f"color: {COLORS['border']}; max-width: 1px;")

        btn_pptx   = _tb_btn("⬢  Import PPTX")
        btn_pptx.clicked.connect(self.import_pptx)

        btn_sort_new = _tb_btn("⬇ Newest First")
        btn_sort_new.clicked.connect(lambda: self.auto_sort_images_by_date(True))
        btn_sort_old = _tb_btn("⬆ Oldest First")
        btn_sort_old.clicked.connect(lambda: self.auto_sort_images_by_date(False))
        btn_random   = _tb_btn("⤡ Shuffle")
        btn_random.clicked.connect(self.set_random_images_order)

        btn_batch_dur = _tb_btn("⏱ Fit to Audio", "Auto-set slide durations to match audio length")
        btn_batch_dur.clicked.connect(self.auto_calc_image_duration)

        sep2 = QFrame(); sep2.setFrameShape(QFrame.VLine)
        sep2.setStyleSheet(f"color: {COLORS['border']}; max-width: 1px;")

        btn_preview = _tb_btn("▶  Preview", "Preview the full slideshow with audio (no export needed)")
        btn_preview.setProperty("accent", "true")
        btn_preview.clicked.connect(self.open_preview_dialog)

        btn_info    = _tb_btn("ℹ Info")
        btn_info.clicked.connect(self.show_info)
        btn_clear   = _tb_btn("✕  Clear")
        btn_clear.clicked.connect(self.clear_project)

        for w in [btn_open, btn_save, btn_saveas, sep1,
                  btn_pptx, btn_sort_new, btn_sort_old, btn_random, btn_batch_dur,
                  sep2, btn_preview, btn_info, btn_clear]:
            tb_layout.addWidget(w)

        tb_layout.addStretch()

        # GPU status badge (shows which encoder will be used)

        # Language switch
        lang_lbl = QLabel("Language:")
        lang_lbl.setStyleSheet(f"color: {COLORS['text_muted']}; font-size: 11px;")
        lang_en = _tb_btn("EN")
        lang_he = _tb_btn("עב")
        lang_en.setToolTip("Switch to English")
        lang_he.setToolTip("Switch to Hebrew (עברית)")
        lang_en.clicked.connect(lambda: self.set_language("en"))
        lang_he.clicked.connect(lambda: self.set_language("he"))

        tb_layout.addWidget(lang_lbl)
        tb_layout.addWidget(lang_en)
        tb_layout.addWidget(lang_he)

        parent_layout.addWidget(tb_widget)

    def _make_taskbar_progress(self):
        try:
            from PyQt5.QtWinExtras import QWinTaskbarButton
            btn = QWinTaskbarButton(self)
            btn.setWindow(self.windowHandle())
            progress = btn.progress()
            progress.setRange(0, 100)
            return progress
        except Exception:
            return _TaskbarProgressStub()

    # ── Image table helpers ───────────────────────────────────────────────────

    def on_edit_on_table(self, item):
        col = item.column()
        row = item.row()
        if col == 2:
            try:
                val = round(float(item.text()), 1)
                if val < 2 or val > 600:
                    raise ValueError(self.tr("duration_out_of_range_error"))
                self.images[row]["duration"] = val
                if self.images[row]["transition_duration"] > val - 1:
                    self.images[row]["transition_duration"] = val - 1
                    self.update_image_table()
            except ValueError:
                cur = self.images[row]["duration"]
                item.setText(str(int(cur)) if float(cur) == int(float(cur)) else f"{float(cur):.1f}")
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
        self._refresh_stats()

    def _fast_reorder(self, from_idx: int, to_idx: int):
        """
        Move images[from_idx] to to_idx and repopulate ONLY the shifted rows.
        This avoids the freeze caused by a full update_image_table() rebuild.
        """
        img = self.images.pop(from_idx)
        self.images.insert(to_idx, img)

        # After reordering, two second-images may have ended up adjacent.
        # Clear the is_second_image flag on the moved image if that happened.
        moved = self.images[to_idx]
        if moved.get("is_second_image"):
            prev_is_second = to_idx > 0 and self.images[to_idx - 1].get("is_second_image", False)
            next_is_second = to_idx + 1 < len(self.images) and self.images[to_idx + 1].get("is_second_image", False)
            if prev_is_second or next_is_second:
                moved["is_second_image"] = False

        lo, hi = min(from_idx, to_idx), max(from_idx, to_idx)
        self.image_table.blockSignals(True)
        self.image_table.setUpdatesEnabled(False)
        for r in range(lo, hi + 1):
            self._populate_row(r, self.images[r])
        self.image_table.setUpdatesEnabled(True)
        self.image_table.blockSignals(False)

        self.image_table.setCurrentCell(to_idx, 1)
        self.filmstrip.set_images(self.images)
        if self._full_timeline_dlg and self._full_timeline_dlg.isVisible():
            self._full_timeline_dlg.refresh(self.images, to_idx)
        self._refresh_stats()
        self.update_preview_with_row(to_idx)

    def _on_filmstrip_reorder(self, from_idx: int, to_idx: int):
        """Called when user drags a card in the filmstrip to a new position."""
        if 0 <= from_idx < len(self.images) and 0 <= to_idx < len(self.images):
            self._fast_reorder(from_idx, to_idx)

    def _on_filmstrip_delete(self, idx: int):
        """Right-click → Delete from filmstrip."""
        if 0 <= idx < len(self.images):
            del self.images[idx]
            self.image_table.removeRow(idx)
            if not self.images:
                self.preview_panel.clear()
            else:
                new_row = max(0, idx - 1)
                self.image_table.setCurrentCell(new_row, 1)
                self.update_preview_with_row(new_row)
            self.filmstrip.set_images(self.images)
            self._refresh_stats()
            if self._full_timeline_dlg and self._full_timeline_dlg.isVisible():
                self._full_timeline_dlg.refresh(self.images, max(0, idx - 1))

    def _on_filmstrip_move_to(self, cur_idx: int, target_pos: int):
        """Right-click → Set Position from filmstrip (target_pos is 1-based)."""
        new_idx = target_pos - 1
        new_idx = max(0, min(new_idx, len(self.images) - 1))
        if new_idx == cur_idx:
            return
        self._fast_reorder(cur_idx, new_idx)

    def _open_full_timeline(self):
        """Open (or bring to front) the full-view timeline dialog."""
        sel = self.image_table.currentRow()
        if self._full_timeline_dlg is None or not self._full_timeline_dlg.isVisible():
            self._full_timeline_dlg = FilmstripFullDialog(
                self.images, selected_idx=max(0, sel), parent=self
            )
            self._full_timeline_dlg.order_changed.connect(self._on_filmstrip_reorder)
            self._full_timeline_dlg.delete_at.connect(self._on_filmstrip_delete)
            self._full_timeline_dlg.move_to.connect(self._on_filmstrip_move_to)
            self._full_timeline_dlg.show()
        else:
            self._full_timeline_dlg.raise_()
            self._full_timeline_dlg.activateWindow()
            self._full_timeline_dlg.refresh(self.images, max(0, sel))

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
            new = [{
                "path": f, "duration": 5, "transition": "fade",
                "transition_duration": self.default_transition_duration,
                "text": "", "rotation": 0, "is_second_image": False,
                "date": datetime.fromtimestamp(os.path.getmtime(f)).strftime("%Y-%m-%d %H:%M:%S"),
                "ken_burns": "none"
            } for f in files]
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
        if state == Qt.Checked:
            # Guard: never allow two consecutive second-images.
            # Check the image immediately before this one.
            prev_is_second = self.images[row - 1].get("is_second_image", False)
            # Check the image immediately after this one (only if it exists).
            next_is_second = (
                self.images[row + 1].get("is_second_image", False)
                if row + 1 < len(self.images)
                else False
            )
            if prev_is_second or next_is_second:
                QMessageBox.warning(self, self.tr("error"), self.tr("subsequent_line"), QMessageBox.Ok)
                self.update_image_row(row)
                return
        is_checked = (state == Qt.Checked)
        self.images[row]["is_second_image"] = is_checked
        item = self.image_table.item(row, 1)
        if item:
            item.setData(Qt.UserRole, is_checked)
        self.update_image_row(row)

    def update_image_table(self, _update_filmstrip: bool = True):
        self.image_table.blockSignals(True)
        self.image_table.setUpdatesEnabled(False)
        self.image_table.setSortingEnabled(False)
        self.image_table.setRowCount(len(self.images))
        for row, img in enumerate(self.images):
            self._populate_row(row, img)
        self.image_table.setSortingEnabled(False)
        self.image_table.setUpdatesEnabled(True)
        self.image_table.blockSignals(False)
        if _update_filmstrip:
            self.filmstrip.set_images(self.images)
        self._refresh_stats()

    def _populate_row(self, row: int, img: dict):
        self.image_table.setRowHeight(row, 34)

        filename_item = QTableWidgetItem(os.path.basename(img["path"]))
        filename_item.setData(Qt.UserRole, img.get("is_second_image", False))
        filename_item.setFlags(filename_item.flags() & ~Qt.ItemIsEditable)

        dur_val = img.get("duration", 5)
        dur_display = str(int(dur_val)) if float(dur_val) == int(float(dur_val)) else f"{float(dur_val):.1f}"
        duration_item = QTableWidgetItem(dur_display)

        transition_cb = QComboBox()
        transition_cb.addItems(self.transitions_types)
        transition_cb.setCurrentText(img.get("transition", "fade"))
        transition_cb.currentTextChanged.connect(lambda text, r=row: self.update_transition(r, text))

        tl_item = QTableWidgetItem(str(img.get("transition_duration", self.default_transition_duration)))
        tl_item.setFlags(tl_item.flags() & ~Qt.ItemIsEditable)

        text_item = QTableWidgetItem(str(img.get("text", "")))
        text_item.setFlags(text_item.flags() | Qt.ItemIsEditable)

        rotation_item = QTableWidgetItem(str(img.get("rotation", 0)))

        second_cb = QCheckBox()
        second_cb.setChecked(img.get("is_second_image", False))
        second_cb.setStyleSheet("QCheckBox { margin-left: 10px; }")
        second_cb.stateChanged.connect(lambda state, r=row: self.set_second_image(r, state))

        date_item = QTableWidgetItem(str(img.get("date", "")))
        date_item.setFlags(date_item.flags() & ~Qt.ItemIsEditable)
        date_item.setForeground(QColor(COLORS["text_muted"]))

        self.image_table.setItem(row, 1, filename_item)
        self.image_table.setItem(row, 2, duration_item)
        self.image_table.setCellWidget(row, 3, transition_cb)
        self.image_table.setItem(row, 4, tl_item)
        self.image_table.setItem(row, 5, text_item)
        self.image_table.setItem(row, 6, rotation_item)
        self.image_table.setCellWidget(row, 7, second_cb)
        self.image_table.setItem(row, 8, date_item)

        kb_cb = QComboBox()
        kb_cb.addItems(["none", "zoom_in", "zoom_out", "pan_left", "pan_right", "pan_up", "pan_down"])
        kb_cb.setCurrentText(img.get("ken_burns", "none"))
        kb_cb.currentTextChanged.connect(lambda text, r=row: self._update_ken_burns(r, text))
        self.image_table.setCellWidget(row, 9, kb_cb)

        # Action buttons – compact icon style
        up_btn   = QPushButton("↑")
        dn_btn   = QPushButton("↓")
        del_btn  = QPushButton("✕")
        up_btn.setProperty("action", "icon")
        dn_btn.setProperty("action", "icon")
        del_btn.setProperty("action", "icon")

        # Crop button — stands out visually; glows orange when a crop is active
        crop_btn = _make_crop_btn(bool(img.get("crop")))

        del_btn.setStyleSheet(f"QPushButton {{ color: {COLORS['danger']}; background: transparent; border: none; padding: 3px 6px; border-radius: 4px; }}"
                              f"QPushButton:hover {{ background: rgba(247,90,90,0.15); }}")
        up_btn.clicked.connect(self.move_image_up)
        dn_btn.clicked.connect(self.move_image_down)
        del_btn.clicked.connect(self.delete_image)
        crop_btn.clicked.connect(lambda _, r=row: self.open_crop_dialog(r))

        btn_widget = QWidget()
        btn_layout = QHBoxLayout(btn_widget)
        btn_layout.addWidget(up_btn)
        btn_layout.addWidget(dn_btn)
        btn_layout.addWidget(crop_btn)
        btn_layout.addWidget(del_btn)
        btn_layout.setContentsMargins(4, 0, 4, 0)
        btn_layout.setSpacing(2)
        self.image_table.setCellWidget(row, 0, btn_widget)

    def move_image_up(self):
        row = self.image_table.currentRow()
        if row > 0:
            self.images[row], self.images[row - 1] = self.images[row - 1], self.images[row]
            self.update_image_row(row)
            self.update_image_row(row - 1)
            self.image_table.setCurrentCell(row - 1, 1)
            self.filmstrip.set_images(self.images)
            self.filmstrip.highlight_index(row - 1)
            self.update_preview_with_row(row - 1)
            if self._full_timeline_dlg and self._full_timeline_dlg.isVisible():
                self._full_timeline_dlg.refresh(self.images, row - 1)

    def move_image_down(self):
        row = self.image_table.currentRow()
        if row < len(self.images) - 1:
            self.images[row], self.images[row + 1] = self.images[row + 1], self.images[row]
            self.update_image_row(row)
            self.update_image_row(row + 1)
            self.image_table.setCurrentCell(row + 1, 1)
            self.filmstrip.set_images(self.images)
            self.filmstrip.highlight_index(row + 1)
            self.update_preview_with_row(row + 1)
            if self._full_timeline_dlg and self._full_timeline_dlg.isVisible():
                self._full_timeline_dlg.refresh(self.images, row + 1)

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
                self.preview_panel.clear()
            else:
                new_row = max(0, row - 1)
                self.image_table.setCurrentCell(new_row, 1)
                self.update_preview_with_row(new_row)
        self.image_table.setSortingEnabled(False)
        self.filmstrip.set_images(self.images)
        self._refresh_stats()

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
            if new_pos == cur_row:
                return
            img = self.images.pop(cur_row)
            self.images.insert(new_pos, img)
            lo, hi = min(cur_row, new_pos), max(cur_row, new_pos)
            self.image_table.blockSignals(True)
            self.image_table.setUpdatesEnabled(False)
            for r in range(lo, hi + 1):
                self._populate_row(r, self.images[r])
            self.image_table.setUpdatesEnabled(True)
            self.image_table.blockSignals(False)
            self.image_table.setCurrentCell(new_pos, 1)

    def update_image_progress(self, value: int):
        self.image_progress_bar.setValue(value)
        self.progress_section._image_pct.setText(f"{value}%")
        self.taskbar_progress.setValue(value)

    def _warn_corrupted_image(self, path: str):
        name = os.path.basename(path)
        QMessageBox.warning(self, "Corrupted Image",
            f"The following image appears to be corrupted and will be skipped:\n\n{name}", QMessageBox.Ok)

    def _store_temp_dirs(self, dirs: list):
        self._pending_temp_dirs = dirs

    def _cleanup_temp_dirs(self):
        for d in self._pending_temp_dirs:
            if d and os.path.isdir(d):
                try:
                    shutil.rmtree(d, ignore_errors=True)
                except Exception as e:
                    print(f"Could not remove temp folder {d}: {e}")
        self._pending_temp_dirs = []

    def on_image_processing_finished(self):
        self.progress_section.hide_image()
        self.taskbar_progress.reset()
        self.taskbar_progress.hide()
        self.continue_with_video_export()

    def _refresh_stats(self):
        n = len(self.images)
        dur = sum(img["duration"] for img in self.images)
        aud = len(self.audio_files)
        self.slide_count_label.setText(str(n))
        self.audio_count_label.setText(str(aud))
        self.stats_bar.update_stats(n, dur, aud)

    # ── Audio ─────────────────────────────────────────────────────────────────

    def add_audio(self):
        file, _ = QFileDialog.getOpenFileName(
            self, "Select Audio", "", "Audio Files (*.mp3 *.wav *.flac)"
        )
        if file:
            self.audio_files.append({"path": file})
            self._audio_dur_cache.pop(file, None)  # ensure fresh duration next time
            self.update_audio_table()

    def update_audio_table(self):
        self.audio_table.setRowCount(len(self.audio_files))
        for row, audio in enumerate(self.audio_files):
            filename_item = QTableWidgetItem(os.path.basename(audio["path"]))
            filename_item.setFlags(filename_item.flags() & ~Qt.ItemIsEditable)
            self.audio_table.setItem(row, 1, filename_item)
            self.audio_table.setRowHeight(row, 30)

            mu_btn = QPushButton("↑"); mu_btn.setProperty("action", "icon")
            md_btn = QPushButton("↓"); md_btn.setProperty("action", "icon")
            del_btn = QPushButton("✕"); del_btn.setProperty("action", "icon")
            del_btn.setStyleSheet(f"QPushButton {{ color: {COLORS['danger']}; background: transparent; border: none; padding: 2px 4px; }}"
                                  f"QPushButton:hover {{ background: rgba(247,90,90,0.15); }}")
            mu_btn.clicked.connect(lambda _, r=row: self.move_audio_up(r))
            md_btn.clicked.connect(lambda _, r=row: self.move_audio_down(r))
            del_btn.clicked.connect(lambda _, r=row: self.delete_audio(r))

            bw = QWidget()
            bl = QHBoxLayout(bw)
            bl.addWidget(mu_btn); bl.addWidget(md_btn); bl.addWidget(del_btn)
            bl.setContentsMargins(4, 0, 4, 0); bl.setSpacing(2)
            self.audio_table.setCellWidget(row, 0, bw)
        self._refresh_stats()

    def move_audio_up(self, row: int):
        if row > 0:
            self.audio_files[row], self.audio_files[row - 1] = self.audio_files[row - 1], self.audio_files[row]
            self.update_audio_table()
            self.audio_table.setCurrentCell(row - 1, 1)

    def move_audio_down(self, row: int):
        if row < len(self.audio_files) - 1:
            self.audio_files[row], self.audio_files[row + 1] = self.audio_files[row + 1], self.audio_files[row]
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

    # Cache for audio durations — keyed by absolute path, cleared on audio list change
    _audio_dur_cache: dict = {}

    def _get_audio_dur(self, path: str) -> float:
        """Return cached audio duration, calling ffprobe only when necessary."""
        if path not in self._audio_dur_cache:
            self._audio_dur_cache[path] = _get_audio_duration(path)
        return self._audio_dur_cache[path]

    def _total_audio_duration(self) -> float:
        return sum(self._get_audio_dur(a["path"]) for a in self.audio_files)

    def continue_with_video_export(self):
        total_img_dur   = sum(img["duration"] for img in self.images)
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

        import shutil as _shutil
        import shlex  as _shlex
        ffmpeg_exe = _ffmpeg_exe()
        import os as _os
        if not _os.path.isfile(ffmpeg_exe):
            ffmpeg_exe = None
        if not ffmpeg_exe:
            QMessageBox.critical(self, "FFmpeg not found",
                "ffmpeg could not be found on your PATH or in the script folder.\n"
                "Please place ffmpeg.exe next to Eventure.py or add it to your PATH.", QMessageBox.Ok)
            return

        raw_args = _shlex.split(command, posix=False)
        args = [a.strip('"') for a in raw_args[1:]]

        self.process = QProcess(self)
        from PyQt5.QtCore import QProcessEnvironment
        self.process.setProcessEnvironment(QProcessEnvironment.systemEnvironment())
        self.process.readyReadStandardError.connect(self.update_progress)
        self.process.finished.connect(self.export_finished)

        self.progress_section.show_export()
        self.progress_bar.setValue(0)
        self.progress_section._export_pct.setText("0%")
        self.taskbar_progress.show()
        self.taskbar_progress.setValue(0)

        self.process.start(ffmpeg_exe, args)
        if self.process.state() == 0:
            err = self.process.errorString()
            QMessageBox.critical(self, "Export failed", f"Failed to launch ffmpeg:\n{err}", QMessageBox.Ok)

    def export_slideshow(self):
        if not self.images or not self.audio_files:
            QMessageBox.critical(self, self.tr("error"), self.tr("error_no_audio"), QMessageBox.Ok)
            return
        QMessageBox.warning(self, self.tr("just_know"), self.tr("no_secondery"), QMessageBox.Ok)

        file_path, _ = QFileDialog.getSaveFileName(self, "Save Slideshow", "", "Video Files (*.mp4);;All Files (*)")
        if not file_path:
            QMessageBox.critical(self, self.tr("error"), self.tr("error_select_location"), QMessageBox.Ok)
            return
        self.output_file = file_path
        output_folder = os.path.join(os.path.dirname(self.images[0]["path"]), "A_Blur")

        self.images_backup = copy.deepcopy(self.images)
        self.backup_state  = True

        self.progress_section.show_image()
        self.image_progress_bar.setValue(0)
        self.progress_section._image_pct.setText("0%")
        self.taskbar_progress.show()
        self.taskbar_progress.setValue(0)

        self.image_worker = ImageProcessingWorker(self.images, output_folder, self.common_width, self.common_height,
                                                   font_family=getattr(self, "text_font", None))
        self.image_worker.progress.connect(self.update_image_progress)
        self.image_worker.cleanup_dirs.connect(self._store_temp_dirs)
        self.image_worker.finished.connect(self.on_image_processing_finished)
        self.image_worker.corrupted_image.connect(self._warn_corrupted_image)
        self.image_worker.start()

    def export_finished(self):
        self.progress_bar.setValue(100)
        self.progress_section._export_pct.setText("100%")
        self.taskbar_progress.setValue(100)

        # ── Pass 2: embed thumbnail as attached_pic via a fast remux ─────────
        # Done as a separate pass so -shortest in pass 1 never sees the
        # zero-duration image stream (which truncated the entire video).
        output_norm = str(self.output_file).replace(chr(92), "/")
        first_img_path = str(self.images[0]["path"]).replace(chr(92), "/") if self.images else ""
        if first_img_path and os.path.exists(first_img_path):
            import tempfile as _tmp2
            tmp_main = _tmp2.NamedTemporaryFile(suffix=".mp4", delete=False).name.replace("\\", "/")
            try:
                os.rename(self.output_file, tmp_main)
                thumb_cmd = (
                    f'ffmpeg -y -i "{tmp_main}" -i "{first_img_path}" '
                    f'-map 0 -map 1:v '
                    f'-c copy -c:v:1 mjpeg '
                    f'-disposition:v:1 attached_pic '
                    f'-metadata:s:v:1 comment=Cover '
                    f'-movflags +faststart '
                    f'"{output_norm}"'
                )
                import subprocess as _sp2
                _sp2.run(
                    thumb_cmd, shell=True, check=False,
                    stdout=_sp2.DEVNULL, stderr=_sp2.DEVNULL,
                    creationflags=_sp2.CREATE_NO_WINDOW if hasattr(_sp2, "CREATE_NO_WINDOW") else 0,
                )
            except Exception as e:
                print(f"Thumbnail embed error (non-fatal): {e}")
                # Restore original if rename succeeded but ffmpeg failed
                if not os.path.exists(self.output_file) and os.path.exists(tmp_main):
                    try: os.rename(tmp_main, self.output_file)
                    except Exception: pass
            finally:
                try: os.remove(tmp_main)
                except OSError: pass

        QMessageBox.information(self, self.tr("success_export_complete_window"), self.tr("success_export_complete"))
        self.progress_section.hide_export()
        self.taskbar_progress.reset()
        self.taskbar_progress.hide()
        self._cleanup_temp_dirs()
        fc = getattr(self, "_fc_script_path", None)
        if fc:
            try: os.remove(fc)
            except OSError: pass
            self._fc_script_path = None

    def build_ffmpeg_command(self) -> str:
        """
        Build the ffmpeg command for the final MP4 export.

        xfade math (correct):
        ─────────────────────
        Each slide has:
          • duration      (d)  = net visible seconds the user set (e.g. 3 s)
          • transition_dur (td) = overlap with the NEXT slide      (e.g. 1 s)

        ffmpeg clip lengths:
          • Every clip (except the last) needs d + td seconds of material so that
            the next slide can begin td seconds before this one finishes.
          • The last clip only needs d seconds (nothing follows it).

        xfade offset for transition i→i+1:
          offset[i] = sum of d[0..i]   (cumulative FULL durations, not net)
          This is the exact moment in the composited timeline where slide i
          should start cross-fading into slide i+1.

        Audio fade-out starts at total_video_dur = sum(d) to match the last frame.
        """
        images = self.images
        n      = len(images)
        inputs, filters = [], []

        # ── Video inputs ──────────────────────────────────────────────────────
        for i, img in enumerate(images):
            d  = img["duration"]
            td = img.get("transition_duration", 1)
            # All clips except the last need extra material for the outgoing transition.
            # The last clip has no outgoing transition so it only needs d seconds.
            clip_len = d + td if i < n - 1 else d

            kb_clip = img.get("_kb_clip_path")
            if kb_clip and os.path.exists(kb_clip):
                kb_clip_norm = str(kb_clip).replace("\\", "/")
                inputs.append(f'-t {clip_len} -i "{kb_clip_norm}"')
                filters.append(
                    f"[{i}:v]fps=25,setpts=PTS-STARTPTS,"
                    f"scale=1920:1080,setsar=1,format=yuv420p[{i}v]"
                )
            else:
                img_path_norm = str(img["path"]).replace("\\", "/")
                inputs.append(f'-loop 1 -t {clip_len} -i "{img_path_norm}"')
                filters.append(
                    f"[{i}:v]fps=25,scale=1920:1080,"
                    f"setsar=1,setpts=PTS-STARTPTS,format=yuv420p[{i}v]"
                )

        # ── xfade transitions ─────────────────────────────────────────────────
        # offset[i] = cumulative sum of full durations for slides 0 … i.
        # This is exactly when slide i should begin transitioning into slide i+1.
        cumulative_offset = 0.0
        for i in range(n - 1):
            cumulative_offset += images[i]["duration"]
            td   = images[i].get("transition_duration", 1)
            prev = f"[{i}v]" if i == 0 else f"[v{i}]"
            filters.append(
                f"{prev}[{i + 1}v]xfade=transition={images[i]['transition']}"
                f":duration={td}:offset={cumulative_offset:.6f}[v{i + 1}]"
            )

        # ── Force yuv420p (required for broad compatibility) ─────────────────
        # Without this, xfade may output yuv444p (H.264 High 4:4:4 Predictive),
        # which is unplayable in Windows Media Player and most mobile players.
        final_v_label = f"v{n - 1}" if n > 1 else "0v"
        filters.append(f"[{final_v_label}]format=yuv420p[vout]")

        # ── Audio inputs + concat + fade-out ─────────────────────────────────
        audio_index   = n
        audio_streams = []
        for i, audio in enumerate(self.audio_files):
            audio_norm = str(audio["path"]).replace("\\", "/")
            inputs.append(f'-i "{audio_norm}"')
            audio_streams.append(f"[{audio_index + i}:a]")

        # Total video duration = sum of each slide's NET duration (what the user set).
        # Audio fade-out begins fade_duration seconds before the video ends.
        total_video_dur = sum(img["duration"] for img in images)
        fade_duration   = 3.0
        fade_start      = max(0.0, total_video_dur - fade_duration)
        print(f"[export] total_video_dur={total_video_dur:.1f}s  fade_start={fade_start:.1f}s")

        if len(audio_streams) > 1:
            filters.append(
                f"{''.join(audio_streams)}concat=n={len(audio_streams)}:v=0:a=1[outa_raw]"
            )
            filters.append(
                f"[outa_raw]afade=t=out:st={fade_start:.3f}:d={fade_duration:.3f}[outa]"
            )
            audio_map = "-map [outa]"
        elif len(audio_streams) == 1:
            filters.append(
                f"[{audio_index}:a]afade=t=out:st={fade_start:.3f}:d={fade_duration:.3f}[outa]"
            )
            audio_map = "-map [outa]"
        else:
            audio_map = ""

        # ── Write filter_complex to a temp file (avoids command-line length limits) ─
        filter_complex = ";".join(filters)

        import tempfile as _tempfile
        fc_file = _tempfile.NamedTemporaryFile(mode="w", suffix=".txt", delete=False, encoding="utf-8")
        fc_file.write(filter_complex)
        fc_file.close()
        self._fc_script_path = fc_file.name

        output_norm = str(self.output_file).replace(chr(92), "/")

        command = (
            f'ffmpeg -y {" ".join(inputs)} '
            f'-filter_complex_script "{fc_file.name}" '
            f'-map [vout] {audio_map} '
            f'-c:v libx264 -preset ultrafast -crf 28 '
            f'-c:a aac '
            f'-movflags +faststart -shortest '
            f'"{output_norm}"'
        )

        # Restore image list to pre-export state (durations may have been adjusted
        # by the audio-mismatch handler in continue_with_video_export).
        if self.backup_state:
            self.images       = self.images_backup
            self.backup_state = False
            self.update_image_table()

        return command

    def update_progress(self):
        output = self.process.readAllStandardError().data().decode("utf-8", errors="ignore")
        for line in output.split("\n"):
            if "time=" in line:
                time_str = line.split("time=")[1].split(" ")[0]
                parts = time_str.split(":")
                if len(parts) == 3:
                    try:
                        h, m, s = map(float, parts)
                        cur   = h * 3600 + m * 60 + s
                        total = sum(img["duration"] for img in self.images)
                        pct   = int(cur / total * 100) if total else 0
                        self.progress_bar.setValue(pct)
                        self.progress_section._export_pct.setText(f"{pct}%")
                        self.taskbar_progress.setValue(pct)
                    except ValueError:
                        pass

    # ── Preview ───────────────────────────────────────────────────────────────

    def _load_pixmap(self, path: str, crop: tuple | None = None) -> QPixmap:
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
            if crop:
                iw, ih = img.size
                cx = max(0, int(crop[0] * iw))
                cy = max(0, int(crop[1] * ih))
                cw = max(1, min(int(crop[2] * iw), iw - cx))
                ch = max(1, min(int(crop[3] * ih), ih - cy))
                img = img.crop((cx, cy, cx + cw, cy + ch))
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
            pixmap   = self._load_pixmap(img_data["path"], crop=img_data.get("crop"))
            if not pixmap.isNull():
                rotation = img_data.get("rotation", 0)
                if rotation:
                    t = QTransform()
                    t.rotate(rotation)
                    pixmap = pixmap.transformed(t, Qt.SmoothTransformation)
                self.preview_panel.set_pixmap(pixmap, os.path.basename(img_data["path"]))
            else:
                self.preview_panel.clear()

    def setup_connections(self):
        self.image_table.itemSelectionChanged.connect(self.update_preview)
        self.image_table.itemSelectionChanged.connect(self._sync_filmstrip_selection)

    def _sync_filmstrip_selection(self):
        selected = self.image_table.selectedItems()
        if selected:
            row = self.image_table.row(selected[0])
            self.filmstrip.highlight_index(row)

    # ── Project ───────────────────────────────────────────────────────────────

    def clear_project(self):
        self.images.clear()
        self.audio_files.clear()
        self.image_table.setRowCount(0)
        self.audio_table.setRowCount(0)
        self.preview_panel.clear()
        self.loaded_project = ""
        self._refresh_stats()

    def _write_project_file(self, path: str):
        with open(path, "w", encoding="utf-8") as f:
            f.write(f"{len(self.audio_files)}\n")
            for audio in self.audio_files:
                f.write(f"{audio['path']}\n")
            for img in self.images:
                text = img.get("text", "").replace("\n", "\\n")
                crop = img.get("crop")
                crop_str = f"{crop[0]:.6f}|{crop[1]:.6f}|{crop[2]:.6f}|{crop[3]:.6f}" if crop else "none"
                f.write(
                    f"{img['path']},{img.get('duration', 5)},{img.get('transition', 'fade')},"
                    f"{img.get('transition_duration', 1)},{text},{img.get('rotation', 0)},"
                    f"{img.get('is_second_image', False)},{img.get('date', '')},"
                    f"{img.get('ken_burns', 'none')},{img.get('text_on_kb', True)},"
                    f"{crop_str}\n"
                )

    # ── Recent Files ──────────────────────────────────────────────────────────

    _RECENT_MAX  = 10
    _RECENT_FILE = BASEPATH / "recent_projects.json"

    def _load_recent(self) -> list:
        try:
            with open(self._RECENT_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            # Filter out paths that no longer exist on disk
            return [p for p in data if isinstance(p, str) and os.path.exists(p)]
        except (FileNotFoundError, json.JSONDecodeError):
            return []

    def _save_recent(self, recent: list) -> None:
        try:
            with open(self._RECENT_FILE, "w", encoding="utf-8") as f:
                json.dump(recent[:self._RECENT_MAX], f, ensure_ascii=False, indent=2)
        except OSError as e:
            print(f"Could not save recent projects list: {e}")

    def _push_recent(self, path: str) -> None:
        """Add *path* to the top of the recent list and persist it."""
        recent = self._load_recent()
        try:
            recent.remove(path)
        except ValueError:
            pass
        recent.insert(0, path)
        self._save_recent(recent)
        self._rebuild_recent_menu()

    def _rebuild_recent_menu(self) -> None:
        """Repopulate the Recent Projects submenu from disk."""
        if not hasattr(self, "recent_menu"):
            return
        self.recent_menu.clear()
        recent = self._load_recent()
        if not recent:
            empty_action = QAction("(no recent projects)", self)
            empty_action.setEnabled(False)
            self.recent_menu.addAction(empty_action)
            return
        for path in recent:
            name = os.path.basename(path)
            display = f"{name}  —  {os.path.dirname(path)}"
            action = QAction(display, self)
            action.setData(path)
            action.triggered.connect(lambda checked, p=path: self._open_recent(p))
            self.recent_menu.addAction(action)
        self.recent_menu.addSeparator()
        clear_action = QAction("Clear Recent Projects", self)
        clear_action.triggered.connect(self._clear_recent)
        self.recent_menu.addAction(clear_action)

    def _open_recent(self, path: str) -> None:
        if not os.path.exists(path):
            QMessageBox.warning(
                self, "File Not Found",
                f"The project file no longer exists:\n{path}\n\nIt will be removed from the recent list."
            )
            recent = self._load_recent()
            try:
                recent.remove(path)
            except ValueError:
                pass
            self._save_recent(recent)
            self._rebuild_recent_menu()
            return
        self._run_in_thread(
            target=lambda: self._parse_project_file(path),
            on_success=lambda result: self._apply_loaded_project(result, path),
            on_error=lambda e: QMessageBox.critical(
                self, self.tr("error"), f"Failed to load project:\n{e}"
            ),
            status_msg=f"Loading {os.path.basename(path)}…",
        )

    def _clear_recent(self) -> None:
        self._save_recent([])
        self._rebuild_recent_menu()

    def save_project(self):
        if self.loaded_project:
            self._write_project_file(self.loaded_project)
            self.statusBar().showMessage(
                f"  ✔  Saved  {os.path.basename(self.loaded_project)}", 3000
            )
        else:
            self.save_project_as()

    def save_project_as(self):
        file_name, _ = QFileDialog.getSaveFileName(
            self, "Save Project", "", "Project Files (*.slideshow);;All Files (*)"
        )
        if file_name:
            self._write_project_file(file_name)
            self.loaded_project = file_name
            self._push_recent(file_name)
            proj_name = os.path.splitext(os.path.basename(file_name))[0]
            self.setWindowTitle(f"{self.tr('window_title')} — {proj_name}")

    def load_project(self):
        file_name, _ = QFileDialog.getOpenFileName(
            self, "Load Project", "", "Project Files (*.slideshow);;All Files (*)"
        )
        if not file_name:
            return
        self._run_in_thread(
            target=lambda: self._parse_project_file(file_name),
            on_success=lambda result: self._apply_loaded_project(result, file_name),
            on_error=lambda e: QMessageBox.critical(
                self, self.tr("error"), f"Failed to load project:\n{e}"
            ),
            status_msg=f"Loading {os.path.basename(file_name)}…",
        )

    def import_pptx(self):
        file_name, _ = QFileDialog.getOpenFileName(
            self, "Select a PowerPoint file", "", "PowerPoint files (*.pptx;*.pptm);;All Files (*)"
        )
        if not file_name:
            return
        self._run_in_thread(
            target=lambda: extract_pptx_content_to_slideshow_file(file_name),
            on_success=lambda slideshow_file: (
                self._load_project_from_path(slideshow_file) if slideshow_file else None
            ),
            on_error=lambda e: QMessageBox.critical(
                self, self.tr("error"), f"Failed to import PPTX:\n{e}"
            ),
            status_msg=f"Importing {os.path.basename(file_name)}…",
        )

    # ── Threaded task runner ──────────────────────────────────────────────────

    def _run_in_thread(self, target, on_success, on_error, status_msg="Working…"):
        """
        Run *target* on a background QThread.
        When it finishes, call *on_success(result)* or *on_error(exception)*
        back on the main thread.  Shows a non-blocking status message while busy.
        """
        # Show busy indicator in the status bar
        self.statusBar().showMessage(f"  ⏳  {status_msg}")
        QApplication.setOverrideCursor(Qt.WaitCursor)

        class _Worker(QThread):
            done    = pyqtSignal(object)   # carries the return value
            failed  = pyqtSignal(Exception)

            def __init__(self, fn):
                super().__init__()
                self._fn = fn

            def run(self):
                try:
                    self.done.emit(self._fn())
                except Exception as exc:
                    self.failed.emit(exc)

        worker = _Worker(target)

        def _on_done(result):
            QApplication.restoreOverrideCursor()
            self.statusBar().clearMessage()
            try:
                on_success(result)
            except Exception as exc:
                on_error(exc)
            # Keep the worker alive until this slot returns
            worker.deleteLater()

        def _on_failed(exc):
            QApplication.restoreOverrideCursor()
            self.statusBar().clearMessage()
            on_error(exc)
            worker.deleteLater()

        worker.done.connect(_on_done)
        worker.failed.connect(_on_failed)

        # Store reference so GC doesn't collect the thread before it finishes
        if not hasattr(self, "_bg_workers"):
            self._bg_workers = []
        self._bg_workers.append(worker)
        worker.finished.connect(lambda: self._bg_workers.remove(worker) if worker in self._bg_workers else None)

        worker.start()

    # ── Project parse (runs in background thread) ─────────────────────────────

    @staticmethod
    def _parse_crop(s: str) -> tuple | None:
        """Parse a crop string like '0.1|0.05|0.8|0.9' → tuple, or None."""
        if not s or s.strip().lower() in ("none", ""):
            return None
        try:
            vals = [float(v) for v in s.strip().split("|")]
            if len(vals) == 4:
                return tuple(vals)
        except ValueError:
            pass
        return None

    def _parse_project_file(self, file_name: str) -> dict:
        """
        Parse a .slideshow file entirely off the UI thread.
        Returns a dict with 'audio_files' and 'images' keys.
        Raises on any error — the caller's on_error handler will show the dialog.
        """
        with open(file_name, "r", encoding="utf-8") as f:
            lines = f.readlines()
        count = int(lines[0].strip())
        if len(lines) < count + 1:
            raise ValueError("Project file is truncated.")
        audio_files = [{"path": lines[i + 1].strip()} for i in range(count)]
        images = []
        for line in lines[count + 1:]:
            parts = line.strip().split(",")
            if len(parts) < 8:
                continue
            path        = parts[0]
            dur         = parts[1]
            transition  = parts[2]
            trans_dur   = parts[3]
            text        = parts[4]
            rotation    = parts[5]
            is_second   = parts[6]
            date        = parts[7] if len(parts) > 7 else ""
            ken_burns   = parts[8].strip() if len(parts) > 8 else "none"
            crop        = self._parse_crop(parts[10]) if len(parts) > 10 else None
            images.append({
                "path":                path,
                "duration":            float(dur),
                "transition":          transition,
                "transition_duration": float(trans_dur),
                "text":                text.replace("\\n", "\n"),
                "rotation":            int(rotation),
                "is_second_image":     is_second.strip().lower() == "true",
                "date":                date,
                "ken_burns":           ken_burns,
                "crop":                crop,
            })
        return {"audio_files": audio_files, "images": images}

    def _apply_loaded_project(self, parsed: dict, file_name: str):
        """
        Apply a parsed project result to the UI.  Must run on the main thread.
        """
        self.audio_files = parsed["audio_files"]
        self.images      = parsed["images"]
        self.update_image_table()
        self.update_audio_table()
        self.loaded_project = file_name
        self._push_recent(file_name)
        # Reflect the project name in the window title
        proj_name = os.path.splitext(os.path.basename(file_name))[0]
        self.setWindowTitle(f"{self.tr('window_title')} — {proj_name}")

    def _load_project_from_path(self, file_name: str):
        """Load a project from *file_name* on the main thread (used after PPTX import)."""
        try:
            parsed = self._parse_project_file(file_name)
            self._apply_loaded_project(parsed, file_name)
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
        KB_OPTIONS = ["none", "zoom_in", "zoom_out", "pan_left", "pan_right", "pan_up", "pan_down"]
        dialog = QInputDialog(self)
        dialog.setWindowTitle(self.tr("dialog_set_ken_burns_title"))
        dialog.setLabelText(self.tr("dialog_set_ken_burns_label"))
        dialog.setComboBoxItems(KB_OPTIONS)
        if dialog.exec_() == QDialog.Accepted:
            effect = dialog.textValue()
            for img in self.images:
                img["ken_burns"] = effect
            self.update_image_table()

    def _set_random_ken_burns_per_image(self):
        KB_OPTIONS = ["zoom_in", "zoom_out", "pan_left", "pan_right", "pan_up", "pan_down"]
        for img in self.images:
            img["ken_burns"] = random.choice(KB_OPTIONS)
        self.update_image_table()

    def _set_smart_ken_burns(self):
        REVERSAL = {
            "zoom_in": "zoom_out", "zoom_out": "zoom_in",
            "pan_left": "pan_right", "pan_right": "pan_left",
            "pan_up": "pan_down", "pan_down": "pan_up",
        }
        ZOOM_EFFECTS = ["zoom_in", "zoom_out"]
        PAN_EFFECTS  = ["pan_left", "pan_right", "pan_up", "pan_down"]

        def _opposite_category(effect):
            return PAN_EFFECTS if effect in ZOOM_EFFECTS else ZOOM_EFFECTS

        prev_effect = None
        reversal_streak = 0

        for img in self.images:
            if img.get("is_second_image", False):
                continue
            if prev_effect is None:
                chosen = random.choice(ZOOM_EFFECTS + PAN_EFFECTS)
            else:
                if random.random() < 0.20:
                    chosen = random.choice(_opposite_category(prev_effect))
                    reversal_streak = 0
                elif reversal_streak >= 2:
                    chosen = random.choice(_opposite_category(prev_effect))
                    reversal_streak = 0
                else:
                    chosen = REVERSAL[prev_effect]
                    reversal_streak += 1
            img["ken_burns"] = chosen
            prev_effect = chosen
        self.update_image_table()

    def set_all_images_transition(self):
        dialog = QInputDialog(self)
        dialog.setWindowTitle(self.tr("dialog_set_transition"))
        dialog.setLabelText(self.tr("dialog_select_transition"))
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
        """
        Smart batch-duration dialog.
        Shows total audio, total slides, and a live preview of
        the per-slide duration before committing anything.
        """
        if not self.images:
            QMessageBox.warning(self, self.tr("error_no_images_title"), self.tr("error_no_images"))
            return

        total_audio = self._total_audio_duration()
        n_all       = len(self.images)
        n_primary   = sum(1 for img in self.images if not img.get("is_second_image"))

        # ── Dialog ────────────────────────────────────────────────────────────
        dlg = QDialog(self)
        dlg.setWindowTitle("Batch Duration from Audio")
        dlg.setMinimumWidth(380)
        layout = QVBoxLayout(dlg)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(14)

        # Title
        title_lbl = QLabel("⏱  Batch Duration from Audio")
        title_lbl.setStyleSheet(
            f"font-size: 15px; font-weight: 700; color: {COLORS['text_primary']};"
        )
        layout.addWidget(title_lbl)
        layout.addWidget(_make_divider())

        # Info rows
        def _info(label, value):
            row = QHBoxLayout()
            l = QLabel(label)
            l.setStyleSheet(f"color: {COLORS['text_secondary']}; font-size: 12px;")
            v = QLabel(value)
            v.setStyleSheet(f"color: {COLORS['text_primary']}; font-size: 13px; font-weight: 600;")
            v.setAlignment(Qt.AlignRight)
            row.addWidget(l); row.addStretch(); row.addWidget(v)
            layout.addLayout(row)

        _info("Total audio duration:", format_time_hms(total_audio))
        _info("Total slides:", str(n_all))
        _info("Primary slides (excl. second-image):", str(n_primary))
        layout.addWidget(_make_divider())

        # Scope radio
        scope_lbl = QLabel("Apply to:")
        scope_lbl.setStyleSheet(f"color: {COLORS['text_secondary']}; font-size: 12px;")
        layout.addWidget(scope_lbl)

        rb_all     = QRadioButton(f"All {n_all} slides equally")
        rb_primary = QRadioButton(f"Primary slides only ({n_primary} slides, skip second-image)")
        rb_all.setChecked(True)
        for rb in (rb_all, rb_primary):
            rb.setStyleSheet(f"color: {COLORS['text_primary']}; font-size: 12px;")
        layout.addWidget(rb_all)
        layout.addWidget(rb_primary)
        layout.addWidget(_make_divider())

        # Tail reserve spinner
        tail_row = QHBoxLayout()
        tail_lbl = QLabel("Reserve at end (seconds):")
        tail_lbl.setStyleSheet(f"color: {COLORS['text_secondary']}; font-size: 12px;")
        tail_spin = QSpinBox()
        tail_spin.setRange(0, 30)
        tail_spin.setValue(2)
        tail_spin.setFixedWidth(70)
        tail_row.addWidget(tail_lbl); tail_row.addStretch(); tail_row.addWidget(tail_spin)
        layout.addLayout(tail_row)

        # Live preview label
        preview_lbl = QLabel()
        preview_lbl.setStyleSheet(
            f"background: {COLORS['bg_card']}; border: 1px solid {COLORS['border']}; "
            f"border-radius: 6px; padding: 10px 14px; "
            f"color: {COLORS['accent']}; font-size: 13px; font-weight: 700;"
        )
        preview_lbl.setAlignment(Qt.AlignCenter)
        layout.addWidget(preview_lbl)

        def _update_preview():
            tail   = tail_spin.value()
            usable = max(0.0, total_audio - tail)
            n      = n_primary if rb_primary.isChecked() else n_all
            if n == 0:
                preview_lbl.setText("No slides to distribute.")
                return
            new_dur = max(2, int(usable / n))
            preview_lbl.setText(
                f"Each slide → {new_dur} s   "
                f"({format_time_hms(usable)} ÷ {n} slides)"
            )

        rb_all.toggled.connect(lambda _: _update_preview())
        rb_primary.toggled.connect(lambda _: _update_preview())
        tail_spin.valueChanged.connect(lambda _: _update_preview())
        _update_preview()

        # Buttons
        btn_row = QHBoxLayout()
        btn_row.setSpacing(8)
        cancel_btn = _styled_btn("Cancel", "")
        apply_btn  = _styled_btn("✔  Apply", "primary")
        apply_btn.setFixedHeight(36)
        cancel_btn.clicked.connect(dlg.reject)
        apply_btn.clicked.connect(dlg.accept)
        btn_row.addStretch()
        btn_row.addWidget(cancel_btn)
        btn_row.addWidget(apply_btn)
        layout.addLayout(btn_row)

        if dlg.exec_() != QDialog.Accepted:
            return

        # ── Apply ─────────────────────────────────────────────────────────────
        tail   = tail_spin.value()
        usable = max(0.0, total_audio - tail)
        apply_primary_only = rb_primary.isChecked()
        targets = (
            [img for img in self.images if not img.get("is_second_image")]
            if apply_primary_only else self.images
        )
        n = len(targets)
        if n == 0:
            return
        new_dur = max(2, int(usable / n))
        for img in targets:
            img["duration"] = new_dur
        self.update_image_table()
        self.statusBar().showMessage(
            f"  ✔  Set {n} slides to {new_dur} s each  "
            f"(total {format_time_hms(n * new_dur)})", 4000
        )

    # ── Premiere Export ───────────────────────────────────────────────────────

    def export_premiere_slideshow(self):
        file_path, _ = QFileDialog.getSaveFileName(self, "Save Premiere Slideshow", "", "Folder")
        if not file_path:
            QMessageBox.critical(self, self.tr("error"), self.tr("error_select_location_premiere"), QMessageBox.Ok)
            return
        self.premiere_project_folder = file_path

        self.progress_section.show_premiere()
        self.image_premiere_progress_bar.setValue(0)
        self.progress_section._premiere_pct.setText("0%")
        self.taskbar_progress.show()
        self.taskbar_progress.setValue(0)

        self.image_premiere_worker = ImageProcessingPremiereWorker(
            self.images, self.premiere_project_folder,
            self.common_width, self.common_height,
            audio_files=self.audio_files,
        )
        self.image_premiere_worker.progress.connect(self.update_image_premiere_progress)
        self.image_premiere_worker.finished.connect(self.on_image_premiere_processing_finished)
        self.image_premiere_worker.xml_ready.connect(self.on_premiere_xml_ready)
        self.image_premiere_worker.start()

    def on_premiere_xml_ready(self, xml_path: str):
        print(f"Premiere XML ready: {xml_path}")
        dst_folder = os.path.join(self.premiere_project_folder, "04_פרוייקט")
        os.makedirs(dst_folder, exist_ok=True)
        dst = os.path.join(dst_folder, "premiere_timeline.xml")
        try:
            shutil.move(xml_path, dst)
            QMessageBox.information(self, "Premiere XML",
                f"Timeline XML saved:\n{dst}\n\nIn Premiere Pro: File → Import → premiere_timeline.xml")
        except Exception as e:
            print(f"Could not move XML to project folder: {e}")

    def update_image_premiere_progress(self, value: int):
        self.image_premiere_progress_bar.setValue(value)
        self.progress_section._premiere_pct.setText(f"{value}%")
        self.taskbar_progress.setValue(value)

    def on_image_premiere_processing_finished(self):
        self.progress_section.hide_premiere()
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
        style_src = APP_DIR / "Premiere_Project" / "טקסט למצגת - עברית.prtextstyle"
        if style_src.exists():
            shutil.copy(str(style_src), os.path.join(folder, style_src.name))
        srt_path     = os.path.join(folder, "exported_texts.srt")
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
        src = APP_DIR / "Premiere_Project" / "Project.prproj"
        if not src.exists():
            print(f"Premiere project template not found at {src}")
            return
        dst_folder = os.path.join(self.premiere_project_folder, "04_פרוייקט")
        os.makedirs(dst_folder, exist_ok=True)
        name = os.path.basename(self.premiere_project_folder) + ".prproj"
        shutil.copy(str(src), os.path.join(dst_folder, name))

    # ── Slideshow Preview ─────────────────────────────────────────────────────

    def open_preview_dialog(self):
        if not self.images:
            QMessageBox.information(self, "Preview",
                "Add some images first before previewing the slideshow.")
            return
        dlg = SlideshowPreviewDialog(
            images=self.images,
            audio_files=self.audio_files,
            parent=self,
        )
        dlg.exec_()

    # ── Crop ──────────────────────────────────────────────────────────────────

    def open_crop_dialog(self, row: int):
        if not (0 <= row < len(self.images)):
            return
        img = self.images[row]
        dlg = CropDialog(
            image_path=img["path"],
            rotation=img.get("rotation", 0),
            existing_crop=img.get("crop") or None,
            parent=self,
        )
        if dlg.exec_() == QDialog.Accepted:
            result = dlg.get_result()
            # If the result is a full-image crop (within 0.5% on all sides), clear it
            if result and (result[0] < 0.005 and result[1] < 0.005
                           and result[2] > 0.99 and result[3] > 0.99):
                result = None
            img["crop"] = result
            # Refresh the row so the ✂ button colour updates
            self.image_table.blockSignals(True)
            self._populate_row(row, img)
            self.image_table.blockSignals(False)
            self.update_preview_with_row(row)

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

    def _open_font_picker(self):
        """Open the font picker dialog and apply the chosen font to text overlays."""
        dlg = FontPickerDialog(current_font=self.text_font, parent=self)
        if dlg.exec_() == QDialog.Accepted:
            family = dlg.get_result()
            if family:
                self.text_font = family
                # Persist choice
                try:
                    cfg_path = BASEPATH / "font_settings.json"
                    with open(cfg_path, "w", encoding="utf-8") as f:
                        json.dump({"text_font": family}, f)
                except Exception as e:
                    print(f"Could not save font setting: {e}")
                # Register with Qt so it's available for rendering
                QFontDatabase.addApplicationFont(
                    str(BASEPATH / "Fonts" / "GoogleFonts" / f"{family.replace(' ', '_')}.ttf")
                )
                QMessageBox.information(
                    self, "Font Updated",
                    f"Text font set to '{family}'.\n\n"
                    "This font will be used for all new text overlays on exported slideshows."
                )

    def _load_font_settings(self):
        """Load saved font settings on startup."""
        try:
            cfg_path = BASEPATH / "font_settings.json"
            with open(cfg_path, "r", encoding="utf-8") as f:
                cfg = json.load(f)
            self.text_font = cfg.get("text_font", "Segoe UI")
        except (FileNotFoundError, json.JSONDecodeError):
            pass


    def export_pdf_book(self):
        if not self.images:
            QMessageBox.warning(self, "No Images", "Add images before exporting.")
            return

        # Check reportlab availability
        try:
            import reportlab  # noqa: F401
        except ImportError:
            QMessageBox.critical(
                self, "Missing Library",
                "The 'reportlab' library is required for PDF export.\n\n"
                "Install it with:\n  pip install reportlab\n\n"
                "Then restart Eventure and try again."
            )
            return

        out_path, _ = QFileDialog.getSaveFileName(
            self, "Save PDF Photo Book", "", "PDF Files (*.pdf)"
        )
        if not out_path:
            return

        progress_dlg = QDialog(self)
        progress_dlg.setWindowTitle("Exporting PDF…")
        progress_dlg.setFixedSize(340, 100)
        progress_dlg.setWindowFlags(
            progress_dlg.windowFlags() & ~Qt.WindowCloseButtonHint
        )
        pl = QVBoxLayout(progress_dlg)
        pl.setContentsMargins(20, 16, 20, 16)
        pdf_lbl = QLabel("Creating PDF photo book…")
        pdf_lbl.setStyleSheet(f"color: {COLORS['text_secondary']}; font-size: 12px;")
        pdf_bar = QProgressBar()
        pdf_bar.setRange(0, 100)
        pl.addWidget(pdf_lbl)
        pl.addWidget(pdf_bar)

        class _PDFWorker(QThread):
            progress = pyqtSignal(int)
            done     = pyqtSignal(bool)
            def __init__(self, images, path):
                super().__init__()
                self._images = images
                self._path   = path
            def run(self):
                ok = export_pdf_photo_book(
                    self._images, self._path,
                    progress_callback=self.progress.emit
                )
                self.done.emit(ok)

        images_snap = copy.deepcopy(self.images)
        worker = _PDFWorker(images_snap, out_path)
        worker.progress.connect(pdf_bar.setValue)
        worker.done.connect(lambda ok: (
            progress_dlg.accept(),
            QMessageBox.information(
                self, "PDF Exported",
                f"✔  PDF saved to:\n{out_path}"
            ) if ok else QMessageBox.critical(
                self, "PDF Export Failed",
                "Could not create the PDF. Check the console for details."
            )
        ))
        worker.start()
        progress_dlg.exec_()

    # ── HTML Gallery Export (Feature 13) ─────────────────────────────────────

    def export_html_gallery_action(self):
        if not self.images:
            QMessageBox.warning(self, "No Images", "Add images before exporting.")
            return

        out_path, _ = QFileDialog.getSaveFileName(
            self, "Save HTML Gallery", "", "HTML Files (*.html)"
        )
        if not out_path:
            return

        proj_title = (
            os.path.splitext(os.path.basename(self.loaded_project))[0]
            if self.loaded_project else "Eventure Gallery"
        )

        progress_dlg = QDialog(self)
        progress_dlg.setWindowTitle("Exporting HTML Gallery…")
        progress_dlg.setFixedSize(360, 100)
        progress_dlg.setWindowFlags(
            progress_dlg.windowFlags() & ~Qt.WindowCloseButtonHint
        )
        pl = QVBoxLayout(progress_dlg)
        pl.setContentsMargins(20, 16, 20, 16)
        html_lbl = QLabel("Building self-contained HTML gallery…")
        html_lbl.setStyleSheet(f"color: {COLORS['text_secondary']}; font-size: 12px;")
        html_bar = QProgressBar()
        html_bar.setRange(0, 100)
        pl.addWidget(html_lbl)
        pl.addWidget(html_bar)

        class _HTMLWorker(QThread):
            progress = pyqtSignal(int)
            done     = pyqtSignal(bool)
            def __init__(self, images, path, title):
                super().__init__()
                self._images = images
                self._path   = path
                self._title  = title
            def run(self):
                ok = export_html_gallery(
                    self._images, self._path, self._title,
                    progress_callback=self.progress.emit
                )
                self.done.emit(ok)

        images_snap = copy.deepcopy(self.images)
        worker = _HTMLWorker(images_snap, out_path, proj_title)
        worker.progress.connect(html_bar.setValue)
        worker.done.connect(lambda ok: (
            progress_dlg.accept(),
            QMessageBox.information(
                self, "HTML Gallery Exported",
                f"✔  Gallery saved to:\n{out_path}\n\n"
                "Open the file in any web browser to view it."
            ) if ok else QMessageBox.critical(
                self, "Export Failed",
                "Could not write the HTML file. Check the console for details."
            )
        ))
        worker.start()
        progress_dlg.exec_()

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
        self.show_info_action.setText(self.tr("action_show_info"))
        self.set_language_english_action.setText(self.tr("action_set_language_english"))
        self.set_language_hebrew_action.setText(self.tr("action_set_language_hebrew"))
        self.open_help_dialog_action.setText(self.tr("action_browse_help_topics"))

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

    # ── Menu ──────────────────────────────────────────────────────────────────

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

        self.import_images      = _action("action_import_images",   self.add_images,          self.import_menu, "import_images")
        self.import_audio       = _action("action_import_audio",    self.add_audio,           self.import_menu, "import_audio")
        self.import_pptx_action = _action("action_import_pptx",     self.import_pptx,         self.import_menu)
        self.load_action        = _action("action_load_project",    self.load_project,        self.file_menu,   "load")
        self.save_action        = _action("action_save_project",    self.save_project,        self.file_menu,   "save")
        self.save_as_action     = _action("action_save_project_as", self.save_project_as,     self.file_menu,   "save_as")
        self.clear_action       = _action("action_clear_project",   self.clear_project,       self.file_menu)

        # ── Recent Projects submenu ────────────────────────────────────────
        self.file_menu.addSeparator()
        self.recent_menu = self.file_menu.addMenu("Recent Projects")
        self._rebuild_recent_menu()
        self.file_menu.addSeparator()
        self.export_slideshow_action = _action("action_export_slideshow", self.export_slideshow,           self.export_menu)
        self.export_premiere_action  = _action("action_export_premiere",  self.export_premiere_slideshow,  self.export_menu)

        self.delete_row_action      = _action("action_delete_row",    self.delete_image,    self.Img_menu, "delete_row")
        self.move_image_up_action   = _action("action_move_image_up", self.move_image_up,   self.Img_menu, "move_image_up")
        self.move_image_down_action = _action("action_move_image_down", self.move_image_down, self.Img_menu, "move_image_down")
        self.set_all_images_duration_action   = _action("action_set_all_image_duration",   self.set_all_images_duration,    self.Img_menu)
        self.set_random_image_order_action    = _action("action_set_random_images_order",  self.set_random_images_order,    self.Img_menu)
        self.auto_set_images_action           = _action("action_auto_calc_image_duration", self.auto_calc_image_duration,   self.Img_menu)
        self.set_image_location_action        = _action("dialog_set_image_location",       self.set_image_location,         self.Img_menu, "set_image_location")
        self.auto_sort_images_by_date_Newest_action = _action("action_auto_sort_newest_first", lambda: self.auto_sort_images_by_date(True),  self.Auto_sort_menu)
        self.auto_sort_images_by_date_Oldest_action = _action("action_auto_sort_oldest_first", lambda: self.auto_sort_images_by_date(False), self.Auto_sort_menu)

        self.set_all_images_transition_type_action       = _action("action_set_all_transition_type",          self.set_all_images_transition,            self.Transitions_menu)
        self.set_random_transition_for_each_image_action = _action("action_set_random_transition_per_image", self.set_random_transition_for_each_image,  self.Transitions_menu)

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
        self.open_help_dialog_action     = _action("action_browse_help_topics",   self.open_help_dialog,           self.help_menu)

        # ── Font Picker ───────────────────────────────────────────────────
        self.settings_menu.addSeparator()
        font_action = QAction("🔤  Choose Text Font…", self)
        font_action.triggered.connect(self._open_font_picker)
        self.settings_menu.addAction(font_action)

         # Tools menu (new)
        self.tools_menu = self.menubar.addMenu("Tools")

        pdf_action = QAction("📄  Export PDF Photo Book…", self)
        pdf_action.triggered.connect(self.export_pdf_book)
        self.tools_menu.addAction(pdf_action)

        # Feature 13: HTML Gallery
        html_action = QAction("🌐  Export HTML Gallery…", self)
        html_action.triggered.connect(self.export_html_gallery_action)
        self.tools_menu.addAction(html_action)


# ── Worker Threads ────────────────────────────────────────────────────────────

class ImageProcessingWorker(QThread):
    progress        = pyqtSignal(int)
    finished        = pyqtSignal()
    corrupted_image = pyqtSignal(str)
    cleanup_dirs    = pyqtSignal(list)

    def __init__(self, images, output_folder, common_width, common_height, font_family=None):
        super().__init__()
        self.images        = images
        self.output_folder = output_folder
        self.common_width  = common_width
        self.common_height = common_height
        self.font_family   = font_family
        self._temp_dirs: list = []

    def _resize_one(self, i: int):
        img_path   = self.images[i]["path"]
        rotation   = self.images[i]["rotation"]
        text       = self.images[i]["text"]
        has_kb     = self.images[i].get("ken_burns", "none") != "none"
        text_on_static = not has_kb
        crop       = self.images[i].get("crop")
        try:
            original = Image.open(img_path)
            original.verify()
            original = Image.open(img_path)
            if original.size != (self.common_width, self.common_height):
                new_path = Image_resizer.process_image(
                    img_path, self.output_folder, text, rotation, text_on_static,
                    crop=crop, font_family=self.font_family
                )
                return i, new_path
        except Exception as e:
            print(f"Corrupted image {img_path}: {e}")
            self.corrupted_image.emit(img_path)
        return i, None

    def run(self):
        total = len(self.images)
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
                self.progress.emit(int(done / total * 50))

        KB_WORKERS = min(2, os.cpu_count() or 1)
        kb_images = [i for i in range(total) if self.images[i].get("ken_burns", "none") != "none"]
        kb_dir = os.path.join(self.output_folder, "kb_clips")
        if kb_images:
            os.makedirs(kb_dir, exist_ok=True)

        self._temp_dirs = [self.output_folder, kb_dir]

        def _render_kb_one(i: int):
            img    = self.images[i]
            effect = img["ken_burns"]
            d      = img.get("duration", 5)
            td     = img.get("transition_duration", 1)
            # Pre-rendered KB clips must match exactly what build_ffmpeg_command
            # will request: all clips except the last get d+td seconds; last gets d.
            if i < len(self.images) - 1:
                clip_duration = d + td
            else:
                clip_duration = d
            kb_out = os.path.join(kb_dir, f"kb_{i}_{effect}.mp4").replace("\\", "/")
            has_kb = self.images[i].get("ken_burns", "none") != "none"
            text_on_static = not has_kb
            success = render_ken_burns_clip(
                img["path"], effect, clip_duration, kb_out,
                text=img.get("text", ""), text_on_kb=text_on_static,
                crop=img.get("crop"),
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
                    else:
                        print(f"  KB render failed for image {i}, will use still image")
                except Exception as e:
                    print(f"  KB render exception: {e}")
                self.progress.emit(50 + int(done_kb / max(len(kb_images), 1) * 50))

        self.cleanup_dirs.emit(self._temp_dirs)
        self.finished.emit()


class ImageProcessingPremiereWorker(QThread):
    progress  = pyqtSignal(int)
    finished  = pyqtSignal()
    xml_ready = pyqtSignal(str)

    def __init__(self, images, output_folder, common_width, common_height, audio_files=None):
        super().__init__()
        self.images        = images
        self.output_folder = output_folder
        self.common_width  = common_width
        self.common_height = common_height
        self.audio_files   = audio_files or []

    def run(self):
        premiere_export.process_images(self.images, self.output_folder, self.progress.emit)
        bg_folder  = os.path.join(self.output_folder, "01_images", "backgrounds")
        img_folder = os.path.join(self.output_folder, "01_images", "foregrounds")
        slide_list = []
        for i, img in enumerate(self.images, start=1):
            if img.get("is_second_image"):
                fg = os.path.join(img_folder, f"img{i}_2nd_of_img{i-1}.png")
                slide_list.append({"bg_path": None, "fg_path": fg if os.path.exists(fg) else None,
                                   "duration": img.get("duration", 5.0), "text": img.get("text", ""),
                                   "is_second_image": True})
            else:
                bg = os.path.join(bg_folder, f"background_img{i}.jpg")
                fg = os.path.join(img_folder, f"img{i}.png")
                slide_list.append({"bg_path": bg if os.path.exists(bg) else None,
                                   "fg_path": fg if os.path.exists(fg) else None,
                                   "duration": img.get("duration", 5.0), "text": img.get("text", ""),
                                   "is_second_image": False})
        try:
            xml_path = premiere_export.generate_premiere_xml(
                slide_list=slide_list, output_folder=self.output_folder, music_paths=self.audio_files)
            self.xml_ready.emit(xml_path)
        except Exception as e:
            print(f"XML generation error: {e}")
        self.finished.emit()



# ── Slideshow Preview ─────────────────────────────────────────────────────────

class _FrameRenderer:
    """
    Pure-PIL/numpy renderer that produces a single preview frame.
    Works in a background thread – no Qt objects created here.
    Resolution is 960×540 (half 1080p) for real-time performance.
    """
    W, H = 960, 540
    ZOOM = 1.10

    def __init__(self, images: list):
        self.images = images
        self._img_cache: dict = {}   # path → numpy array (H, W, 3)

    # ── image loading ────────────────────────────────────────────────────────

    def _load(self, img_data: dict):
        """Load, EXIF-correct, crop and rotate; return numpy uint8 RGB array."""
        key = (img_data["path"], img_data.get("rotation", 0),
               str(img_data.get("crop")))
        if key in self._img_cache:
            return self._img_cache[key]

        try:
            import numpy as np
            pil = Image.open(img_data["path"])
            # EXIF
            try:
                if hasattr(pil, "_getexif"):
                    exif = pil._getexif()
                    if exif and _ORIENTATION_TAG:
                        v = exif.get(_ORIENTATION_TAG)
                        if v == 3:   pil = pil.rotate(180, expand=True)
                        elif v == 6: pil = pil.rotate(270, expand=True)
                        elif v == 8: pil = pil.rotate(90,  expand=True)
            except Exception:
                pass
            rot = img_data.get("rotation", 0)
            if rot:
                pil = pil.rotate(rot, expand=True)
            crop = img_data.get("crop")
            if crop:
                iw, ih = pil.size
                cx = max(0, int(crop[0]*iw)); cy = max(0, int(crop[1]*ih))
                cw = max(1, min(int(crop[2]*iw), iw-cx))
                ch = max(1, min(int(crop[3]*ih), ih-cy))
                pil = pil.crop((cx, cy, cx+cw, cy+ch))
            pil = pil.convert("RGB")
            arr = np.array(pil)
            # keep cache small
            if len(self._img_cache) > 30:
                self._img_cache.pop(next(iter(self._img_cache)))
            self._img_cache[key] = arr
            return arr
        except Exception as e:
            import numpy as np
            print(f"Preview load error: {e}")
            return np.zeros((self.H, self.W, 3), dtype=np.uint8)

    # ── static frame (no KB) ─────────────────────────────────────────────────

    def render_static(self, img_data: dict, text_override: str | None = None) -> "np.ndarray":
        import numpy as np
        import cv2

        arr = self._load(img_data)
        ih, iw = arr.shape[:2]
        W, H = self.W, self.H

        # aspect-fit
        if iw/ih > W/H:
            fw, fh = W, int(W * ih / iw)
        else:
            fh, fw = H, int(H * iw / ih)

        fitted = cv2.resize(arr, (fw, fh), interpolation=cv2.INTER_LINEAR)

        # blurred background
        bg = cv2.resize(arr, (W, H), interpolation=cv2.INTER_LINEAR)
        bg = cv2.GaussianBlur(bg, (0, 0), 15)

        frame = bg.copy()
        fg_w, fg_h = int(fw*0.9), int(fh*0.9)
        fg = cv2.resize(fitted, (fg_w, fg_h), interpolation=cv2.INTER_LINEAR)
        ox, oy = (W-fg_w)//2, (H-fg_h)//2
        frame[oy:oy+fg_h, ox:ox+fg_w] = fg

        # text
        text = text_override if text_override is not None else img_data.get("text", "")
        if text and text.strip():
            frame = self._draw_text(frame, text)

        return frame

    # ── Ken Burns frame ──────────────────────────────────────────────────────

    def render_kb_frame(self, img_data: dict, t: float) -> "np.ndarray":
        """t = 0.0 … 1.0 within the slide duration."""
        import numpy as np
        import cv2

        arr = self._load(img_data)
        ih, iw = arr.shape[:2]
        W, H   = self.W, self.H
        effect = img_data.get("ken_burns", "zoom_in")
        ZOOM   = self.ZOOM

        def smooth(x):
            x = max(0.0, min(1.0, x))
            return x*x*(3.0-2.0*x)
        t = smooth(t)

        # ── Build the same blurred-background composite as render_static ──────
        # 1. aspect-fit the image (same letterbox logic as static)
        if iw / ih > W / H:
            fw, fh = W, int(W * ih / iw)
        else:
            fh, fw = H, int(H * iw / ih)

        fitted = cv2.resize(arr, (fw, fh), interpolation=cv2.INTER_LINEAR)

        # 2. blurred full-frame background
        bg = cv2.resize(arr, (W, H), interpolation=cv2.INTER_LINEAR)
        bg = cv2.GaussianBlur(bg, (0, 0), 15)

        # 3. composite: blurred bg + 90%-scaled foreground centred
        composite = bg.copy()
        fg_w, fg_h = int(fw * 0.9), int(fh * 0.9)
        fg = cv2.resize(fitted, (fg_w, fg_h), interpolation=cv2.INTER_LINEAR)
        ox, oy = (W - fg_w) // 2, (H - fg_h) // 2
        composite[oy:oy + fg_h, ox:ox + fg_w] = fg

        # ── Apply Ken Burns zoom/pan on top of the full composite ─────────────
        # Scale the composite canvas to W*ZOOM x H*ZOOM so the animated crop
        # window never spills outside the image.
        cov_w = int(round(W * ZOOM))
        cov_h = int(round(H * ZOOM))
        canvas = cv2.resize(composite, (cov_w, cov_h), interpolation=cv2.INTER_LINEAR)

        # Travel clamped to actual canvas headroom per axis — vertical headroom
        # (cov_h-H)/2 is smaller than horizontal, causing pan_up/down to clamp
        # and appear frozen at the start without this.
        h_travel = min((cov_w - W) / 2.0, min(W, H) * 0.08)
        v_travel = min((cov_h - H) / 2.0, min(W, H) * 0.08)

        sw, sh = float(W), float(H)
        sx = (cov_w - W) / 2.0
        sy = (cov_h - H) / 2.0

        if effect == "zoom_in":
            z = ZOOM - (ZOOM - 1.0) * t        # ZOOM → 1.0
            sw, sh = W * z, H * z
            sx = (cov_w - sw) / 2.0
            sy = (cov_h - sh) / 2.0
        elif effect == "zoom_out":
            z = 1.0 + (ZOOM - 1.0) * t         # 1.0 → ZOOM
            sw, sh = W * z, H * z
            sx = (cov_w - sw) / 2.0
            sy = (cov_h - sh) / 2.0
        elif effect == "pan_left":
            sx = (cov_w - W) / 2.0 + h_travel * (1.0 - t)
            sy = (cov_h - H) / 2.0
        elif effect == "pan_right":
            sx = (cov_w - W) / 2.0 - h_travel * (1.0 - t)
            sy = (cov_h - H) / 2.0
        elif effect == "pan_up":
            sx = (cov_w - W) / 2.0
            sy = (cov_h - H) / 2.0 + v_travel * (1.0 - t)
        elif effect == "pan_down":
            sx = (cov_w - W) / 2.0
            sy = (cov_h - H) / 2.0 - v_travel * (1.0 - t)

        sx = max(0.0, min(float(cov_w) - sw, sx))
        sy = max(0.0, min(float(cov_h) - sh, sy))

        scale_x, scale_y = W / sw, H / sh
        M = np.array([[scale_x, 0, -sx * scale_x],
                      [0, scale_y, -sy * scale_y]], dtype=np.float64)
        frame = cv2.warpAffine(canvas, M, (W, H),
                               flags=cv2.INTER_LINEAR,
                               borderMode=cv2.BORDER_REFLECT_101)
        text = img_data.get("text", "")
        if text and text.strip():
            frame = self._draw_text(frame, text)
        return frame

    # ── transition blend ─────────────────────────────────────────────────────

    def render_transition(self, frame_a: "np.ndarray", frame_b: "np.ndarray",
                          t: float, transition: str) -> "np.ndarray":
        import numpy as np
        t = max(0.0, min(1.0, t))

        # ── Smoothstep easing — removes the mechanical stutter at start/end ──
        t_e = t * t * (3.0 - 2.0 * t)

        # ── Fade family ───────────────────────────────────────────────────────
        if transition in ("fade", "fadeblack", "fadewhite", "dissolve"):
            return (frame_a.astype(np.float32) * (1.0 - t_e) +
                    frame_b.astype(np.float32) * t_e).clip(0, 255).astype(np.uint8)

        # ── Wipe helpers ──────────────────────────────────────────────────────
        # FEATHER: blend a narrow band around the cut so the edge glides
        # rather than snapping — eliminates the visual stutter on hard cuts.
        FEATHER = max(4, min(24, self.W // 60))   # ~1.5% of width, min 4 px

        def _wipe_h(pos: int) -> np.ndarray:
            """Horizontal wipe: frame_b reveals from the left at pixel column pos."""
            # Fast path: before or after the feather zone
            lo = max(0, pos - FEATHER)
            hi = min(self.W, pos + FEATHER)
            if lo == 0 and hi == self.W:
                return frame_b
            # Build output: hard regions + blended strip
            parts = []
            if lo > 0:
                parts.append(frame_b[:, :lo])          # fully revealed
            if lo < hi:
                # Blend strip — ramp alpha from 1→0 (frame_b→frame_a)
                strip_w = hi - lo
                alpha = np.linspace(1.0, 0.0, strip_w, dtype=np.float32)
                alpha = alpha[np.newaxis, :, np.newaxis]   # (1, W, 1)
                b_strip = frame_b[:, lo:hi].astype(np.float32)
                a_strip = frame_a[:, lo:hi].astype(np.float32)
                blend = (b_strip * alpha + a_strip * (1.0 - alpha)).clip(0, 255).astype(np.uint8)
                parts.append(blend)
            if hi < self.W:
                parts.append(frame_a[:, hi:])           # not yet revealed
            return np.concatenate(parts, axis=1)

        def _wipe_v(pos: int) -> np.ndarray:
            """Vertical wipe: frame_b reveals from the top at pixel row pos."""
            lo = max(0, pos - FEATHER)
            hi = min(self.H, pos + FEATHER)
            if lo == 0 and hi == self.H:
                return frame_b
            parts = []
            if lo > 0:
                parts.append(frame_b[:lo, :])
            if lo < hi:
                strip_h = hi - lo
                alpha = np.linspace(1.0, 0.0, strip_h, dtype=np.float32)
                alpha = alpha[:, np.newaxis, np.newaxis]   # (H, 1, 1)
                b_strip = frame_b[lo:hi, :].astype(np.float32)
                a_strip = frame_a[lo:hi, :].astype(np.float32)
                blend = (b_strip * alpha + a_strip * (1.0 - alpha)).clip(0, 255).astype(np.uint8)
                parts.append(blend)
            if hi < self.H:
                parts.append(frame_a[hi:, :])
            return np.concatenate(parts, axis=0)

        # ── Directional wipes ─────────────────────────────────────────────────
        if transition == "wipeleft":
            # frame_b sweeps in from the left
            return _wipe_h(int(self.W * t_e))

        elif transition == "wiperight":
            # frame_b sweeps in from the right — mirror of wipeleft
            pos = int(self.W * (1.0 - t_e))
            lo = max(0, pos - FEATHER)
            hi = min(self.W, pos + FEATHER)
            parts = []
            if lo > 0:
                parts.append(frame_a[:, :lo])
            if lo < hi:
                strip_w = hi - lo
                alpha = np.linspace(0.0, 1.0, strip_w, dtype=np.float32)
                alpha = alpha[np.newaxis, :, np.newaxis]
                b_strip = frame_b[:, lo:hi].astype(np.float32)
                a_strip = frame_a[:, lo:hi].astype(np.float32)
                blend = (b_strip * alpha + a_strip * (1.0 - alpha)).clip(0, 255).astype(np.uint8)
                parts.append(blend)
            if hi < self.W:
                parts.append(frame_b[:, hi:])
            return np.concatenate(parts, axis=1)

        elif transition == "wipeup":
            return _wipe_v(int(self.H * t_e))

        elif transition == "wipedown":
            pos = int(self.H * (1.0 - t_e))
            lo = max(0, pos - FEATHER)
            hi = min(self.H, pos + FEATHER)
            parts = []
            if lo > 0:
                parts.append(frame_a[:lo, :])
            if lo < hi:
                strip_h = hi - lo
                alpha = np.linspace(0.0, 1.0, strip_h, dtype=np.float32)
                alpha = alpha[:, np.newaxis, np.newaxis]
                b_strip = frame_b[lo:hi, :].astype(np.float32)
                a_strip = frame_a[lo:hi, :].astype(np.float32)
                blend = (b_strip * alpha + a_strip * (1.0 - alpha)).clip(0, 255).astype(np.uint8)
                parts.append(blend)
            if hi < self.H:
                parts.append(frame_b[hi:, :])
            return np.concatenate(parts, axis=0)

        else:
            # All other transitions (slide*, smooth*, circle*, etc.) → fade
            return (frame_a.astype(np.float32) * (1.0 - t_e) +
                    frame_b.astype(np.float32) * t_e).clip(0, 255).astype(np.uint8)

    # ── text overlay ─────────────────────────────────────────────────────────

    def _draw_text(self, frame_arr: "np.ndarray", text: str) -> "np.ndarray":
        try:
            import numpy as np
            from bidi.algorithm import get_display as _bidi
            pil = Image.fromarray(frame_arr)
            draw = ImageDraw.Draw(pil)
            try:
                font_path = BASEPATH / "Fonts" / "Birzia-Black.otf"
                font = ImageFont.truetype(str(font_path), 42)
            except Exception:
                font = ImageFont.load_default()
            hebrew = _bidi(text)
            bbox = draw.textbbox((0,0), hebrew, font=font)
            tw, th = bbox[2]-bbox[0], bbox[3]-bbox[1]
            bg_w, bg_h = tw+30, th+16
            bg_x = (self.W-bg_w)//2
            bg_y = self.H - bg_h - 30
            draw.rounded_rectangle((bg_x, bg_y, bg_x+bg_w, bg_y+bg_h),
                                   radius=8, fill="white")
            draw.text(((self.W-tw)//2, bg_y-2), hebrew, font=font, fill="black")
            return np.array(pil)
        except Exception:
            return frame_arr

    # ── main entry point ─────────────────────────────────────────────────────

    def get_frame(self, global_t: float) -> "np.ndarray":
        """
        Compute which slide + transition we're in and render the frame.
        global_t is time in seconds from the start.
        """
        if not self.images:
            import numpy as np
            return np.zeros((self.H, self.W, 3), dtype=np.uint8)

        # Timeline matches the ffmpeg export exactly: cursor advances by full dur.
        # This keeps get_frame in sync with _slide_at, _slide_start, and the
        # marker bar (all use simple sum-of-durations).
        #
        # KB continuity: when slide i is entered after a transition, its KB
        # animation should already be td seconds in (not starting from 0).
        # We achieve this by passing t_in + td_prev to _render_slide, where
        # td_prev is the incoming transition duration. During the transition we
        # pass next_t_in = t_trans * td, so the animation is perfectly continuous
        # at the moment the transition ends (next_t_in -> td = t_in_kb at t=0).
        cursor = 0.0
        for i, img in enumerate(self.images):
            dur = float(img.get("duration", 5))
            td  = float(img.get("transition_duration", 1)) if i < len(self.images)-1 else 0.0
            slide_end = cursor + dur

            if global_t < slide_end or i == len(self.images)-1:
                t_in = global_t - cursor  # raw seconds since this slide's cursor

                # KB offset: slide i was already rendered for td_prev seconds during
                # the incoming transition, so KB starts from td_prev not 0.
                td_prev = float(self.images[i-1].get("transition_duration", 1)) if i > 0 else 0.0
                t_in_kb = t_in + td_prev

                # Are we in the outgoing transition of this slide?
                trans_start = dur - td
                if td > 0 and t_in >= trans_start and i < len(self.images)-1:
                    t_trans = (t_in - trans_start) / td   # 0->1 over the transition
                    fa = self._render_slide(img, t_in_kb, dur + td_prev)
                    next_img = self.images[i + 1]
                    next_dur = float(next_img.get("duration", 5))
                    # next slide KB: 0 at trans start, td at trans end —
                    # matches t_in_kb = td_prev = td when it becomes active.
                    next_t_in_kb = t_trans * td
                    fb = self._render_slide(next_img, next_t_in_kb, next_dur + td)
                    transition = img.get("transition", "fade")
                    return self.render_transition(fa, fb, t_trans, transition)

                return self._render_slide(img, t_in_kb, dur + td_prev)

            cursor += dur   # matches export timeline and all helper functions

        import numpy as np
        return np.zeros((self.H, self.W, 3), dtype=np.uint8)

    def _render_slide(self, img_data: dict, t_in: float, dur: float) -> "np.ndarray":
        effect = img_data.get("ken_burns", "none")
        if effect and effect != "none":
            t_norm = max(0.0, min(1.0, t_in / max(dur, 0.001)))
            return self.render_kb_frame(img_data, t_norm)
        else:
            return self.render_static(img_data)

    @property
    def total_duration(self) -> float:
        return sum(float(img.get("duration", 5)) for img in self.images)


class _PreviewRenderThread(QThread):
    """
    Background thread that pre-renders frames around the current playhead.
    Signals frame_ready with (global_t_ms, QImage).
    """
    frame_ready = pyqtSignal(float, object)  # (t_seconds, QImage)

    def __init__(self, renderer: _FrameRenderer, parent=None):
        super().__init__(parent)
        self._renderer = renderer
        self._queue: list = []
        self._lock = __import__("threading").Lock()
        self._stop = False
        self._cond = __import__("threading").Condition(self._lock)

    def request_frame(self, t: float):
        with self._cond:
            # Drop stale requests, keep only latest
            self._queue = [t]
            self._cond.notify()

    def stop(self):
        with self._cond:
            self._stop = True
            self._cond.notify()

    def run(self):
        import numpy as np
        while True:
            with self._cond:
                while not self._queue and not self._stop:
                    self._cond.wait(timeout=0.1)
                if self._stop:
                    break
                t = self._queue.pop(0) if self._queue else None

            if t is None:
                continue
            try:
                arr = self._renderer.get_frame(t)
                h, w = arr.shape[:2]
                qimg = QImage(arr.tobytes(), w, h, w*3, QImage.Format_RGB888).copy()
                self.frame_ready.emit(t, qimg)
            except Exception as e:
                print(f"Preview render error: {e}")


class SlideshowPreviewDialog(QDialog):
    """
    Full-window slideshow preview with seek, play/pause, audio and time display.
    Opens instantly — no export required.
    """

    FPS    = 25
    W, H   = 960, 540
    FRAME_MS = int(1000 / FPS)

    def __init__(self, images: list, audio_files: list, parent=None):
        super().__init__(parent)
        self.setWindowTitle("▶  Slideshow Preview")
        self.setMinimumSize(1050, 720)
        self.resize(1100, 760)
        self.setWindowFlags(self.windowFlags() | Qt.WindowMaximizeButtonHint)
        self.setStyleSheet(f"background: {COLORS['bg_deep']}; color: {COLORS['text_primary']};")

        self._images      = images
        self._audio_files = audio_files
        self._renderer    = _FrameRenderer(images)
        self._total_dur   = self._renderer.total_duration

        # State
        self._playing      = False
        self._current_t    = 0.0          # playhead in seconds
        self._current_qimg: QImage | None = None
        self._audio_proc: QProcess | None = None
        self._audio_offset = 0.0          # what second audio started from

        self._build_ui()

        # Background render thread
        self._render_thread = _PreviewRenderThread(self._renderer, self)
        self._render_thread.frame_ready.connect(self._on_frame_ready)
        self._render_thread.start()

        # Playback timer
        self._timer = QTimer(self)
        self._timer.setInterval(self.FRAME_MS)
        self._timer.timeout.connect(self._tick)

        # Render first frame
        self._seek(0.0, start_audio=False)

    # ── UI construction ──────────────────────────────────────────────────────

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # ── Canvas ────────────────────────────────────────────────────────────
        self._canvas = QLabel()
        self._canvas.setAlignment(Qt.AlignCenter)
        self._canvas.setStyleSheet(f"background: #000;")
        self._canvas.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        root.addWidget(self._canvas, 1)

        # ── Slide indicator strip ─────────────────────────────────────────────
        self._slide_label = QLabel("Slide 1 / 1")
        self._slide_label.setAlignment(Qt.AlignCenter)
        self._slide_label.setStyleSheet(
            f"color: {COLORS['text_muted']}; font-size: 11px; "
            f"background: {COLORS['bg_deep']}; padding: 3px;"
        )
        root.addWidget(self._slide_label)

        # ── Timeline scrubber ─────────────────────────────────────────────────
        scrub_row = QHBoxLayout()
        scrub_row.setContentsMargins(16, 4, 16, 2)
        scrub_row.setSpacing(10)

        self._time_label = QLabel("00:00:00")
        self._time_label.setStyleSheet(
            f"color: {COLORS['accent']}; font-size: 12px; font-weight: 700; "
            f"font-family: 'Consolas', monospace; min-width: 68px;"
        )
        self._dur_label = QLabel(f"/ {format_time_hms(self._total_dur)}")
        self._dur_label.setStyleSheet(
            f"color: {COLORS['text_muted']}; font-size: 12px; min-width: 68px;"
        )

        self._scrubber = QSlider(Qt.Horizontal)
        self._scrubber.setRange(0, max(1, int(self._total_dur * 100)))
        self._scrubber.setValue(0)
        self._scrubber.setStyleSheet(f"""
            QSlider::groove:horizontal {{
                height: 6px; background: {COLORS['bg_hover']};
                border-radius: 3px;
            }}
            QSlider::sub-page:horizontal {{
                background: {COLORS['accent']}; border-radius: 3px;
            }}
            QSlider::handle:horizontal {{
                background: #fff; border: 2px solid {COLORS['accent']};
                width: 14px; height: 14px; margin: -5px 0; border-radius: 7px;
            }}
        """)
        self._scrubber.sliderMoved.connect(self._on_scrub)
        self._scrubber.sliderPressed.connect(self._on_scrub_start)
        self._scrubber.sliderReleased.connect(self._on_scrub_end)
        self._scrubbing = False

        scrub_row.addWidget(self._time_label)
        scrub_row.addWidget(self._dur_label)
        scrub_row.addWidget(self._scrubber, 1)
        root.addLayout(scrub_row)

        # ── Slide markers on timeline ─────────────────────────────────────────
        self._marker_bar = _SlideMarkerBar(self._images, self._total_dur)
        self._marker_bar.seek_to.connect(lambda t: self._seek(t))
        root.addWidget(self._marker_bar)

        # ── Transport controls ────────────────────────────────────────────────
        ctrl = QHBoxLayout()
        ctrl.setContentsMargins(16, 6, 16, 12)
        ctrl.setSpacing(10)

        def _ctrl_btn(text, tip="", accent=False):
            b = QPushButton(text)
            b.setFixedSize(42, 36)
            color = COLORS["accent"] if accent else COLORS["text_secondary"]
            b.setStyleSheet(
                f"QPushButton {{ background: {COLORS['bg_card']}; color: {color}; "
                f"border: 1px solid {COLORS['border']}; border-radius: 6px; "
                f"font-size: 15px; font-weight: 700; }}"
                f"QPushButton:hover {{ background: {COLORS['bg_hover']}; }}"
            )
            if tip: b.setToolTip(tip)
            return b

        self._btn_prev  = _ctrl_btn("⏮", "Previous slide")
        self._btn_back  = _ctrl_btn("◀◀", "Back 5 s")
        self._btn_play  = _ctrl_btn("▶", "Play / Pause  (Space)", accent=True)
        self._btn_fwd   = _ctrl_btn("▶▶", "Forward 5 s")
        self._btn_next  = _ctrl_btn("⏭", "Next slide")

        self._btn_play.setFixedSize(52, 40)

        self._btn_prev.clicked.connect(self._prev_slide)
        self._btn_back.clicked.connect(lambda: self._seek(max(0, self._current_t - 5)))
        self._btn_play.clicked.connect(self._toggle_play)
        self._btn_fwd.clicked.connect(lambda: self._seek(min(self._total_dur, self._current_t + 5)))
        self._btn_next.clicked.connect(self._next_slide)

        # Speed selector
        speed_lbl = QLabel("Speed:")
        speed_lbl.setStyleSheet(f"color: {COLORS['text_muted']}; font-size: 11px;")
        self._speed_box = QComboBox()
        self._speed_box.addItems(["0.5×", "1×", "1.5×", "2×"])
        self._speed_box.setCurrentIndex(1)
        self._speed_box.setFixedWidth(68)
        self._speed_box.setStyleSheet(f"background: {COLORS['bg_card']}; color: {COLORS['text_primary']}; border-radius: 4px;")
        self._speed_box.currentIndexChanged.connect(self._on_speed_change)
        self._speed = 1.0

        close_btn = QPushButton("✕  Close")
        close_btn.setFixedHeight(36)
        close_btn.setStyleSheet(
            f"QPushButton {{ background: {COLORS['bg_card']}; color: {COLORS['text_muted']}; "
            f"border: 1px solid {COLORS['border']}; border-radius: 6px; padding: 0 16px; }}"
            f"QPushButton:hover {{ background: {COLORS['danger']}; color: #fff; }}"
        )
        close_btn.clicked.connect(self.close)

        # Volume control
        vol_lbl = QLabel("🔊")
        vol_lbl.setStyleSheet(f"color: {COLORS['text_muted']}; font-size: 13px;")
        self._vol_slider = QSlider(Qt.Horizontal)
        self._vol_slider.setRange(0, 100)
        self._vol_slider.setValue(45)
        self._vol_slider.setFixedWidth(90)
        self._vol_slider.setToolTip("Music volume")
        self._vol_slider.setStyleSheet(f"""
            QSlider::groove:horizontal {{
                height: 4px; background: {COLORS['bg_hover']}; border-radius: 2px;
            }}
            QSlider::sub-page:horizontal {{
                background: {COLORS['success']}; border-radius: 2px;
            }}
            QSlider::handle:horizontal {{
                background: #fff; border: 2px solid {COLORS['success']};
                width: 12px; height: 12px; margin: -5px 0; border-radius: 6px;
            }}
        """)
        self._vol_slider.valueChanged.connect(self._on_volume_change)

        ctrl.addWidget(self._btn_prev)
        ctrl.addWidget(self._btn_back)
        ctrl.addWidget(self._btn_play)
        ctrl.addWidget(self._btn_fwd)
        ctrl.addWidget(self._btn_next)
        ctrl.addSpacing(16)
        ctrl.addWidget(speed_lbl)
        ctrl.addWidget(self._speed_box)
        ctrl.addSpacing(16)
        ctrl.addWidget(vol_lbl)
        ctrl.addWidget(self._vol_slider)
        ctrl.addStretch()
        ctrl.addWidget(close_btn)
        root.addLayout(ctrl)

    # ── Keyboard shortcut ─────────────────────────────────────────────────────

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Space:
            self._toggle_play()
        elif event.key() == Qt.Key_Left:
            self._seek(max(0, self._current_t - 5))
        elif event.key() == Qt.Key_Right:
            self._seek(min(self._total_dur, self._current_t + 5))
        elif event.key() == Qt.Key_Escape:
            self.close()
        else:
            super().keyPressEvent(event)

    # ── Playback ─────────────────────────────────────────────────────────────

    def _toggle_play(self):
        if self._playing:
            self._pause()
        else:
            self._play()

    def _play(self):
        if self._current_t >= self._total_dur:
            self._seek(0.0, start_audio=False)
        self._playing = True
        self._btn_play.setText("⏸")
        self._playback_start_wall = __import__("time").perf_counter()
        self._playback_start_t    = self._current_t
        self._start_audio(self._current_t)
        self._timer.start()

    def _pause(self):
        self._playing = False
        self._btn_play.setText("▶")
        self._timer.stop()
        self._stop_audio()

    def _tick(self):
        import time
        elapsed = (time.perf_counter() - self._playback_start_wall) * self._speed
        t = self._playback_start_t + elapsed
        if t >= self._total_dur:
            t = self._total_dur
            self._pause()
        self._current_t = t
        self._update_ui_position(t)
        self._render_thread.request_frame(t)

    # ── Seeking ───────────────────────────────────────────────────────────────

    def _seek(self, t: float, start_audio: bool = True):
        was_playing = self._playing
        if self._playing:
            self._pause()
        t = max(0.0, min(self._total_dur, t))
        self._current_t = t
        self._update_ui_position(t)
        # Render synchronously for instant feedback (cheap at half-res)
        try:
            import numpy as np
            arr = self._renderer.get_frame(t)
            h, w = arr.shape[:2]
            self._current_qimg = QImage(arr.tobytes(), w, h, w*3,
                                        QImage.Format_RGB888).copy()
            self._paint_frame()
        except Exception as e:
            print(f"Seek render error: {e}")

        if was_playing and start_audio:
            self._play()
        elif start_audio and not was_playing:
            pass  # keep paused

    def _update_ui_position(self, t: float):
        # scrubber (block signals to avoid recursive seek)
        self._scrubber.blockSignals(True)
        self._scrubber.setValue(int(t * 100))
        self._scrubber.blockSignals(False)
        self._time_label.setText(format_time_hms(t))
        self._marker_bar.set_playhead(t)
        # slide indicator
        idx, _ = self._slide_at(t)
        self._slide_label.setText(f"Slide {idx+1} / {len(self._images)}"
                                   + (f"  —  {os.path.basename(self._images[idx]['path'])}"
                                      if self._images else ""))

    def _slide_at(self, t: float) -> tuple:
        """Return (slide_index, t_within_slide)."""
        cursor = 0.0
        for i, img in enumerate(self._images):
            dur = float(img.get("duration", 5))
            if t < cursor + dur or i == len(self._images) - 1:
                return i, t - cursor
            cursor += dur
        return 0, 0.0

    def _slide_start(self, idx: int) -> float:
        return sum(float(img.get("duration", 5)) for img in self._images[:idx])

    def _prev_slide(self):
        idx, t_in = self._slide_at(self._current_t)
        if t_in > 0.5:
            self._seek(self._slide_start(idx))
        else:
            self._seek(self._slide_start(max(0, idx-1)))

    def _next_slide(self):
        idx, _ = self._slide_at(self._current_t)
        self._seek(self._slide_start(min(len(self._images)-1, idx+1)))

    # ── Scrubber ──────────────────────────────────────────────────────────────

    def _on_scrub_start(self):
        self._scrubbing = True
        if self._playing:
            self._pause()

    def _on_scrub(self, value: int):
        t = value / 100.0
        self._current_t = t
        self._update_ui_position(t)
        self._render_thread.request_frame(t)

    def _on_scrub_end(self):
        self._scrubbing = False
        t = self._scrubber.value() / 100.0
        self._seek(t, start_audio=False)

    # ── Speed ─────────────────────────────────────────────────────────────────

    def _on_volume_change(self, value: int):
        """Apply volume (0-100) to the active QMediaPlayer."""
        if self._audio_proc is not None:
            try:
                self._audio_proc.setVolume(value)
            except Exception:
                pass

    def _on_speed_change(self, idx: int):
        speeds = [0.5, 1.0, 1.5, 2.0]
        self._speed = speeds[idx]
        if self._playing:
            # Restart timing reference
            self._playback_start_wall = __import__("time").perf_counter()
            self._playback_start_t    = self._current_t

    # ── Frame rendering ───────────────────────────────────────────────────────

    def _on_frame_ready(self, t: float, qimg: QImage):
        # Only paint if this frame is still relevant (within 200ms of playhead)
        if abs(t - self._current_t) < 0.2 or self._playing:
            self._current_qimg = qimg
            self._paint_frame()

    def _paint_frame(self):
        if self._current_qimg is None:
            return
        cw, ch = self._canvas.width(), self._canvas.height()
        pix = QPixmap.fromImage(self._current_qimg).scaled(
            cw, ch, Qt.KeepAspectRatio, Qt.SmoothTransformation
        )
        self._canvas.setPixmap(pix)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._paint_frame()

    # ── Audio via ffmpeg pipe → system audio ─────────────────────────────────

    def _start_audio(self, offset: float):
        """Play audio starting from *offset* seconds using QMediaPlayer.
        All audio files are played in sequence; offset is applied across the
        concatenated timeline so seeking mid-playlist works correctly.
        """
        self._stop_audio()
        if not self._audio_files:
            return
        try:
            from PyQt5.QtMultimedia import QMediaPlayer, QMediaContent

            paths = [a["path"] for a in self._audio_files
                     if os.path.exists(str(a.get("path", "")))]
            if not paths:
                return

            # Pre-compute per-file durations so we can find which file
            # contains the requested offset.
            durations = [_get_audio_duration(p) for p in paths]

            # Walk the playlist to find which file the offset falls in.
            remaining = offset
            start_index = 0
            file_offset = 0.0
            for i, dur in enumerate(durations):
                if remaining < dur or i == len(durations) - 1:
                    start_index = i
                    file_offset = remaining
                    break
                remaining -= dur

            self._audio_paths    = paths
            self._audio_durations = durations
            self._audio_index    = start_index
            self._audio_offset   = offset

            self._audio_proc = QMediaPlayer(self)
            self._audio_file_offset = int(file_offset * 1000)  # seek after load
            self._audio_proc.mediaStatusChanged.connect(self._on_audio_status)
            # Apply current volume immediately
            vol = self._vol_slider.value() if hasattr(self, "_vol_slider") else 100
            self._audio_proc.setVolume(vol)
            url = QUrl.fromLocalFile(str(paths[start_index]))
            self._audio_proc.setMedia(QMediaContent(url))
            self._audio_proc.setPlaybackRate(self._speed)
            # Don't call play() here – seek + play happen in _on_audio_status
            # once the media is loaded (avoids the setPosition-before-load bug)
        except Exception as e:
            print(f"Audio preview not available: {e}")
            self._audio_proc = None

    def _on_audio_status(self, status):
        """Seek-then-play on first load; advance playlist on EndOfMedia."""
        try:
            from PyQt5.QtMultimedia import QMediaPlayer, QMediaContent
            if self._audio_proc is None:
                return

            if status == QMediaPlayer.LoadedMedia or status == QMediaPlayer.BufferedMedia:
                # Media is ready – apply the pending seek offset and start playing.
                # We only do this once per load (offset is reset to 0 after use).
                offset_ms = getattr(self, "_audio_file_offset", 0)
                if offset_ms > 0:
                    self._audio_proc.setPosition(offset_ms)
                    self._audio_file_offset = 0
                # Guard: don't auto-play if we're paused (seek while paused)
                if self._playing:
                    self._audio_proc.play()

            elif status == QMediaPlayer.EndOfMedia:
                if not hasattr(self, "_audio_paths"):
                    return
                self._audio_index += 1
                if self._audio_index >= len(self._audio_paths):
                    return  # all songs done
                next_path = self._audio_paths[self._audio_index]
                self._audio_file_offset = 0  # next file starts from beginning
                url = QUrl.fromLocalFile(str(next_path))
                self._audio_proc.setMedia(QMediaContent(url))
                self._audio_proc.setPlaybackRate(self._speed)
                # play() will be called by the LoadedMedia branch above
        except Exception as e:
            print(f"Audio status error: {e}")

    def _stop_audio(self):
        if self._audio_proc is not None:
            try:
                self._audio_proc.stop()
            except Exception:
                pass
            self._audio_proc = None

    # ── Cleanup ───────────────────────────────────────────────────────────────

    def closeEvent(self, event):
        self._timer.stop()
        self._stop_audio()
        self._render_thread.stop()
        self._render_thread.wait(2000)
        super().closeEvent(event)


class _SlideMarkerBar(QWidget):
    """
    Thin bar below the scrubber showing slide boundaries and the current
    playhead. Click anywhere to seek.
    """
    seek_to = pyqtSignal(float)

    def __init__(self, images: list, total_dur: float, parent=None):
        super().__init__(parent)
        self._images    = images
        self._total_dur = max(total_dur, 1.0)
        self._playhead  = 0.0
        self.setFixedHeight(18)
        self.setCursor(Qt.PointingHandCursor)
        # Pre-compute slide start times (simple sum, matches get_frame cursor).
        self._starts: list[float] = []
        t = 0.0
        for img in images:
            self._starts.append(t)
            t += float(img.get("duration", 5))

    def set_playhead(self, t: float):
        self._playhead = t
        self.update()

    def paintEvent(self, _):
        p = QPainter(self)
        p.fillRect(self.rect(), QColor(COLORS["bg_deep"]))
        W = self.width()

        # slide regions alternating tint
        for i, st in enumerate(self._starts):
            x1 = int(st / self._total_dur * W)
            end = self._starts[i+1] if i+1 < len(self._starts) else self._total_dur
            x2 = int(end / self._total_dur * W)
            color = QColor(COLORS["bg_card"]) if i % 2 == 0 else QColor(COLORS["bg_hover"])
            p.fillRect(x1, 0, x2-x1, self.height(), color)
            # tick mark
            p.setPen(QPen(QColor(COLORS["border_light"]), 1))
            p.drawLine(x1, 0, x1, self.height())

        # playhead
        px = int(self._playhead / self._total_dur * W)
        p.setPen(QPen(QColor(COLORS["accent"]), 2))
        p.drawLine(px, 0, px, self.height())
        p.end()

    def mousePressEvent(self, event):
        t = event.x() / max(self.width(), 1) * self._total_dur
        self.seek_to.emit(max(0.0, min(self._total_dur, t)))

    def mouseMoveEvent(self, event):
        if event.buttons() & Qt.LeftButton:
            t = event.x() / max(self.width(), 1) * self._total_dur
            self.seek_to.emit(max(0.0, min(self._total_dur, t)))


# ── Crop Dialog ───────────────────────────────────────────────────────────────

class CropCanvas(QWidget):
    """
    Displays an image and lets the user drag a crop rectangle over it.
    The crop rect is stored internally in *pixel* coordinates relative to the
    displayed (scaled) image, but converted to/from normalised 0-1 coords for
    the outside world so it is resolution-independent.
    """
    crop_changed = pyqtSignal()

    _HANDLE  = 10          # handle square half-size in px
    _MIN_DIM = 20          # minimum crop size in display pixels

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setMouseTracking(True)
        self.setCursor(Qt.CrossCursor)

        self._pixmap:   QPixmap | None = None   # scaled-to-widget pixmap
        self._img_rect  = None                  # QRect: where the image sits inside the widget
        self._crop_rect = None                  # QRect: current crop in image-space pixels
        self._orig_w = 0
        self._orig_h = 0

        # drag state
        self._drag_mode   = None   # "new" | "move" | "tl"|"tr"|"bl"|"br"|"t"|"b"|"l"|"r"
        self._drag_origin = None
        self._drag_rect_start = None

    # ── public API ────────────────────────────────────────────────────────────

    def load_image(self, path: str, rotation: int = 0,
                   norm_crop: tuple | None = None):
        """Load *path*, apply EXIF + manual rotation, set initial crop."""
        try:
            img = Image.open(path)
            # EXIF
            try:
                exif = img._getexif() if hasattr(img, "_getexif") else None
                _tag = next((k for k, v in ExifTags.TAGS.items() if v == "Orientation"), None)
                if exif and _tag:
                    v = exif.get(_tag)
                    if v == 3:   img = img.rotate(180, expand=True)
                    elif v == 6: img = img.rotate(270, expand=True)
                    elif v == 8: img = img.rotate(90,  expand=True)
            except Exception:
                pass
            if rotation:
                img = img.rotate(rotation, expand=True)
            img = img.convert("RGB")
            self._orig_w, self._orig_h = img.size

            # Convert to QPixmap
            data  = img.tobytes("raw", "RGB")
            qimg  = QImage(data, img.width, img.height, img.width * 3, QImage.Format_RGB888)
            self._pixmap = QPixmap.fromImage(qimg)
        except Exception as e:
            print(f"CropCanvas load error: {e}")
            self._pixmap = None
            return

        if norm_crop:
            x  = int(norm_crop[0] * self._orig_w)
            y  = int(norm_crop[1] * self._orig_h)
            w  = int(norm_crop[2] * self._orig_w)
            h  = int(norm_crop[3] * self._orig_h)
            self._crop_rect = QRect(x, y, w, h)
        else:
            self._crop_rect = QRect(0, 0, self._orig_w, self._orig_h)

        self._layout_image()
        self.update()

    def get_norm_crop(self) -> tuple | None:
        """Return (x, y, w, h) in 0-1 coords, or None if no image."""
        if self._crop_rect is None or self._orig_w == 0:
            return None
        r = self._crop_rect.normalized()
        return (
            max(0.0, r.x()      / self._orig_w),
            max(0.0, r.y()      / self._orig_h),
            min(1.0, r.width()  / self._orig_w),
            min(1.0, r.height() / self._orig_h),
        )

    def reset_crop(self):
        if self._orig_w:
            self._crop_rect = QRect(0, 0, self._orig_w, self._orig_h)
            self.crop_changed.emit()
            self.update()

    # ── layout ────────────────────────────────────────────────────────────────

    def _layout_image(self):
        if not self._pixmap:
            return
        pw, ph = self.width(), self.height()
        iw, ih = self._pixmap.width(), self._pixmap.height()
        scale   = min(pw / iw, ph / ih)
        dw, dh  = int(iw * scale), int(ih * scale)
        ox, oy  = (pw - dw) // 2, (ph - dh) // 2
        self._img_rect   = QRect(ox, oy, dw, dh)
        self._scale      = scale   # orig → display

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._layout_image()
        self.update()

    # ── coordinate helpers ────────────────────────────────────────────────────

    def _to_display(self, orig_pt):
        """QPoint in orig-image space → display space."""
        r = self._img_rect
        return QPoint(int(r.x() + orig_pt.x() * self._scale),
                      int(r.y() + orig_pt.y() * self._scale))

    def _to_orig(self, disp_pt):
        """QPoint in display space → orig-image space (clamped)."""
        r = self._img_rect
        x = (disp_pt.x() - r.x()) / self._scale
        y = (disp_pt.y() - r.y()) / self._scale
        x = max(0, min(self._orig_w, x))
        y = max(0, min(self._orig_h, y))
        return QPoint(int(x), int(y))

    def _display_crop(self) -> QRect | None:
        if self._crop_rect is None or self._img_rect is None:
            return None
        c  = self._crop_rect.normalized()
        tl = self._to_display(c.topLeft())
        br = self._to_display(c.bottomRight())
        return QRect(tl, br)

    def _handle_rects(self, dc: QRect) -> dict:
        h = self._HANDLE
        cx = dc.center().x();  cy = dc.center().y()
        return {
            "tl": QRect(dc.left()  - h, dc.top()    - h, h*2, h*2),
            "tr": QRect(dc.right() - h, dc.top()    - h, h*2, h*2),
            "bl": QRect(dc.left()  - h, dc.bottom() - h, h*2, h*2),
            "br": QRect(dc.right() - h, dc.bottom() - h, h*2, h*2),
            "t":  QRect(cx - h,         dc.top()    - h, h*2, h*2),
            "b":  QRect(cx - h,         dc.bottom() - h, h*2, h*2),
            "l":  QRect(dc.left()  - h, cy          - h, h*2, h*2),
            "r":  QRect(dc.right() - h, cy          - h, h*2, h*2),
        }

    def _hit_test(self, pos) -> str | None:
        dc = self._display_crop()
        if dc is None:
            return None
        for name, rect in self._handle_rects(dc).items():
            if rect.contains(pos):
                return name
        if dc.contains(pos):
            return "move"
        return "new"

    # ── painting ──────────────────────────────────────────────────────────────

    def paintEvent(self, _):
        p = QPainter(self)
        p.setRenderHint(QPainter.SmoothPixmapTransform)
        p.fillRect(self.rect(), QColor(COLORS["bg_deep"]))

        if not self._pixmap or not self._img_rect:
            p.setPen(QColor(COLORS["text_muted"]))
            p.drawText(self.rect(), Qt.AlignCenter, "No image")
            p.end()
            return

        p.drawPixmap(self._img_rect, self._pixmap)

        dc = self._display_crop()
        if dc is None:
            p.end()
            return

        # Dim outside crop
        outer = self._img_rect
        dim = QColor(0, 0, 0, 140)
        p.fillRect(QRect(outer.left(), outer.top(),  outer.width(), dc.top() - outer.top()),   dim)
        p.fillRect(QRect(outer.left(), dc.bottom(), outer.width(), outer.bottom() - dc.bottom()), dim)
        p.fillRect(QRect(outer.left(), dc.top(),    dc.left() - outer.left(), dc.height()),    dim)
        p.fillRect(QRect(dc.right(),   dc.top(),    outer.right() - dc.right(), dc.height()), dim)

        # Crop border
        pen = QPen(QColor(COLORS["accent"]), 2)
        pen.setStyle(Qt.SolidLine)
        p.setPen(pen)
        p.setBrush(Qt.NoBrush)
        p.drawRect(dc)

        # Rule-of-thirds grid (subtle)
        grid_pen = QPen(QColor(255, 255, 255, 55), 1, Qt.DashLine)
        p.setPen(grid_pen)
        for frac in (1/3, 2/3):
            x = int(dc.left() + dc.width()  * frac)
            y = int(dc.top()  + dc.height() * frac)
            p.drawLine(x, dc.top(), x, dc.bottom())
            p.drawLine(dc.left(), y, dc.right(), y)

        # Handles
        p.setPen(QPen(QColor(COLORS["accent"]), 1))
        p.setBrush(QBrush(QColor("#FFFFFF")))
        for rect in self._handle_rects(dc).values():
            p.drawEllipse(rect)

        p.end()

    # ── mouse ─────────────────────────────────────────────────────────────────

    def mousePressEvent(self, event):
        if event.button() != Qt.LeftButton or not self._pixmap:
            return
        mode = self._hit_test(event.pos())
        self._drag_mode   = mode
        self._drag_origin = event.pos()
        self._drag_rect_start = QRect(self._crop_rect) if self._crop_rect else None
        if mode == "new":
            orig = self._to_orig(event.pos())
            self._crop_rect = QRect(orig, orig)
        self.update()

    def mouseMoveEvent(self, event):
        if not self._pixmap:
            return

        # Cursor shape on hover
        if not (event.buttons() & Qt.LeftButton):
            hit = self._hit_test(event.pos())
            cursors = {
                "move": Qt.SizeAllCursor,
                "tl": Qt.SizeFDiagCursor, "br": Qt.SizeFDiagCursor,
                "tr": Qt.SizeBDiagCursor, "bl": Qt.SizeBDiagCursor,
                "t":  Qt.SizeVerCursor,   "b":  Qt.SizeVerCursor,
                "l":  Qt.SizeHorCursor,   "r":  Qt.SizeHorCursor,
                "new": Qt.CrossCursor,
            }
            self.setCursor(cursors.get(hit, Qt.CrossCursor))
            return

        if self._drag_mode is None or self._drag_rect_start is None:
            return

        dx = int((event.x() - self._drag_origin.x()) / self._scale)
        dy = int((event.y() - self._drag_origin.y()) / self._scale)
        r  = QRect(self._drag_rect_start)

        def _clamp(v, lo, hi): return max(lo, min(hi, v))
        W, H = self._orig_w, self._orig_h

        if self._drag_mode == "new":
            orig = self._to_orig(event.pos())
            start = self._to_orig(self._drag_origin)
            self._crop_rect = QRect(start, orig).normalized()

        elif self._drag_mode == "move":
            nx = _clamp(r.x() + dx, 0, W - r.width())
            ny = _clamp(r.y() + dy, 0, H - r.height())
            self._crop_rect = QRect(nx, ny, r.width(), r.height())

        else:
            x1, y1, x2, y2 = r.left(), r.top(), r.right(), r.bottom()
            if "l" in self._drag_mode:
                x1 = _clamp(r.left() + dx, 0, x2 - self._MIN_DIM)
            if "r" in self._drag_mode:
                x2 = _clamp(r.right() + dx, x1 + self._MIN_DIM, W)
            if "t" in self._drag_mode:
                y1 = _clamp(r.top()  + dy, 0, y2 - self._MIN_DIM)
            if "b" in self._drag_mode:
                y2 = _clamp(r.bottom() + dy, y1 + self._MIN_DIM, H)
            self._crop_rect = QRect(QPoint(x1, y1), QPoint(x2, y2))

        self.crop_changed.emit()
        self.update()

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.LeftButton:
            if self._crop_rect:
                self._crop_rect = self._crop_rect.normalized()
            self._drag_mode = None
            self.crop_changed.emit()
            self.update()


class CropDialog(QDialog):
    """
    Full crop editor.  Opens for one image at a time; returns the normalised
    crop tuple (x, y, w, h) via get_result() after exec_() == Accepted.
    """

    def __init__(self, image_path: str, rotation: int = 0,
                 existing_crop: tuple | None = None, parent=None):
        super().__init__(parent)
        self.setWindowTitle("✂  Crop Image")
        self.setMinimumSize(800, 600)
        self.resize(1000, 700)
        self.setWindowFlags(self.windowFlags() | Qt.WindowMaximizeButtonHint)
        self._result: tuple | None = None

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # ── Header bar ────────────────────────────────────────────────────────
        header = QWidget()
        header.setFixedHeight(48)
        header.setStyleSheet(
            f"background: {COLORS['toolbar_bg']}; border-bottom: 1px solid {COLORS['border']};"
        )
        hl = QHBoxLayout(header)
        hl.setContentsMargins(16, 0, 16, 0)
        hl.setSpacing(12)

        title_lbl = QLabel("✂  Crop Image")
        title_lbl.setStyleSheet(
            f"color: {COLORS['text_primary']}; font-size: 14px; font-weight: 700;"
        )
        hint_lbl = QLabel(
            "Drag to draw a new crop  •  Drag edges/corners to resize  •  Drag inside to move"
        )
        hint_lbl.setStyleSheet(f"color: {COLORS['text_muted']}; font-size: 11px;")

        self._info_lbl = QLabel()
        self._info_lbl.setStyleSheet(
            f"color: {COLORS['accent']}; font-size: 12px; font-weight: 600; min-width: 220px;"
        )
        self._info_lbl.setAlignment(Qt.AlignRight | Qt.AlignVCenter)

        hl.addWidget(title_lbl)
        hl.addWidget(hint_lbl)
        hl.addStretch()
        hl.addWidget(self._info_lbl)
        root.addWidget(header)

        # ── Canvas ────────────────────────────────────────────────────────────
        self._canvas = CropCanvas()
        self._canvas.crop_changed.connect(self._update_info)
        root.addWidget(self._canvas, 1)

        # ── Footer bar ────────────────────────────────────────────────────────
        footer = QWidget()
        footer.setFixedHeight(52)
        footer.setStyleSheet(
            f"background: {COLORS['toolbar_bg']}; border-top: 1px solid {COLORS['border']};"
        )
        fl = QHBoxLayout(footer)
        fl.setContentsMargins(16, 0, 16, 0)
        fl.setSpacing(10)

        reset_btn  = _styled_btn("↺  Reset Crop", "")
        cancel_btn = _styled_btn("Cancel", "")
        apply_btn  = _styled_btn("✔  Apply Crop", "primary")
        apply_btn.setFixedHeight(36)
        reset_btn.setFixedHeight(36)
        cancel_btn.setFixedHeight(36)

        reset_btn.clicked.connect(self._canvas.reset_crop)
        cancel_btn.clicked.connect(self.reject)
        apply_btn.clicked.connect(self._accept)

        fl.addWidget(reset_btn)
        fl.addStretch()
        fl.addWidget(cancel_btn)
        fl.addWidget(apply_btn)
        root.addWidget(footer)

        # Load image last so the canvas has its size
        self._canvas.load_image(image_path, rotation, existing_crop)
        self._update_info()

    def _update_info(self):
        nc = self._canvas.get_norm_crop()
        if nc is None:
            self._info_lbl.setText("")
            return
        w = int(nc[2] * self._canvas._orig_w)
        h = int(nc[3] * self._canvas._orig_h)
        ar = f"{w/h:.2f}" if h else "—"
        self._info_lbl.setText(f"  {w} × {h} px   ratio {ar}  ")

    def _accept(self):
        self._result = self._canvas.get_norm_crop()
        self.accept()

    def get_result(self) -> tuple | None:
        """Returns (x, y, w, h) in 0-1 coords, or None if reset/cancelled."""
        return self._result


# ── Dialogs ───────────────────────────────────────────────────────────────────

class HelpDialog(QDialog):
    def __init__(self, parent=None, language="en"):
        super().__init__(parent)
        self.setWindowTitle("Help Topics")
        self.resize(640, 460)
        self.language = language
        self.setStyleSheet(f"background: {COLORS['bg_panel']}; color: {COLORS['text_primary']};")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        title = QLabel("Help Topics")
        title.setStyleSheet(f"font-size: 16px; font-weight: 700; color: {COLORS['text_primary']};")
        layout.addWidget(title)

        splitter = QSplitter(Qt.Horizontal)
        self.topic_list = QListWidget()
        self.info_display = QTextEdit()
        self.info_display.setReadOnly(True)
        splitter.addWidget(self.topic_list)
        splitter.addWidget(self.info_display)
        splitter.setSizes([200, 400])
        layout.addWidget(splitter)

        close_btn = _styled_btn("Close", "")
        close_btn.clicked.connect(self.close)
        close_btn.setFixedWidth(80)
        btn_row = QHBoxLayout()
        btn_row.addStretch()
        btn_row.addWidget(close_btn)
        layout.addLayout(btn_row)

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
        self.setMinimumSize(500, 380)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(14)

        # Header row
        header = QHBoxLayout()
        self.counter_label = QLabel()
        self.counter_label.setStyleSheet(f"color: {COLORS['text_muted']}; font-size: 12px;")
        header.addStretch()
        header.addWidget(self.counter_label)
        layout.addLayout(header)

        # Image preview
        self.image_label = QLabel()
        self.image_label.setAlignment(Qt.AlignCenter)
        self.image_label.setMinimumHeight(200)
        self.image_label.setStyleSheet(
            f"background: {COLORS['bg_card']}; border: 1px solid {COLORS['border']}; border-radius: 8px;"
        )
        layout.addWidget(self.image_label)

        # Text input
        self.text_input = QTextEdit()
        self.text_input.setPlaceholderText(self.tr("enter_text_for_image"))
        self.text_input.setAlignment(Qt.AlignRight)
        self.text_input.setLayoutDirection(Qt.RightToLeft)
        self.text_input.setPlainText(self.images[self.current_index].get("text", ""))
        self.text_input.setFixedHeight(80)
        self.text_input.installEventFilter(self)
        layout.addWidget(self.text_input)

        # Buttons
        btn_row = QHBoxLayout()
        btn_row.setSpacing(8)
        close_btn = _styled_btn(self.tr("close"), "")
        next_btn  = _styled_btn(f"{self.tr('next')}  →", "primary")
        close_btn.clicked.connect(self.close)
        next_btn.clicked.connect(self.next_image)
        next_btn.setFixedHeight(36)
        btn_row.addWidget(close_btn)
        btn_row.addStretch()
        btn_row.addWidget(next_btn)
        layout.addLayout(btn_row)

        self.text_input.moveCursor(QTextCursor.Start)
        self.update_image()

    def update_image(self):
        n = len(self.images)
        self.counter_label.setText(f"{self.current_index + 1} / {n}")
        if 0 <= self.current_index < n:
            data = self.images[self.current_index]
            px   = QPixmap(data["path"])
            rot  = data.get("rotation", 0)
            if rot:
                t = QTransform(); t.rotate(rot)
                px = px.transformed(t, Qt.SmoothTransformation)
            self.image_label.setPixmap(px.scaled(460, 190, Qt.KeepAspectRatio, Qt.SmoothTransformation))
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
        self.setMinimumWidth(340)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 24, 24, 20)
        layout.setSpacing(16)

        title = QLabel("Project Info")
        title.setStyleSheet(f"font-size: 16px; font-weight: 700; color: {COLORS['text_primary']};")
        layout.addWidget(title)
        layout.addWidget(_make_divider())

        dur_with    = sum(img["duration"] for img in images)
        dur_without = sum(img["duration"] for img in images if not img.get("is_second_image"))
        audio_dur   = sum(_get_audio_duration(a["path"]) for a in audio_files)

        def _info_row(label: str, value: str):
            row = QHBoxLayout()
            lbl = QLabel(label)
            lbl.setStyleSheet(f"color: {COLORS['text_secondary']}; font-size: 12px;")
            val = QLabel(value)
            val.setStyleSheet(f"color: {COLORS['text_primary']}; font-size: 13px; font-weight: 600;")
            val.setAlignment(Qt.AlignRight)
            row.addWidget(lbl)
            row.addStretch()
            row.addWidget(val)
            return row

        layout.addLayout(_info_row(self.tr("info_total_images"), str(len(images))))
        layout.addLayout(_info_row(self.tr("info_duration_with_second"), format_time_hms(dur_with)))
        layout.addLayout(_info_row(self.tr("info_duration_without_second"), format_time_hms(dur_without)))
        layout.addLayout(_info_row(self.tr("info_audio_duration"), format_time_hms(audio_dur)))
        layout.addWidget(_make_divider())

        close_btn = _styled_btn(self.tr("close"), "primary")
        close_btn.clicked.connect(self.close)
        close_btn.setFixedHeight(34)
        btn_row = QHBoxLayout()
        btn_row.addStretch()
        btn_row.addWidget(close_btn)
        layout.addLayout(btn_row)


class FontPickerDialog(QDialog):
    """
    Font picker:
    • Left panel  – installed Windows/system fonts (from QFontDatabase)
    • Right panel – Google Fonts browser (search + one-click download)
    Selected font is returned via get_result().
    """
    font_selected = pyqtSignal(str)   # font family name

    # A curated list of popular Google Fonts families
    GOOGLE_FONTS_POPULAR = [
        "Roboto", "Open Sans", "Lato", "Montserrat", "Oswald", "Raleway",
        "Poppins", "Merriweather", "Ubuntu", "Nunito", "Playfair Display",
        "Source Sans Pro", "PT Sans", "Noto Sans", "Inter", "Rubik",
        "Work Sans", "Fira Sans", "Quicksand", "Cabin", "Barlow",
        "Josefin Sans", "Mulish", "Karla", "Exo 2", "Titillium Web",
        "Heebo", "Libre Franklin", "Arimo", "Tinos", "Alef", "David Libre",
        "Frank Ruhl Libre", "Varela Round", "Assistant", "Secular One",
    ]

    def __init__(self, current_font: str = "Segoe UI", parent=None):
        super().__init__(parent)
        self.setWindowTitle("Choose Font")
        self.setMinimumSize(760, 520)
        self.resize(820, 560)
        self.setWindowFlags(self.windowFlags() | Qt.WindowMaximizeButtonHint)
        self._result: str | None = None
        self._download_dir = BASEPATH / "Fonts" / "GoogleFonts"
        self._download_dir.mkdir(parents=True, exist_ok=True)
        self._build_ui(current_font)

    # ── UI ────────────────────────────────────────────────────────────────────

    def _build_ui(self, current_font: str):
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # Header
        header = QWidget()
        header.setFixedHeight(48)
        header.setStyleSheet(
            f"background: {COLORS['toolbar_bg']}; border-bottom: 1px solid {COLORS['border']};"
        )
        hl = QHBoxLayout(header)
        hl.setContentsMargins(16, 0, 16, 0)
        title = QLabel("  Choose Font")
        title.setStyleSheet(f"color: {COLORS['text_primary']}; font-size: 14px; font-weight: 700;")
        self._preview_label = QLabel("The quick brown fox  אבגדהוזחטי")
        self._preview_label.setStyleSheet(
            f"color: {COLORS['text_secondary']}; font-size: 18px; padding: 0 16px;"
        )
        hl.addWidget(title)
        hl.addStretch()
        hl.addWidget(self._preview_label)
        root.addWidget(header)

        # Split: system fonts | Google Fonts
        splitter = QSplitter(Qt.Horizontal)
        splitter.setHandleWidth(1)

        # ── Left: system fonts ────────────────────────────────────────────
        left = QWidget()
        ll = QVBoxLayout(left)
        ll.setContentsMargins(12, 12, 6, 12)
        ll.setSpacing(8)
        ll.addWidget(_make_section_label("System / Installed Fonts"))

        self._sys_search = QLineEdit()
        self._sys_search.setPlaceholderText("Search…")
        self._sys_search.textChanged.connect(self._filter_system)
        ll.addWidget(self._sys_search)

        self._sys_list = QListWidget()
        db = QFontDatabase()
        self._all_system_fonts = sorted(db.families())
        self._sys_list.addItems(self._all_system_fonts)
        # pre-select current
        matches = self._sys_list.findItems(current_font, Qt.MatchExactly)
        if matches:
            self._sys_list.setCurrentItem(matches[0])
            self._sys_list.scrollToItem(matches[0])
        self._sys_list.currentTextChanged.connect(self._on_system_select)
        ll.addWidget(self._sys_list)
        splitter.addWidget(left)

        # ── Right: Google Fonts ───────────────────────────────────────────
        right = QWidget()
        rl = QVBoxLayout(right)
        rl.setContentsMargins(6, 12, 12, 12)
        rl.setSpacing(8)
        rl.addWidget(_make_section_label("Google Fonts  (downloads to your computer)"))

        gf_top = QHBoxLayout()
        self._gf_search = QLineEdit()
        self._gf_search.setPlaceholderText("Search Google Fonts…")
        self._gf_search.textChanged.connect(self._filter_google)
        refresh_btn = QPushButton("↻")
        refresh_btn.setFixedSize(30, 28)
        refresh_btn.setToolTip("Refresh list from internet")
        refresh_btn.clicked.connect(self._fetch_google_fonts_list)
        gf_top.addWidget(self._gf_search)
        gf_top.addWidget(refresh_btn)
        rl.addLayout(gf_top)

        self._gf_list = QListWidget()
        self._all_google_fonts = list(self.GOOGLE_FONTS_POPULAR)
        self._populate_google_list()
        self._gf_list.currentTextChanged.connect(self._on_google_select)
        rl.addWidget(self._gf_list)

        self._dl_status = QLabel("")
        self._dl_status.setStyleSheet(f"color: {COLORS['text_muted']}; font-size: 11px;")
        self._dl_status.setWordWrap(True)
        rl.addWidget(self._dl_status)

        self._dl_btn = _styled_btn("⬇  Download & Use Selected Font", "primary")
        self._dl_btn.setFixedHeight(34)
        self._dl_btn.setEnabled(False)
        self._dl_btn.clicked.connect(self._download_google_font)
        rl.addWidget(self._dl_btn)
        splitter.addWidget(right)

        splitter.setSizes([380, 380])
        root.addWidget(splitter, 1)

        # Footer
        footer = QWidget()
        footer.setFixedHeight(52)
        footer.setStyleSheet(
            f"background: {COLORS['toolbar_bg']}; border-top: 1px solid {COLORS['border']};"
        )
        fl = QHBoxLayout(footer)
        fl.setContentsMargins(16, 0, 16, 0)
        fl.setSpacing(10)
        cancel_btn = _styled_btn("Cancel", "")
        self._apply_btn = _styled_btn("✔  Use This Font", "primary")
        self._apply_btn.setFixedHeight(36)
        self._apply_btn.setEnabled(False)
        cancel_btn.clicked.connect(self.reject)
        self._apply_btn.clicked.connect(self._accept)
        self._sel_label = QLabel("No font selected")
        self._sel_label.setStyleSheet(
            f"color: {COLORS['accent']}; font-size: 13px; font-weight: 700;"
        )
        fl.addWidget(self._sel_label)
        fl.addStretch()
        fl.addWidget(cancel_btn)
        fl.addWidget(self._apply_btn)
        root.addWidget(footer)

    # ── system font handling ──────────────────────────────────────────────────

    def _filter_system(self, text: str):
        self._sys_list.clear()
        lo = text.lower()
        self._sys_list.addItems([f for f in self._all_system_fonts if lo in f.lower()])

    def _on_system_select(self, family: str):
        if not family:
            return
        self._gf_list.clearSelection()
        self._result = family
        self._sel_label.setText(f"Selected: {family}")
        self._apply_btn.setEnabled(True)
        self._preview_label.setFont(QFont(family, 14))

    # ── Google Fonts handling ─────────────────────────────────────────────────

    def _populate_google_list(self, filter_text: str = ""):
        self._gf_list.clear()
        lo = filter_text.lower()
        for name in sorted(self._all_google_fonts):
            if lo in name.lower():
                # Mark downloaded fonts
                ttf = self._download_dir / f"{name.replace(' ', '_')}.ttf"
                label = f"✓ {name}" if ttf.exists() else name
                item = QListWidgetItem(label)
                item.setData(Qt.UserRole, name)
                if ttf.exists():
                    item.setForeground(QColor(COLORS["success"]))
                self._gf_list.addItem(item)

    def _filter_google(self, text: str):
        self._populate_google_list(text)

    def _on_google_select(self, _):
        item = self._gf_list.currentItem()
        if not item:
            return
        family = item.data(Qt.UserRole) or item.text().lstrip("✓ ").strip()
        self._sys_list.clearSelection()
        ttf = self._download_dir / f"{family.replace(' ', '_')}.ttf"
        if ttf.exists():
            self._result = family
            self._sel_label.setText(f"Selected: {family}")
            self._apply_btn.setEnabled(True)
            # Register font with Qt if not already registered
            fdb = QFontDatabase()
            fdb.addApplicationFont(str(ttf))
            self._preview_label.setFont(QFont(family, 14))
            self._dl_btn.setText("✔  Font already downloaded — Use It")
        else:
            self._result = None
            self._apply_btn.setEnabled(False)
            self._sel_label.setText(f"Download required: {family}")
            self._dl_btn.setText(f"⬇  Download & Use  '{family}'")
        self._dl_btn.setEnabled(True)

    def _fetch_google_fonts_list(self):
        """Background fetch of broader font list from Google Fonts API (if available)."""
        self._dl_status.setText("Fetching Google Fonts list…")

        def _fetch():
            try:
                url = "https://fonts.google.com/metadata/fonts"
                req = urllib.request.Request(url, headers={"User-Agent": "Eventure-App"})
                with urllib.request.urlopen(req, timeout=8) as resp:
                    raw = resp.read().decode("utf-8")
                # The metadata JSON has a ")]}'" prefix (XSSI protection)
                if raw.startswith(")]}'"):
                    raw = raw[4:].lstrip()
                import json as _json
                data = _json.loads(raw)
                families = [f["family"] for f in data.get("familyMetadataList", [])]
                if families:
                    return families
            except Exception as e:
                print(f"Google Fonts metadata fetch failed: {e}")
            return None

        class _W(QThread):
            done = pyqtSignal(object)
            def run(self): self.done.emit(_fetch())

        w = _W(self)
        def _on_done(families):
            if families:
                self._all_google_fonts = families
                self._populate_google_list(self._gf_search.text())
                self._dl_status.setText(f"Loaded {len(families)} Google Fonts families.")
            else:
                self._dl_status.setText("Could not load font list — check internet connection.")
            w.deleteLater()
        w.done.connect(_on_done)
        w.start()

    def _download_google_font(self):
        item = self._gf_list.currentItem()
        if not item:
            return
        family = item.data(Qt.UserRole) or item.text().lstrip("✓ ").strip()
        ttf_path = self._download_dir / f"{family.replace(' ', '_')}.ttf"
        if ttf_path.exists():
            # Already downloaded — just apply
            self._result = family
            self._sel_label.setText(f"Selected: {family}")
            self._apply_btn.setEnabled(True)
            fdb = QFontDatabase()
            fdb.addApplicationFont(str(ttf_path))
            self._preview_label.setFont(QFont(family, 14))
            return

        self._dl_status.setText(f"Downloading '{family}'…")
        self._dl_btn.setEnabled(False)
        QApplication.processEvents()

        def _do_download():
            try:
                # Google Fonts CSS API → extract TTF URL
                css_url = (
                    "https://fonts.googleapis.com/css2"
                    f"?family={family.replace(' ', '+')}:wght@400;700"
                    "&display=swap"
                )
                req = urllib.request.Request(
                    css_url,
                    headers={"User-Agent":
                        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
                )
                with urllib.request.urlopen(req, timeout=10) as resp:
                    css = resp.read().decode("utf-8")
                # Extract first TTF/WOFF2 URL from CSS
                import re as _re
                urls = _re.findall(r"url\((https://[^)]+)\)", css)
                if not urls:
                    return None, "No font URL found in Google Fonts CSS."
                font_url = urls[0]
                with urllib.request.urlopen(
                    urllib.request.Request(font_url,
                        headers={"User-Agent": "Mozilla/5.0"}), timeout=15
                ) as resp:
                    data = resp.read()
                with open(ttf_path, "wb") as f:
                    f.write(data)
                return str(ttf_path), None
            except Exception as e:
                return None, str(e)

        class _DL(QThread):
            done = pyqtSignal(object, object)
            def run(self): self.done.emit(*_do_download())

        dl = _DL(self)
        def _on_dl(path, err):
            if path:
                fdb = QFontDatabase()
                fdb.addApplicationFont(path)
                self._result = family
                self._sel_label.setText(f"Selected: {family}")
                self._apply_btn.setEnabled(True)
                self._preview_label.setFont(QFont(family, 14))
                self._dl_status.setText(f"✓ '{family}' downloaded successfully.")
                self._populate_google_list(self._gf_search.text())
                self._dl_btn.setText("✔  Font downloaded — Use It")
            else:
                self._dl_status.setText(f"Download failed: {err}")
            self._dl_btn.setEnabled(True)
            dl.deleteLater()
        dl.done.connect(_on_dl)
        dl.start()

    # ── accept / result ───────────────────────────────────────────────────────

    def _accept(self):
        if self._result:
            self.font_selected.emit(self._result)
            self.accept()

    def get_result(self) -> str | None:
        return self._result


class AudioLibraryDialog(QDialog):
    """
    Audio library that streams its song index from GitHub.
    Songs are downloaded on-demand to BASEPATH/Songs/ — no need to bundle them.
    The index is re-fetched every time the dialog opens (with a local cache
    fallback if offline).
    """
    def __init__(self, tr_function=None, parent=None):
        super().__init__(parent)
        self.tr    = tr_function
        self.songs = []
        self._downloading: set = set()   # track in-progress downloads
        self.setWindowTitle(self.tr("label_audio_library"))
        self.setMinimumSize(720, 520)
        self.resize(800, 560)
        self._songs_dir = BASEPATH / "Songs"
        self._songs_dir.mkdir(parents=True, exist_ok=True)
        self._init_ui()
        self._fetch_index()   # async load

    # ── index fetching ────────────────────────────────────────────────────────

    def _fetch_index(self):
        """Fetch songs.json from GitHub in the background."""
        self._set_status("Checking for music library updates…")

        class _W(QThread):
            done = pyqtSignal(object, object)   # (songs_list, error_str)
            def run(self):
                try:
                    req = urllib.request.Request(
                        GITHUB_AUDIO_INDEX,
                        headers={"User-Agent": "Eventure-App",
                                 "Cache-Control": "no-cache"}
                    )
                    with urllib.request.urlopen(req, timeout=8) as resp:
                        import json as _j
                        data = _j.loads(resp.read().decode("utf-8"))
                    # Persist locally so we can work offline next time
                    local = Path(str(BASEPATH / "Songs" / "songs.json"))
                    with open(local, "w", encoding="utf-8") as f:
                        _j.dump(data, f, ensure_ascii=False, indent=2)
                    self.done.emit(data, None)
                except Exception as e:
                    self.done.emit(None, str(e))

        w = _W(self)
        def _on_done(songs, err):
            if songs is not None:
                self.songs = songs
                self._set_status(f"{len(songs)} songs in library  •  click ⬇ to download")
            else:
                # Fall back to local cache
                self._load_local_songs()
                if self.songs:
                    self._set_status(f"Offline mode — {len(self.songs)} cached songs")
                else:
                    self._set_status(f"Could not load library: {err}")
            self._populate(filter_text="")
            w.deleteLater()
        w.done.connect(_on_done)
        w.start()

    def _load_local_songs(self):
        songs_file = self._songs_dir / "songs.json"
        try:
            import json as _j
            with open(songs_file, "r", encoding="utf-8") as f:
                self.songs = _j.load(f)
        except Exception as e:
            print(f"Songs local load error: {e}")
            self.songs = []

    # ── UI ────────────────────────────────────────────────────────────────────

    def _init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        title = QLabel(self.tr("label_audio_library"))
        title.setStyleSheet(f"font-size: 16px; font-weight: 700; color: {COLORS['text_primary']};")
        layout.addWidget(title)

        # Status / info bar
        self._status_lbl = QLabel("Loading…")
        self._status_lbl.setStyleSheet(f"color: {COLORS['text_muted']}; font-size: 11px;")
        layout.addWidget(self._status_lbl)

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText(self.tr("search_songs"))
        self.search_input.textChanged.connect(self._filter_songs)
        layout.addWidget(self.search_input)

        content_row = QHBoxLayout()
        self.song_list = QListWidget()
        content_row.addWidget(self.song_list, 1)

        # Info panel
        info_panel = QFrame()
        info_panel.setProperty("class", "panel")
        info_panel.setMinimumWidth(240)
        info_layout = QVBoxLayout(info_panel)
        info_layout.setContentsMargins(14, 14, 14, 14)
        info_layout.setSpacing(8)
        info_layout.addWidget(_make_section_label("Song Details"))
        self.info_label = QLabel(self.tr("song_info_label"))
        self.info_label.setStyleSheet(
            f"color: {COLORS['text_secondary']}; font-size: 12px; line-height: 1.6;"
        )
        self.info_label.setWordWrap(True)
        self.info_label.setAlignment(Qt.AlignTop)
        info_layout.addWidget(self.info_label)

        # Download progress inside info panel
        _dl_row_widget = QWidget()
        _dl_row = QHBoxLayout(_dl_row_widget)
        _dl_row.setContentsMargins(0, 0, 0, 0)
        _dl_row.setSpacing(6)
        self._dl_bar = QProgressBar()
        self._dl_bar.setRange(0, 100)
        self._dl_bar.setValue(0)
        self._dl_bar.setFixedHeight(4)
        self._dl_pct_lbl = QLabel("0%")
        self._dl_pct_lbl.setStyleSheet(
            f"color: {COLORS['text_secondary']}; font-size: 11px; font-weight: 600;"
        )
        self._dl_pct_lbl.setFixedWidth(34)
        self._dl_pct_lbl.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        _dl_row.addWidget(self._dl_bar)
        _dl_row.addWidget(self._dl_pct_lbl)
        _dl_row_widget.setVisible(False)
        self._dl_bar_container = _dl_row_widget
        info_layout.addWidget(_dl_row_widget)

        self._dl_lbl = QLabel("")
        self._dl_lbl.setStyleSheet(f"color: {COLORS['warning']}; font-size: 11px;")
        self._dl_lbl.setWordWrap(True)
        info_layout.addWidget(self._dl_lbl)

        info_layout.addStretch()
        content_row.addWidget(info_panel)
        layout.addLayout(content_row)

        btn_row = QHBoxLayout()
        btn_row.setSpacing(8)
        close_btn = _styled_btn(self.tr("close"), "")
        refresh_btn = _styled_btn("↻  Refresh Library", "")
        self.add_btn = _styled_btn(f"＋  {self.tr('add_selected')}", "primary")
        self.add_btn.setFixedHeight(36)
        close_btn.clicked.connect(self.close)
        refresh_btn.clicked.connect(self._fetch_index)
        self.add_btn.clicked.connect(self._add_selected)
        btn_row.addWidget(close_btn)
        btn_row.addWidget(refresh_btn)
        btn_row.addStretch()
        btn_row.addWidget(self.add_btn)
        layout.addLayout(btn_row)

        self.song_list.itemSelectionChanged.connect(self._update_info)

    def _set_status(self, text: str):
        self._status_lbl.setText(text)

    def _populate(self, filter_text: str = ""):
        self.song_list.clear()
        lo = filter_text.lower()
        for song in self.songs:
            if (lo in song["name"].lower() or lo in song["author"].lower()
                    or lo in song.get("fits_for", "").lower()):
                local_path = self._local_path(song)
                downloaded = local_path.exists()
                icon = "✓" if downloaded else "⬇"
                item = QListWidgetItem(f"{icon}  {song['name']} — {song['author']}")
                item.setData(Qt.UserRole, song)
                if downloaded:
                    item.setForeground(QColor(COLORS["success"]))
                self.song_list.addItem(item)

    def _filter_songs(self):
        self._populate(self.search_input.text())

    def _local_path(self, song: dict) -> Path:
        """
        Return the local file path for a song.
        Songs are stored flat in self._songs_dir (no sub-folders).
        Works whether the JSON entry uses:
          {BASE_PATH}/Songs/AudioFiles/Artist - Title.mp3   (old format)
          AudioFiles/Artist - Title.mp3                     (github_path)
          Artist - Title.mp3                                (filename only)
        """
        raw = song.get("path", song["name"] + ".mp3")
        raw = raw.replace("{BASE_PATH}", "").replace("\\", "/")
        filename = os.path.basename(raw.replace("\\", "/"))
        return self._songs_dir / filename

    def _github_url(self, song: dict) -> str:
        """
        Build the raw GitHub download URL.
        Handles all three path formats in the JSON:
          1. github_path = "AudioFiles/Artist - Title.mp3"   → use as-is (subfolder aware)
          2. path        = "{BASE_PATH}/Songs/AudioFiles/..."  → strip prefix, keep subfolder
          3. neither     → just the filename, assume root of Songs/
        GITHUB_AUDIO_BASE already ends with /Songs so we append the relative part.
        """
        # Prefer explicit github_path (cleanest)
        gp = song.get("github_path", "")
        if gp:
            # Already relative to Songs/ on GitHub
            rel = gp.replace("\\", "/").lstrip("/")
            return f"{GITHUB_AUDIO_BASE}/{urllib.parse.quote(rel, safe='/')}"

        # Fall back to "path" — strip {BASE_PATH}/Songs/ prefix to get relative part
        raw = song.get("path", song["name"] + ".mp3")
        raw = raw.replace("{BASE_PATH}", "").replace("\\", "/")
        # raw might now be  /Songs/AudioFiles/X.mp3  or  Songs/AudioFiles/X.mp3
        # Strip leading /Songs/ so we don't double it with GITHUB_AUDIO_BASE
        for prefix in ("/Songs/", "Songs/"):
            if raw.lstrip("/").startswith(prefix.lstrip("/")):
                raw = raw.lstrip("/")[len(prefix.lstrip("/")):]
                break
        raw = raw.lstrip("/")
        if not raw:
            raw = song["name"] + ".mp3"
        return f"{GITHUB_AUDIO_BASE}/{urllib.parse.quote(raw, safe='/')}"

    def _fmt_dur(self, seconds: float) -> str:
        m = int(seconds // 60); s = int(seconds % 60)
        return f"{m}:{s:02d}"

    def _update_info(self):
        selected = self.song_list.selectedItems()
        if selected:
            s = selected[0].data(Qt.UserRole)
            local = self._local_path(s)
            dl_status = "✓ Downloaded" if local.exists() else "⬇ Not yet downloaded"
            self.info_label.setText(
                f"<b>Name:</b> {s['name']}<br><br>"
                f"<b>Artist:</b> {s['author']}<br><br>"
                f"<b>Duration:</b> {self._fmt_dur(s['duration'])}<br><br>"
                f"<b>Fits for:</b> {s.get('fits_for', '')}<br><br>"
                f"<b>Status:</b> {dl_status}<br>"
            )

    def _download_song(self, song: dict, on_done):
        """Download a song from GitHub asynchronously."""
        local_path = self._local_path(song)
        if local_path.exists():
            on_done(str(local_path), None)
            return

        url = self._github_url(song)
        self._dl_lbl.setText(f"Downloading: {song['name']}…")
        self._dl_bar_container.setVisible(True)
        self._dl_bar.setValue(0)
        self._dl_pct_lbl.setText("0%")

        class _DL(QThread):
            progress = pyqtSignal(int)
            done = pyqtSignal(object, object)

            def __init__(self, _url, _path):
                super().__init__()
                self._url  = _url
                self._path = _path

            def run(self):
                try:
                    req = urllib.request.Request(
                        self._url, headers={"User-Agent": "Eventure-App"}
                    )
                    with urllib.request.urlopen(req, timeout=30) as resp:
                        total = int(resp.headers.get("Content-Length", 0))
                        downloaded = 0
                        chunk_size = 32768
                        buf = b""
                        while True:
                            chunk = resp.read(chunk_size)
                            if not chunk:
                                break
                            buf += chunk
                            downloaded += len(chunk)
                            if total:
                                self.progress.emit(int(downloaded / total * 100))
                    with open(self._path, "wb") as f:
                        f.write(buf)
                    self.done.emit(str(self._path), None)
                except Exception as e:
                    self.done.emit(None, str(e))

        dl = _DL(url, str(local_path))
        self._downloading.add(dl)

        def _on_progress(pct: int):
            self._dl_bar.setValue(pct)
            self._dl_pct_lbl.setText(f"{pct}%")

        dl.progress.connect(_on_progress)

        def _finish(path, err):
            self._dl_bar_container.setVisible(False)
            if path:
                self._dl_lbl.setText(f"✓ Downloaded: {song['name']}")
                self._populate(self.search_input.text())
            else:
                self._dl_lbl.setText(f"Download failed: {err}")
            self._downloading.discard(dl)
            dl.deleteLater()
            on_done(path, err)
        dl.done.connect(_finish)
        dl.start()

    def _add_selected(self):
        items = self.song_list.selectedItems()
        if not items:
            return
        song = items[0].data(Qt.UserRole)
        local_path = self._local_path(song)
        if local_path.exists():
            self._do_add(str(local_path))
        else:
            def _on_dl(path, err):
                if path:
                    self._do_add(path)
                else:
                    QMessageBox.warning(self, "Download Failed",
                        f"Could not download '{song['name']}':\n{err}\n\n"
                        "Please check your internet connection.")
            self._download_song(song, _on_dl)

    def _do_add(self, path: str):
        parent = self.parent()
        if parent and not any(str(a["path"]) == path for a in parent.audio_files):
            parent.audio_files.append({"path": path})
            parent.update_audio_table()
        self.close()


def export_pdf_photo_book(images: list, output_path: str,
                          progress_callback=None) -> bool:
    """
    Create a printable PDF with one image per page and its caption as footer.
    Uses only reportlab (pip install reportlab) — no cloud, no subscription.
    Falls back with a friendly error if reportlab is not installed.
    """
    try:
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.pdfgen import canvas as rl_canvas
        from reportlab.lib.units import mm
        from reportlab.lib.utils import ImageReader
    except ImportError:
        return False  # caller should show an install hint

    PAGE_W, PAGE_H = landscape(A4)
    MARGIN = 12 * mm

    c = rl_canvas.Canvas(output_path, pagesize=(PAGE_W, PAGE_H))

    total = len(images)
    for idx, img_data in enumerate(images):
        path     = img_data["path"]
        caption  = img_data.get("text", "").strip()
        rotation = img_data.get("rotation", 0)

        try:
            pil_img = Image.open(path)
            # ── Apply EXIF orientation ────────────────────────────────────────
            try:
                if hasattr(pil_img, "_getexif"):
                    exif = pil_img._getexif()
                    if exif and _ORIENTATION_TAG:
                        val = exif.get(_ORIENTATION_TAG)
                        if val == 3:
                            pil_img = pil_img.rotate(180, expand=True)
                        elif val == 6:
                            pil_img = pil_img.rotate(270, expand=True)
                        elif val == 8:
                            pil_img = pil_img.rotate(90, expand=True)
            except Exception:
                pass
            # ── Apply manual rotation from the table ──────────────────────────
            if rotation:
                pil_img = pil_img.rotate(rotation, expand=True)
            pil_img = pil_img.convert("RGB")
            from io import BytesIO
            buf = BytesIO()
            pil_img.save(buf, format="JPEG", quality=88)
            buf.seek(0)
            ir = ImageReader(buf)

            # Available space (leave room for optional caption at bottom)
            avail_w = PAGE_W - 2 * MARGIN
            avail_h = PAGE_H - 2 * MARGIN - (14 * mm if caption else 0)

            iw, ih = pil_img.size
            scale = min(avail_w / iw, avail_h / ih)
            draw_w = iw * scale
            draw_h = ih * scale

            x = (PAGE_W - draw_w) / 2
            y = MARGIN + (14 * mm if caption else 0) + (avail_h - draw_h) / 2

            c.drawImage(ir, x, y, draw_w, draw_h, preserveAspectRatio=True)

            if caption:
                c.setFont("Helvetica", 11)
                c.setFillColorRGB(0.3, 0.3, 0.3)
                c.drawCentredString(PAGE_W / 2, MARGIN + 3 * mm, caption)

        except Exception as e:
            print(f"PDF export: skipping {path}: {e}")

        c.showPage()
        if progress_callback:
            progress_callback(int((idx + 1) / total * 100))

    c.save()
    return True


# ── Feature 13: Export as HTML Gallery ───────────────────────────────────────

def export_html_gallery(images: list, output_path: str,
                        title: str = "Eventure Gallery",
                        progress_callback=None) -> bool:
    """
    Generate a fully self-contained index.html with all images embedded
    as base64 data URIs — no server, no internet connection needed.
    Each slide shows its slide number in the top-left corner.
    Images are rendered with EXIF orientation, manual rotation, and crop
    settings from the table.
    """
    from io import BytesIO
    slides_html = []
    total = len(images)

    for idx, img_data in enumerate(images):
        path     = img_data["path"]
        caption  = img_data.get("text", "").strip()
        rotation = img_data.get("rotation", 0)
        crop     = img_data.get("crop")       # tuple (x, y, w, h) in 0-1 coords or None
        try:
            pil_img = Image.open(path)
            # ── Apply EXIF orientation ────────────────────────────────────────
            try:
                if hasattr(pil_img, "_getexif"):
                    exif = pil_img._getexif()
                    if exif and _ORIENTATION_TAG:
                        val = exif.get(_ORIENTATION_TAG)
                        if val == 3:
                            pil_img = pil_img.rotate(180, expand=True)
                        elif val == 6:
                            pil_img = pil_img.rotate(270, expand=True)
                        elif val == 8:
                            pil_img = pil_img.rotate(90, expand=True)
            except Exception:
                pass
            # ── Apply manual rotation from the table ──────────────────────────
            if rotation:
                pil_img = pil_img.rotate(rotation, expand=True)
            # ── Apply crop (normalised 0-1 coords relative to post-rotation size) ─
            if crop:
                iw, ih = pil_img.size
                cx = max(0, int(crop[0] * iw))
                cy = max(0, int(crop[1] * ih))
                cw = max(1, min(int(crop[2] * iw), iw - cx))
                ch = max(1, min(int(crop[3] * ih), ih - cy))
                pil_img = pil_img.crop((cx, cy, cx + cw, cy + ch))
            pil_img = pil_img.convert("RGB")
            # Downscale large images to reduce file size (max 1280px wide)
            if pil_img.width > 1280:
                ratio = 1280 / pil_img.width
                pil_img = pil_img.resize(
                    (1280, int(pil_img.height * ratio)), Image.LANCZOS
                )
            buf = BytesIO()
            pil_img.save(buf, format="JPEG", quality=82, optimize=True)
            b64 = base64.b64encode(buf.getvalue()).decode("ascii")
            src = f"data:image/jpeg;base64,{b64}"
        except Exception as e:
            print(f"HTML gallery: skipping {path}: {e}")
            continue

        slide_num = idx + 1
        cap_html = (
            f'<div class="caption">{caption}</div>' if caption else ""
        )
        slides_html.append(
            f'<div class="slide" id="slide-{idx}">'
            f'<div class="img-wrap">'
            f'<img src="{src}" alt="Slide {slide_num}" />'
            f'<span class="slide-num">{slide_num}</span>'
            f'</div>'
            f'{cap_html}'
            f'</div>'
        )
        if progress_callback:
            progress_callback(int((idx + 1) / total * 80))

    gallery_items = "\n".join(slides_html)
    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>{title}</title>
<style>
  *{{box-sizing:border-box;margin:0;padding:0}}
  body{{background:#111;color:#eee;font-family:sans-serif;}}
  header{{padding:20px 32px;background:#1a1a2e;border-bottom:1px solid #333;display:flex;align-items:center;gap:16px;flex-wrap:wrap;}}
  header h1{{font-size:22px;font-weight:700;color:#5b9bff;flex:1;}}
  header p{{font-size:12px;color:#666;margin-top:4px;}}
  .header-info{{flex:1;}}
  .gallery{{display:grid;grid-template-columns:repeat(auto-fill,minmax(320px,1fr));gap:16px;padding:24px;}}
  .slide{{background:#1e2228;border-radius:10px;overflow:hidden;border:1px solid #363e4a;
          transition:transform 0.2s,box-shadow 0.2s;cursor:pointer;}}
  .slide:hover{{transform:translateY(-3px);box-shadow:0 8px 24px rgba(0,0,0,.5);}}
  /* Wrapper so the badge is relative to the image */
  .img-wrap{{position:relative;}}
  .img-wrap img{{width:100%;display:block;aspect-ratio:16/9;object-fit:cover;}}
  /* Slide number badge */
  .slide-num{{
    position:absolute;top:8px;left:8px;
    background:rgba(0,0,0,0.65);color:#fff;
    font-size:11px;font-weight:700;line-height:1;
    padding:4px 7px;border-radius:5px;
    pointer-events:none;
  }}
  .caption{{padding:10px 14px;font-size:13px;color:#95a0b4;background:#252b34;}}
  /* Lightbox */
  #lb{{display:none;position:fixed;inset:0;background:rgba(0,0,0,.9);
       align-items:center;justify-content:center;z-index:999;flex-direction:column;}}
  #lb.open{{display:flex;}}
  #lb img{{max-width:92vw;max-height:85vh;border-radius:6px;}}
  #lb .lb-cap{{color:#ccc;font-size:14px;margin-top:12px;max-width:800px;text-align:center;}}
  #lb-num{{position:absolute;top:18px;left:24px;font-size:13px;font-weight:700;
           background:rgba(0,0,0,.6);color:#fff;padding:4px 9px;border-radius:6px;}}
  #lb-close{{position:absolute;top:18px;right:24px;font-size:28px;color:#aaa;cursor:pointer;
             background:none;border:none;line-height:1;}}
  #lb-prev,#lb-next{{position:absolute;top:50%;transform:translateY(-50%);font-size:32px;
                     color:#aaa;cursor:pointer;background:none;border:none;padding:8px 16px;}}
  #lb-prev{{left:12px;}} #lb-next{{right:12px;}}
</style>
</head>
<body>
<header>
  <div class="header-info">
    <h1>🎬 {title}</h1>
    <p>Generated by Eventure • {len(slides_html)} slides • Click any image to enlarge</p>
  </div>
</header>
<div class="gallery" id="gallery">
{gallery_items}
</div>
<!-- Lightbox -->
<div id="lb">
  <span id="lb-num"></span>
  <button id="lb-close">✕</button>
  <button id="lb-prev">&#8249;</button>
  <img id="lb-img" src="" alt=""/>
  <div class="lb-cap" id="lb-cap"></div>
  <button id="lb-next">&#8250;</button>
</div>
<script>
  const slides = Array.from(document.querySelectorAll('.slide'));
  let cur = 0;
  function openLB(i){{
    cur = i;
    const img = slides[i].querySelector('img');
    const cap = slides[i].querySelector('.caption');
    const num = slides[i].querySelector('.slide-num');
    document.getElementById('lb-img').src = img.src;
    document.getElementById('lb-cap').textContent = cap ? cap.textContent : '';
    document.getElementById('lb-num').textContent = num ? 'Slide ' + num.textContent : '';
    document.getElementById('lb').classList.add('open');
  }}
  slides.forEach((s,i) => s.addEventListener('click', () => openLB(i)));
  document.getElementById('lb-close').addEventListener('click', () =>
    document.getElementById('lb').classList.remove('open'));
  document.getElementById('lb-prev').addEventListener('click', () =>
    openLB((cur - 1 + slides.length) % slides.length));
  document.getElementById('lb-next').addEventListener('click', () =>
    openLB((cur + 1) % slides.length));
  document.addEventListener('keydown', e => {{
    if(e.key==='Escape') document.getElementById('lb').classList.remove('open');
    if(e.key==='ArrowLeft') document.getElementById('lb-prev').click();
    if(e.key==='ArrowRight') document.getElementById('lb-next').click();
  }});
</script>
</body>
</html>"""

    try:
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(html)
        if progress_callback:
            progress_callback(100)
        return True
    except Exception as e:
        print(f"HTML gallery write error: {e}")
        return False


# ── Entry Point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import Image_resizer
    Image_resizer.sync_app_folders()

    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon("logo.ico"))
    app.setStyleSheet(STYLESHEET)

    window = SlideshowCreator()
    window.create_menu()
    window.setup_connections()
    window.show()

    if len(sys.argv) > 1:
        arg = sys.argv[1]
        if arg.endswith(".slideshow") and os.path.exists(arg):
            window._load_project_from_path(arg)

    check_for_updates(window, APP_VERSION)
    sys.exit(app.exec_())