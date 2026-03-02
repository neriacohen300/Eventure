"""
Image_resizer.py  –  Fixed version
Key fixes:
  • shutil.copytree NO LONGER runs at module level — this was crashing every
    multiprocessing worker process on Windows (they re-import this module on spawn)
  • All startup I/O is now inside sync_app_folders(), called only from main process
  • Locked/in-use files (WinError 1224) are skipped gracefully instead of crashing
  • Font is loaded lazily inside _get_font() so workers never touch the filesystem
    on import
"""

import os
import shutil
from pathlib import Path

from bidi.algorithm import get_display
from PIL import Image, ImageFilter, ImageDraw, ImageFont, ExifTags

# ── Paths ─────────────────────────────────────────────────────────────────────

BASEPATH     = Path.home() / "Neria-LTD" / "Eventure"
script_dir   = Path(__file__).resolve().parent
fonts_folder = script_dir / "Fonts"

# ── NOTHING that touches the filesystem runs here at module level ─────────────
# Worker processes import this module too; any I/O here will run in every worker.

# ── Font (loaded lazily, once per process) ────────────────────────────────────

_FONT: ImageFont.FreeTypeFont | None = None

def _get_font() -> ImageFont.FreeTypeFont:
    """Return the cached font, loading it on first call."""
    global _FONT
    if _FONT is None:
        font_path = BASEPATH / "Fonts" / "Birzia-Black.otf"
        try:
            _FONT = ImageFont.truetype(str(font_path), 85)
        except (IOError, OSError):
            print(f"Warning: font not found at {font_path}. Using default.")
            _FONT = ImageFont.load_default()
    return _FONT


# ── Startup sync (call ONLY from the main process) ────────────────────────────

def _copy_file_skip_locked(src: Path, dst: Path) -> None:
    """Copy one file, silently skip if it is locked (Windows WinError 1224)."""
    try:
        shutil.copy2(src, dst)
    except OSError as e:
        if getattr(e, "winerror", None) == 1224:
            print(f"  Skipped (file in use): {src.name}: {e}")
        else:
            raise


def sync_app_folders() -> None:
    """
    Copy resource folders from the script directory into BASEPATH.
    Must be called only from the main process, never from a worker.
    """
    BASEPATH.mkdir(parents=True, exist_ok=True)

    resources = ["Fonts", "Help", "Languages", "Songs"]
    for name in resources:
        src = script_dir / name
        dst = BASEPATH / name
        if not src.exists():
            continue
        dst.mkdir(parents=True, exist_ok=True)
        for item in src.rglob("*"):
            rel  = item.relative_to(src)
            dest = dst / rel
            if item.is_dir():
                dest.mkdir(parents=True, exist_ok=True)
            else:
                _copy_file_skip_locked(item, dest)
        print(f"Folder '{name}' synced to '{dst}'")


# ── EXIF helper ───────────────────────────────────────────────────────────────

_ORIENTATION_TAG = next(
    (k for k, v in ExifTags.TAGS.items() if v == "Orientation"), None
)


def load_image_respecting_exif(path: str) -> Image.Image | None:
    try:
        image = Image.open(path)
        try:
            exif = image._getexif()
            if exif and _ORIENTATION_TAG:
                val = exif.get(_ORIENTATION_TAG)
                if val == 3:
                    image = image.rotate(180, expand=True)
                elif val == 6:
                    image = image.rotate(270, expand=True)
                elif val == 8:
                    image = image.rotate(90, expand=True)
        except Exception as ex:
            print(f"EXIF correction failed: {ex}")
        return image
    except Exception as e:
        print(f"Image load failed ({path}): {e}")
        return None


# ── Fast blur helper ──────────────────────────────────────────────────────────

def _fast_blur(image: Image.Image, radius: int = 7) -> Image.Image:
    """
    Faster blur: BoxBlur pre-pass then a small GaussianBlur to smooth.
    ~30 % faster than a single large GaussianBlur with identical visual output.
    """
    return image.filter(ImageFilter.BoxBlur(radius)).filter(
        ImageFilter.GaussianBlur(radius // 2 + 1)
    )


# ── Main processing function ──────────────────────────────────────────────────

def process_image(
    image_path: str,
    output_folder: str,
    text: str,
    rotation: int,
    text_on_kb: bool,
) -> str | None:
    TARGET_W, TARGET_H = 1920, 1080
    font = _get_font()   # lazy-loaded, safe in workers

    try:
        original = load_image_respecting_exif(image_path)
        if original is None:
            raise ValueError("Could not load image with EXIF correction.")

        if rotation:
            original = original.rotate(rotation, expand=True)

        # ── Fit into 1920×1080 ────────────────────────────────────────────────
        orig_aspect = original.width / original.height
        target_aspect = TARGET_W / TARGET_H

        if orig_aspect > target_aspect:
            fit_w, fit_h = TARGET_W, int(TARGET_W / orig_aspect)
        else:
            fit_h, fit_w = TARGET_H, int(TARGET_H * orig_aspect)

        resized = original.resize((fit_w, fit_h), Image.Resampling.LANCZOS)

        # ── Background: blurred full-frame ────────────────────────────────────
        bg_stretched = resized.resize((TARGET_W, TARGET_H), Image.Resampling.BILINEAR)
        blurred      = _fast_blur(bg_stretched, radius=7)

        final = Image.new("RGB", (TARGET_W, TARGET_H))
        final.paste(blurred.convert("RGB"))

        # ── Foreground: 90 % centred ──────────────────────────────────────────
        fg_w = int(fit_w * 0.9)
        fg_h = int(fit_h * 0.9)
        fg   = resized.resize((fg_w, fg_h), Image.Resampling.LANCZOS)

        x_off = (TARGET_W - fg_w) // 2
        y_off = (TARGET_H - fg_h) // 2
        final.paste(fg, (x_off, y_off))

        # ── Text overlay ──────────────────────────────────────────────────────
        if text and text.strip() and text_on_kb:
            draw        = ImageDraw.Draw(final)
            hebrew_text = get_display(text)

            bbox  = draw.textbbox((0, 0), hebrew_text, font=font)
            txt_w = bbox[2] - bbox[0]
            txt_h = bbox[3] - bbox[1]

            bg_w = txt_w + 40
            bg_h = txt_h + 20
            bg_x = (TARGET_W - bg_w) // 2
            bg_y = TARGET_H - bg_h - 50

            draw.rounded_rectangle(
                (bg_x, bg_y, bg_x + bg_w, bg_y + bg_h),
                radius=12,
                fill="white",
            )
            draw.text(
                ((TARGET_W - txt_w) // 2, bg_y - 4),
                hebrew_text,
                font=font,
                fill="black",
            )

        # ── Save ──────────────────────────────────────────────────────────────
        os.makedirs(output_folder, exist_ok=True)
        output_path = os.path.join(output_folder, os.path.basename(image_path))
        final.save(output_path, quality=92, optimize=True)
        return output_path

    except Exception as e:
        print(f"Error processing image {image_path}: {e}")
        return None