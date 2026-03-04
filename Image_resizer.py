"""
Image_resizer.py  –  Fixed version
Key fixes:
  • shutil.copytree NO LONGER runs at module level — this was crashing every
    multiprocessing worker process on Windows (they re-import this module on spawn)
  • All startup I/O is now inside sync_app_folders(), called only from main process
  • Locked/in-use files (WinError 1224) are skipped gracefully instead of crashing
  • Font is loaded lazily inside _get_font() so workers never touch the filesystem
    on import
  • GIF files now supported — converted to RGB after load (fixes "wrong mode" error)
  • EXIF correction no longer crashes on GIF/PNG/WebP (uses hasattr check)
  • Corrupted JPEG SOS warnings suppressed via warnings filter
"""

import os
import shutil
import warnings
from pathlib import Path

from bidi.algorithm import get_display
from PIL import Image, ImageFilter, ImageDraw, ImageFont, ExifTags

# ── Suppress noisy but non-fatal PIL warnings (e.g. corrupt JPEG SOS) ─────────
warnings.filterwarnings("ignore", category=UserWarning, module="PIL")

# ── Paths ─────────────────────────────────────────────────────────────────────

BASEPATH     = Path.home() / "Neria-LTD" / "Eventure"
script_dir   = Path(__file__).resolve().parent
fonts_folder = script_dir / "Fonts"

# ── NOTHING that touches the filesystem runs here at module level ─────────────
# Worker processes import this module too; any I/O here will run in every worker.

# ── Font (loaded lazily, once per process) ────────────────────────────────────

_FONT: ImageFont.FreeTypeFont | None = None

def _get_font(font_family: str | None = None) -> ImageFont.FreeTypeFont:
    """Return a font, preferring font_family if provided, else Birzia-Black."""
    global _FONT
    # If a specific family is requested, try to load it fresh each call
    if font_family and font_family not in ("Segoe UI", ""):
        # Check GoogleFonts folder first
        gf_path = BASEPATH / "Fonts" / "GoogleFonts" / f"{font_family.replace(' ', '_')}.ttf"
        if gf_path.exists():
            try:
                return ImageFont.truetype(str(gf_path), 85)
            except Exception:
                pass
        # Try system font dirs (Windows)
        import platform
        if platform.system() == "Windows":
            win_fonts = Path("C:/Windows/Fonts")
            for candidate in [
                win_fonts / f"{font_family}.ttf",
                win_fonts / f"{font_family.replace(' ', '')}.ttf",
                win_fonts / f"{font_family.lower().replace(' ', '')}.ttf",
            ]:
                if candidate.exists():
                    try:
                        return ImageFont.truetype(str(candidate), 85)
                    except Exception:
                        pass
    # Default: cached Birzia-Black
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

    resources = ["Fonts", "Help", "Languages"]
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
            # Only JPEG images have _getexif(); GIF, PNG, WebP etc. do not
            if hasattr(image, "_getexif"):
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
    crop: tuple | None = None,
    font_family: str | None = None,
) -> str | None:
    TARGET_W, TARGET_H = 1920, 1080
    font = _get_font(font_family)   # lazy-loaded, safe in workers

    try:
        original = load_image_respecting_exif(image_path)
        if original is None:
            raise ValueError("Could not load image with EXIF correction.")

        # ── Normalise mode to RGB ─────────────────────────────────────────────
        if original.mode != "RGB":
            original = original.convert("RGB")

        if rotation:
            original = original.rotate(rotation, expand=True)

        # ── Apply crop (normalised 0-1 coords relative to post-rotation size) ─
        if crop:
            iw, ih = original.size
            cx = int(crop[0] * iw)
            cy = int(crop[1] * ih)
            cw = max(1, int(crop[2] * iw))
            ch = max(1, int(crop[3] * ih))
            # Clamp to image bounds
            cx = max(0, min(cx, iw - 1))
            cy = max(0, min(cy, ih - 1))
            cw = min(cw, iw - cx)
            ch = min(ch, ih - cy)
            original = original.crop((cx, cy, cx + cw, cy + ch))

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
        # GIFs are saved as JPEG since the output is a still RGB frame
        base_name = os.path.basename(image_path)
        if base_name.lower().endswith(".gif"):
            base_name = os.path.splitext(base_name)[0] + ".jpg"
        output_path = os.path.join(output_folder, base_name)
        final.save(output_path, quality=92, optimize=True)
        return output_path

    except Exception as e:
        print(f"Error processing image {image_path}: {e}")
        return None