"""
slideshow_io.py  –  Unified save / load for .slideshow project files
=====================================================================

FORMAT v2  (JSON, introduced in Eventure 1.0.8)
---------------------------------------------
{
  "version": 2,
  "created": "2025-03-05 10:00:00",
  "audio": [
    {"path": "C:/music/song.mp3"}
  ],
  "slides": [
    {
      "path":               "C:/photos/img.jpg",
      "duration":           5.0,
      "transition":         "fade",
      "transition_duration": 1.0,
      "text":               "Hello world",
      "rotation":           0,
      "is_second_image":    false,
      "date":               "2025-03-05 10:00:00",
      "ken_burns":          "zoom_in",
      "text_on_kb":         true,
      "crop":               [0.1, 0.05, 0.8, 0.9]   // null when no crop
    }
  ]
}

FORMAT v1  (legacy CSV, read-only — written by older Eventure versions)
------------------------------------------------------------------------
<audio_count>
<audio_path_1>
...
<image_path>,<dur>,<transition>,<trans_dur>,<text>,<rot>,<is_second>,<date>,<ken_burns>,<text_on_kb>,<crop>
...

Backward-compat rules
---------------------
* load() auto-detects the format by peeking at the first character.
  '{' → JSON v2.   digit / anything else → legacy CSV v1.
* save() always writes JSON v2.
* Legacy files are never modified in-place; they only produce a v2 file when
  the user performs "Save" or "Save As".
"""

from __future__ import annotations

import json
import os
from datetime import datetime
from pathlib import Path
from typing import Any

# ── Public constants ──────────────────────────────────────────────────────────

SLIDESHOW_VERSION = 2
SLIDESHOW_EXT     = ".slideshow"


# ── Internal helpers ──────────────────────────────────────────────────────────

def _crop_to_list(crop) -> list | None:
    """Convert a crop tuple/list to a plain list for JSON, or None."""
    if crop is None:
        return None
    try:
        vals = list(crop)
        if len(vals) == 4:
            return [round(float(v), 6) for v in vals]
    except (TypeError, ValueError):
        pass
    return None


def _parse_crop_str(s: str) -> tuple | None:
    """Parse legacy '0.1|0.05|0.8|0.9' crop string → tuple, or None."""
    if not s or s.strip().lower() in ("none", ""):
        return None
    try:
        vals = [float(v) for v in s.strip().split("|")]
        if len(vals) == 4:
            return tuple(vals)
    except ValueError:
        pass
    return None


def _parse_crop_value(v) -> tuple | None:
    """Accept either a list/tuple (JSON) or a '|'-delimited string (legacy)."""
    if v is None:
        return None
    if isinstance(v, (list, tuple)):
        try:
            vals = [float(x) for x in v]
            if len(vals) == 4:
                return tuple(vals)
        except (TypeError, ValueError):
            pass
        return None
    return _parse_crop_str(str(v))


# ── Save ──────────────────────────────────────────────────────────────────────

def save(
    path: str | Path,
    audio_files: list[dict],
    images: list[dict],
) -> None:
    """
    Write *audio_files* and *images* to *path* in JSON v2 format.

    Parameters
    ----------
    path        : Destination .slideshow file path.
    audio_files : List of dicts with at least a 'path' key.
    images      : List of image dicts as used internally by Eventure.
    """
    slides = []
    for img in images:
        text = img.get("text", "")
        # Normalise newlines to literal \n in JSON strings (they round-trip fine)
        slides.append({
            "path":                img.get("path", ""),
            "duration":            float(img.get("duration", 5)),
            "transition":          img.get("transition", "fade"),
            "transition_duration": float(img.get("transition_duration", 1)),
            "text":                text,
            "rotation":            int(img.get("rotation", 0)),
            "is_second_image":     bool(img.get("is_second_image", False)),
            "date":                img.get("date", ""),
            "ken_burns":           img.get("ken_burns", "none"),
            "text_on_kb":          bool(img.get("text_on_kb", True)),
            "crop":                _crop_to_list(img.get("crop")),
            # future-proof: any extra keys already in the dict are silently
            # dropped here so we never accidentally corrupt the file with
            # Qt-internal objects, etc.
        })

    data: dict[str, Any] = {
        "version": SLIDESHOW_VERSION,
        "created": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "audio":   [{"path": a["path"]} for a in audio_files],
        "slides":  slides,
    }

    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


# ── Load ──────────────────────────────────────────────────────────────────────

def load(
    path: str | Path,
    default_transition_duration: float = 1.0,
) -> dict:
    """
    Load a .slideshow file (v1 CSV or v2 JSON) and return a normalised dict:

        {
            "audio_files": [{"path": ...}, ...],
            "images":      [{...}, ...],
        }

    Raises
    ------
    ValueError  if the file is truncated or cannot be parsed.
    OSError     if the file cannot be opened.
    """
    path = Path(path)
    with open(path, "r", encoding="utf-8") as f:
        first_char = f.read(1)
        rest = f.read()

    full_text = first_char + rest

    if first_char == "{":
        return _load_v2(full_text, default_transition_duration)
    else:
        return _load_v1(full_text, default_transition_duration)


# ── v2 JSON loader ────────────────────────────────────────────────────────────

def _load_v2(text: str, default_td: float) -> dict:
    try:
        data = json.loads(text)
    except json.JSONDecodeError as e:
        raise ValueError(f"Invalid JSON in .slideshow file: {e}") from e

    audio_files = [{"path": a["path"]} for a in data.get("audio", [])]
    images = []
    for s in data.get("slides", []):
        images.append({
            "path":                s.get("path", ""),
            "duration":            float(s.get("duration", 5)),
            "transition":          s.get("transition", "fade"),
            "transition_duration": float(s.get("transition_duration", default_td)),
            "text":                s.get("text", ""),
            "rotation":            int(s.get("rotation", 0)),
            "is_second_image":     bool(s.get("is_second_image", False)),
            "date":                s.get("date", ""),
            "ken_burns":           s.get("ken_burns", "none"),
            "text_on_kb":          bool(s.get("text_on_kb", True)),
            "crop":                _parse_crop_value(s.get("crop")),
        })
    return {"audio_files": audio_files, "images": images}


# ── v1 CSV (legacy) loader ────────────────────────────────────────────────────

def _load_v1(text: str, default_td: float) -> dict:
    lines = text.splitlines()
    if not lines:
        raise ValueError("Project file is empty.")

    try:
        count = int(lines[0].strip())
    except ValueError:
        raise ValueError("Project file header is not a valid integer.")

    if len(lines) < count + 1:
        raise ValueError("Project file is truncated (audio section).")

    audio_files = [{"path": lines[i + 1].strip()} for i in range(count)]
    images = []
    for line in lines[count + 1:]:
        parts = line.strip().split(",")
        if len(parts) < 8:
            continue
        path       = parts[0]
        dur        = parts[1]
        transition = parts[2]
        trans_dur  = parts[3]
        text       = parts[4]
        rotation   = parts[5]
        is_second  = parts[6]
        date       = parts[7] if len(parts) > 7 else ""
        ken_burns  = parts[8].strip() if len(parts) > 8 else "none"
        text_on_kb = parts[9].strip().lower() != "false" if len(parts) > 9 else True
        crop       = _parse_crop_str(parts[10]) if len(parts) > 10 else None
        images.append({
            "path":                path,
            "duration":            float(dur),
            "transition":          transition,
            "transition_duration": default_td,
            "text":                text.replace("\\n", "\n"),
            "rotation":            int(rotation),
            "is_second_image":     is_second.strip().lower() == "true",
            "date":                date,
            "ken_burns":           ken_burns,
            "text_on_kb":          text_on_kb,
            "crop":                crop,
        })
    return {"audio_files": audio_files, "images": images}


# ── pptx_export helper ────────────────────────────────────────────────────────

def write_from_pptx(
    slideshow_path: str | Path,
    slides: list[dict],
) -> str:
    """
    Write a brand-new v2 .slideshow file from PPTX-extracted slide data.

    Each dict in *slides* must have:
        image   : str   – absolute path to the extracted image file
        text    : str   – slide text (may be empty)
        index   : int   – 0-based index within the slide's image list
                          (used to decide is_second_image)

    Returns the path as a string.
    """
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    slide_records = []
    for s in slides:
        slide_records.append({
            "path":                s["image"],
            "duration":            5.0,
            "transition":          "fade",
            "transition_duration": 1.0,
            "text":                s.get("text", ""),
            "rotation":            0,
            "is_second_image":     s.get("index", 0) == 1,
            "date":                now,
            "ken_burns":           "none",
            "text_on_kb":          True,
            "crop":                None,
        })

    data: dict[str, Any] = {
        "version": SLIDESHOW_VERSION,
        "created": now,
        "audio":   [],
        "slides":  slide_records,
    }

    slideshow_path = Path(slideshow_path)
    slideshow_path.parent.mkdir(parents=True, exist_ok=True)
    with open(slideshow_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    return str(slideshow_path)