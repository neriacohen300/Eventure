"""
premiere_export.py  –  v2.0
Improvements over v1:
  • Generates a Final Cut Pro 7 XML timeline importable by Adobe Premiere Pro
  • Each clip gets a Ken Burns effect (zoom + pan) via keyframed filters
  • Background images placed on a lower video track (V1), foreground on V2
  • Optional background music audio track (V-audio / A1)
  • True parallel image processing via ProcessPoolExecutor
  • Accurate progress tracking
  • EXIF orientation handled correctly
  • Blurred background crop (better quality than naive stretch)
"""

import os
import uuid
import random
import xml.etree.ElementTree as ET
from xml.dom import minidom
from PIL import Image, ImageFilter, ExifTags
import concurrent.futures
from concurrent.futures import ProcessPoolExecutor, as_completed
from typing import Optional
from mutagen import File as MutagenFile



# ── EXIF helpers ─────────────────────────────────────────────────────────────

_ORIENTATION_TAG = next(
    (k for k, v in ExifTags.TAGS.items() if v == "Orientation"), None
)


def _apply_exif_rotation(image: Image.Image) -> Image.Image:
    try:
        exif = image._getexif()
        if exif and _ORIENTATION_TAG:
            val = exif.get(_ORIENTATION_TAG)
            if val == 3:
                return image.rotate(180, expand=True)
            elif val == 6:
                return image.rotate(270, expand=True)
            elif val == 8:
                return image.rotate(90, expand=True)
    except Exception:
        pass
    return image


def load_image_respecting_exif(path: str) -> Optional[Image.Image]:
    try:
        img = Image.open(path)
        return _apply_exif_rotation(img)
    except Exception as e:
        print(f"Image load failed ({path}): {e}")
        return None


# ── Single-image worker ───────────────────────────────────────────────────────

def process_single_image(
    index: int,
    img_data: dict,
    bg_folder: str,
    img_folder: str,
) -> bool:
    TARGET_W, TARGET_H = 1920, 1080

    original = load_image_respecting_exif(img_data["path"])
    if original is None:
        return False

    original = original.convert("RGBA")
    rotation = img_data.get("rotation", 0)
    if rotation:
        original = original.rotate(rotation, expand=True)

    # Apply crop (normalised 0-1 coords, relative to post-rotation size)
    crop = img_data.get("crop")
    if crop:
        iw, ih = original.size
        cx = max(0, int(crop[0] * iw))
        cy = max(0, int(crop[1] * ih))
        cw = max(1, min(int(crop[2] * iw), iw - cx))
        ch = max(1, min(int(crop[3] * ih), ih - cy))
        original = original.crop((cx, cy, cx + cw, cy + ch))

    orig_aspect = original.width / original.height
    target_aspect = TARGET_W / TARGET_H

    if orig_aspect > target_aspect:
        fit_w, fit_h = TARGET_W, int(TARGET_W / orig_aspect)
    else:
        fit_h, fit_w = TARGET_H, int(TARGET_H * orig_aspect)

    fitted = original.resize((fit_w, fit_h), Image.Resampling.BILINEAR)

    # Background: blurred + slightly darkened
    if not img_data.get("is_second_image", False):
        bg = fitted.resize((TARGET_W, TARGET_H), Image.Resampling.BILINEAR)
        bg = bg.filter(ImageFilter.GaussianBlur(radius=10))
        bg_rgb = bg.convert("RGB")
        bg_final = Image.new("RGB", (TARGET_W, TARGET_H))
        bg_final.paste(bg_rgb)
        bg_path = os.path.join(bg_folder, f"background_img{index}.jpg")
        bg_final.save(bg_path, quality=82, optimize=True, subsampling=2)

    # Foreground: 90%, centred, transparent
    fg_w = int(fit_w * 0.9)
    fg_h = int(fit_h * 0.9)
    fg = fitted.resize((fg_w, fg_h), Image.Resampling.LANCZOS)

    canvas = Image.new("RGBA", (TARGET_W, TARGET_H), (0, 0, 0, 0))
    x_off = (TARGET_W - fg_w) // 2
    y_off = (TARGET_H - fg_h) // 2
    canvas.paste(fg, (x_off, y_off), fg)

    if img_data.get("is_second_image", False):
        fg_path = os.path.join(img_folder, f"img{index}_2nd_of_img{index - 1}.png")
    else:
        fg_path = os.path.join(img_folder, f"img{index}.png")

    canvas.save(fg_path, optimize=True)
    return True


# ── Batch processor ───────────────────────────────────────────────────────────

def process_images(
    image_paths: list[dict],
    output_folder: str,
    progress_callback=None,
) -> bool:
    bg_folder  = os.path.join(output_folder, "01_images", "backgrounds")
    img_folder = os.path.join(output_folder, "01_images", "foregrounds")
    os.makedirs(bg_folder,  exist_ok=True)
    os.makedirs(img_folder, exist_ok=True)

    total = len(image_paths)
    if total == 0:
        return True

    completed = 0

    with ProcessPoolExecutor() as executor:
        future_map = {
            executor.submit(process_single_image, i, img, bg_folder, img_folder): i
            for i, img in enumerate(image_paths, start=1)
        }
        for future in as_completed(future_map):
            completed += 1
            try:
                future.result()
            except Exception as e:
                idx = future_map[future]
                print(f"Error processing image index {idx}: {e}")
            if progress_callback:
                progress_callback(int(completed / total * 100))

    return True


# ── Ken Burns keyframe helpers ────────────────────────────────────────────────

def _random_ken_burns_params() -> dict:
    """
    Returns start/end scale and position for a subtle Ken Burns move.
    Scale is a percentage zoom (100 = no zoom, 115 = 15% zoom in).
    Position offset is in pixels from centre (1920x1080).
    """
    styles = [
        # Zoom in, drift up-left
        {"scale_start": 100, "scale_end": 115, "x_start": 0,    "y_start": 0,    "x_end": -40,  "y_end": -30},
        # Zoom in, drift down-right
        {"scale_start": 100, "scale_end": 112, "x_start": 0,    "y_start": 0,    "x_end": 30,   "y_end": 25},
        # Zoom out, drift right
        {"scale_start": 115, "scale_end": 100, "x_start": -30,  "y_start": 0,    "x_end": 30,   "y_end": 0},
        # Zoom in, no drift
        {"scale_start": 100, "scale_end": 110, "x_start": 0,    "y_start": 0,    "x_end": 0,    "y_end": 0},
        # Zoom out from top-right
        {"scale_start": 118, "scale_end": 100, "x_start": 40,   "y_start": -25,  "x_end": 0,    "y_end": 0},
    ]
    return random.choice(styles)


# ── FCP7 XML builder ──────────────────────────────────────────────────────────

FPS        = 25          # timeline frame rate
TIMEBASE   = str(FPS)

def _frames(seconds: float) -> int:
    return int(round(seconds * FPS))

def _add_keyframe(parent: ET.Element, time: int, value: str) -> None:
    kf = ET.SubElement(parent, "keyframe")
    ET.SubElement(kf, "when").text = str(time)
    ET.SubElement(kf, "value").text = value
    ET.SubElement(kf, "interp").text = "linear"


def _make_clip_id() -> str:
    return "clip-" + uuid.uuid4().hex[:8]


def _build_video_clip(
    clip_id: str,
    file_path: str,
    start: int,        # timeline in-point (frames)
    duration: int,     # duration (frames)
    track_index: int,  # 1 = background, 2 = foreground
    ken_burns: Optional[dict] = None,
) -> ET.Element:
    """Build a <clipitem> element for the XML timeline."""
    item = ET.Element("clipitem", id=clip_id)
    ET.SubElement(item, "name").text = os.path.basename(file_path)
    ET.SubElement(item, "duration").text   = str(duration)
    ET.SubElement(item, "start").text      = str(start)
    ET.SubElement(item, "end").text        = str(start + duration)
    ET.SubElement(item, "in").text         = "0"
    ET.SubElement(item, "out").text        = str(duration)

    # File reference
    file_el = ET.SubElement(item, "file", id="file-" + clip_id)
    ET.SubElement(file_el, "name").text    = os.path.basename(file_path)
    ET.SubElement(file_el, "pathurl").text = "file:" + file_path.replace("\\", "/")
    ET.SubElement(file_el, "duration").text = str(duration)
    rate = ET.SubElement(file_el, "rate")
    ET.SubElement(rate, "timebase").text   = TIMEBASE
    ET.SubElement(rate, "ntsc").text       = "FALSE"
    media = ET.SubElement(file_el, "media")
    video = ET.SubElement(media, "video")
    ET.SubElement(video, "duration").text  = str(duration)

    # Ken Burns via Motion filter
    if ken_burns:
        kb = ken_burns
        filters = ET.SubElement(item, "filters")

        # ── Scale (Motion > Scale) ────────────────────────────────────────
        scale_filter = ET.SubElement(filters, "filter")
        ET.SubElement(scale_filter, "name").text   = "Motion"
        ET.SubElement(scale_filter, "effectid").text = "basic motion"
        ET.SubElement(scale_filter, "effecttype").text = "motion"
        scale_param = ET.SubElement(scale_filter, "parameter")
        ET.SubElement(scale_param, "parameterid").text = "scale"
        ET.SubElement(scale_param, "name").text = "Scale"
        ET.SubElement(scale_param, "valuemin").text = "0"
        ET.SubElement(scale_param, "valuemax").text = "10000"
        kfs = ET.SubElement(scale_param, "keyframe_list")
        _add_keyframe(kfs, 0,        str(kb["scale_start"]))
        _add_keyframe(kfs, duration, str(kb["scale_end"]))

        # ── Centre offset X ───────────────────────────────────────────────
        px_filter = ET.SubElement(filters, "filter")
        ET.SubElement(px_filter, "name").text = "Motion"
        ET.SubElement(px_filter, "effectid").text = "basic motion"
        ET.SubElement(px_filter, "effecttype").text = "motion"
        px_param = ET.SubElement(px_filter, "parameter")
        ET.SubElement(px_param, "parameterid").text = "cenx"
        ET.SubElement(px_param, "name").text = "Centre X"
        kfsx = ET.SubElement(px_param, "keyframe_list")
        _add_keyframe(kfsx, 0,        str(kb["x_start"]))
        _add_keyframe(kfsx, duration, str(kb["x_end"]))

        # ── Centre offset Y ───────────────────────────────────────────────
        py_filter = ET.SubElement(filters, "filter")
        ET.SubElement(py_filter, "name").text = "Motion"
        ET.SubElement(py_filter, "effectid").text = "basic motion"
        ET.SubElement(py_filter, "effecttype").text = "motion"
        py_param = ET.SubElement(py_filter, "parameter")
        ET.SubElement(py_param, "parameterid").text = "ceny"
        ET.SubElement(py_param, "name").text = "Centre Y"
        kfsy = ET.SubElement(py_param, "keyframe_list")
        _add_keyframe(kfsy, 0,        str(kb["y_start"]))
        _add_keyframe(kfsy, duration, str(kb["y_end"]))

    return item


def get_audio_duration_frames(path: str) -> int:
    audio = MutagenFile(path)
    return _frames(audio.info.length) if audio else _frames(FPS)


def _build_audio_clip(
    music_path: str,
    start_frame: int,
    duration_frames: int,
    total_sequence_frames: int,
) -> ET.Element:
    clip_id = _make_clip_id()
    item = ET.Element("clipitem", id=clip_id)
    ET.SubElement(item, "name").text       = os.path.basename(music_path)
    ET.SubElement(item, "duration").text   = str(get_audio_duration_frames(str(music_path)))
    ET.SubElement(item, "start").text      = str(start_frame)
    ET.SubElement(item, "end").text        = str(start_frame + get_audio_duration_frames(str(music_path)))
    ET.SubElement(item, "in").text         = "0"
    ET.SubElement(item, "out").text        = str(get_audio_duration_frames(str(music_path)))

    file_el = ET.SubElement(item, "file", id="file-" + clip_id)
    ET.SubElement(file_el, "name").text    = os.path.basename(music_path)
    ET.SubElement(file_el, "pathurl").text = "file:" + music_path.replace("\\", "/")
    ET.SubElement(file_el, "duration").text = str(total_sequence_frames)
    rate = ET.SubElement(file_el, "rate")
    ET.SubElement(rate, "timebase").text   = TIMEBASE
    ET.SubElement(rate, "ntsc").text       = "FALSE"
    media = ET.SubElement(file_el, "media")
    audio_m = ET.SubElement(media, "audio")
    ET.SubElement(audio_m, "duration").text = str(total_sequence_frames)

    is_last = (start_frame + get_audio_duration_frames(str(music_path)) >= total_sequence_frames)
    if is_last:
        fade_start = max(0, get_audio_duration_frames(str(music_path)) - _frames(2))
        filters = ET.SubElement(item, "filters")
        vol_filter = ET.SubElement(filters, "filter")
        ET.SubElement(vol_filter, "name").text     = "Audio Levels"
        ET.SubElement(vol_filter, "effectid").text = "audiolevels"
        vol_param = ET.SubElement(vol_filter, "parameter")
        ET.SubElement(vol_param, "parameterid").text = "level"
        ET.SubElement(vol_param, "name").text        = "Level"
        kf_vol = ET.SubElement(vol_param, "keyframe_list")
        _add_keyframe(kf_vol, 0,                    "100")
        _add_keyframe(kf_vol, fade_start,           "100")
        _add_keyframe(kf_vol, total_sequence_frames, "0")

    return item


def generate_premiere_xml(
    slide_list: list[dict],
    output_folder: str,
    music_paths: list[str] = None,   # ← now a list
    default_duration_sec: float = 5.0,
    transition_duration_sec: float = 1.0,
) -> str:
    """
    Build an FCP7 XML file importable by Adobe Premiere Pro.

    slide_list items:
        {
          "bg_path":      str,    # path to background JPEG (None for second images)
          "fg_path":      str,    # path to foreground PNG
          "duration":     float,  # per-slide duration in seconds
          "text":         str,    # optional (for reference)
          "is_second_image": bool # if True, overlaid as PiP on the previous slide (V3)
        }

    Second images are placed on V3 at the same timeline position as their
    parent slide so they appear as a picture-in-picture overlay.

    Returns the path to the saved .xml file.
    """
    seq_id  = "seq-" + uuid.uuid4().hex[:8]

    # Only primary (non-PiP) slides count toward total sequence duration
    primary_slides = [s for s in slide_list if not s.get("is_second_image", False)]
    seq_dur = sum(
        _frames(s.get("duration", default_duration_sec)) for s in primary_slides
    )

    # ── Root ──────────────────────────────────────────────────────────────────
    root = ET.Element("xmeml", version="5")
    project = ET.SubElement(root, "project")
    ET.SubElement(project, "name").text = "Slideshow"

    children = ET.SubElement(project, "children")
    seq = ET.SubElement(children, "sequence", id=seq_id)
    ET.SubElement(seq, "name").text     = "Slideshow Sequence"
    ET.SubElement(seq, "duration").text = str(seq_dur)

    rate = ET.SubElement(seq, "rate")
    ET.SubElement(rate, "timebase").text = TIMEBASE
    ET.SubElement(rate, "ntsc").text     = "FALSE"

    media = ET.SubElement(seq, "media")
    video = ET.SubElement(media, "video")

    # Format track
    fmt_track = ET.SubElement(video, "format")
    sample_chars = ET.SubElement(fmt_track, "samplecharacteristics")
    rate2 = ET.SubElement(sample_chars, "rate")
    ET.SubElement(rate2, "timebase").text = TIMEBASE
    ET.SubElement(rate2, "ntsc").text     = "FALSE"
    ET.SubElement(sample_chars, "width").text  = "1920"
    ET.SubElement(sample_chars, "height").text = "1080"
    ET.SubElement(sample_chars, "anamorphic").text   = "FALSE"
    ET.SubElement(sample_chars, "pixelaspectratio").text = "square"
    ET.SubElement(sample_chars, "fielddominance").text   = "none"

    # V1 = backgrounds, V2 = foregrounds, V3 = PiP second images
    track_bg  = ET.SubElement(video, "track")   # V1
    track_fg  = ET.SubElement(video, "track")   # V2
    track_pip = ET.SubElement(video, "track")   # V3 – second images (PiP)

    cursor       = 0   # timeline cursor in frames (primary slides only)
    prev_cursor  = 0   # cursor before the last primary slide (for PiP placement)

    # We need to iterate slide_list in order; keep track of the parent slide's
    # start frame so PiP images can be placed at the same position.
    last_primary_start = 0

    for i, slide in enumerate(slide_list):
        is_pip = slide.get("is_second_image", False)
        dur    = _frames(slide.get("duration", default_duration_sec))
        kb     = _random_ken_burns_params()

        if is_pip:
            # ── PiP slide: goes on V3 at the same start as its parent ─────────
            # In Premiere's FCP7 XML, position is in pixels offset from the
            # frame centre (960, 540).  For a true side-by-side layout:
            #   • Primary image (V2) will be repositioned to the LEFT half.
            #   • PiP image (V3) sits in the RIGHT half.
            # Each half is 960 px wide.  At scale=45 a 1920-wide frame fits
            # exactly in 864 px, which comfortably fills a 960-wide cell.
            # We use scale=47 so there's a small breathing margin.
            #
            # FCP7 cenx/ceny are offsets from the frame centre (in pixels,
            # positive = right / down).
            #   Left  centre = frame_centre_x - cell_w/2 = 960 - 480 = 480 → cenx = -480
            #   Right centre = frame_centre_x + cell_w/2 = 960 + 480 = 480 → cenx = +480
            pip_start = last_primary_start
            if slide.get("fg_path"):
                pip_clip = _build_video_clip(
                    clip_id     = _make_clip_id(),
                    file_path   = slide["fg_path"],
                    start       = pip_start,
                    duration    = dur,
                    track_index = 3,
                    ken_burns   = {
                        # Static: scale to 47 %, positioned in right half
                        "scale_start": 47,  "scale_end": 47,
                        "x_start":     480, "x_end":    480,   # right half
                        "y_start":     0,   "y_end":    0,
                    },
                )
                track_pip.append(pip_clip)

            # Also reposition the V2 primary foreground for this slide to the
            # LEFT half by patching the last-appended fg clip's ken_burns params.
            # We do this by replacing the last fg clip with a repositioned version.
            if track_fg:
                # Remove the last fg clip (the parent primary slide just added)
                last_fg = track_fg[-1]
                track_fg.remove(last_fg)
                # Re-add it with left-half positioning
                parent_slide  = slide_list[i - 1]   # the slide just before this PiP
                parent_dur_fr = _frames(parent_slide.get("duration", default_duration_sec))
                if parent_slide.get("fg_path"):
                    left_clip = _build_video_clip(
                        clip_id     = _make_clip_id(),
                        file_path   = parent_slide["fg_path"],
                        start       = last_primary_start,
                        duration    = parent_dur_fr,
                        track_index = 2,
                        ken_burns   = {
                            # Static: scale to 47 %, positioned in left half
                            "scale_start": 47,  "scale_end": 47,
                            "x_start":    -480, "x_end":   -480,   # left half
                            "y_start":     0,   "y_end":    0,
                        },
                    )
                    track_fg.append(left_clip)
            # PiP slides do not advance the cursor
        else:
            # ── Primary slide ─────────────────────────────────────────────────
            last_primary_start = cursor

            # Background clip (V1) – static
            if slide.get("bg_path"):
                bg_clip = _build_video_clip(
                    clip_id     = _make_clip_id(),
                    file_path   = slide["bg_path"],
                    start       = cursor,
                    duration    = dur,
                    track_index = 1,
                    ken_burns   = None,
                )
                track_bg.append(bg_clip)

            # Foreground clip (V2) – Ken Burns effect
            if slide.get("fg_path"):
                fg_clip = _build_video_clip(
                    clip_id     = _make_clip_id(),
                    file_path   = slide["fg_path"],
                    start       = cursor,
                    duration    = dur,
                    track_index = 2,
                    ken_burns   = kb,
                )
                track_fg.append(fg_clip)

            cursor += dur

    # ── Audio track (background music) ────────────────────────────────────────
    audio = ET.SubElement(media, "audio")
    if music_paths:
        audio_track = ET.SubElement(audio, "track")
        cursor_audio = 0
        for item in music_paths:
            path = item["path"] if isinstance(item, dict) else item
            if not os.path.exists(path):
                continue
            if cursor_audio >= seq_dur:
                break
            audio_dur = get_audio_duration_frames(path)  # real file duration
            audio_clip = _build_audio_clip(path, cursor_audio, audio_dur, seq_dur)
            audio_track.append(audio_clip)
            cursor_audio += audio_dur 

    # ── Pretty-print & save ───────────────────────────────────────────────────
    xml_str = minidom.parseString(
        ET.tostring(root, encoding="unicode")
    ).toprettyxml(indent="  ")

    # Remove the extra <?xml ...?> declaration minidom adds (keep only first line)
    lines = xml_str.split("\n")
    clean = "\n".join(lines)

    xml_path = os.path.join(output_folder, "premiere_timeline.xml")
    with open(xml_path, "w", encoding="utf-8") as f:
        f.write(clean)

    print(f"Premiere XML saved at: {xml_path}")
    return xml_path


# ── High-level entry point ────────────────────────────────────────────────────

def export_slideshow(
    image_paths: list[dict],
    output_folder: str,
    music_path: list[str] = None,
    default_duration_sec: float = 5.0,
    progress_callback=None,
) -> str:
    """
    Full pipeline:
      1. Process images (resize, background, foreground)
      2. Generate Premiere XML timeline

    Returns path to the generated XML file.
    """
    bg_folder  = os.path.join(output_folder, "01_images", "backgrounds")
    img_folder = os.path.join(output_folder, "01_images", "foregrounds")

    # Step 1 – process images
    process_images(image_paths, output_folder, progress_callback)

    # Step 2 – build slide list from processed files
    slide_list = []
    for i, img in enumerate(image_paths, start=1):
        if img.get("is_second_image"):
            fg = os.path.join(img_folder, f"img{i}_2nd_of_img{i-1}.png")
            slide_list.append({
                "bg_path":  None,
                "fg_path":  fg if os.path.exists(fg) else None,
                "duration": img.get("duration", default_duration_sec),
                "text":     img.get("text", ""),
            })
        else:
            bg = os.path.join(bg_folder,  f"background_img{i}.jpg")
            fg = os.path.join(img_folder, f"img{i}.png")
            slide_list.append({
                "bg_path":  bg if os.path.exists(bg) else None,
                "fg_path":  fg if os.path.exists(fg) else None,
                "duration": img.get("duration", default_duration_sec),
                "text":     img.get("text", ""),
            })

    # Step 3 – generate XML
    xml_path = generate_premiere_xml(
        slide_list          = slide_list,
        output_folder       = output_folder,
        music_path          = music_path,
        default_duration_sec= default_duration_sec,
    )

    return xml_path