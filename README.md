# 🎬 Eventure — Slideshow Creator

![Python](https://img.shields.io/badge/Python-3.10%2B-blue?style=flat&logo=python)
![PyQt5](https://img.shields.io/badge/Framework-PyQt5-lightgrey?style=flat&logo=pyqt5)
![Version](https://img.shields.io/badge/Version-1.0.5-informational?style=flat)
![Status](https://img.shields.io/badge/Status-Active_Development-green)
![License](https://img.shields.io/badge/License-MIT-yellow?style=flat)


<img width="1920" height="600" alt="Eventure-Banner" src="https://github.com/user-attachments/assets/73462cb1-0f8a-4e00-ab22-bc73f1834d6e" />
<img width="1396" height="850" alt="image" src="https://github.com/user-attachments/assets/a080ea24-2a3c-49d5-81d3-951a11b01917" />







## ✨ Features

### 🖼️ Image Management
- Import images in bulk (JPG, PNG, GIF, WebP, and more)
- Drag-and-drop reordering via an interactive filmstrip timeline
- Per-image customisable display duration
- EXIF-aware image loading (auto-corrects rotation)
- Right-click context menu on timeline cards

### 🎞️ Transitions & Ken Burns
- Set transition type and duration per image or for all images at once
- Random transition assignment across the slideshow
- **Ken Burns effects**: pan, zoom, and smart auto-detection
  - Set all images to the same Ken Burns effect
  - Assign a random Ken Burns effect per image
  - Smart Ken Burns mode for automatic effect selection

### 🔤 Text Overlays
- Add text captions to individual slides
- Full **RTL (right-to-left)** language support via `python-bidi`
- Custom font rendering with `Birzia-Black.otf`
- Easy-text writing dialog for rapid caption entry

### 🎵 Audio
- Import and manage multiple audio files
- Built-in audio library browser with search and duration info
- Audio duration displayed in the stats bar
- Audio files exported alongside Premiere projects

### 📤 Export Options

| Format | Details |
|---|---|
| **Video** | Renders each slide as a processed image with Ken Burns clips via ffmpeg |
| **Adobe Premiere Pro** | Full project export: centered images, audio, SRT subtitle file, and duration Excel sheet |
| **PowerPoint (.pptx)** | Import `.pptx` files and convert slides to the Eventure `.slideshow` format |
| **Slideshow file** | Save/load projects as `.slideshow` files |

### 🌐 Internationalisation
- Full multi-language UI (English and Hebrew included)
- Language files stored as `.json` in the `Languages/` folder
- Switchable at runtime via the Settings menu

### 🎨 Themes & UI
- Dark cinema aesthetic with a card-based layout
- Modern toolbar, sidebar preview panel, and filmstrip timeline
- Stats bar showing slide count, total duration, and audio count
- Dual progress bars for image processing and Premiere export
- Customisable keyboard shortcuts

---

## 📁 Project Structure

```
Eventure/
├── Eventure.py            # Main application entry point & UI
├── Image_resizer.py       # Image processing (resize, blur bg, text overlay)
├── premiere_export.py     # Adobe Premiere Pro XML project generator
├── pptx_export.py         # PowerPoint import & .slideshow converter
├── Fonts/                 # Bundled fonts (e.g. Birzia-Black.otf)
├── Help/                  # Help documentation (per language)
├── Languages/             # Translation JSON files
├── Premiere_Project/      # Premiere Pro project template
├── Songs/                 # Bundled audio library
├── logo.ico               # Application icon
└── requirements.txt       # Python dependencies
```

---

## 🚀 Getting Started

### Prerequisites

- Python 3.10+
- [ffmpeg](https://ffmpeg.org/) installed and available on your system `PATH`
- (Optional) Adobe Premiere Pro for `.prproj` export workflow

### Installation

```bash
# Clone the repository
git clone https://github.com/neriacohen300/Eventure.git
cd Eventure

# Install dependencies
pip install -r requirements.txt

# Run the application
python Eventure.py
```

### Or download the installer

Download the latest `setup.exe` from the [Releases](https://github.com/neriacohen300/Eventure/releases) page.
Optionally, also download the Adobe Premiere Pro preset pack for the best results.

---

## 🎯 Recommended Workflow (with Adobe Premiere Pro)

1. Import your images into Eventure and arrange them on the filmstrip.
2. Set durations, transitions, and Ken Burns effects per slide.
3. Add text captions and audio files as needed.
4. Click **Export → Adobe Premiere** — Eventure will generate a folder containing:
   - `01_תמונות/` — Centred and resized images (1920×1080)
   - `02_אודיו/` — Your audio files
   - `03_טקסט/` — SRT subtitle file + Excel duration sheet
   - `04_פרוייקט/` — Ready-to-open `.prproj` file
5. Open the project in Premiere Pro, apply your preset effects, and render.

---

## 🖼️ Image Processing Pipeline (`Image_resizer.py`)

Each image is processed to a **1920×1080** frame using the following pipeline:

1. Load image with EXIF orientation correction
2. Convert to RGB (handles GIF, RGBA, palette-mode images)
3. Apply optional rotation
4. Fit image proportionally within 1920×1080
5. Generate a blurred, stretched background (BoxBlur + GaussianBlur)
6. Composite the foreground at 90% scale, centred
7. Optionally render a rounded-rectangle text caption overlay
8. Save as JPEG (quality 92, optimised)

Processing runs in parallel using Python's `multiprocessing` / `ThreadPoolExecutor` for performance.

---

## 📊 PowerPoint Import (`pptx_export.py`)

Eventure can import `.pptx` / `.pptm` files and convert them into its native `.slideshow` format:

- Extracts all images from each slide
- Extracts text content (skips text overlapping images)
- Cleans illegal control characters for safe file output
- Generates a `.slideshow` file with 5-second default duration and fade transition per slide
- Can also export slide content to a structured Excel (`.xlsx`) file

---

## ⌨️ Keyboard Shortcuts

| Action | Default Shortcut |
|---|---|
| Import Audio | `Ctrl+Shift+A` |
| Import Images | `Ctrl+Shift+I` |
| Save | `Ctrl+S` |
| Save As | `Ctrl+Shift+S` |
| Load Project | | `Ctrl+L` |
| Easy Text Writing | | `Ctrl + T` |
| Show Info | | `Alt + I` |
| Delete Image | | `Delete` |
| Set Image Location | | `Ctrl + Q` |
| Move Image Up | | `Ctrl + UP` |
| Mve Image Down | | `Ctrl + Down` |



Shortcuts are fully customisable and saved between sessions.

---

## 🔧 Dependencies

| Package | Purpose |
|---|---|
| `PyQt5` | Desktop UI framework |
| `Pillow` | Image loading, resizing, drawing |
| `python-pptx` | PowerPoint file parsing |
| `openpyxl` | Excel file generation |
| `python-bidi` | Right-to-left text rendering |
| `ffmpeg` (system) | Video clip rendering & audio duration probing |

---

## 🗂️ Data Storage

Eventure stores its runtime data (fonts, help files, language files, songs) in:

```
~/Neria-LTD/Eventure/
```

This folder is automatically created and synced from the application directory on first launch.

---

## 📝 License

This project is licensed under the [MIT License](LICENSE).

---

## 📬 Contact

For questions, bug reports, or feature requests, please open an [issue](https://github.com/neriacohen300/Eventure/issues) on GitHub.
