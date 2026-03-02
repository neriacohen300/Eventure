# BUILD.md — Packaging Your App as a Standalone Executable

## Goal
One `.exe` (Windows) or `.app` (macOS) that users can run with zero installs.

---

## Tool: PyInstaller (recommended)

PyInstaller bundles Python + all dependencies into a single file.

### 1. Install PyInstaller in your dev environment

```bash
pip install pyinstaller
```

### 2. One-file build (simplest)

```bash
pyinstaller --onefile --windowed main.py
```

| Flag | What it does |
|---|---|
| `--onefile` | Packs everything into a single `.exe` / binary |
| `--windowed` | No console window (for GUI apps with tkinter) |
| `--name "SlideshowMaker"` | Sets the output filename |
| `--icon icon.ico` | Adds a custom icon |

The output is in `dist/main.exe` (Windows) or `dist/main` (macOS/Linux).

---

### 3. Recommended full command for your app

```bash
pyinstaller \
  --onefile \
  --windowed \
  --name "SlideshowMaker" \
  --icon icon.ico \
  --add-data "assets;assets" \
  main.py
```

On macOS/Linux, use `:` instead of `;` in `--add-data`:
```bash
--add-data "assets:assets"
```

---

### 4. Hidden imports (important for your app)

PIL/Pillow and python-pptx sometimes need explicit imports declared.
Create a `slideshowmaker.spec` or pass these flags:

```bash
pyinstaller \
  --onefile \
  --windowed \
  --hidden-import "PIL._imaging" \
  --hidden-import "pptx" \
  --hidden-import "openpyxl" \
  --hidden-import "concurrent.futures" \
  main.py
```

---

### 5. Full `.spec` file approach (most reliable)

After the first run, PyInstaller creates a `.spec` file you can edit.
Use this for fine-grained control:

```python
# main.spec
block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        'PIL._imaging',
        'PIL.Image',
        'PIL.ImageFilter',
        'PIL.ExifTags',
        'pptx',
        'pptx.oxml',
        'openpyxl',
        'openpyxl.styles',
        'xml.etree.ElementTree',
        'concurrent.futures',
        'tkinter',
        'tkinter.filedialog',
    ],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    cipher=block_cipher,
)

pyz = PYZ(a.pure, a.zlib, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    name='SlideshowMaker',
    debug=False,
    strip=False,
    upx=True,            # compress with UPX for smaller size
    console=False,       # no terminal window
    icon='icon.ico',
)
```

Then build with:
```bash
pyinstaller main.spec
```

---

## Platform notes

### Windows
- Build on Windows to get a `.exe`. The output runs on any Windows 10/11 PC.
- The `.exe` will be **30–80 MB** depending on what's bundled (Pillow is large).
- UPX compression (`--upx-dir`) can cut size by ~30%.

### macOS
- Build on macOS to get a `.app` bundle.
- Use `--windowed` to create a proper `.app`.
- For distribution outside the App Store, sign with:
  ```bash
  codesign --deep --sign "Developer ID Application: YOUR NAME" dist/SlideshowMaker.app
  ```
- Or just tell users to right-click → Open to bypass Gatekeeper.

### Linux
- Produces a single binary. Works on distros with glibc ≥ your build machine's version.
- Build on the **oldest** Linux you want to support (e.g. Ubuntu 20.04) for max compatibility.

---

## Alternative: Nuitka (faster runtime, smaller binary)

```bash
pip install nuitka
python -m nuitka \
  --standalone \
  --onefile \
  --windows-disable-console \
  --enable-plugin=tk-inter \
  --enable-plugin=pylint-warnings \
  main.py
```

Nuitka compiles Python to C, making the app **2-3× faster** and slightly smaller,
but build time is much longer (5–15 min).

---

## Checklist before distributing

- [ ] Test on a clean VM / machine with no Python installed
- [ ] Check that file dialogs open correctly (tkinter works in `--windowed`)
- [ ] Verify multiprocessing works: add this to the top of `main.py`:

```python
import multiprocessing
multiprocessing.freeze_support()   # REQUIRED for --onefile on Windows
```

- [ ] Test with a real `.pptx` file end-to-end

---

## Quick-start summary

```bash
pip install pyinstaller
# Add multiprocessing.freeze_support() to main.py first!
pyinstaller --onefile --windowed --name SlideshowMaker main.py
# Your app is now at:  dist/SlideshowMaker.exe  (Windows)
#                      dist/SlideshowMaker       (macOS/Linux)
```
