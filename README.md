
# WP Converter

Convert WordPerfect files (`.wpd`) to Word documents (`.docx`) using LibreOffice.

---

## Applications

- **WP Converter (CLI & Tkinter GUI)**  
  A native Python app that offers both an interactive Tkinter GUI and a command-line interface for batch conversions.

- **WP Converter Web UI**  
  A modern, dark-mode web-based interface powered by HTML/CSS (Tailwind) and PyWebView, packaged as its own macOS `.app`.

---

## Key Features

### Common (Both Apps)

- **Single-file & Batch conversion**  
  Convert an individual `.wpd` or all `.wpd` files in a folder.

- **Optional recursion**  
  Enable “Search sub-folders” to process nested directories.

- **Automatic or custom output**  
  - **Same location** (default)  
  - **‘Converted’ sub-folder** (`--organize` CLI flag or UI toggle)  
  - **Custom destination** via `--dest /path` CLI flag or UI, with optional `--retain-structure` / “Preserve folder structure” to mirror the source hierarchy.

- **Cross-platform LibreOffice detection**  
  Auto-locates `soffice` on macOS (including Homebrew installations) and Windows, with clear error prompts if missing.

- **Robust error handling**  
  Comprehensive validation, timeout protection, and detailed error messages for failed conversions.

- **Conversion statistics**  
  Detailed reporting of successful, failed, and skipped files with summary statistics.

- **File conflict handling**  
  Automatically skips files that already exist in the destination to prevent overwriting.

### CLI & Tkinter GUI App

- **Interactive prompt**  
  Fallback terminal prompt if no GUI libraries are available.

- **Tkinter GUI**  
  Standalone window for non-technical users.

### Web UI App

- **Dark-mode styling**  
  Beautiful Tailwind-based dark theme with system fonts.

- **Progress visualization**  
  Animated progress bar with percentage display and checkmark completion indicator.

- **Dynamic UX**  
  - Contextual controls: hide/show inputs based on user choices  
  - Live summary: displays selected source, destination, and conversion results  
  - File dialogs via native OS dialogs (no additional dependencies)  
  - Visual feedback with color-coded success/error states

---

## Requirements

- **macOS** (Intel or Apple Silicon), **Windows**, or **Linux**  
- **Python 3.9+**  
- **LibreOffice** (headless) – platform-specific installation:
  - **macOS:** [Download](https://www.libreoffice.org/download/download/) or `brew install --cask libreoffice`
  - **Windows:** [Download](https://www.libreoffice.org/download/download/) from libreoffice.org
  - **Linux:** `sudo apt install libreoffice` (Ubuntu/Debian) or equivalent for your distribution
- **Python dependencies**: Install via `pip install -r requirements.txt`
  - PyWebView (for Web UI)
  - PyInstaller (for building executables)

---

## Windows Installation (Step-by-Step)

### Prerequisites
1. **Install Python 3.9+**
   - Download from [python.org](https://www.python.org/downloads/)
   - ⚠️ **Important**: Check "Add Python to PATH" during installation

2. **Install LibreOffice**
   - Download from [libreoffice.org](https://www.libreoffice.org/download/download/)
   - Choose the Windows x86-64 version
   - Install with default settings

### Quick Start (Web UI)
1. **Download/Clone this repository**
   ```cmd
   git clone <repository-url>
   cd "WP Converter"
   ```

2. **Install Python dependencies**
   ```cmd
   pip install -r requirements.txt
   ```

3. **Run the Web UI**
   ```cmd
   python web_gui.py
   ```
   Or double-click `Convert WP.bat`

### Building Executable (Optional)
```cmd
python rebuild.py
```
This creates `WP Converter.exe` and `WP Converter Web UI.exe` in the `dist/` folder.

---

## Installation & Usage

### Build & Package (All-in-One)

From the project root:

```bash
pip install -r requirements.txt
python3 rebuild.py
```

This generates platform-specific executables in `dist/`:

- **macOS:** `WP Converter.app` and `WP Converter Web UI.app`
- **Windows:** `WP Converter.exe` and `WP Converter Web UI.exe`  
- **Linux:** Equivalent executable files

The build process automatically selects appropriate icons and configurations for each platform.

### CLI Usage

```bash
# Interactive prompt
$ python3 wpd_to_docx.py

# Single file
$ python3 wpd_to_docx.py /path/to/file.wpd

# Folder (non-recursive)
$ python3 wpd_to_docx.py /path/to/folder

# Folder with recursive search
$ python3 wpd_to_docx.py /path/to/folder --recursive

# Folder with 'Converted' sub-folder
$ python3 wpd_to_docx.py /path/to/folder --organize

# Custom destination, preserving structure
$ python3 wpd_to_docx.py /path/to/folder --dest /path/to/output --retain-structure
```

### Tkinter GUI

```bash
python3 wpd_to_docx.py
```

**Platform-specific executables:**
- **macOS:** Double-click **WP Converter.app**
- **Windows:** Double-click **WP Converter.exe** or run `Convert WP.bat`
- **Linux:** Run the generated executable

### Web UI

#### Development

```bash
pip install -r requirements.txt
python3 web_gui.py
```

#### Packaged App

Double-click **WP Converter Web UI.app** (macOS) or run the `.exe` on Windows.

---

## License & Distribution

- **LGPL v3** for LibreOffice (external dependency).  
- **MIT** (or your chosen) license for WP Converter code.  
- Keep the LibreOffice bundle external by default to respect upstream licensing and minimize bundle size.

---

> Designed for seamless, flexible WordPerfect→Word workflows across command-line, desktop, and web interfaces.
