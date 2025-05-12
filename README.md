
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

- **Smart LibreOffice detection**  
  Auto-locates `soffice` on `$PATH` or `/Applications/LibreOffice.app/...` (macOS), with clear error prompts if missing.

### CLI & Tkinter GUI App

- **Interactive prompt**  
  Fallback terminal prompt if no GUI libraries are available.

- **Tkinter GUI**  
  Standalone window for non-technical users.

### Web UI App

- **Dark-mode styling**  
  Beautiful Tailwind-based dark theme with system fonts.

- **Dynamic UX**  
  - Contextual controls: hide/show inputs based on user choices  
  - Live summary: displays selected source, destination, and conversion results  
  - File dialogs via native OS dialogs (no additional dependencies)

---

## Requirements

- **macOS** (Intel or Apple Silicon) or **Windows**  
- **Python 3.9+**  
- **LibreOffice** (headless) – install via Homebrew:  
  ```bash
  brew install --cask libreoffice
  ```
- For Web UI: **PyWebView** (`pip install pywebview`)

---

## Installation & Usage

### Build & Package (All-in-One)

From the project root:

```bash
python3 -m pip install -r requirements.txt
python3 rebuild.py
```

This generates in `dist/`:

- **WP Converter.app**  
- **WP Converter Web UI.app**

You can zip or DMG these bundles for distribution.

### CLI Usage

```bash
# Interactive prompt
$ python3 wpd_to_docx.py

# Single file
$ python3 wpd_to_docx.py /path/to/file.wpd

# Folder (non-recursive)
$ python3 wpd_to_docx.py /path/to/folder

# Folder with 'Converted' sub-folder
$ python3 wpd_to_docx.py /path/to/folder --organize

# Custom destination, preserving structure
$ python3 wpd_to_docx.py /path/to/folder --dest /path/to/output --retain-structure
```

### Tkinter GUI

```bash
python3 wpd_to_docx.py
```

Double-click **WP Converter.app** (macOS) or run the generated executable on Windows.

### Web UI

#### Development

```bash
pip install pywebview
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