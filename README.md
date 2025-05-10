# WP Converter

Converts WordPerfect files (.wpd) to Word files (.docx) using LibreOffice.

## Key features

* **One-click conversion (GUI)**
  * Launch the bundled WP Converter.app (macOS) or packaged .exe (Windows) to open a minimal Tkinter window—no Terminal or Python setup required.
* **WordPerfect → Word in one step**
  * Converts any .wpd file to modern .docx via LibreOffice, preserving formatting.
* **Batch mode with optional recursion**
  * Select a folder to convert every .wpd inside; enable “Search sub-folders” to process the entire directory tree.
* **Automatic file organization**
  * Enable “Place converted files in ‘Converted’ sub-folder” to route all generated .docx files into a tidy Converted/ directory.
* **Command-line interface for power users**
* **Smart LibreOffice detection**
  * Finds soffice on $PATH or at /Applications/LibreOffice.app/... (macOS) and prompts clearly if not found.
* **Cross-platform packaging**
  * Built with PyInstaller (--windowed --onedir), so the app includes its own Python runtime—users only need LibreOffice.
* **LGPL-friendly distribution**
  * LibreOffice remains external by default, respecting license terms and keeping the download footprint small.

## Requirements

* OS X (Apple Silicon)
* Python3
* LibreOffice ([get it here](https://www.libreoffice.org))

## How to install & use

### MacOS App

1. Download LibreOffice ([get it here](https://www.libreoffice.org))
2. Download WP Converter.app located in `dist/` folder
3. Copy into Applications folder
4. Double click to run

### Python Script

1. Download LibreOffice ([get it here](https://www.libreoffice.org))
2. Download repo
3. Run `python3 wpd_to_docx.py`

#### Python Script Usage

```
# interactive prompt
$ python wpd_to_docx.py 

# Pass an individual file to convert as an argument 
$ python wpd_to_docx.py /path/to/file.wpd

# Pass an folder to recursively convert as an argument 
$ python wpd_to_docx.py /path/to/folder

```
