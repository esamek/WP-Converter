# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Development Commands

### Building Applications
```bash
# Build both applications (CLI+GUI and Web UI versions)
python3 rebuild.py

# Install dependencies (if needed)
python3 -m pip install pyinstaller pywebview
```

### Running Applications
```bash
# CLI with interactive prompt or Tkinter GUI
python3 wpd_to_docx.py

# Web UI development mode
python3 web_gui.py

# Using the platform-specific launcher
./Convert\ WP.command    # macOS/Linux
Convert\ WP.bat          # Windows
```

### CLI Usage Examples
```bash
# Single file conversion
python3 wpd_to_docx.py /path/to/file.wpd

# Folder conversion with recursion
python3 wpd_to_docx.py /path/to/folder --recursive

# Organized output in 'Converted' subfolder
python3 wpd_to_docx.py /path/to/folder --organize

# Custom destination with structure preservation
python3 wpd_to_docx.py /path/to/folder --dest /path/to/output --retain-structure
```

## Architecture Overview

### Core Components

**wpd_to_docx.py** - Main conversion engine and CLI/Tkinter GUI
- Contains LibreOffice detection logic (`ensure_soffice()`)
- File conversion functions (`convert_file()`, `walk_and_convert()`)
- Tkinter GUI implementation with file dialogs and progress tracking
- CLI argument parsing and interactive prompts

**api.py** - Web UI backend API bridge
- Exposes conversion functions to JavaScript via pywebview
- Provides file dialog helpers (`choose_file()`, `choose_folder()`, `choose_dest()`)
- Handles progress tracking and statistics for web interface
- Returns JSON-serializable data for frontend consumption

**web_gui.py** - Web UI launcher
- Creates pywebview window displaying `ui/index.html`
- Injects API class as JavaScript bridge
- Configures window properties (680x720, non-resizable)

**rebuild.py** - PyInstaller build automation
- Builds both CLI/GUI and Web UI versions
- Handles icon embedding and asset bundling
- Cleans previous builds and generates .app bundles in `dist/`

### UI Components

**ui/index.html** - Modern web interface
- Tailwind CSS dark theme with responsive design
- Progressive form with contextual controls
- Real-time progress visualization with animated progress bar
- JavaScript integration with Python API via pywebview bridge

### Build System

The project uses PyInstaller for packaging with platform-aware configuration:
- **Platform-specific icons:** `.icns` for macOS, `.ico` for Windows
- **WP Converter** - CLI + Tkinter GUI version
- **WP Converter Web UI** - Web-based interface version
- Both versions create self-contained applications for the target platform

**Cross-Platform Build Notes:**
- `rebuild.py` automatically selects appropriate icon format based on `sys.platform`
- Windows builds use `icon.ico`, macOS builds use `icon.icns`
- Linux builds proceed without icons if neither format is available

### Dependencies

**Core Requirements:**
- Python 3.9+
- LibreOffice (headless mode) - auto-detected in common install locations
- PyInstaller (for building applications)
- pywebview (for Web UI version)

**LibreOffice Detection:**
The application automatically searches for LibreOffice in:
- System PATH (cross-platform)
- **macOS:** `/Applications/LibreOffice.app/Contents/MacOS/soffice`, Homebrew paths
- **Windows:** Program Files, user installations, Microsoft Store, portable installations
- **Linux:** Standard package manager installation paths

**Windows-Specific Paths:**
- `C:/Program Files/LibreOffice/program/soffice.exe`
- `C:/Program Files (x86)/LibreOffice/program/soffice.exe`
- `%USERPROFILE%/AppData/Local/Programs/LibreOffice/program/soffice.exe`
- `%USERPROFILE%/AppData/Local/Microsoft/WindowsApps/soffice.exe` (Microsoft Store)
- Portable installation paths on C: and D: drives

### Conversion Architecture

The conversion process uses LibreOffice's headless mode:
1. Source validation (file/directory existence, .wpd extension)
2. Destination path calculation based on user options
3. LibreOffice subprocess execution with timeout protection
4. Error handling and statistics collection
5. Progress reporting (for GUI versions)

The system supports three output modes:
- **Same location** - Output .docx files alongside .wpd sources
- **Organized** - Create 'Converted' subfolder for outputs
- **Custom destination** - User-specified folder with optional structure preservation

## Planning Document Organization

### Structure
All planning and design documents are organized in the `planning/` directory:

```
planning/
├── feature-name-plan.md      # Feature planning documents
├── improvement-name-plan.md  # Major improvement plans
└── completed-plans/          # Archive for reference
```

### Naming Convention
- Use descriptive, hyphenated names: `windows-support-plan.md`
- Include plan type in name: `feature-`, `improvement-`, `refactor-`
- Keep names concise but clear

### Planning Workflow
1. **Create Planning Document**: Before starting any non-trivial feature or improvement, create a planning document in `planning/`
2. **Document Structure**: Include Executive Summary, Current State Analysis, Required Changes, Implementation Steps, and Success Criteria
3. **Track Progress**: Use TodoWrite to track implementation progress against the plan
4. **Archive Completed Plans**: Move completed plans to `planning/completed-plans/` for reference

### When to Create Planning Documents
- Multi-step features requiring 3+ distinct tasks
- Major architectural changes or refactoring
- Cross-platform compatibility improvements
- Performance optimization efforts
- Security enhancements
- Integration of new technologies or frameworks

This ensures systematic approach to development and maintains institutional knowledge for future reference.

## Claude Memories

- Always update plan .md files to mark that tasks are completed after they are completed and ONLY when they are completed.