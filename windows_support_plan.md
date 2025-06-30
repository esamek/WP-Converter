# Windows Support Plan for WP Converter

## Executive Summary

The WP Converter codebase already has good cross-platform foundations, but requires several specific changes to ensure full Windows compatibility. The main issues are icon format support, executable detection enhancement, platform-specific help text, and the creation of Windows-equivalent launcher scripts.

## Current Windows Compatibility Status

### ✅ Already Windows-Compatible
- **Path handling**: Uses `pathlib.Path` throughout, which handles Windows paths correctly
- **LibreOffice detection**: Already includes Windows-specific paths for LibreOffice installation
- **File operations**: All file I/O uses cross-platform Python standard library functions
- **GUI frameworks**: Tkinter and pywebview both work on Windows
- **Build system**: PyInstaller supports Windows builds
- **Subprocess execution**: Uses `subprocess.run()` which works on Windows

### ❌ Needs Windows-Specific Changes
- **Icon format**: Currently uses `.icns` (macOS only), needs `.ico` for Windows
- **Executable detection**: May need enhancement for `.exe` extension handling
- **Help text**: Contains macOS-specific installation instructions (Homebrew)
- **Launcher script**: `Convert WP.command` is Unix-specific, needs Windows equivalent
- **Build configuration**: Icon selection needs to be platform-aware

## Required Changes

### 1. Icon Format Support (HIGH PRIORITY)

**Issue**: Build system hardcodes `icon.icns` which won't work on Windows PyInstaller builds.

**Tasks**:
- [ ] Create Windows `.ico` version of the icon from existing `icon.png`
- [ ] Modify `rebuild.py` to select appropriate icon format based on platform
- [ ] Test PyInstaller builds on Windows with correct icon

**Files to modify**: `rebuild.py`

**Implementation approach**:
```python
# Platform-specific icon selection
if sys.platform == "win32":
    ICON = ROOT / "icon.ico"
elif sys.platform == "darwin":
    ICON = ROOT / "icon.icns"
else:
    ICON = None  # Linux/other platforms
```

### 2. Enhanced LibreOffice Detection (MEDIUM PRIORITY)

**Issue**: Current detection should work but could be more robust for Windows environments.

**Tasks**:
- [ ] Verify `shutil.which("soffice")` finds `soffice.exe` on Windows
- [ ] Add additional Windows LibreOffice paths if needed (portable installations, Microsoft Store version)
- [ ] Test with different LibreOffice installation methods on Windows

**Files to modify**: `wpd_to_docx.py`

**Potential additional Windows paths**:
```python
# Additional Windows paths to consider
Path("C:/Program Files/LibreOffice 7.x/program/soffice.exe"),  # Version-specific
Path(f"{Path.home()}/AppData/Roaming/LibreOffice/program/soffice.exe"),  # Roaming profile
```

### 3. Platform-Specific Help Text (MEDIUM PRIORITY)

**Issue**: Help text includes macOS-specific installation instructions that confuse Windows users.

**Tasks**:
- [ ] Update error messages in `wpd_to_docx.py` to provide Windows-appropriate installation guidance
- [ ] Update Tkinter GUI error dialog to show Windows-specific help
- [ ] Ensure installation URLs and instructions are platform-appropriate

**Files to modify**: `wpd_to_docx.py`

**Implementation approach**:
```python
# Platform-specific installation guidance
if sys.platform == "win32":
    install_msg += "\n\nDownload and install from: https://www.libreoffice.org/download/download/"
    install_msg += "\n\nAfter installation, restart this application."
```

### 4. Windows Launcher Script (LOW PRIORITY)

**Issue**: `Convert WP.command` is a Unix bash script that won't work on Windows.

**Tasks**:
- [ ] Create `Convert WP.bat` Windows batch file equivalent
- [ ] Test batch file works correctly with Python path detection
- [ ] Update documentation to reference appropriate launcher for each platform

**New file**: `Convert WP.bat`

**Implementation approach**:
```batch
@echo off
cd /d "%~dp0"
python wpd_to_docx.py
pause
```

### 5. Build System Enhancements (MEDIUM PRIORITY)

**Issue**: Build system needs platform-aware configuration for optimal Windows experience.

**Tasks**:
- [ ] Add platform detection to `rebuild.py`
- [ ] Configure Windows-specific PyInstaller options if needed
- [ ] Ensure proper file associations and metadata for Windows builds
- [ ] Test build process on Windows environment

**Files to modify**: `rebuild.py`

**Implementation approach**:
```python
# Platform-specific PyInstaller arguments
def get_platform_args():
    args = []
    if sys.platform == "win32":
        args.extend(["--console"])  # or --windowed depending on preference
    return args
```

### 6. Documentation Updates (LOW PRIORITY)

**Tasks**:
- [ ] Update README.md with Windows-specific installation and usage instructions
- [ ] Update CLAUDE.md with Windows development notes
- [ ] Create Windows-specific troubleshooting section

## Testing Strategy

### Windows Testing Requirements
1. **Test environments**:
   - Windows 10/11 with standard LibreOffice installation
   - Windows with LibreOffice installed via Microsoft Store
   - Windows with portable LibreOffice installation

2. **Test scenarios**:
   - CLI usage with various file paths and options
   - Tkinter GUI functionality and file dialogs
   - Web UI functionality and file operations
   - PyInstaller build process and executable functionality
   - LibreOffice auto-detection across different installation types

### Validation Checklist
- [ ] All file paths resolve correctly on Windows
- [ ] LibreOffice detection works with standard Windows installations
- [ ] File dialogs open and return valid Windows paths
- [ ] Conversion process completes successfully
- [ ] Error messages display appropriate Windows guidance
- [ ] Built executables run without console dependency errors
- [ ] Icons display correctly in Windows Explorer and taskbar

## Implementation Priority

1. **Phase 1 (Essential)**: Icon format support and enhanced LibreOffice detection
2. **Phase 2 (Important)**: Platform-specific help text and build system enhancements  
3. **Phase 3 (Nice-to-have)**: Windows launcher script and documentation updates

## Risk Assessment

**Low Risk**: The codebase already has strong cross-platform foundations. Most changes are additive and won't break existing macOS functionality.

**Potential Issues**:
- Different LibreOffice installation paths on Windows corporate environments
- Windows file permission handling differences
- PyInstaller executable size and startup time on Windows

## Success Criteria

1. WP Converter runs successfully on Windows without modification after build
2. LibreOffice is detected automatically on standard Windows installations
3. All GUI elements work correctly with Windows look-and-feel
4. Error messages provide helpful, platform-appropriate guidance
5. Built Windows executables are professional and user-friendly