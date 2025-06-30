#!/usr/bin/env python3
"""
Rebuild WP Converter.app using PyInstaller.

Usage:
    python rebuild.py
"""

import shutil
import subprocess
import sys
from pathlib import Path
import os

ROOT = Path(__file__).resolve().parent

# Define build targets: (app name, entry-point script, extra PyInstaller args)
APPS = [
    ("WP Converter", ROOT / "wpd_to_docx.py", []),
    ("WP Converter Web UI", ROOT / "web_gui.py", ["--add-data", "ui" + os.pathsep + "ui"]),
]

# Select appropriate icon based on platform
if sys.platform == "win32":
    ICON = ROOT / "icon.ico"
elif sys.platform == "darwin":
    ICON = ROOT / "icon.icns"
else:
    ICON = None  # Linux/other platforms

def clean():
    """Remove PyInstaller artifacts for all builds."""
    for name, _, _ in APPS:
        spec = ROOT / f"{name}.spec"
        for path in (ROOT / "build", ROOT / "dist", spec):
            if path.exists():
                if path.is_dir():
                    shutil.rmtree(path)
                else:
                    path.unlink()

def build():
    """Invoke PyInstaller for each app in APPS."""
    for name, script, extras in APPS:
        args = [
            sys.executable, "-m", "PyInstaller",
            "--windowed", "--onedir", "--name", name
        ]
        if ICON and ICON.exists():
            args += ["--icon", str(ICON)]
        args += extras
        args.append(str(script))
        print("• Building:", name)
        print("  Command:", " ".join(args))
        subprocess.run(args, check=True)

if __name__ == "__main__":
    clean()
    build()
    print(f"\nDone! The rebuilt app is in {ROOT/'dist'}")