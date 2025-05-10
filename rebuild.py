

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

ROOT = Path(__file__).resolve().parent
SCRIPT = ROOT / "wpd_to_docx.py"
ICON   = ROOT / "icon.icns"      # change if your .icns is elsewhere
APP    = "WP Converter"

def clean():
    """Remove PyInstaller artefacts from previous builds."""
    for path in (ROOT / "build", ROOT / "dist", ROOT / f"{APP}.spec"):
        if path.exists():
            if path.is_dir():
                shutil.rmtree(path)
            else:
                path.unlink()

def build():
    """Invoke PyInstaller in windowed onedir mode."""
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--windowed",
        "--onedir",
        "--name", APP,
    ]
    if ICON.exists():
        cmd.extend(["--icon", str(ICON)])
    cmd.append(str(SCRIPT))
    print("• Running:", " ".join(cmd))
    subprocess.run(cmd, check=True)

if __name__ == "__main__":
    clean()
    build()
    print(f"\nDone! The rebuilt app is in {ROOT/'dist'}")