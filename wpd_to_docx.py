#!/usr/bin/env python3
"""
Convert a WordPerfect (.wpd) file (or every .wpd inside a directory tree)
to Microsoft Word (.docx) using LibreOffice in headless mode.

Usage
-----
$ python wpd_to_docx.py                # interactive prompt
$ python wpd_to_docx.py /path/to/file.wpd
$ python wpd_to_docx.py /path/to/folder
"""
import subprocess, sys, shutil
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

SOFFICE = shutil.which("soffice")  # LibreOffice executable

def ensure_soffice():
    """
    Make sure the global ``SOFFICE`` variable points to a LibreOffice
    binary. First tries whatever is on PATH, then falls back to the
    default macOS install location /Applications/LibreOffice.app.
    """
    global SOFFICE
    if SOFFICE:  # already found via shutil.which
        return
    alt = Path("/Applications/LibreOffice.app/Contents/MacOS/soffice")
    if alt.exists():
        SOFFICE = str(alt)
    else:
        sys.exit(
            "LibreOffice is required.  Install it from libreoffice.org "
            "or with Homebrew:\n\n  brew install --cask libreoffice"
        )

def convert_file(wpd: Path, organize: bool = False):
    print(f"→ {wpd.name}", end="  ")
    out_dir = (wpd.parent / "Converted") if organize else wpd.parent
    out_dir.mkdir(exist_ok=True)
    cmd = [
        SOFFICE, "--headless", "--convert-to", "docx",
        "--outdir", str(out_dir), str(wpd)
    ]
    res = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
    if res.returncode:
        print("failed"); print(res.stdout.strip())
    else:
        print("✓")

def walk_and_convert(path: Path, organize: bool = False):
    if path.is_file() and path.suffix.lower() == ".wpd":
        convert_file(path, organize)
    elif path.is_dir():
        for wpd in path.rglob("*.wpd"):
            convert_file(wpd, organize)
    else:
        print("No .wpd files found.")


# ------------- GUI LAUNCHER -------------
def launch_gui():
    """
    Launch a simple Tkinter GUI that lets the user pick either
    a single .wpd file or a folder and optionally recurse into
    sub‑folders before converting to .docx.
    """
    root = tk.Tk()
    root.title("WordPerfect → Word Converter")

    target_path = tk.StringVar()
    recurse = tk.BooleanVar(value=False)
    organize = tk.BooleanVar(value=False)

    # ---------- helper callbacks ----------
    def choose_file():
        path = filedialog.askopenfilename(
            title="Select WordPerfect file",
            filetypes=[("WordPerfect files", "*.wpd")])
        if path:
            target_path.set(path)

    def choose_folder():
        path = filedialog.askdirectory(title="Select Folder")
        if path:
            target_path.set(path)

    def run_conversion():
        ensure_soffice()  # try to auto‑locate LibreOffice
        # make sure LibreOffice (soffice) is available *after* GUI has loaded,
        # otherwise the app would exit before the window appears when launched
        if SOFFICE is None:
            messagebox.showerror(
                "LibreOffice not found",
                "The converter needs LibreOffice. "
                "Install it from libreoffice.org or with Homebrew:\n\n"
                "  brew install --cask libreoffice\n\n"
                "Then try again."
            )
            return
        if not target_path.get():
            messagebox.showerror("Error", "Please select a file or folder first.")
            return
        path = Path(target_path.get()).expanduser()
        if path.is_file():
            convert_file(path, organize.get())
        elif path.is_dir():
            if recurse.get():
                for wpd in path.rglob("*.wpd"):
                    convert_file(wpd, organize.get())
            else:
                for wpd in path.glob("*.wpd"):
                    convert_file(wpd, organize.get())
        messagebox.showinfo("Finished", "Conversion complete.")

    # ---------- UI layout ----------
    tk.Label(root, text="Pick a .wpd file or a folder containing .wpd files").pack(
        padx=20, pady=(15, 5))
    tk.Button(root, text="Browse File…", command=choose_file, width=25).pack(pady=2)
    tk.Button(root, text="Browse Folder…", command=choose_folder, width=25).pack(pady=2)
    tk.Checkbutton(root, text="Search sub‑folders (recursive)",
                   variable=recurse).pack(pady=6)
    tk.Checkbutton(root, text="Place converted files in 'Converted' sub‑folder",
                   variable=organize).pack(pady=0)
    tk.Button(root, text="Convert", command=run_conversion, width=25).pack(pady=(0, 15))

    root.mainloop()

def main():
    if len(sys.argv) > 1:                          # called from CLI → keep old behaviour
        ensure_soffice()
        target = Path(sys.argv[1]).expanduser()
        if not target.exists():
            sys.exit("Path does not exist.")
        walk_and_convert(target, organize=False)
    else:                                          # double‑clicked .app → show GUI first
        launch_gui()

if __name__ == "__main__":
    main()