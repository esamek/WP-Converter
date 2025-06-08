#!/usr/bin/env python3
"""
Convert a WordPerfect (.wpd) file (or every .wpd inside a directory tree)
to Microsoft Word (.docx) using LibreOffice in headless mode.

Usage
-----
$ python wpd_to_docx.py                # interactive prompt
$ python wpd_to_docx.py /path/to/file.wpd
$ python wpd_to_docx.py /path/to/folder
$ python wpd_to_docx.py /path/to/folder --recursive  # search sub-folders recursively
$ python wpd_to_docx.py /path/to/folder --organize  # place files in 'Converted' subfolder
$ python wpd_to_docx.py /path/to/folder --dest /path/to/destination  # custom destination
$ python wpd_to_docx.py /path/to/folder --dest /path/to/destination --retain-structure  # keep folder structure
"""
import subprocess, sys, shutil
from pathlib import Path
from typing import Optional
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk

SOFFICE = shutil.which("soffice")  # LibreOffice executable

def ensure_soffice():
    """
    Make sure the global ``SOFFICE`` variable points to a LibreOffice
    binary. First tries whatever is on PATH, then falls back to the
    default install locations for macOS and Windows.
    """
    global SOFFICE
    if SOFFICE:  # already found via shutil.which
        return
    
    # Try common installation paths
    possible_paths = []
    
    # macOS paths
    possible_paths.extend([
        Path("/Applications/LibreOffice.app/Contents/MacOS/soffice"),
        Path("/opt/homebrew/bin/soffice"),  # Homebrew on Apple Silicon
        Path("/usr/local/bin/soffice"),     # Homebrew on Intel
    ])
    
    # Windows paths
    possible_paths.extend([
        Path("C:/Program Files/LibreOffice/program/soffice.exe"),
        Path("C:/Program Files (x86)/LibreOffice/program/soffice.exe"),
        Path(f"{Path.home()}/AppData/Local/Programs/LibreOffice/program/soffice.exe"),
    ])
    
    for path in possible_paths:
        if path.exists():
            SOFFICE = str(path)
            return
    
    # If we get here, LibreOffice wasn't found
    install_msg = (
        "LibreOffice is required. Install it from libreoffice.org"
    )
    if sys.platform == "darwin":  # macOS
        install_msg += " or with Homebrew:\n\n  brew install --cask libreoffice"
    elif sys.platform == "win32":  # Windows
        install_msg += "\n\nDownload from: https://www.libreoffice.org/download/download/"
    
    sys.exit(install_msg)

def convert_file(
    wpd: Path,
    organize: bool = False,
    dest_folder: Optional[Path] = None,
    retain_structure: bool = False,
    src_root: Optional[Path] = None,
):
    print(f"→ {wpd.name}", end="  ")
    
    try:
        # Check if source file exists and is readable
        if not wpd.exists():
            print("failed - file not found")
            return False
        
        if not wpd.is_file():
            print("failed - not a file")
            return False
        
        # Validate file extension
        if wpd.suffix.lower() != ".wpd":
            print("failed - not a .wpd file")
            return False
        
        # Determine output directory
        if dest_folder:
            if retain_structure and src_root:
                # Preserve full relative path from the original source root
                rel_path = wpd.parent.relative_to(src_root)
                out_dir = dest_folder / rel_path
            else:
                out_dir = dest_folder
        else:
            # Original behavior: same folder or "Converted" subfolder
            out_dir = (wpd.parent / "Converted") if organize else wpd.parent
        
        # Create output directory with error handling
        try:
            out_dir.mkdir(exist_ok=True, parents=True)
        except PermissionError:
            print("failed - permission denied creating output directory")
            return False
        except OSError as e:
            print(f"failed - cannot create output directory: {e}")
            return False
        
        # Check if output file already exists
        output_file = out_dir / f"{wpd.stem}.docx"
        if output_file.exists():
            print("skipped - output file already exists")
            return True
        
        # Run LibreOffice conversion
        cmd = [
            SOFFICE, "--headless", "--convert-to", "docx",
            "--outdir", str(out_dir), str(wpd)
        ]
        
        try:
            res = subprocess.run(
                cmd, 
                stdout=subprocess.PIPE, 
                stderr=subprocess.STDOUT, 
                text=True,
                timeout=60  # 60 second timeout
            )
        except subprocess.TimeoutExpired:
            print("failed - conversion timeout")
            return False
        except FileNotFoundError:
            print("failed - LibreOffice not found")
            return False
        
        if res.returncode:
            error_msg = res.stdout.strip() if res.stdout.strip() else "unknown error"
            print(f"failed - {error_msg}")
            return False
        else:
            # Verify the output file was created
            if output_file.exists():
                print("✓")
                return True
            else:
                print("failed - output file not created")
                return False
                
    except Exception as e:
        print(f"failed - unexpected error: {e}")
        return False

def walk_and_convert(
    path: Path,
    organize: bool = False,
    dest_folder: Optional[Path] = None,
    retain_structure: bool = False,
    src_root: Optional[Path] = None,
    recursive: bool = True,
):
    """
    Convert WPD files and return conversion statistics.
    
    Returns:
        dict: Statistics with 'total', 'successful', 'failed', 'skipped' counts
    """
    stats = {'total': 0, 'successful': 0, 'failed': 0, 'skipped': 0}
    
    if src_root is None:
        src_root = path if path.is_dir() else path.parent
        
    if path.is_file() and path.suffix.lower() == ".wpd":
        stats['total'] = 1
        result = convert_file(path, organize, dest_folder, retain_structure, src_root)
        if result is True:
            stats['successful'] = 1
        elif result is False:
            stats['failed'] = 1
    elif path.is_dir():
        # Use rglob for recursive search or glob for non-recursive search
        finder = path.rglob if recursive else path.glob
        wpd_files = list(finder("*.wpd"))
        if not wpd_files:
            print("No .wpd files found.")
            return stats
            
        stats['total'] = len(wpd_files)
        for wpd in wpd_files:
            result = convert_file(wpd, organize, dest_folder, retain_structure, src_root)
            if result is True:
                stats['successful'] += 1
            elif result is False:
                stats['failed'] += 1
            # Note: skipped files are counted in convert_file when it returns True for existing files
    else:
        print("No .wpd files found.")
    
    return stats


# ------------- GUI LAUNCHER -------------
def launch_gui():
    """
    Launch a simple Tkinter GUI that lets the user pick either
    a single .wpd file or a folder and optionally recurse into
    sub‑folders before converting to .docx.
    """
    root = tk.Tk()
    root.title("WordPerfect → Word Converter")
    ttk.Style().configure("Header.TLabel", font=("", 12, "bold"))
    
    # Add padding around the entire window
    main_frame = tk.Frame(root, padx=20, pady=15)
    main_frame.pack(fill=tk.BOTH, expand=True)

    target_path = tk.StringVar()
    dest_path = tk.StringVar()
    recurse = tk.BooleanVar(value=False)
    retain_structure = tk.BooleanVar(value=False)
    
    # Radio button variable for destination type
    destination_type = tk.StringVar(value="same")  # default: same location as source

    # ---------- helper callbacks ----------
    def choose_file():
        path = filedialog.askopenfilename(
            title="Select WordPerfect file",
            filetypes=[("WordPerfect files", "*.wpd")])
        if path:
            target_path.set(path)
            update_selection_display()

    def choose_folder():
        path = filedialog.askdirectory(title="Select Folder")
        if path:
            target_path.set(path)
            update_selection_display()

    def choose_dest_folder():
        path = filedialog.askdirectory(title="Select Destination Folder")
        if path:
            dest_path.set(path)
            update_selection_display()
    
    def on_destination_change():
        # Enable/disable controls based on the selected destination type
        if destination_type.get() == "custom":
            dest_folder_button.config(state=tk.NORMAL)
            retain_structure_cb.config(state=tk.NORMAL)
        else:
            dest_folder_button.config(state=tk.DISABLED)
            retain_structure_cb.config(state=tk.DISABLED)
        update_selection_display()

    def update_selection_display():
        # Clear existing text
        selection_display.config(state=tk.NORMAL)
        selection_display.delete(1.0, tk.END)
        
        # Show selected source
        if target_path.get():
            selection_display.insert(tk.END, f"Source: {target_path.get()}\n")
        
        # Show selected destination based on radio button selection
        dest_type = destination_type.get()
        if dest_type == "same":
            selection_display.insert(tk.END, "Destination: Same as source")
        elif dest_type == "converted":
            selection_display.insert(tk.END, "Destination: 'Converted' subfolder")
        elif dest_type == "custom" and dest_path.get():
            selection_display.insert(tk.END, f"Destination: {dest_path.get()}")
            if retain_structure.get():
                selection_display.insert(tk.END, " (preserving folder structure)")
        elif dest_type == "custom":
            selection_display.insert(tk.END, "Destination: Custom folder (not yet selected)")
            
        selection_display.config(state=tk.DISABLED)

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
            
        # Check if custom destination is selected but no path provided
        if destination_type.get() == "custom" and not dest_path.get():
            messagebox.showerror("Error", "Please select a destination folder.")
            return
        
        # Set destination folder and options based on radio selection
        dest_folder = None
        organize_files = False
        
        if destination_type.get() == "converted":
            organize_files = True
        elif destination_type.get() == "custom":
            dest_folder = Path(dest_path.get()).expanduser()
        
        path = Path(target_path.get()).expanduser()
        
        # Use walk_and_convert for consistent behavior
        stats = walk_and_convert(
            path,
            organize=organize_files,
            dest_folder=dest_folder,
            retain_structure=retain_structure.get(),
            recursive=recurse.get()
        )
        
        # Show detailed results
        if stats['total'] > 0:
            message = f"Conversion complete!\n\n"
            message += f"Total files: {stats['total']}\n"
            message += f"Successful: {stats['successful']}\n"
            if stats['failed'] > 0:
                message += f"Failed: {stats['failed']}\n"
            messagebox.showinfo("Finished", message)
        else:
            messagebox.showinfo("Finished", "No .wpd files found to convert.")

    # ---------- UI layout ----------
    # Title section
    ttk.Style().configure("Title.TLabel", font=("", 14, "bold"))
    title_label_text = "WordPerfect to Word Converter"
    title_label = ttk.Label(
        main_frame,
        text=title_label_text,
        style="Title.TLabel"
    )
    title_label.pack(pady=(0, 15), anchor="center")

    # Create a section frame with a header and content using ttk
    def create_section(parent, title, pady=(12, 6)):
        """Create a titled section using ttk so text adopts system theme."""
        frame = tk.Frame(parent)
        frame.pack(fill=tk.X, pady=pady)

        # Top separator
        ttk.Separator(frame, orient="horizontal").pack(fill=tk.X, pady=(0, 3))

        # Title label (system text colour)
        ttk.Label(frame, text=title, style="Header.TLabel").pack()

        # Bottom separator
        ttk.Separator(frame, orient="horizontal").pack(fill=tk.X, pady=(3, 6))

        # Container for section content
        content_frame = tk.Frame(frame)
        content_frame.pack(fill=tk.X, padx=5, pady=2)
        return content_frame
    
    # 1. SOURCE SECTION
    source_frame = create_section(main_frame, "SOURCE FILES")
    
    tk.Label(source_frame, text="Pick a .wpd file or a folder containing .wpd files:").pack(anchor="w")
    
    button_frame = tk.Frame(source_frame)
    button_frame.pack(fill=tk.X, pady=5)
    tk.Button(button_frame, text="Browse File…", command=choose_file, width=15).pack(side=tk.LEFT, padx=(0, 5))
    tk.Button(button_frame, text="Browse Folder…", command=choose_folder, width=15).pack(side=tk.LEFT)
    
    # Source options
    recursive_cb = tk.Checkbutton(source_frame, text="Search sub‑folders (recursive)",
                                variable=recurse, command=update_selection_display)
    recursive_cb.pack(pady=5, anchor="w")
    
    # 2. DESTINATION SECTION
    dest_frame = create_section(main_frame, "DESTINATION OPTIONS")
    
    # Radio button group for destination options
    dest_type_frame = tk.Frame(dest_frame)
    dest_type_frame.pack(fill=tk.X, anchor="w")
    
    # Same location radio button
    same_rb = tk.Radiobutton(dest_type_frame, text="Save in the same location as source files", 
                          variable=destination_type, value="same", command=on_destination_change)
    same_rb.pack(anchor="w", pady=2)
    
    # Converted subfolder radio button
    converted_rb = tk.Radiobutton(dest_type_frame, text="Place converted files in 'Converted' sub‑folder", 
                                variable=destination_type, value="converted", command=on_destination_change)
    converted_rb.pack(anchor="w", pady=2)
    
    # Custom destination radio button
    custom_rb = tk.Radiobutton(dest_type_frame, text="Use custom destination folder", 
                             variable=destination_type, value="custom", command=on_destination_change)
    custom_rb.pack(anchor="w", pady=2)
    
    # Custom destination folder options
    custom_dest_frame = tk.Frame(dest_frame)
    custom_dest_frame.pack(fill=tk.X, pady=3, padx=20)
    
    dest_folder_button = tk.Button(custom_dest_frame, text="Browse Destination…", 
                                 command=choose_dest_folder, width=20, state=tk.DISABLED)
    dest_folder_button.pack(anchor="w")
    
    retain_structure_cb = tk.Checkbutton(custom_dest_frame, text="Preserve folder structure in destination",
                                       variable=retain_structure, command=update_selection_display, 
                                       state=tk.DISABLED)
    retain_structure_cb.pack(anchor="w", pady=3)
    
    # 3. SUMMARY SECTION
    summary_frame = create_section(main_frame, "SELECTION SUMMARY")
    
    selection_display = tk.Text(summary_frame, height=3, width=50, state=tk.DISABLED, 
                             wrap=tk.WORD, bg="#f0f0f0", relief=tk.SUNKEN, padx=5, pady=5)
    selection_display.pack(fill=tk.X, pady=5)
    
    # 4. ACTION SECTION
    action_frame = create_section(main_frame, "CONVERT", pady=(15, 10))
    
    convert_button = tk.Button(action_frame, text="Convert Files", command=run_conversion, 
                            width=20, bg="#4CAF50", fg="white", relief=tk.RAISED, 
                            font=("", 11, "bold"))
    convert_button.pack(pady=5)

    # Initialize the display
    update_selection_display()
    root.mainloop()

def main():
    if len(sys.argv) > 1:                          # called from CLI → keep old behaviour
        ensure_soffice()
        target = Path(sys.argv[1]).expanduser()
        if not target.exists():
            sys.exit("Path does not exist.")
        
        # Use simple flags for CLI options
        organize = "--organize" in sys.argv
        recursive = "--recursive" in sys.argv  # New flag for recursive search
        
        # Check for destination folder parameter
        dest_folder = None
        retain_structure = False
        
        for i, arg in enumerate(sys.argv):
            if arg == "--dest" and i+1 < len(sys.argv):
                dest_folder = Path(sys.argv[i+1]).expanduser()
            elif arg == "--retain-structure":
                retain_structure = True
        
        stats = walk_and_convert(
            target, 
            organize=organize, 
            dest_folder=dest_folder, 
            retain_structure=retain_structure,
            recursive=recursive
        )
        
        # Print summary statistics
        if stats['total'] > 0:
            print(f"\nConversion Summary:")
            print(f"Total files: {stats['total']}")
            print(f"Successful: {stats['successful']}")
            if stats['failed'] > 0:
                print(f"Failed: {stats['failed']}")
        else:
            print("\nNo .wpd files found to convert.")
    else:                                          # double‑clicked .app → show GUI first
        launch_gui()

if __name__ == "__main__":
    main()
