"""
Back-end API exposed to the Web UI via pywebview.

All methods return plain strings or JSON-serialisable objects so the
JavaScript side can use them directly.
"""
from pathlib import Path
from typing import Optional
import webview
from wpd_to_docx import ensure_soffice, convert_file, walk_and_convert


class API:
    # -------- file-dialog helpers -------------------------------------------
    def choose_file(self) -> Optional[str]:
        paths = webview.windows[0].create_file_dialog(
            webview.OPEN_DIALOG,
            allow_multiple=False
        )
        return paths[0] if paths else None

    def choose_folder(self) -> Optional[str]:
        paths = webview.windows[0].create_file_dialog(
            webview.FOLDER_DIALOG
        )
        return paths[0] if paths else None

    def choose_dest(self) -> Optional[str]:
        paths = webview.windows[0].create_file_dialog(
            webview.FOLDER_DIALOG
        )
        return paths[0] if paths else None

    # -------- conversion logic ----------------------------------------------
    def convert(self, src_path: str, opts: dict) -> str:
        """
        src_path : the file/folder chosen on the JS side
        opts     : dict with keys:
                    - recursive: bool
                    - destType: "same"|"converted"|"custom"
                    - destPath: str (empty if not custom)
                    - preserve: bool
        Returns  : newline-separated log of converted files
        """
        ensure_soffice()

        # Parse options
        recursive = bool(opts.get("recursive"))
        dest_type = opts.get("destType", "same")
        dest_root = Path(opts.get("destPath", "")).expanduser() if dest_type == "custom" else None
        organize = dest_type == "converted"
        preserve = bool(opts.get("preserve", False))

        src = Path(src_path).expanduser()
        log = []

        if src.is_file():
            convert_file(src, organize, dest_root, preserve, src.parent)
            log.append(f"{src.name} ✓")
        elif src.is_dir():
            # Use walk_and_convert to handle recursion and structure
            walk_and_convert(src, organize, dest_root, preserve, src)
            # For logging, re-walk and capture names
            pattern = "**/*.wpd" if recursive else "*.wpd"
            for wpd in src.rglob("*.wpd") if recursive else src.glob("*.wpd"):
                log.append(f"{wpd.relative_to(src)} ✓")
        else:
            return "No .wpd files found."

        return "\n".join(log) or "No .wpd files found."