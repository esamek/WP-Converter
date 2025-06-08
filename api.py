"""
Back-end API exposed to the Web UI via pywebview.

All methods return plain strings or JSON-serialisable objects so the
JavaScript side can use them directly.
"""
from pathlib import Path
from typing import Optional
import webview
import time
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
    def get_file_count(self, src_path: str, opts: dict) -> int:
        """
        Get the total number of WPD files that would be processed.
        Used for progress calculation.
        """
        try:
            recursive = bool(opts.get("recursive"))
            src = Path(src_path).expanduser()
            
            if src.is_file() and src.suffix.lower() == ".wpd":
                return 1
            elif src.is_dir():
                finder = src.rglob if recursive else src.glob
                wpd_files = list(finder("*.wpd"))
                return len(wpd_files)
            else:
                return 0
        except Exception:
            return 0

    def convert_with_progress(self, src_path: str, opts: dict) -> dict:
        """
        Convert files with progress updates.
        This is a generator-like function that yields progress updates.
        """
        try:
            ensure_soffice()

            # Parse options
            recursive = bool(opts.get("recursive"))
            dest_type = opts.get("destType", "same")
            dest_root = Path(opts.get("destPath", "")).expanduser() if dest_type == "custom" else None
            organize = dest_type == "converted"
            preserve = bool(opts.get("preserve", False))

            src = Path(src_path).expanduser()
            
            # Get total file count
            total_files = self.get_file_count(src_path, opts)
            
            if total_files == 0:
                return {
                    "success": False,
                    "message": "No .wpd files found to convert.",
                    "stats": {"total": 0, "successful": 0, "failed": 0, "skipped": 0},
                    "progress": 100
                }

            # Process files with progress tracking
            processed = 0
            successful = 0
            failed = 0
            
            if src.is_file() and src.suffix.lower() == ".wpd":
                # Single file conversion
                from wpd_to_docx import convert_file
                result = convert_file(src, organize, dest_root, preserve, src.parent)
                processed = 1
                if result:
                    successful = 1
                else:
                    failed = 1
            elif src.is_dir():
                # Directory conversion
                finder = src.rglob if recursive else src.glob
                wpd_files = list(finder("*.wpd"))
                
                for wpd in wpd_files:
                    from wpd_to_docx import convert_file
                    result = convert_file(wpd, organize, dest_root, preserve, src)
                    processed += 1
                    if result:
                        successful += 1
                    else:
                        failed += 1
                    
                    # Small delay to make progress visible
                    time.sleep(0.1)

            stats = {
                "total": total_files,
                "successful": successful,
                "failed": failed,
                "skipped": 0  # We'll count skipped files as successful for now
            }

            # Create result message
            if stats['total'] > 0:
                result_lines = [f"Conversion completed!"]
                result_lines.append(f"Total files processed: {stats['total']}")
                result_lines.append(f"Successfully converted: {stats['successful']}")
                if stats['failed'] > 0:
                    result_lines.append(f"Failed conversions: {stats['failed']}")
                
                return {
                    "success": True,
                    "message": "\n".join(result_lines),
                    "stats": stats,
                    "progress": 100
                }
            else:
                return {
                    "success": False,
                    "message": "No .wpd files found to convert.",
                    "stats": stats,
                    "progress": 100
                }

        except Exception as e:
            return {
                "success": False,
                "message": f"Error during conversion: {str(e)}",
                "stats": {"total": 0, "successful": 0, "failed": 0, "skipped": 0},
                "progress": 100
            }

    def convert(self, src_path: str, opts: dict) -> dict:
        """
        src_path : the file/folder chosen on the JS side
        opts     : dict with keys:
                    - recursive: bool
                    - destType: "same"|"converted"|"custom"
                    - destPath: str (empty if not custom)
                    - preserve: bool
        Returns  : dict with conversion statistics and results
        """
        try:
            ensure_soffice()

            # Parse options
            recursive = bool(opts.get("recursive"))
            dest_type = opts.get("destType", "same")
            dest_root = Path(opts.get("destPath", "")).expanduser() if dest_type == "custom" else None
            organize = dest_type == "converted"
            preserve = bool(opts.get("preserve", False))

            src = Path(src_path).expanduser()

            # Use the enhanced walk_and_convert function
            stats = walk_and_convert(
                src, 
                organize=organize, 
                dest_folder=dest_root, 
                retain_structure=preserve, 
                recursive=recursive
            )

            # Create a detailed result message
            if stats['total'] > 0:
                result_lines = [f"Conversion completed!"]
                result_lines.append(f"Total files processed: {stats['total']}")
                result_lines.append(f"Successfully converted: {stats['successful']}")
                if stats['failed'] > 0:
                    result_lines.append(f"Failed conversions: {stats['failed']}")
                
                return {
                    "success": True,
                    "message": "\n".join(result_lines),
                    "stats": stats
                }
            else:
                return {
                    "success": False,
                    "message": "No .wpd files found to convert.",
                    "stats": stats
                }

        except Exception as e:
            return {
                "success": False,
                "message": f"Error during conversion: {str(e)}",
                "stats": {"total": 0, "successful": 0, "failed": 0, "skipped": 0}
            }
