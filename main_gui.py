"""
PyWebView front‑end for WordPerfect→Word converter.
Retains all functionality of the original Tk GUI.
"""
import pathlib, os
import webview
from wpd_to_docx import convert_file, ensure_soffice
from tkinter import Tk, filedialog  # for native dialogs


def pick_file(pattern="*.wpd", folder=False):
    Tk().withdraw()  # hide root
    if folder:
        path = filedialog.askdirectory()
    else:
        path = filedialog.askopenfilename(filetypes=[("WordPerfect", pattern)])
    return path or None


class API:
    """Methods exposed to JavaScript."""
    def choose_file(self):
        return pick_file()

    def choose_folder(self):
        return pick_file(folder=True)

    def choose_dest(self):
        return pick_file(folder=True)

    def convert(self, src_path, opts: dict):
        ensure_soffice()

        dest_type = opts.get("destType")
        recursive = bool(opts.get("recursive"))
        preserve  = bool(opts.get("preserve"))
        dest_root = pathlib.Path(opts.get("destPath")).expanduser() if dest_type=="custom" else None
        organize  = dest_type == "converted"

        src = pathlib.Path(src_path).expanduser()
        log=[]
        if src.is_file():
            convert_file(src, organize, dest_root, preserve)
            log.append(f"{src.name} ✓")
        else:
            walker = src.rglob("*.wpd") if recursive else src.glob("*.wpd")
            for wpd in walker:
                convert_file(wpd, organize, dest_root, preserve)
                log.append(f"{wpd.relative_to(src)} ✓")

        return "\n".join(log) or "No .wpd files found."


def main():
    # Resolve path to bundled index.html whether running frozen or not
    base_dir = pathlib.Path(os.path.abspath(os.path.dirname(__file__)))
    html_path = str(base_dir / "ui" / "index.html")

    api = API()
    window = webview.create_window(
        "WordPerfect → Word Converter",
        html_path,
        js_api=api,
        width=680,
        height=720,
        resizable=False
    )
    webview.start()


if __name__ == "__main__":
    main()