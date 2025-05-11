"""
web_gui.py
==========

Launches the HTML Web UI (ui/index.html) in a pywebview window and
injects the API bridge (api.API). All conversion logic remains in
api.py and wpd_to_docx.py.
"""
from pathlib import Path
import webview
from api import API


def main() -> None:
    """Create and start the pywebview window."""
    base_dir = Path(__file__).resolve().parent
    html_path = base_dir / "ui" / "index.html"

    webview.create_window(
        title="WordPerfect → Word Converter",
        url=str(html_path),
        js_api=API(),
        width=680,
        height=720,
        resizable=False,
    )
    webview.start()


if __name__ == "__main__":
    main()