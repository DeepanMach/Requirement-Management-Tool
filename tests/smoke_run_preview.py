import os
import sys
import re
import json
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PyQt6.QtWidgets import QApplication
import pandas as pd

def main() -> int:
    app = QApplication(sys.argv)
    # Add repo root to sys.path
    repo_root = Path(__file__).resolve().parents[1]
    if str(repo_root) not in sys.path:
        sys.path.insert(0, str(repo_root))
    from requirement_tool.ui.main_window import MainWindow
    from requirement_tool.data_manager import RequirementDataManager

    # Instantiate main window (offscreen)
    try:
        win = MainWindow()
    except Exception as exc:
        print(f"INIT_ERROR: {exc}")
        return 2

    # Verify UI controls exist
    has_paged = hasattr(win, "paged_preview_chk") and win.paged_preview_chk is not None
    has_browser_btn = hasattr(win, "open_browser_btn") and win.open_browser_btn is not None
    print(f"HAS_PAGED_PREVIEW_CHK={has_paged}")
    print(f"HAS_OPEN_BROWSER_BTN={has_browser_btn}")

    # Build a small dataframe including hyperlinks
    df = pd.DataFrame([
        {"Object Type": "Text", "Object Text": "Visit https://example.com for more info."},
        {"Object Type": "Text", "Attachment Type": "html", "Attachment Data": "<p>An <strong>example</strong> with <a href=\"https://ex.com\">link</a>.</p>", "Object Text": ""},
    ])

    try:
        win.data_manager.dataframe = win.data_manager.finalize_dataframe(df)
        html_preview = win.compose_preview_html(win.data_manager.dataframe)
    except Exception as exc:
        print(f"PREVIEW_ERROR: {exc}")
        return 3

    # Check that URLs are shown in parentheses
    if re.search(r"<a href=\"https://example.com\">https://example.com</a> \(<span class=\"link-href\">https://example.com</span>\)", html_preview):
        print("LINK_PARENS_OK=1")
    else:
        print("LINK_PARENS_OK=0")

    # Ensure anchor link in HTML attachment also shows parentheses
    if re.search(r"<a href=\"https://ex.com\">.*?</a> \(<span class=\"link-href\">https://ex.com</span>\)", html_preview):
        print("HTML_LINK_PARENS_OK=1")
    else:
        print("HTML_LINK_PARENS_OK=0")

    # Write paged HTML to temp (simulate Open in Browser)
    try:
        html_paged = html_preview.replace("<body>", "<body data-paged='1'>", 1)
        tmp = Path(os.environ.get("TMP", os.getcwd())) / "rmt_smoke_preview.html"
        tmp.write_text(html_paged, encoding="utf-8")
        print(f"PREVIEW_FILE={tmp}")
    except Exception as exc:
        print(f"WRITE_PREVIEW_ERROR: {exc}")
        return 4

    # Clean up the Qt app
    win.close()
    app.quit()
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
