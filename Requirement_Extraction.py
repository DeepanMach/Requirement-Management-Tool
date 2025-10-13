import os, re, sys, tempfile, webbrowser, traceback, datetime, shutil
from typing import List, Dict
import pandas as pd
from collections import defaultdict

from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QTextEdit, QFileDialog, QMessageBox, QListWidget, QListWidgetItem, QTableWidget,
    QTableWidgetItem, QHeaderView, QProgressBar, QSplitter, QLineEdit, QTabWidget,
    QCheckBox, QMainWindow, QInputDialog, QScrollArea, QGridLayout
)
from PyQt6.QtGui import QPixmap, QAction
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QSize

# ------------------- Debug -------------------
DEBUG_ENABLED = False
DEBUG_LOG_FILE = os.path.join(os.path.dirname(__file__), "debug_log.txt")

def debug_log(*args):
    if not DEBUG_ENABLED: return
    ts = datetime.datetime.now().strftime("[%H:%M:%S]")
    msg = " ".join(str(a) for a in args)
    print(f"{ts} {msg}")
    sys.stdout.flush()
    try:
        with open(DEBUG_LOG_FILE, "a", encoding="utf-8") as f:
            f.write(f"{ts} {msg}\n")
    except Exception:
        pass

# ------------------- Optional Imports -------------------
try:
    from docx import Document
except Exception:
    Document = None
try:
    from lxml import etree
except Exception:
    etree = None

# ------------------- Helpers -------------------
def clean_id(text: str) -> str:
    """Normalize extracted ID; preserve one decimal level like CAP-SRS-0012.1"""
    if not text:
        return ""
    text = text.strip("[]()").rstrip(",:;.").strip()
    text = re.sub(r"\s+", "", text)  # remove internal spaces

    # Trim long numeric tails, preserve up to 4 digits + optional .X
    text = re.sub(
        r"([A-Za-z0-9\-_]+-)(\d{5,})(\.\d+)?$",
        lambda m: m.group(1) + m.group(2)[:4] + (m.group(3) or ""),
        text
    )

    # Keep only ONE decimal layer, e.g. CAP-SRS-0012.1 → CAP-SRS-0012.1
    # But CAP-SRS-0012.1.3 → CAP-SRS-0012
    text = re.sub(r"([A-Za-z0-9\-_]+-\d{1,4})(?:\.\d{1,2}\.\d+)+$", r"\1", text)

    return text


def build_regex_from_prefix(prefix: str) -> str:
    """Builds base regex allowing only one .X after the numeric part."""
    prefix = prefix.rstrip("-_")
    prefix = re.escape(prefix)
    return rf"{prefix}[A-Za-z0-9\-_]*[-_]\d+(?:\.\d+)?"


def build_regexes_from_input(user_input: str, pdf_mode: bool = False):
    """Return {prefix: compiled_regex} — PDF mode makes it whitespace-tolerant."""
    prefixes = [p.strip() for p in user_input.replace(",", "\n").splitlines() if p.strip()]
    regex_map = {}
    for p in prefixes:
        simple_pat = build_regex_from_prefix(p)
        if pdf_mode:
            simple_pat = simple_pat.replace(
                r"\d+(?:\.\d+)?",
                r"\d+(?:\s*\d+)*(?:\.\d+)?(?=[^\dA-Za-z]|$)"
            )
        full_pat = rf"(?<![A-Za-z0-9]){simple_pat}(?![A-Za-z0-9])"
        regex_map[p] = re.compile(full_pat)
    return regex_map


# ------------------- Image Extraction -------------------
def extract_images_with_anchor_docx(doc, out_folder='images'):
    os.makedirs(out_folder, exist_ok=True)
    images = []
    idx = 0
    debug_log("Scanning for image anchors...")
    for child in doc.element.body.iterchildren():
        try:
            xml = etree.tostring(child, encoding="unicode") if etree else ""
        except Exception:
            xml = ""
        for match in re.finditer(r"r:embed=\"(rId[0-9]+)\"", xml):
            rid = match.group(1)
            try:
                rel = doc.part.rels[rid]
                blob = rel.target_part.blob
                filename = os.path.basename(rel.target_ref)
                out = os.path.join(out_folder, f"{idx}_{filename}")
                base, ext = os.path.splitext(out)
                i = 1
                while os.path.exists(out):
                    out = f"{base}_{i}{ext}"
                    i += 1
                with open(out, "wb") as f:
                    f.write(blob)
                images.append((idx, out))
                debug_log("Saved image:", out)
            except Exception as e:
                debug_log("Image extraction error:", e)
        idx += 1
    debug_log("Total images extracted:", len(images))
    return images

# ------------------- DOCX Extraction -------------------
def extract_full_requirements_docx(path, regex_map, include_assets, progress_callback=None, status_callback=None):
    """
    Extract requirements, tables, and images from a DOCX file using flexible regex patterns.
    Improved version with:
      - Safe per-block error handling
      - Duplicate detection via set()
      - Smarter progress updates
      - Configurable image output directory
      - UTF-8 text safety
    """
    if Document is None:
        raise RuntimeError("python-docx missing")

    doc = Document(path)
    id_patterns = list(regex_map.values()) if regex_map else [
        re.compile(r"Requirement\s*ID\s*[:\-]?\s*([A-Z0-9\-_\.]+)", re.IGNORECASE)
    ]

    # Use a temp folder near document for image extraction
    out_folder = os.path.join(os.path.dirname(path), "images")
    os.makedirs(out_folder, exist_ok=True)

    images_by_index = defaultdict(list)
    if include_assets:
        try:
            imgs = extract_images_with_anchor_docx(doc, out_folder)
            for idx, p in imgs:
                images_by_index[idx].append(p)
        except Exception as e:
            debug_log("Image extraction failed:", traceback.format_exc())

    # Collect block elements
    from docx.text.paragraph import Paragraph
    from docx.table import Table
    blocks = []
    for idx, child in enumerate(doc.element.body.iterchildren()):
        tag = child.tag
        if tag.endswith('}p'):
            blocks.append((idx, 'p', Paragraph(child, doc)))
        elif tag.endswith('}tbl'):
            blocks.append((idx, 'tbl', Table(child, doc)))

    results = []
    seen_ids = set()
    n = len(blocks)

    for i in range(n):
        try:
            bidx, kind, element = blocks[i]
            text = element.text.strip() if kind == 'p' else "\n".join(
                [" | ".join(c.text.strip() for c in r.cells) for r in element.rows]
            )

            # Match requirement definition line more broadly
            matched = None
            for rx in id_patterns:
                m = rx.search(text)
                if m and re.match(r"^\s*(Requirement\s*(ID|No)|Req(\.|uirement)?(\s*ID|\s*No)?)", text, re.IGNORECASE):
                    matched = m
                    break
            if not matched:
                if progress_callback and i % 5 == 0:
                    progress_callback(int(i / max(1, n) * 100))
                continue

            # Extract and clean ID
            try:
                raw_id = matched.group(1)
            except IndexError:
                raw_id = matched.group(0)
            req_id = clean_id(raw_id)

            # Skip duplicates efficiently
            if req_id in seen_ids:
                debug_log("Skipping duplicate requirement:", req_id)
                continue
            seen_ids.add(req_id)

            # Gather related content
            j = i + 1
            chunk_texts, associated_tables = [], []
            while j < n:
                bj, bkind, belem = blocks[j]
                next_text = belem.text.strip() if bkind == 'p' else "\n".join(
                    " | ".join(c.text.strip() for c in r.cells) for r in belem.rows
                )

                # Stop at next requirement definition
                if any(rx.search(next_text) and re.match(r"^\s*(Requirement\s*(ID|No)|Req(\.|uirement)?(\s*ID|\s*No)?)",
                                                         next_text, re.IGNORECASE)
                       for rx in id_patterns):
                    break
                if bkind == 'p':
                    chunk_texts.append(belem.text.strip())
                elif bkind == 'tbl':
                    tbl_rows = [" | ".join(c.text.strip() for c in r.cells) for r in belem.rows]
                    associated_tables.append("\n".join(tbl_rows))
                j += 1

            block_text = "\n".join(chunk_texts).strip()
            ref_tables = ", ".join(sorted(set(re.findall(r"Table\s+\d+", block_text, re.IGNORECASE))))
            ref_figs = ", ".join(sorted(set(re.findall(r"Figure\s+\d+", block_text, re.IGNORECASE))))
            ref_reqs = [clean_id(r) for r in re.findall(
                r"\b([A-Z0-9]{2,}(?:-[A-Z0-9]+)*[-_]\d{1,4}(?:\.\d+)?)\b", block_text
            ) if r and r != req_id]

            imgs = []
            for k in range(i, j):
                imgs.extend(images_by_index.get(blocks[k][0], []))
            imgs.extend(images_by_index.get(bidx, []))

            results.append({
                "Requirement ID": req_id,
                "Requirement Text": block_text,
                "Referenced Tables": ref_tables,
                "Referenced Figures": ref_figs,
                "Referenced Requirements": ", ".join(sorted(set(ref_reqs))),
                "Table Content": "\n\n".join(associated_tables),
                "Image Path(s)": ", ".join(imgs)
            })

        except Exception as e:
            debug_log(f"Error processing block {i}:", traceback.format_exc())

        if progress_callback and i % 3 == 0:
            progress_callback(int(i / max(1, n) * 100))
        if status_callback and i % 10 == 0:
            status_callback(f"Processing {i}/{n}...")

    if progress_callback:
        progress_callback(100)
    debug_log(f"✅ Extraction completed with {len(results)} unique requirements.")
    return results


# ------------------- Worker -------------------
class ExtractWorker(QThread):
    progress = pyqtSignal(int)
    status = pyqtSignal(str)
    finished = pyqtSignal(list)
    def __init__(self, path, prefix_input, include_assets):
        super().__init__()
        self.path = path
        self.prefix_input = prefix_input
        self.include_assets = include_assets

    def run(self):
        try:
            regex_map = build_regexes_from_input(self.prefix_input)
            results = extract_full_requirements_docx(
                self.path, regex_map, self.include_assets,
                progress_callback=self.progress.emit
            )
            self.finished.emit(results)
        except Exception as e:
            debug_log("Worker error:", traceback.format_exc())
            self.status.emit(f"Error: {e}")
            self.finished.emit([])

# ------------------- Main UI -------------------
class MainUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Requirement Extractor v2 (Unique IDs + Merge)")
        self.resize(1300, 800)
        self.central = QWidget()
        self.setCentralWidget(self.central)
        layout = QVBoxLayout(self.central)

        ctrl = QHBoxLayout()
        self.btn_open = QPushButton("Open Document")
        self.prefix_input = QTextEdit()
        self.prefix_input.setFixedHeight(60)
        self.chk_assets = QCheckBox("Extract Tables & Images")
        self.chk_debug = QCheckBox("Verbose Debug")
        self.search_box = QLineEdit()
        self.search_box.setPlaceholderText("Search requirement...")
        self.btn_extract = QPushButton("Extract")
        self.btn_save_excel = QPushButton("Save to Excel")
        ctrl.addWidget(self.btn_open)
        ctrl.addWidget(QLabel("Prefixes:"))
        ctrl.addWidget(self.prefix_input)
        ctrl.addWidget(self.chk_assets)
        ctrl.addWidget(self.chk_debug)
        ctrl.addWidget(self.search_box)
        ctrl.addWidget(self.btn_extract)
        ctrl.addWidget(self.btn_save_excel)
        layout.addLayout(ctrl)

        pr = QHBoxLayout()
        self.progress = QProgressBar()
        self.status_lbl = QLabel("Idle.")
        pr.addWidget(self.progress, 1)
        pr.addWidget(self.status_lbl)
        layout.addLayout(pr)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        self.req_list = QListWidget()
        splitter.addWidget(self.req_list)

        right = QWidget()
        right_l = QVBoxLayout(right)
        self.req_text = QTextEdit()
        self.req_text.setReadOnly(True)
        right_l.addWidget(QLabel("Requirement Text:"))
        right_l.addWidget(self.req_text, 2)

        self.table_view = QTableWidget()
        right_l.addWidget(QLabel("Table Content:"))
        right_l.addWidget(self.table_view, 1)

        img_scroll = QScrollArea()
        self.img_container = QWidget()
        self.img_layout = QGridLayout(self.img_container)
        img_scroll.setWidgetResizable(True)
        img_scroll.setWidget(self.img_container)
        right_l.addWidget(QLabel("Images:"))
        right_l.addWidget(img_scroll, 2)
        splitter.addWidget(right)
        layout.addWidget(splitter)

        self.btn_open.clicked.connect(self.open_file)
        self.btn_extract.clicked.connect(self.start_extraction)
        self.btn_save_excel.clicked.connect(self.save_excel)
        self.req_list.currentRowChanged.connect(self.show_selected)
        self.search_box.textChanged.connect(self.filter_list)

        self.current_path = None
        self.extracted = []
        self.filtered = []

    def open_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select File", "", "DOCX Files (*.docx)")
        if not path:
            return
        self.current_path = path
        self.status_lbl.setText(f"Selected: {os.path.basename(path)}")

    def start_extraction(self):
        if not self.current_path:
            QMessageBox.warning(self, "Error", "Select a document first.")
            return
        global DEBUG_ENABLED
        DEBUG_ENABLED = self.chk_debug.isChecked()
        if DEBUG_ENABLED and os.path.exists(DEBUG_LOG_FILE):
            os.remove(DEBUG_LOG_FILE)
        self.worker = ExtractWorker(
            self.current_path,
            self.prefix_input.toPlainText(),
            self.chk_assets.isChecked()
        )
        self.worker.progress.connect(self.progress.setValue)
        self.worker.status.connect(self.status_lbl.setText)
        self.worker.finished.connect(self.on_done)
        self.worker.start()

    def on_done(self, results):
        self.extracted = results
        self.filtered = results
        self.req_list.clear()
        for r in results:
            item = QListWidgetItem(f"{r['Requirement ID']} — {(r.get('Requirement Text') or '')[:100]}")
            self.req_list.addItem(item)
        self.status_lbl.setText(f"Extracted {len(results)} unique requirements")

    def show_selected(self, idx):
        if idx < 0 or idx >= len(self.filtered):
            return
        rec = self.filtered[idx]
        self.req_text.setPlainText(rec.get("Requirement Text", ""))
        self.show_table(rec.get("Table Content", ""))
        self.show_images(rec.get("Image Path(s)", ""))

    def show_table(self, content):
        self.table_view.clear()
        rows = [r for r in content.splitlines() if "|" in r]
        parsed = [r.split("|") for r in rows]
        self.table_view.setRowCount(len(parsed))
        self.table_view.setColumnCount(max((len(p) for p in parsed), default=0))
        for i, row in enumerate(parsed):
            for j, val in enumerate(row):
                self.table_view.setItem(i, j, QTableWidgetItem(val.strip()))
        self.table_view.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)

    def show_images(self, imgs):
        """Display a limited number of images to prevent memory overload."""
        for i in reversed(range(self.img_layout.count())):
            w = self.img_layout.itemAt(i).widget()
            if w:
                w.deleteLater()

        img_list = [p.strip() for p in imgs.split(",") if p.strip() and os.path.exists(p)]
        MAX_IMAGES = 20
        for idx, path in enumerate(img_list[:MAX_IMAGES]):
            lbl = QLabel()
            pix = QPixmap(path)
            if not pix.isNull():
                lbl.setPixmap(pix.scaled(300, 200, Qt.AspectRatioMode.KeepAspectRatio,
                                         Qt.TransformationMode.SmoothTransformation))
                self.img_layout.addWidget(lbl, idx // 2, idx % 2)
        if len(img_list) > MAX_IMAGES:
            more_lbl = QLabel(f"... ({len(img_list) - MAX_IMAGES} more images omitted)")
            self.img_layout.addWidget(more_lbl, (MAX_IMAGES // 2) + 1, 0)


    def filter_list(self, text):
        text = text.lower().strip()
        self.req_list.clear()
        if not text:
            self.filtered = self.extracted
        else:
            self.filtered = [
                r for r in self.extracted
                if text in r.get("Requirement ID", "").lower()
                or text in r.get("Requirement Text", "").lower()
            ]
        for r in self.filtered:
            item = QListWidgetItem(f"{r['Requirement ID']} — {(r.get('Requirement Text') or '')[:100]}")
            self.req_list.addItem(item)

    def save_excel(self):
        if not self.extracted:
            QMessageBox.warning(self, "No Data", "No extracted data.")
            return
        path, _ = QFileDialog.getSaveFileName(self, "Save Excel", "", "Excel Files (*.xlsx)")
        if not path:
            return

        df = pd.DataFrame(self.extracted)
        try:
            df.to_excel(path, index=False, engine="openpyxl")  # openpyxl handles UTF-8 by default
        except Exception as e:
            QMessageBox.critical(self, "Save Error", f"Failed to write Excel: {e}")
            return

        img_src = os.path.join(os.path.dirname(self.current_path or ""), "images")
        if os.path.exists(img_src):
            try:
                shutil.copytree(img_src, os.path.join(os.path.dirname(path), "images"), dirs_exist_ok=True)
            except Exception as e:
                debug_log("Image copy failed:", e)

        QMessageBox.information(self, "Saved", f"✅ Results and images saved to:\n{path}")


def main():
    app = QApplication(sys.argv)
    w = MainUI()
    w.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
