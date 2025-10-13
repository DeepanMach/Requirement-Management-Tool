"""Main window for the requirement management tool."""
from __future__ import annotations

import logging
from typing import Optional

import pandas as pd
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QKeySequence
from PyQt6.QtWidgets import (
    QApplication,
    QFileDialog,
    QHBoxLayout,
    QMainWindow,
    QMenu,
    QMessageBox,
    QPushButton,
    QShortcut,
    QTableWidget,
    QTableWidgetItem,
    QTextBrowser,
    QTextEdit,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from ..data_manager import RequirementDataError, RequirementDataManager
from .trace_view import TraceMatrixView

LOGGER = logging.getLogger(__name__)


class MainWindow(QMainWindow):
    """Top-level window coordinating UI widgets with the data manager."""

    def __init__(self, data_manager: Optional[RequirementDataManager] = None):
        super().__init__()
        self.data_manager = data_manager or RequirementDataManager()
        self.setWindowTitle("Mach Requirement Management Tool")
        self.resize(1300, 800)

        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)

        button_row = QHBoxLayout()
        self.load_btn = QPushButton("📂 Load Excel")
        self.load_btn.clicked.connect(self.load_excels)
        self.run_btn = QPushButton("▶️ Run (Convert to Word)")
        self.run_btn.clicked.connect(self.run_convert_preview)
        self.back_btn = QPushButton("🔙 Back to Table View")
        self.back_btn.clicked.connect(self.show_table_view)
        self.save_btn = QPushButton("💾 Save Word")
        self.save_btn.clicked.connect(self.save_word)
        self.trace_btn = QPushButton("🔗 Traceability Matrix")
        self.trace_btn.clicked.connect(self.show_trace_view)

        for button in (
            self.load_btn,
            self.run_btn,
            self.back_btn,
            self.save_btn,
            self.trace_btn,
        ):
            button_row.addWidget(button)
        main_layout.addLayout(button_row)

        center_row = QHBoxLayout()
        main_layout.addLayout(center_row)

        self.nav_tree = QTreeWidget()
        self.nav_tree.setHeaderLabel("Navigation")
        self.nav_tree.itemClicked.connect(self.on_nav_item_clicked)
        center_row.addWidget(self.nav_tree, 2)

        from PyQt6.QtWidgets import QStackedWidget

        self.view_stack = QStackedWidget()
        center_row.addWidget(self.view_stack, 8)

        self.table = QTableWidget()
        self.table.setEditTriggers(QTableWidget.EditTrigger.DoubleClicked)
        self.table.itemChanged.connect(self.on_cell_changed)
        self.table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.table_context_menu)
        self.view_stack.addWidget(self.table)

        self.word_preview = QTextEdit()
        self.word_preview.setReadOnly(True)
        self.view_stack.addWidget(self.word_preview)

        self.trace_view = TraceMatrixView(self)
        self.view_stack.addWidget(self.trace_view)

        self.console = QTextBrowser()
        self.console.setFixedHeight(120)
        main_layout.addWidget(self.console)

        QShortcut(QKeySequence("Ctrl+Z"), self, activated=self.undo_last)

        self._undo_stack: list[pd.DataFrame] = []
        self._loading = False
        self.console.append("✅ Application Ready.")

    # ------------------------------------------------------------------
    def log_console(self, message: str) -> None:
        LOGGER.info(message)
        self.console.append(message)

    # ------------------------------------------------------------------
    def load_excels(self) -> None:
        files, _ = QFileDialog.getOpenFileNames(
            self, "Select Excel Files", "", "Excel Files (*.xlsx *.xls)"
        )
        if not files:
            return
        try:
            df = self.data_manager.load_workbooks(files)
        except RequirementDataError as exc:
            LOGGER.exception("Failed to load workbooks")
            QMessageBox.critical(self, "Load Error", str(exc))
            return
        self._undo_stack = [df.copy()]
        self.populate_table()
        self.populate_navigation()
        self.log_console("✅ Excel files loaded and numbering applied.")

    # ------------------------------------------------------------------
    def populate_table(self) -> None:
        df = self.data_manager.dataframe
        if df.empty:
            return

        columns = self.data_manager.visible_columns
        self._loading = True
        self.table.clear()
        self.table.setRowCount(len(df))
        self.table.setColumnCount(len(columns))
        self.table.setHorizontalHeaderLabels(columns)
        self.table.horizontalHeader().setStretchLastSection(True)

        for row_index in range(len(df)):
            for column_index, column_name in enumerate(columns):
                value = str(df.iloc[row_index][column_name])
                item = QTableWidgetItem(value)
                if "heading" in str(df.iloc[row_index].get("Object Type", "")).lower():
                    font = item.font()
                    font.setBold(True)
                    item.setFont(font)
                self.table.setItem(row_index, column_index, item)

        self.table.resizeColumnsToContents()
        self._loading = False
        self.view_stack.setCurrentIndex(0)
        self.log_console("📊 Table populated with Excel-style formatting.")

    # ------------------------------------------------------------------
    def populate_navigation(self) -> None:
        self.nav_tree.clear()
        current_h1: QTreeWidgetItem | None = None
        current_h2: QTreeWidgetItem | None = None
        for obj_type, section, text in self.data_manager.iter_navigation_items():
            if not text:
                continue
            label = f"{section} {text}".strip()
            if obj_type == "heading 1":
                current_h1 = QTreeWidgetItem([label])
                self.nav_tree.addTopLevelItem(current_h1)
                current_h2 = None
            elif obj_type == "heading 2":
                if current_h1 is None:
                    continue
                current_h2 = QTreeWidgetItem([label])
                current_h1.addChild(current_h2)
            elif obj_type == "heading 3":
                if current_h2 is None:
                    continue
                current_h2.addChild(QTreeWidgetItem([label]))
        self.nav_tree.expandAll()

    # ------------------------------------------------------------------
    def on_nav_item_clicked(self, item: QTreeWidgetItem) -> None:
        target = item.text(0).split(" ", 1)[-1].strip()
        from PyQt6.QtGui import QTextCursor

        self.word_preview.moveCursor(QTextCursor.MoveOperation.Start)
        cursor = self.word_preview.textCursor()
        if self.word_preview.find(target):
            self.word_preview.setTextCursor(cursor)
            self.word_preview.ensureCursorVisible()
            self.log_console(f"🧭 Navigated to: {target}")
        else:
            self.log_console(f"⚠️ Could not find heading: {target}")

    # ------------------------------------------------------------------
    def on_cell_changed(self, item: QTableWidgetItem) -> None:
        if self._loading:
            return
        row = item.row()
        column = item.column()
        column_name = self.table.horizontalHeaderItem(column).text()
        try:
            self.data_manager.update_cell(row, column_name, item.text().strip())
        except RequirementDataError as exc:
            QMessageBox.warning(self, "Edit Error", str(exc))
            return
        self._undo_stack.append(self.data_manager.dataframe.copy())
        self.populate_navigation()

    # ------------------------------------------------------------------
    def table_context_menu(self, position) -> None:
        menu = QMenu(self)
        add_row = menu.addAction("Add Row Below")
        delete_row = menu.addAction("Delete Row")
        action = menu.exec(self.table.viewport().mapToGlobal(position))

        current_row = self.table.currentRow()
        if action == add_row and current_row >= 0:
            self.table.insertRow(current_row + 1)
        elif action == delete_row and current_row >= 0:
            self.table.removeRow(current_row)

    # ------------------------------------------------------------------
    def undo_last(self) -> None:
        if len(self._undo_stack) <= 1:
            return
        self._undo_stack.pop()
        self.data_manager.dataframe = self._undo_stack[-1].copy()
        self.populate_table()
        self.populate_navigation()
        self.log_console("↩️ Undo applied.")

    # ------------------------------------------------------------------
    def run_convert_preview(self) -> None:
        html = self.data_manager.to_html_preview()
        if not html:
            QMessageBox.information(self, "Empty", "No data available for preview.")
            return
        self.word_preview.setHtml(html)
        self.view_stack.setCurrentIndex(1)
        self.log_console("📝 Word preview generated.")

    # ------------------------------------------------------------------
    def show_table_view(self) -> None:
        self.view_stack.setCurrentIndex(0)
        self.log_console("📊 Returned to table view.")

    # ------------------------------------------------------------------
    def save_word(self) -> None:
        if not self.word_preview.toPlainText().strip():
            QMessageBox.warning(self, "Empty", "Nothing to save!")
            return

        file_name, _ = QFileDialog.getSaveFileName(
            self, "Save Word File", "", "Word Document (*.docx)"
        )
        if not file_name:
            return
        if not file_name.endswith(".docx"):
            file_name += ".docx"

        try:
            from docx import Document
            from bs4 import BeautifulSoup
        except Exception as exc:  # pragma: no cover - optional dependency
            QMessageBox.critical(
                self,
                "Dependency Error",
                "python-docx and beautifulsoup4 are required to export Word files.\n"
                f"Details: {exc}",
            )
            return

        doc = Document()
        soup = BeautifulSoup(self.word_preview.toHtml(), "html.parser")
        for tag in soup.find_all(["h1", "h2", "h3", "p", "b"]):
            text = tag.get_text()
            if tag.name == "h1":
                doc.add_heading(text, level=1)
            elif tag.name == "h2":
                doc.add_heading(text, level=2)
            elif tag.name == "h3":
                doc.add_heading(text, level=3)
            elif tag.name == "b":
                doc.add_paragraph(text, style="Strong")
            elif tag.name == "p":
                doc.add_paragraph(text)
        doc.save(file_name)
        self.log_console(f"💾 Word file saved: {file_name}")

    # ------------------------------------------------------------------
    def show_trace_view(self) -> None:
        df = self.data_manager.to_trace_dataframe()
        if df.empty:
            QMessageBox.information(self, "Empty", "No data available.")
            return
        self.trace_view.load_data(df)
        self.view_stack.setCurrentIndex(2)
        self.log_console("🔗 Switched to Traceability Matrix view.")


def run_app() -> None:
    """Entry point used by scripts to launch the GUI."""
    import sys

    logging.basicConfig(level=logging.INFO)
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
