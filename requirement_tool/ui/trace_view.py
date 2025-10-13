"""Traceability matrix view used by the main window."""
from __future__ import annotations

from typing import Iterable, List, Sequence

import logging

import pandas as pd
from PyQt6.QtCore import QPoint, Qt
from PyQt6.QtWidgets import (
    QFileDialog,
    QLabel,
    QMessageBox,
    QPushButton,
    QSizePolicy,
    QTableWidget,
    QTableWidgetItem,
    QHBoxLayout,
    QVBoxLayout,
    QWidget,
)

LOGGER = logging.getLogger(__name__)


class TraceMatrixView(QWidget):
    """Widget to display traceability matrices."""

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self._df = pd.DataFrame()
        self.current_rows: List[Sequence[str]] = []
        self.current_cols: List[str] = []
        self.active_filters: dict[str, set[str]] = {}

        layout = QVBoxLayout(self)
        self.info_label = QLabel(
            "Traceability Matrix (Forward/Backward). Use header clicks to filter."
        )
        layout.addWidget(self.info_label)

        button_row = QHBoxLayout()
        self.btn_forward = QPushButton("Forward Trace")
        self.btn_forward.clicked.connect(self.show_forward_trace)
        self.btn_backward = QPushButton("Backward Trace")
        self.btn_backward.clicked.connect(self.show_backward_trace)
        self.btn_save = QPushButton("Save to Excel")
        self.btn_save.clicked.connect(self.save_current_view)
        button_row.addWidget(self.btn_forward)
        button_row.addWidget(self.btn_backward)
        button_row.addStretch()
        button_row.addWidget(self.btn_save)
        layout.addLayout(button_row)

        self.table = QTableWidget()
        self.table.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding
        )
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.horizontalHeader().setSectionsClickable(True)
        self.table.horizontalHeader().sectionClicked.connect(self._on_header_clicked)
        layout.addWidget(self.table)

    # ------------------------------------------------------------------
    def load_data(self, df: pd.DataFrame) -> None:
        self._df = df.copy() if df is not None else pd.DataFrame()
        self.active_filters.clear()
        self.show_forward_trace()

    # ------------------------------------------------------------------
    def show_forward_trace(self) -> None:
        if self._df.empty:
            self._clear_table("No data loaded.")
            return

        req_col = self._detect_column(["requirement id", "req id", "reqid"])
        up_col = self._detect_column(["up trace", "uptrace"])
        if req_col is None or up_col is None:
            self._clear_table("Required columns not found for forward trace.")
            return

        cols = [req_col, up_col, "Object Text"]
        rows = [tuple(self._df[c].iloc[i] for c in cols) for i in range(len(self._df))]

        self._update_table(cols, rows, "Forward Trace")

    # ------------------------------------------------------------------
    def show_backward_trace(self) -> None:
        if self._df.empty:
            self._clear_table("No data loaded.")
            return

        up_col = self._detect_column(["up trace", "uptrace"])
        req_col = self._detect_column(["requirement id", "req id", "reqid"])
        if req_col is None or up_col is None:
            self._clear_table("Required columns not found for backward trace.")
            return

        cols = [up_col, req_col, "Object Text"]
        rows = [tuple(self._df[c].iloc[i] for c in cols) for i in range(len(self._df))]
        self._update_table(cols, rows, "Backward Trace")

    # ------------------------------------------------------------------
    def _detect_column(self, candidates: Iterable[str]) -> str | None:
        lower_map = {c.lower(): c for c in self._df.columns}
        for cand in candidates:
            if cand in lower_map:
                return lower_map[cand]
        for column in self._df.columns:
            if any(part in column.lower() for part in candidates):
                return column
        return None

    # ------------------------------------------------------------------
    def _update_table(
        self, columns: Sequence[str], rows: Sequence[Sequence[str]], mode: str
    ) -> None:
        self.current_cols = list(columns)
        self.current_rows = [tuple(str(v) for v in row) for row in rows]
        self.info_label.setText(f"Traceability Matrix ({mode})")
        self._populate_table_from_current()

    # ------------------------------------------------------------------
    def _clear_table(self, message: str) -> None:
        self.table.clear()
        self.table.setRowCount(0)
        self.table.setColumnCount(0)
        self.info_label.setText(message)

    # ------------------------------------------------------------------
    def _populate_table_from_current(self) -> None:
        rows = self.current_rows
        cols = self.current_cols
        if self.active_filters:
            filtered: List[Sequence[str]] = []
            for row in rows:
                include = True
                for idx, column in enumerate(cols):
                    selected = self.active_filters.get(column)
                    if selected and row[idx] not in selected:
                        include = False
                        break
                if include:
                    filtered.append(row)
            rows = filtered

        self.table.clear()
        self.table.setColumnCount(len(cols))
        self.table.setRowCount(len(rows))
        self.table.setHorizontalHeaderLabels(cols)

        for i, row in enumerate(rows):
            for j, value in enumerate(row):
                item = QTableWidgetItem(str(value))
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                self.table.setItem(i, j, item)

        self.table.resizeColumnsToContents()

    # ------------------------------------------------------------------
    def _on_header_clicked(self, index: int) -> None:
        if index < 0 or index >= len(self.current_cols):
            return

        column = self.current_cols[index]
        values = sorted({row[index] for row in self.current_rows})

        from PyQt6.QtWidgets import (
            QCheckBox,
            QLineEdit,
            QMenu,
            QPushButton,
            QScrollArea,
            QVBoxLayout,
            QWidget,
            QWidgetAction,
        )

        menu = QMenu(self)
        search = QLineEdit()
        search.setPlaceholderText("Search...")
        act_search = QWidgetAction(menu)
        act_search.setDefaultWidget(search)
        menu.addAction(act_search)
        menu.addSeparator()

        scroll = QScrollArea()
        scroll.setMinimumWidth(320)
        scroll.setMaximumHeight(380)
        scroll.setWidgetResizable(True)

        container = QWidget()
        vbox = QVBoxLayout(container)
        vbox.setContentsMargins(6, 6, 6, 6)
        vbox.setSpacing(4)

        checkboxes: List[QCheckBox] = []
        selected_values = self.active_filters.get(column, set(values))
        for value in values:
            checkbox = QCheckBox(value)
            checkbox.setChecked(value in selected_values)
            vbox.addWidget(checkbox)
            checkboxes.append(checkbox)

        scroll.setWidget(container)
        act_scroll = QWidgetAction(menu)
        act_scroll.setDefaultWidget(scroll)
        menu.addAction(act_scroll)
        menu.addSeparator()

        apply_button = QPushButton("Apply")
        act_apply = QWidgetAction(menu)
        act_apply.setDefaultWidget(apply_button)
        menu.addAction(act_apply)
        clear_action = menu.addAction("Clear Filter")
        show_all_action = menu.addAction("Show All")

        def on_search(text: str) -> None:
            lower_text = text.lower()
            for checkbox in checkboxes:
                checkbox.setVisible(lower_text in checkbox.text().lower())

        search.textChanged.connect(on_search)

        def on_apply() -> None:
            selected = {cb.text() for cb in checkboxes if cb.isChecked()}
            if len(selected) == len(checkboxes):
                self.active_filters.pop(column, None)
            else:
                self.active_filters[column] = selected
            self._populate_table_from_current()
            menu.close()

        apply_button.clicked.connect(on_apply)

        def on_clear() -> None:
            self.active_filters.pop(column, None)
            self._populate_table_from_current()
            menu.close()

        clear_action.triggered.connect(on_clear)

        def on_show_all() -> None:
            self.active_filters.clear()
            self._populate_table_from_current()
            menu.close()

        show_all_action.triggered.connect(on_show_all)

        header = self.table.horizontalHeader()
        x = header.sectionPosition(index)
        y = header.height()
        menu.exec(self.table.mapToGlobal(QPoint(x + 12, y + 20)))

    # ------------------------------------------------------------------
    def save_current_view(self) -> None:
        if not self.current_rows or not self.current_cols:
            QMessageBox.information(self, "Nothing", "No trace data to save.")
            return

        file_name, _ = QFileDialog.getSaveFileName(
            self, "Save Trace to Excel", "", "Excel Files (*.xlsx)"
        )
        if not file_name:
            return

        if not file_name.lower().endswith(".xlsx"):
            file_name += ".xlsx"

        try:
            pd.DataFrame(self.current_rows, columns=self.current_cols).to_excel(
                file_name, index=False
            )
        except Exception as exc:  # pragma: no cover - relies on pandas IO
            LOGGER.exception("Failed to save trace view")
            QMessageBox.critical(self, "Save Error", str(exc))
        else:
            QMessageBox.information(self, "Saved", f"Trace saved to {file_name}")
            parent = self.parent()
            if parent and hasattr(parent, "log_console"):
                parent.log_console(f"💾 Trace saved to {file_name}")
