"""Data management layer for the requirement management tool.

This module encapsulates all operations that read, transform and validate
requirement data. Keeping the logic here ensures we can unit test it without
bringing up the graphical user interface, which is important for DO-178C style
verification.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
import logging
from typing import Iterable, List, Sequence

import pandas as pd

LOGGER = logging.getLogger(__name__)


class RequirementDataError(RuntimeError):
    """Raised when the requirement data cannot be processed safely."""


REQUIRED_COLUMNS: Sequence[str] = (
    "Object Type",
    "Object Text",
)
OPTIONAL_METADATA_COLUMNS: Sequence[str] = (
    "Requirement ID",
    "Up Trace",
    "SourceFile",
    "SheetName",
)


@dataclass
class RequirementDataManager:
    """Manage loading and transformation of requirement data."""

    dataframe: pd.DataFrame = field(default_factory=pd.DataFrame)
    section_column_name: str = "Section Number"

    def load_workbooks(self, paths: Iterable[str]) -> pd.DataFrame:
        """Load and combine requirement spreadsheets.

        Parameters
        ----------
        paths:
            Paths to Excel workbooks selected by the user.

        Returns
        -------
        pandas.DataFrame
            The combined dataframe. ``self.dataframe`` is also updated.

        Raises
        ------
        RequirementDataError
            If no valid workbook could be loaded or the resulting dataframe is
            missing mandatory columns.
        """

        workbooks: List[pd.DataFrame] = []
        for raw_path in paths:
            path = Path(raw_path)
            if not path.exists():
                LOGGER.warning("Workbook %s does not exist", path)
                continue

            try:
                xls = pd.ExcelFile(path)
            except Exception as exc:  # pragma: no cover - relies on pandas IO
                LOGGER.exception("Failed to open workbook: %s", path)
                raise RequirementDataError(f"Failed to open workbook {path}: {exc}")

            for sheet in xls.sheet_names:
                try:
                    df = pd.read_excel(path, sheet_name=sheet).fillna("")
                except Exception as exc:  # pragma: no cover - relies on pandas IO
                    LOGGER.error("Failed to read sheet %s in %s: %s", sheet, path, exc)
                    continue

                df["SourceFile"] = path.name
                df["SheetName"] = sheet
                workbooks.append(df)

        if not workbooks:
            raise RequirementDataError("No valid Excel worksheets were loaded.")

        combined = pd.concat(workbooks, ignore_index=True)
        combined = combined.drop_duplicates(ignore_index=True)

        self._validate_columns(combined)

        numbered = self._apply_section_numbering(combined)
        self.dataframe = numbered
        LOGGER.info("Loaded %s rows from %s workbooks", len(numbered), len(workbooks))
        return self.dataframe

    # ------------------------------------------------------------------
    def _validate_columns(self, df: pd.DataFrame) -> None:
        missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
        if missing:
            raise RequirementDataError(
                "Missing required columns: " + ", ".join(missing)
            )

        for col in OPTIONAL_METADATA_COLUMNS:
            if col not in df.columns:
                df[col] = ""

    # ------------------------------------------------------------------
    def _apply_section_numbering(self, df: pd.DataFrame) -> pd.DataFrame:
        numbers: List[str] = []
        num_h1 = num_h2 = num_h3 = 0

        for _, row in df.iterrows():
            obj_type = str(row.get("Object Type", "")).strip().lower()
            if obj_type == "heading 1":
                num_h1 += 1
                num_h2 = num_h3 = 0
                numbers.append(f"{num_h1}.")
            elif obj_type == "heading 2":
                if num_h1 == 0:
                    num_h1 = 1
                num_h2 += 1
                num_h3 = 0
                numbers.append(f"{num_h1}.{num_h2}")
            elif obj_type == "heading 3":
                if num_h1 == 0:
                    num_h1 = 1
                if num_h2 == 0:
                    num_h2 = 1
                num_h3 += 1
                numbers.append(f"{num_h1}.{num_h2}.{num_h3}")
            else:
                numbers.append("")

        df = df.copy()
        if self.section_column_name in df.columns:
            df[self.section_column_name] = numbers
        else:
            df.insert(0, self.section_column_name, numbers)
        return df

    # ------------------------------------------------------------------
    @property
    def visible_columns(self) -> List[str]:
        if self.dataframe.empty:
            return []
        return [
            c for c in self.dataframe.columns if c.strip().lower() != "object type"
        ]

    # ------------------------------------------------------------------
    def update_cell(self, row: int, column: str, value: str) -> None:
        if self.dataframe.empty:
            raise RequirementDataError("No data loaded")
        if column not in self.dataframe.columns:
            raise RequirementDataError(f"Unknown column: {column}")
        if row < 0 or row >= len(self.dataframe.index):
            raise RequirementDataError(f"Row {row} is outside of the dataframe range")
        self.dataframe.at[row, column] = value

    # ------------------------------------------------------------------
    def to_html_preview(self) -> str:
        if self.dataframe.empty:
            return ""
        parts: List[str] = ["<div>"]
        for _, row in self.dataframe.iterrows():
            obj_type = str(row.get("Object Type", "")).lower()
            section = str(row.get(self.section_column_name, "")).strip()
            text = str(row.get("Object Text", "")).strip()
            req_id = str(row.get("Requirement ID", "")).strip()

            if obj_type == "heading 1":
                parts.append(f"<h1>{section} {text}</h1>")
            elif obj_type == "heading 2":
                parts.append(f"<h2>{section} {text}</h2>")
            elif obj_type == "heading 3":
                parts.append(f"<h3>{section} {text}</h3>")
            elif req_id:
                parts.append(f"<b>Requirement ID:</b> {req_id}<br><p>{text}</p>")
            else:
                parts.append(f"<p>{text}</p>")

        parts.append("</div>")
        return "\n".join(parts)

    # ------------------------------------------------------------------
    def iter_navigation_items(self) -> Iterable[tuple[str, str, str]]:
        if self.dataframe.empty:
            return []
        for _, row in self.dataframe.iterrows():
            yield (
                str(row.get("Object Type", "")).strip().lower(),
                str(row.get(self.section_column_name, "")).strip(),
                str(row.get("Object Text", "")).strip(),
            )

    # ------------------------------------------------------------------
    def to_trace_dataframe(self) -> pd.DataFrame:
        return self.dataframe.copy()
