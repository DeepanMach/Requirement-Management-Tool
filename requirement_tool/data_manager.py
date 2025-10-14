"""Data management layer for the requirement management tool.

This module encapsulates all operations that read, transform and validate
requirement data. Keeping the logic here ensures we can unit test it without
bringing up the graphical user interface, which is important for DO-178C style
verification.
"""
from __future__ import annotations

from dataclasses import dataclass, field
import logging
import base64
import html
import json
import re
from pathlib import Path
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
    "SourceType",
    "Attachment Type",
    "Attachment Data",
)


@dataclass
class RequirementDataManager:
    """Manage loading and transformation of requirement data."""

    dataframe: pd.DataFrame = field(default_factory=pd.DataFrame)
    section_column_name: str = "Section Number"

    _REQ_ID_PATTERNS: Sequence[re.Pattern[str]] = (
        re.compile(
            r"^\s*Requirement\s*ID\s*[:\-]\s*(?P<id>[A-Za-z0-9_.\-]+)\s*(?P<body>.*)$",
            re.IGNORECASE,
        ),
        re.compile(
            r"^\s*(?P<id>[A-Za-z0-9]{2,}[A-Za-z0-9_.\-]*\d+)\s*[:\-–]\s*(?P<body>.+)$"
        ),
        re.compile(
            r"^\s*(?P<id>[A-Za-z0-9]{2,}[A-Za-z0-9_.\-]*\d+)\s+(?P<body>shall.+)$",
            re.IGNORECASE,
        ),
    )

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
                df["SourceType"] = "excel"
                workbooks.append(df)

        if not workbooks:
            raise RequirementDataError("No valid Excel worksheets were loaded.")

        combined = pd.concat(workbooks, ignore_index=True)
        combined = combined.drop_duplicates(ignore_index=True)

        self._validate_columns(combined)

        numbered = self._apply_section_numbering(combined)
        LOGGER.info("Loaded %s rows from %s workbooks", len(numbered), len(workbooks))
        return numbered

    # ------------------------------------------------------------------
    def load_word_documents(self, paths: Iterable[str]) -> pd.DataFrame:
        """Load requirement data from one or more Word documents (.docx)."""
        try:
            from docx import Document  # type: ignore
        except Exception as exc:  # pragma: no cover - optional dependency
            raise RequirementDataError(
                "Loading Word documents requires the python-docx package."
            ) from exc

        records: List[dict[str, str]] = []
        for raw_path in paths:
            path = Path(raw_path)
            if not path.exists():
                LOGGER.warning("Word document %s does not exist", path)
                continue

            try:
                document = Document(path)
            except Exception as exc:  # pragma: no cover - relies on python-docx IO
                LOGGER.exception("Failed to open Word document: %s", path)
                raise RequirementDataError(
                    f"Failed to open Word document {path}: {exc}"
                ) from exc

            records.extend(
                self._extract_records_from_document(document, path.name, path.stem)
            )

        if not records:
            raise RequirementDataError("No requirement content detected in Word files.")

        df = pd.DataFrame(records)
        self._validate_columns(df)
        numbered = self._apply_section_numbering(df)
        LOGGER.info("Loaded %s rows from %s Word files", len(numbered), len(records))
        return numbered

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
        current_source = None
        num_h1 = num_h2 = num_h3 = 0

        for _, row in df.iterrows():
            source = row.get("SourceFile", "")
            if source != current_source:
                current_source = source
                num_h1 = num_h2 = num_h3 = 0

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
            c
            for c in self.dataframe.columns
            if c.strip().lower()
            not in {"object type", "attachment data"}
        ]

    # ------------------------------------------------------------------
    def insert_attachment(
        self,
        *,
        object_type: str,
        attachment_type: str,
        attachment_data: str,
        object_text: str = "",
        requirement_id: str = "",
        insert_at: int | None = None,
        source_file: str = "Manual",
        sheet_name: str = "Manual",
        source_type: str = "manual",
    ) -> pd.DataFrame:
        """Insert an attachment row into the dataframe."""

        record = {
            self.section_column_name: "",
            "Object Type": object_type,
            "Requirement ID": requirement_id,
            "Object Text": object_text,
            "Up Trace": "",
            "SourceFile": source_file,
            "SheetName": sheet_name,
            "SourceType": source_type,
            "Attachment Type": attachment_type,
            "Attachment Data": attachment_data,
        }

        if self.dataframe.empty:
            df = pd.DataFrame([record])
        else:
            df = self.dataframe.copy()
            index = (
                len(df)
                if insert_at is None or insert_at < 0 or insert_at > len(df)
                else insert_at
            )
            upper = df.iloc[:index]
            lower = df.iloc[index:]
            df = pd.concat([upper, pd.DataFrame([record]), lower], ignore_index=True)

        self._validate_columns(df)
        self.dataframe = self._apply_section_numbering(df)
        return self.dataframe

    # ------------------------------------------------------------------
    def merge_new_dataframe(self, new_df: pd.DataFrame) -> pd.DataFrame:
        """Merge ``new_df`` into the current dataframe, replacing duplicate sources."""
        if new_df is None or new_df.empty:
            return self.dataframe

        working = new_df.copy()
        if "SourceFile" not in working.columns:
            working["SourceFile"] = ""
        self._validate_columns(working)

        if self.dataframe.empty:
            merged = working
        else:
            existing = self.dataframe.copy()
            sources = working.get("SourceFile", pd.Series(dtype=str)).unique()
            existing = existing[
                ~existing.get("SourceFile", pd.Series(dtype=str)).isin(sources)
            ]
            merged = pd.concat([existing, working], ignore_index=True)

        merged = self._apply_section_numbering(merged)
        self.dataframe = merged
        return self.dataframe

    # ------------------------------------------------------------------
    def _extract_records_from_document(
        self, document, source_name: str, sheet_name: str
    ) -> List[dict[str, str]]:
        """Extract heading/requirement content (tables/images included) from DOCX."""
        from docx.table import Table  # type: ignore
        from docx.text.paragraph import Paragraph  # type: ignore

        records: List[dict[str, str]] = []
        capture_started = False

        for block in self._iter_doc_blocks(document):
            if isinstance(block, Paragraph):
                text = block.text.strip()
                style_name = (block.style.name or "").lower()
                has_image = self._paragraph_has_image(block)

                start_section = self._is_capture_start(text)
                if start_section:
                    capture_started = True

                if not capture_started:
                    # Ignore prefatory content and TOC before the capture gate.
                    continue

                if self._is_toc_entry(style_name, text):
                    # Skip table of contents entries.
                    continue

                if not text and not has_image:
                    continue

                obj_type = self._classify_paragraph(style_name)
                req_id = ""
                body_text = text

                if obj_type == "Requirement":
                    req_id, body_text = self._maybe_extract_requirement(text)
                    if not req_id:
                        obj_type = "Text"
                        body_text = text
                elif obj_type == "Text":
                    req_id, possible = self._maybe_extract_requirement(text)
                    if req_id:
                        obj_type = "Requirement"
                        body_text = possible
                else:
                    body_text = text

                if obj_type.startswith("Heading") and not body_text:
                    body_text = text

                if body_text or obj_type.startswith("Heading"):
                    records.append(
                        {
                            "Object Type": obj_type,
                            "Requirement ID": req_id,
                            "Object Text": body_text,
                            "Up Trace": "",
                            "SourceFile": source_name,
                            "SheetName": sheet_name,
                            "SourceType": "docx",
                            "Attachment Type": "",
                            "Attachment Data": "",
                        }
                    )

                if has_image:
                    records.extend(
                        self._collect_paragraph_images(block, source_name, sheet_name)
                    )

            elif isinstance(block, Table):
                if not capture_started:
                    continue
                html_table = self._table_to_html(block)
                if not html_table.strip():
                    continue
                records.append(
                    {
                        "Object Type": "Table",
                        "Requirement ID": "",
                        "Object Text": "",
                        "Up Trace": "",
                        "SourceFile": source_name,
                        "SheetName": sheet_name,
                        "SourceType": "docx",
                        "Attachment Type": "table",
                        "Attachment Data": html_table,
                    }
                )

        return records

    # ------------------------------------------------------------------
    def _iter_doc_blocks(self, document) -> Iterable[object]:
        """Yield paragraphs and tables in document order."""
        from docx.oxml.text.paragraph import CT_P  # type: ignore
        from docx.oxml.table import CT_Tbl  # type: ignore
        from docx.table import Table  # type: ignore
        from docx.text.paragraph import Paragraph  # type: ignore

        for child in document.element.body.iterchildren():
            if isinstance(child, CT_P):
                yield Paragraph(child, document)
            elif isinstance(child, CT_Tbl):
                yield Table(child, document)

    # ------------------------------------------------------------------
    def _is_capture_start(self, text: str) -> bool:
        lowered = text.lower().replace("-", " ")
        return "software high level requirement" in lowered or "software low level requirement" in lowered

    # ------------------------------------------------------------------
    def _is_toc_entry(self, style_name: str, text: str) -> bool:
        lowered_style = style_name.lower()
        lowered_text = text.lower()
        if "toc" in lowered_style:
            return True
        return "table of contents" in lowered_text

    # ------------------------------------------------------------------
    def _classify_paragraph(self, style_name: str) -> str:
        style_lower = style_name.lower()
        if "heading 1" in style_lower:
            return "Heading 1"
        if "heading 2" in style_lower:
            return "Heading 2"
        if "heading 3" in style_lower:
            return "Heading 3"
        return "Requirement"

    # ------------------------------------------------------------------
    def _paragraph_has_image(self, paragraph) -> bool:
        blip_tag = "{http://schemas.openxmlformats.org/drawingml/2006/main}blip"
        for run in paragraph.runs:
            if run._element.findall(f".//{blip_tag}"):
                return True
        return False

    # ------------------------------------------------------------------
    def _collect_paragraph_images(
        self, paragraph, source_name: str, sheet_name: str
    ) -> List[dict[str, str]]:
        from docx.oxml.ns import qn  # type: ignore

        records: List[dict[str, str]] = []
        seen: set[str] = set()

        for run in paragraph.runs:
            for blip in run._element.findall(
                ".//{http://schemas.openxmlformats.org/drawingml/2006/main}blip"
            ):
                rel_id = blip.get(qn("r:embed"))
                if not rel_id or rel_id in seen:
                    continue
                seen.add(rel_id)
                part = paragraph.part.related_parts.get(rel_id)
                if part is None:
                    continue
                mime = getattr(part, "content_type", "image/png")
                filename = Path(str(part.partname)).name
                payload = json.dumps(
                    {
                        "mime": mime,
                        "data": base64.b64encode(part.blob).decode("ascii"),
                        "filename": filename,
                    }
                )
                records.append(
                    {
                        "Object Type": "Image",
                        "Requirement ID": "",
                        "Object Text": filename,
                        "Up Trace": "",
                        "SourceFile": source_name,
                        "SheetName": sheet_name,
                        "SourceType": "docx",
                        "Attachment Type": "image",
                        "Attachment Data": payload,
                    }
                )

        return records

    # ------------------------------------------------------------------
    def _table_to_html(self, table) -> str:
        rows_html: List[str] = []
        for row in table.rows:
            cells_html: List[str] = []
            for cell in row.cells:
                fragments = [
                    html.escape(p.text.strip())
                    for p in cell.paragraphs
                    if p.text and p.text.strip()
                ]
                cell_content = "<br/>".join(fragments) if fragments else "&nbsp;"
                cells_html.append(f"<td>{cell_content}</td>")
            rows_html.append("<tr>" + "".join(cells_html) + "</tr>")
        if not rows_html:
            return ""
        return (
            '<table border="1" style="border-collapse: collapse; width: 100%;">'
            + "".join(rows_html)
            + "</table>"
        )

    # ------------------------------------------------------------------
    def dataframe_to_html_table(self, df: pd.DataFrame) -> str:
        headers = "".join(f"<th>{html.escape(str(col))}</th>" for col in df.columns)
        rows = [f"<tr>{headers}</tr>"] if headers else []

        for _, series in df.iterrows():
            cells: List[str] = []
            for value in series:
                if pd.isna(value):
                    display = "&nbsp;"
                else:
                    text = str(value).strip()
                    display = html.escape(text) if text else "&nbsp;"
                cells.append(f"<td>{display}</td>")
            rows.append("<tr>" + "".join(cells) + "</tr>")

        if not rows:
            return ""

        return (
            '<table border="1" style="border-collapse: collapse; width: 100%;">'
            + "".join(rows)
            + "</table>"
        )

    # ------------------------------------------------------------------
    def _maybe_extract_requirement(self, text: str) -> tuple[str, str]:
        """Attempt to split a paragraph into requirement ID and body text."""
        for pattern in self._REQ_ID_PATTERNS:
            match = pattern.match(text)
            if match:
                req_id = match.group("id").strip()
                body = match.group("body").strip()
                if not body:
                    body = text.strip()
                return req_id, body
        return "", text

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
            attachment_type = str(row.get("Attachment Type", "")).strip().lower()
            attachment_data = str(row.get("Attachment Data", "")).strip()

            heading_prefix = f"{section} " if section else ""

            if obj_type == "heading 1":
                parts.append(f"<h1>{heading_prefix}{html.escape(text)}</h1>")
            elif obj_type == "heading 2":
                parts.append(f"<h2>{heading_prefix}{html.escape(text)}</h2>")
            elif obj_type == "heading 3":
                parts.append(f"<h3>{heading_prefix}{html.escape(text)}</h3>")
            elif attachment_type == "table" and attachment_data:
                parts.append(attachment_data)
            elif attachment_type == "image" and attachment_data:
                try:
                    payload = json.loads(attachment_data)
                    mime = payload.get("mime", "image/png")
                    data = payload.get("data", "")
                    filename = payload.get("filename", "image")
                except json.JSONDecodeError:
                    mime = "image/png"
                    data = attachment_data
                    filename = "image"
                if data:
                    parts.append(
                        (
                            '<div class="image-block" style="text-align:center; margin: 12px 0;">'
                            f'<img src="data:{mime};base64,{data}" alt="{html.escape(filename)}" '
                            'style="max-width:100%; height:auto;"/>'
                            "</div>"
                        )
                    )
            elif req_id:
                parts.append(
                    f"<p><b>Requirement ID:</b> {html.escape(req_id)}<br/>"
                    f"{html.escape(text)}</p>"
                )
            else:
                if text:
                    parts.append(f"<p>{html.escape(text)}</p>")

        parts.append("</div>")
        return "\n".join(parts)

    # ------------------------------------------------------------------
    def iter_navigation_items(self, source: str | None = None) -> Iterable[tuple[str, str, str]]:
        if self.dataframe.empty:
            return []
        df = self.dataframe
        if source:
            if "SourceFile" in df.columns:
                df = df[df["SourceFile"] == source]
            else:
                df = df.iloc[0:0]
        for _, row in df.iterrows():
            yield (
                str(row.get("Object Type", "")).strip().lower(),
                str(row.get(self.section_column_name, "")).strip(),
                str(row.get("Object Text", "")).strip(),
            )

    # ------------------------------------------------------------------
    def to_trace_dataframe(self) -> pd.DataFrame:
        return self.dataframe.copy()
