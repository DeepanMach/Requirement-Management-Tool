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
import copy
import re
from pathlib import Path
# at top with other typing imports
from typing import Iterable, List, Sequence, Optional, Dict, Tuple, ClassVar, Pattern


import pandas as pd
import logging
logging.basicConfig(level=logging.DEBUG)

LOGGER = logging.getLogger(__name__)

# ----------------------------------------------------------------------
# 🔧 Helper Functions (taken and simplified from ReqIDTool)
# ----------------------------------------------------------------------
def clean_id(text: str) -> str:
    """Return trimmed text. Requirement detection now preserves exact matches."""
    return str(text or "").strip()

def build_regex_from_prefix(prefix: str) -> str:
    prefix = prefix.rstrip("-_")
    prefix = re.escape(prefix)
    return rf"{prefix}[A-Za-z0-9\-_]*[-_]\d+(?:\.\d+)?"

def build_regexes_from_input(user_input: str) -> list[re.Pattern[str]]:
    """Build compiled regex patterns for user-provided prefixes."""
    prefixes = [p.strip() for p in user_input.replace(",", "\n").splitlines() if p.strip()]
    regexes = []
    for p in prefixes:
        pat = build_regex_from_prefix(p)
        full_pat = rf"(?<![A-Za-z0-9]){pat}(?![A-Za-z0-9])"
        try:
            regexes.append(re.compile(full_pat, re.IGNORECASE))
        except re.error as exc:
            LOGGER.warning("Invalid regex for prefix %s: %s", p, exc)
    return regexes

def compile_user_patterns(user_input: str, manual_regex: bool) -> list[re.Pattern[str]]:
    """
    Compile patterns strictly from user input.
    - manual_regex=True: each non-empty line is compiled as-is.
    - manual_regex=False: each line is compiled as a literal substring (escaped).
    Returns: list of compiled regex objects.
    """
    lines = [ln.strip() for ln in str(user_input or "").replace(",", "\n").splitlines() if ln.strip()]
    out: list[re.Pattern[str]] = []
    for idx, line in enumerate(lines, start=1):
        pat = line if manual_regex else re.escape(line)
        try:
            out.append(re.compile(pat))
        except re.error as exc:
            LOGGER.warning("Invalid pattern on line %s: %s -> %s", idx, line, exc)
    return out


# -------------------- Prefix-based ID regex builder --------------------
def _build_id_core_from_prefix(prefix: str) -> str:
    """Return an ID-core regex (without group names) tailored from a prefix.

    Heuristics:
    - 'CAP-SRS' or contains '-SRS': digits with optional decimal, e.g. CAP-SRS-123 or CAP-SRS-123.4
    - 'H' + digits (e.g., H398): three to four dash groups of A-Z0-9, e.g., H398-XXX-XXX-XXX
    - 'DAU' (case-insensitive): DAUxxx-xxx-(HLR|LLR)-xxx (tolerant A-Z0-9 counts)
    - Fallback: prefix + alnum tail + final numeric with optional .X, tolerant of '_' or '-'
    """
    p = str(prefix or "").strip()
    if not p:
        return ""
    up = p.upper()
    if "-SRS" in up or up.endswith("SRS") or up.startswith("SRS-"):
        # System requirements like CAP-SRS-123 or CAP-SRS-123.4
        return rf"{re.escape(p)}-\d{{1,4}}(?:\.\d+)?"
    if up.startswith("H") and up[1:].isdigit():
        # H###-XXX-XXX-XXX style
        return rf"{re.escape(p)}-[A-Z0-9]{{2,10}}(?:-[A-Z0-9]{{2,10}}){{2,3}}"
    if up.startswith("DAU"):
        # DAUxxx-xxx-(HLR|LLR)-xxx pattern (tolerant segment sizes)
        return r"DAU[A-Z0-9]{2,}-[A-Z0-9]{2,}-(?:HLR|LLR)-[A-Z0-9]{2,}"
    # Generic prefix fallback: <prefix><alnum/-/_>*[-_]digits[.digits]?
    return rf"{re.escape(p)}[A-Z0-9\-_]*[-_]\d+(?:\.\d+)?"


def build_requirement_patterns_from_prefixes(user_input: str) -> list[re.Pattern[str]]:
    """Build compiled regex patterns from comma/line-separated prefixes.

    For each prefix, produce patterns with named groups:
    - ID anywhere: (?P<id>...core...)
    - Requirement line: ^(Requirement|Req...) ... (?P<id>...core...) (optional body)
    """
    prefixes = [p.strip() for p in str(user_input or "").replace(",", "\n").splitlines() if p.strip()]
    patterns: list[re.Pattern[str]] = []
    for p in prefixes:
        core = _build_id_core_from_prefix(p)
        if not core:
            continue
        try:
            pat_id_anywhere = re.compile(rf"(?P<id>{core})", re.IGNORECASE)
            # Capture any trailing text as body even without a delimiter after the ID
            pat_req_line = re.compile(
                rf"^\s*(?:Requirement|Req(?:uirement)?(?:\s*(?:ID|No))?)\s*[:\-–]?\s*(?P<id>{core})\s*(?P<body>.*)$",
                re.IGNORECASE,
            )
            patterns.extend([pat_req_line, pat_id_anywhere])
        except re.error as exc:
            LOGGER.warning("Invalid prefix-built regex for '%s': %s", p, exc)
    return patterns


def preview_prefix_patterns(user_input: str) -> list[str]:
    """Return pattern strings for UI preview from prefix input."""
    out: list[str] = []
    prefixes = [p.strip() for p in str(user_input or "").replace(",", "\n").splitlines() if p.strip()]
    for p in prefixes:
        core = _build_id_core_from_prefix(p)
        if not core:
            continue
        out.append(f"{p} → (?P<id>{core})")
        out.append(f"{p} → ^(?:Requirement|Req...) (?P<id>{core})(?::|\u2013|-)? (?P<body>.+)")
    return out


def _dedup_section_label(s: str) -> str:
    """Collapse immediate duplicate numeric section labels at start of string.
    Example: '10.1.1 10.1.1 Title' -> '10.1.1 Title'.
    """
    if not s:
        return s
    txt = str(s).strip()
    # Remove repeated leading identical numeric labels (once or multiple)
    # Loop to catch triple repeats, etc.
    for _ in range(3):
        m = re.match(r"^\s*(\d+(?:\.\d+)*)\s+\1\b(.*)$", txt)
        if not m:
            break
        txt = f"{m.group(1)}{m.group(2)}".strip()
    return txt





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
    _FIGURE_PREFIX: ClassVar[Pattern[str]] = re.compile(r"^\s*figure\s+(\d+)\s*:\s*", re.IGNORECASE)
    _TABLE_PREFIX: ClassVar[Pattern[str]]  = re.compile(r"^\s*table\s+(\d+)\s*:\s*",  re.IGNORECASE)

    # Add a canonical visible order the UI/grid expects
    DEFAULT_COLUMN_ORDER: ClassVar[List[str]] = [
        "Object Type",
        "Requirement ID",
        "Object Text",
        "Up Trace",
        "Down Trace",
        "Linked ID / Description",
        "SourceFile",
        "SheetName",
        "SourceType",
        "Attachment Type",
        "Attachment Data",
        "Trace Direction",
    ]

    dataframe: pd.DataFrame = field(default_factory=pd.DataFrame)
    section_column_name: str = "Section Number"
    # Controls whether Word import should drop front-matter (content before first Heading)
    skip_front_matter_for_word: bool = False
    front_matter_records: Dict[str, List[Dict[str, str]]] = field(default_factory=dict, init=False)
    front_matter_count: Dict[str, int] = field(default_factory=dict, init=False)
    # When True, ignore any numeric prefixes already present in Object Text
    # and re-number headings purely from Object Type (Heading 1..6).
    ignore_explicit_heading_numbers: bool = False
    _custom_patterns: list[re.Pattern[str]] = field(default_factory=list, init=False)
    _REQ_ID_PATTERNS: ClassVar[Sequence[re.Pattern[str]]] = (
        re.compile(r"^\s*Requirement\s*ID\s*[:\-]\s*(?P<id>[A-Za-z0-9_.\-]+)\s*(?P<body>.*)$", re.IGNORECASE),
        re.compile(r"^\s*(?P<id>[A-Za-z0-9]{2,}[A-Za-z0-9_.\-]*\d+)\s*[:\-–]\s*(?P<body>.+)$"),
        re.compile(r"^\s*(?P<id>[A-Za-z0-9]{2,}[A-Za-z0-9_.\-]*\d+)\s+(?P<body>shall.+)$", re.IGNORECASE),
    )

    def _strip_caption_prefix(self, text: str, kind: str) -> tuple[str, str]:
        text = str(text or "")
        rx = self._FIGURE_PREFIX if kind == "figure" else self._TABLE_PREFIX
        m = rx.match(text)
        if not m:
            cleaned = re.sub(rf"^\s*{kind}\s*:\s*", "", text, flags=re.IGNORECASE)
            return cleaned.strip(), ""
        rest = text[m.end():].strip()
        return rest, m.group(1)

    def _renumber_captions_per_source(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Renumber 'Figure n:' and 'Table n:' per SourceFile (like the VBA logic).
        Works for rows that either:
        - have Attachment Type == image/table, or
        - already start with 'Figure'/'Table' in Object Text.
        """
        if df.empty:
            return df

        out = df.copy()
        if "Object Text" not in out.columns:
            return out

        if "SourceFile" not in out.columns:
            out["SourceFile"] = ""

        # work per source
        for source, sub_idx in out.groupby("SourceFile").groups.items():
            idx = list(sub_idx)
            fig_n = 1
            tbl_n = 1

            for i in idx:
                row = out.loc[i]
                att = str(row.get("Attachment Type", "") or "").strip().lower()
                text = str(row.get("Object Text", "") or "")

                # normalize (strip any old numbering)
                txt_no_fig, _ = self._strip_caption_prefix(text, "figure")
                txt_no_tbl, _ = self._strip_caption_prefix(text, "table")

                # Decide whether this row should be treated as Figure or Table caption
                is_fig = (att == "image") or self._FIGURE_PREFIX.match(text) is not None
                is_tbl = (att == "table") or self._TABLE_PREFIX.match(text) is not None

                if is_fig:
                    cap = txt_no_fig or txt_no_tbl or text
                    out.at[i, "Object Text"] = f"Figure {fig_n}: {cap}".strip()
                    fig_n += 1
                    # If the user mis-set Object Type, don’t force it, but you can:
                    # out.at[i, "Object Type"] = "Image"
                elif is_tbl:
                    cap = txt_no_tbl or txt_no_fig or text
                    # Keep original caption/numbering; do not inject auto numbering
                    out.at[i, "Object Text"] = cap.strip()
                    tbl_n += 1  # advance for legacy callers
                    # Optionally normalize:
                    # out.at[i, "Object Type"] = "Table"
                else:
                    # leave non-caption rows alone
                    pass

        return out

    def renumber_figures_and_tables(self) -> pd.DataFrame:
        """
        Public API to renumber captions on the current dataframe.
        Call this after edits or inserts (mirrors VBA 'RenumberFigures' & 'RenumberTables').
        """
        self.dataframe = self._renumber_captions_per_source(self.dataframe)
        return self.dataframe

    def build_lof(self) -> pd.DataFrame:
        """
        Build a 'List of Figures' DataFrame with columns:
        ['Figure', 'Caption', 'SourceFile']
        (No page numbers – consistent with current docx preview.)
        """
        df = self.dataframe
        if df.empty:
            return pd.DataFrame(columns=["Figure","Caption","SourceFile"])

        rows = []
        for _, r in df.iterrows():
            text = str(r.get("Object Text", "") or "")
            m = self._FIGURE_PREFIX.match(text)
            if m:
                num = m.group(1)
                cap = text[m.end():].strip()
                rows.append({"Figure": num, "Caption": cap, "SourceFile": r.get("SourceFile","")})
        return pd.DataFrame(rows, columns=["Figure","Caption","SourceFile"])

    def build_lot(self) -> pd.DataFrame:
        """
        Build a 'List of Tables' DataFrame with columns:
        ['Table', 'Caption', 'SourceFile']
        """
        df = self.dataframe
        if df.empty:
            return pd.DataFrame(columns=["Table","Caption","SourceFile"])

        rows = []
        for _, r in df.iterrows():
            text = str(r.get("Object Text", "") or "")
            m = self._TABLE_PREFIX.match(text)
            if m:
                num = m.group(1)
                cap = text[m.end():].strip()
                rows.append({"Table": num, "Caption": cap, "SourceFile": r.get("SourceFile","")})
                continue
            m2 = TABLE_NO_ANY_RE.match(text)
            if m2:
                num = m2.group(1)
                cap = text[m2.end():].strip()
                rows.append({"Table": num, "Caption": cap, "SourceFile": r.get("SourceFile","")})
        return pd.DataFrame(rows, columns=["Table","Caption","SourceFile"])

    def lof_lot_as_html(self) -> str:
        """
        Convenience HTML block for LOF + LOT (no page numbers).
        Respects your preview style: this is separate so you can inject it
        into a 'front matter' tab if you want, while the main preview
        continues to skip LOF/LOT-generated sections found in the body.
        """
        lof = self.build_lof()
        lot = self.build_lot()

        parts = ["<div>"]
        if not lof.empty:
            parts.append("<h2>List of Figures</h2><ul>")
            for _, r in lof.iterrows():
                parts.append(f"<li>Figure {r['Figure']}: {html.escape(str(r['Caption']))}</li>")
            parts.append("</ul>")
        if not lot.empty:
            parts.append("<h2>List of Tables</h2><ul>")
            for _, r in lot.iterrows():
                parts.append(f"<li>Table {r['Table']}: {html.escape(str(r['Caption']))}</li>")
            parts.append("</ul>")
        parts.append("</div>")
        return "\n".join(parts)

    def insert_image_with_caption(self, *, file_bytes: bytes, filename: str, caption: str = "", insert_at: int | None = None, source_file: str = "Manual") -> pd.DataFrame:
        """
        Insert a captioned 'Image' row (your model stores image in one row).
        - Auto-assign 'Figure n:' (per SourceFile).
        - If caption is empty, we still number it: 'Figure n:'
        """
        # Build a temp row with attachment only (caption filled later by renumber)
        payload = json.dumps({
            "mime": self._infer_mime_from_name(filename),
            "data": base64.b64encode(file_bytes).decode("ascii"),
            "filename": filename,
        })
        # If user gave caption, keep it; numbering will be fixed by renumber pass.
        text = caption.strip()
        self.insert_attachment(
            object_type="Image",
            attachment_type="image",
            attachment_data=payload,
            object_text=(f"Figure : {text}" if text else "Figure :"),
            requirement_id="",
            insert_at=insert_at,
            source_file=source_file,
            sheet_name="Manual",
            source_type="manual",
        )
        # normalize + renumber like VBA order
        self.dataframe = self.finalize_dataframe(self.dataframe)
        self.dataframe = self._renumber_captions_per_source(self.dataframe)
        return self.dataframe

    def _canonicalize_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Ensure required/optional columns exist with sane dtypes and names.
        Also guarantees the 'Section Number' (or custom) column exists.
        """
        if df is None or df.empty:
            # Create an empty frame with all canonical columns so downstream logic is stable.
            cols = list(self.DEFAULT_COLUMN_ORDER)
            return pd.DataFrame(columns=cols)

        out = df.copy()

        # Ensure the section column exists
        if self.section_column_name not in out.columns:
            out[self.section_column_name] = ""

        # Ensure required + optional metadata columns exist
        for c in REQUIRED_COLUMNS:
            if c not in out.columns:
                out[c] = ""
        for c in OPTIONAL_METADATA_COLUMNS:
            if c not in out.columns:
                out[c] = ""

        # Common additional columns used later
        for c in ("Down Trace", "Linked ID / Description", "Trace Direction"):
            if c not in out.columns:
                out[c] = ""

        # Normalize dtypes for text-y columns (avoid NaN propagation later)
        text_cols = [
            self.section_column_name,
            "Object Type", "Requirement ID", "Object Text",
            "Up Trace", "Down Trace", "Linked ID / Description",
            "SourceFile", "SheetName", "SourceType",
            "Attachment Type", "Attachment Data", "Trace Direction",
        ]
        for c in text_cols:
            if c in out.columns:
                out[c] = out[c].fillna("").astype(str)

        return out

    def _reorder_visible_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Reorder columns for display/export: canonical order first, then any extras.
        """
        if df is None or df.empty:
            return df

        ordered = [c for c in self.DEFAULT_COLUMN_ORDER if c in df.columns]
        tail = [c for c in df.columns if c not in ordered]
        return df.loc[:, ordered + tail]


    def _infer_mime_from_name(self, name: str) -> str:
        ext = str(name or "").lower().rsplit(".", 1)[-1] if "." in str(name or "") else ""
        return {
            "png": "image/png",
            "jpg": "image/jpeg",
            "jpeg": "image/jpeg",
            "gif": "image/gif",
            "bmp": "image/bmp",
            "webp": "image/webp",
        }.get(ext, "image/png")

    def _normalize_title(self, s: str) -> str:
        s = (s or "").lower()
        s = re.sub(r"[^a-z\s]", " ", s)
        s = re.sub(r"\s+", " ", s).strip()
        return s

    def _is_auto_list_heading(self, text: str) -> Optional[str]:
        """
        Returns 'toc' | 'lof' | 'lot' if this line is a List/TOC heading,
        else None. Works even if not styled as Heading.
        """
        norm = self._normalize_title(text)
        if norm == "table of contents":
            return "toc"
        if norm in {"list of figures", "list of figure"}:
            return "lof"
        if norm in {"list of tables", "list of table"}:
            return "lot"
        return None

    def _has_trailing_page_num(self, text: str) -> bool:
        """
        True if line ends with a page number, with possible dot leaders or tabs.
        Examples:
        '... 113' | '...\t113' | '   113'
        """
        t = str(text or "")
        return bool(re.search(r"(?:\t|\s{2,}|\.{2,}|·{2,})?\s\d{1,4}\s*$", t))

    def _is_toc_entry_line(self, text: str) -> bool:
        """
        Examples: '1.2.3 Subsection .......... 17' or '1 Introduction    3'
        """
        t = str(text or "")
        if not self._has_trailing_page_num(t):
            return False
        # section numbering at start
        if re.match(r"^\s*\d+(?:\.\d+)*\s+\S", t):
            return True
        return False

    def _is_lof_entry_line(self, text: str) -> bool:
        """
        Examples: 'Figure 4 – Title ... 113', 'Fig. 3-1 Something\t42'
        """
        t = str(text or "")
        if not self._has_trailing_page_num(t):
            return False
        return bool(re.match(r"^\s*(figure|fig\.)\s+\S+", t, flags=re.I))

    def _is_lot_entry_line(self, text: str) -> bool:
        """
        Examples: 'Table 186 - Title ... 251'
        """
        t = str(text or "")
        if not self._has_trailing_page_num(t):
            return False
        return bool(re.match(r"^\s*table\s+\S+", t, flags=re.I))

    def load_workbooks(self, paths: Iterable[str]) -> pd.DataFrame:
        """Load and combine requirement spreadsheets."""
        workbooks: List[pd.DataFrame] = []
        for raw_path in paths:
            path = Path(raw_path)
            if not path.exists():
                LOGGER.warning("Workbook %s does not exist", path)
                continue

            try:
                xls = pd.ExcelFile(path)
            except Exception as exc:  # pragma: no cover
                LOGGER.exception("Failed to open workbook: %s", path)
                raise RequirementDataError(f"Failed to open workbook {path}: {exc}")

            for sheet in xls.sheet_names:
                try:
                    df = pd.read_excel(path, sheet_name=sheet).fillna("")
                except Exception as exc:  # pragma: no cover
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

        finalized = self.finalize_dataframe(combined)
        LOGGER.info("Loaded %s rows from %s workbooks", len(finalized), len(workbooks))
        return finalized

    def _is_capture_start(self, text: str) -> bool:
        """Detects the start of a capture region (optional future use)."""
        lowered = text.lower().replace("-", " ")
        return ("software high level requirement" in lowered
                or "software low level requirement" in lowered)

    def _is_toc_entry(self, style_name: str, text: str) -> bool:
        """Detects table-of-contents or auto-generated lists to skip."""
        lowered_style = (style_name or "").lower()
        lowered_text = (text or "").lower()
        if "toc" in lowered_style:
            return True
        return (
            "table of contents" in lowered_text
            or "list of figures" in lowered_text
            or "list of tables" in lowered_text
        )

    # ---------------- Inline formatting and list helpers ----------------
    def _paragraph_inlines_to_html(self, paragraph) -> str:
        try:
            from docx.oxml.ns import qn  # type: ignore
        except Exception:
            qn = None
        chunks: list[str] = []
        import urllib.parse as _url
        import re as _re

        def _sanitize_href(h: str) -> str:
            s = str(h or "").strip()
            if not s:
                return ""
            # Convert backslashes to forward slashes and percent-encode spaces
            s = s.replace("\\", "/")
            if s.startswith("http://") or s.startswith("https://"):
                try:
                    parts = _url.urlsplit(s)
                    path = _url.quote(parts.path or "", safe="/%@:+,._-~")
                    query = _url.quote_plus(parts.query, safe="=&;%@:+,._-~") if parts.query else ""
                    frag = _url.quote(parts.fragment, safe="@:+,._-~") if parts.fragment else ""
                    s = _url.urlunsplit((parts.scheme, parts.netloc, path, query, frag))
                except Exception:
                    pass
            return s
        for run in paragraph.runs:
            t = run.text or ""
            if not t:
                continue
            s = html.escape(t)
            href = ""
            try:
                if qn is not None:
                    links = run._r.xpath('ancestor::w:hyperlink', namespaces={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
                    if links:
                        link_el = links[0]
                        rid = link_el.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                        if rid and rid in paragraph.part.rels:
                            href = paragraph.part.rels[rid]._target
                        else:
                            anchor = link_el.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}anchor') or ""
                            if anchor:
                                href = f"#{anchor}"
            except Exception:
                href = ""
            href = _sanitize_href(href)
            # Bold/italic detection also via character styles
            is_bold = False
            try:
                is_bold = bool(run.bold) or bool(getattr(run.font, 'bold', False))
            except Exception:
                pass
            try:
                stname = run.style.name.lower() if run.style and run.style.name else ""
                if ('bold' in stname) or ('strong' in stname):
                    is_bold = True
            except Exception:
                pass
            if is_bold:
                s = f"<strong>{s}</strong>"
            is_italic = False
            try:
                is_italic = bool(run.italic)
            except Exception:
                pass
            try:
                stname = run.style.name.lower() if run.style and run.style.name else ""
                if ('emphasis' in stname) or ('italic' in stname):
                    is_italic = True
            except Exception:
                pass
            if is_italic:
                s = f"<em>{s}</em>"
            try:
                if bool(getattr(run.font, 'underline', False)):
                    s = f"<u>{s}</u>"
            except Exception:
                pass
            if href:
                display = s if s.strip() else html.escape(str(href))
                href_txt = html.escape(str(href))
                if href_txt:
                    # Show the target in parentheses for clarity
                    s = f"<a href=\"{href_txt}\">{display}</a> (<span class=\"link-href\">{href_txt}</span>)"
                else:
                    # If after sanitation it's empty, just keep the text
                    s = display
            chunks.append(s)
        content = "".join(chunks).strip()
        try:
            left = paragraph.paragraph_format.left_indent
            if left and hasattr(left, 'pt'):
                px = int(float(left.pt) * 1.33)
                return f'<p style="margin-left:{px}px;">{content}</p>'
        except Exception:
            pass
        # Fallback auto-link: add anchors for visible http(s) URLs in plain text
        if '<a ' not in content.lower():
            def _repl(m):
                raw = m.group(0)
                safe = _sanitize_href(raw)
                safe_html = html.escape(safe)
                disp = html.escape(raw)
                return f'<a href="{safe_html}">{disp}</a> (<span class="link-href">{safe_html}</span>)'
            content = _re.sub(r'https?://[^\s<>]+', _repl, content)
        return f"<p>{content}</p>"

    def _resolve_numbering_format(self, document, num_id: int, ilvl: int) -> tuple[bool, str]:
        cache = getattr(self, '_numfmt_cache', None)
        if cache is None:
            cache = {}
            self._numfmt_cache = cache
        key = (int(num_id), int(ilvl))
        if key in cache:
            return cache[key]
        try:
            part = getattr(document.part, 'numbering_part', None) or getattr(document.part, '_numbering_part', None)
            numbering_el = getattr(part, 'element', None) if part else None
            if numbering_el is not None:
                ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
                num_nodes = numbering_el.xpath(f".//w:num[@w:numId='{num_id}']/w:abstractNumId", namespaces=ns)
                if num_nodes:
                    abs_id = num_nodes[0].get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
                    lvl_nodes = numbering_el.xpath(f".//w:abstractNum[@w:abstractNumId='{abs_id}']/w:lvl[@w:ilvl='{ilvl}']/w:numFmt", namespaces=ns)
                    if lvl_nodes:
                        fmt = lvl_nodes[0].get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val') or 'decimal'
                        ordered = (fmt != 'bullet')
                        cache[key] = (ordered, fmt)
                        return cache[key]
        except Exception:
            pass
        cache[key] = (True, 'decimal')
        return cache[key]

    def _paragraph_list_info(self, paragraph, document) -> Optional[dict]:
        try:
            ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
            numPr = paragraph._p.xpath('w:pPr/w:numPr', namespaces=ns)
            if not numPr:
                raise AttributeError("no numPr")
            ilvl_nodes = paragraph._p.xpath('w:pPr/w:numPr/w:ilvl', namespaces=ns)
            num_nodes = paragraph._p.xpath('w:pPr/w:numPr/w:numId', namespaces=ns)
            ilvl = int(ilvl_nodes[0].get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')) if ilvl_nodes else 0
            num_id = int(num_nodes[0].get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')) if num_nodes else 0
            ordered, fmt = self._resolve_numbering_format(document, num_id, ilvl)
            return {"level": ilvl + 1, "ordered": bool(ordered), "fmt": str(fmt)}
        except Exception:
            # Heuristic fallback: inspect text prefix and left indent
            try:
                s = (paragraph.style.name or '').lower()
                text = (paragraph.text or '').strip()
                # Determine level from left indent (approx 18pt per level)
                level = 1
                try:
                    li = paragraph.paragraph_format.left_indent
                    if li and hasattr(li, 'pt'):
                        level = max(1, min(10, 1 + int(float(li.pt) / 18.0)))
                except Exception:
                    level = 1
                # Prefix checks
                import re as _re
                if _re.match(r"^\s*[\-\*•]\s+", text) or ('bullet' in s):
                    return {"level": level, "ordered": False, "fmt": 'bullet'}
                m = _re.match(r"^\s*(\d+)\s*[\).]", text)
                if m:
                    return {"level": level, "ordered": True, "fmt": 'decimal'}
                m = _re.match(r"^\s*([ivxlcdm]+)\s*[\).]", text, flags=_re.I)
                if m:
                    fmt = 'lowerRoman' if m.group(1).islower() else 'upperRoman'
                    return {"level": level, "ordered": True, "fmt": fmt}
                m = _re.match(r"^\s*([a-zA-Z])\s*[\).]", text)
                if m:
                    fmt = 'lowerLetter' if m.group(1).islower() else 'upperLetter'
                    return {"level": level, "ordered": True, "fmt": fmt}
            except Exception:
                pass
            return None
    # ------------------------------------------------------------------
    def load_word_documents(self, paths: Iterable[str], progress_callback=None, *, interactive: bool = True) -> pd.DataFrame:
        """Load requirement data from one or more Word documents (.docx)."""
        try:
            from docx import Document  # type: ignore
        except Exception as exc:
            raise RequirementDataError(
                "Loading Word documents requires the python-docx package."
            ) from exc

        all_records: List[dict[str, str]] = []
        self.front_matter_records = {}
        self.front_matter_count = {}

        for raw_path in paths:
            path = Path(raw_path)
            if not path.exists():
                LOGGER.warning("Word document %s does not exist", path)
                continue

            try:
                document = Document(path)
            except Exception as exc:
                LOGGER.exception("Failed to open Word document: %s", path)
                raise RequirementDataError(f"Failed to open Word document {path}: {exc}") from exc

            LOGGER.info("Parsing Word file: %s", path.name)
            # 1) Paragraphs + Tables (with image capture inside paragraphs)
            records = self._extract_records_from_document(document, path.name, path.stem, progress_callback)

            # 2) Full-package image sweep (body, headers, footers, etc.), de-duplicated
            extra_images = []
            # Avoid re-introducing front-page images when skipping front matter.
            # Only perform full-package sweep when we are not skipping the first page.
            if not getattr(self, "skip_front_matter_for_word", False):
                extra_images = self.extract_all_images_from_docx(document, source_name=path.name, sheet_name=path.stem)
            if extra_images:
                # de-dup by filename already present from paragraph-scan
                existing_names: set[str] = set()
                for r in records:
                    if r.get("Attachment Type", "").lower() == "image":
                        try:
                            payload = json.loads(r.get("Attachment Data", "") or "{}")
                            name = str(payload.get("filename", "")).strip()
                            if name:
                                existing_names.add(name)
                        except Exception:
                            pass

                # exclude images that were part of front-matter when that content is skipped
                front_image_names: set[str] = set()
                if getattr(self, "skip_front_matter_for_word", False):
                    try:
                        for fr in self.get_front_matter_records(path.name):
                            if str(fr.get("Attachment Type", "")).lower() == "image":
                                nm = self._safe_payload_name(fr.get("Attachment Data", ""))
                                if nm:
                                    front_image_names.add(nm)
                    except Exception:
                        pass

                def include_pkg_image(img: dict[str, str]) -> bool:
                    if str(img.get("Attachment Type", "")).lower() != "image":
                        return False
                    nm = self._safe_payload_name(img.get("Attachment Data", ""))
                    if not nm:
                        return False
                    if nm in existing_names:
                        return False
                    if nm in front_image_names:
                        return False
                    return True

                dedup = [img for img in extra_images if include_pkg_image(img)]
                if dedup:
                    LOGGER.info("Added %d more image(s) from package sweep", len(dedup))
                    records.extend(dedup)

            all_records.extend(records)

        if not all_records:
            raise RequirementDataError("No requirement content detected in Word files.")

        df = pd.DataFrame(all_records)
        df = self._clean_word_dataframe(df)
        self._validate_columns(df)
        finalized = self.finalize_dataframe(df)

        # Count total requirements
        req_count = finalized["Object Type"].str.lower().eq("requirement").sum()
        LOGGER.info("Total detected requirements: %d", req_count)

        # Ask user for Trace Direction (Up, Down, Bi-Directional)
        direction = "Bi-Directional"
        ok = True
        if interactive:
            try:
                from PyQt6.QtWidgets import QInputDialog
                direction, ok = QInputDialog.getItem(
                    None,
                    "Trace Direction",
                    f"{req_count} requirements found.\nSelect traceability direction:",
                    ["Up Trace", "Down Trace", "Bi-Directional"],
                    2,  # default to Bi-Directional
                    False,
                )
            except Exception:
                LOGGER.warning("PyQt not available for trace direction prompt; defaulting to Bi-Directional.")
                direction = "Bi-Directional"
                ok = True

        if not ok:
            direction = "Bi-Directional"

        trace_mode = direction.strip()
        LOGGER.info("User selected trace direction: %s", trace_mode)

        # Annotate
        finalized["Trace Direction"] = trace_mode
        if trace_mode == "Up Trace":
            finalized["Linked ID / Description"] = finalized["Up Trace"]
        elif trace_mode == "Down Trace":
            finalized["Linked ID / Description"] = finalized["Requirement ID"]
        else:
            finalized["Linked ID / Description"] = finalized["Up Trace"] + " / " + finalized["Requirement ID"]

        LOGGER.info("Traceability configuration complete: %s | Total Requirements: %d", trace_mode, req_count)
        return finalized

    # ------------------------------------------------------------------
    def _validate_columns(self, df: pd.DataFrame) -> None:
        missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
        if missing:
            raise RequirementDataError("Missing required columns: " + ", ".join(missing))
        for col in OPTIONAL_METADATA_COLUMNS:
            if col not in df.columns:
                df[col] = ""

    # ------------------------------------------------------------------
    def _apply_section_numbering(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Auto-number headings while keeping Object Type clean.
        Adds numbering prefix (like '1.2 Scope') into Object Text.
        Supports both explicit numbering already present and automatic numbering.
        """
        import re
        if df is None or df.empty:
            return df

        df = df.copy()
        num_re = re.compile(r'^\s*(\d+(?:\s*\.\s*\d+)*)\.?\s+')

        def strip_leading_number(s: str) -> tuple[str, str | None]:
            """
            Return (clean_text, section_label or None).
            Example:
                '1.2  Scope'  -> ('Scope', '1.2')
                '  3.  Intro' -> ('Intro', '3')
                'Text only'   -> ('Text only', None)
            """
            s = str(s or '')
            m = num_re.match(s)
            if not m:
                return s.strip(), None
            label = re.sub(r'\s*\.\s*', '.', m.group(1)).strip('.')  # normalize spaces
            rest = s[m.end():].strip()
            return rest, label

        current_source = None
        # Track counters for Heading levels 1..10
        h = [0]*10  # indices 0..9 represent levels 1..10
        new_obj_types: list[str] = []
        new_obj_texts: list[str] = []

        for _, row in df.iterrows():
            source = row.get("SourceFile", "")
            if source != current_source:
                current_source = source
                h = [0]*10

            raw_type = str(row.get("Object Type", "") or "").strip()
            raw_type_l = raw_type.lower()

            text = str(row.get("Object Text", "") or "")
            clean_text, explicit_label = strip_leading_number(text)

            # --- Explicit numbering in text (like "1.2 Scope" or "1.2.3.4 Title") ---
            if explicit_label and not self.ignore_explicit_heading_numbers:
                parts = explicit_label.split('.')
                level = min(len(parts), 10)
                try:
                    nums = [int(p) for p in parts]
                except ValueError:
                    nums = []
                # Reset all counters, then set according to explicit parts
                h = [0]*10
                for i, n in enumerate(nums[:10]):
                    h[i] = n

                obj_type_label = f"Heading {level}"     # keep Object Type clean
                new_obj_types.append(obj_type_label)
                new_obj_texts.append(f"{explicit_label} {clean_text}".strip())
                continue

            # --- Auto-numbering case (or explicit labels ignored) ---
            # Generic handler for Heading 1..10
            m_h = re.match(r"heading\s*(\d+)", raw_type_l)
            if m_h:
                lvl = max(1, min(10, int(m_h.group(1))))
                # Ensure parent levels initialized
                if h[0] == 0:
                    h[0] = 1
                for j in range(1, lvl-1):
                    if h[j] == 0:
                        h[j] = 1
                h[lvl-1] += 1
                for j in range(lvl, 10):
                    h[j] = 0
                label = ".".join(str(h[i]) for i in range(lvl))
                new_obj_types.append(f"Heading {lvl}")
                new_obj_texts.append(f"{label} {clean_text}".strip())
            else:
                # Non-heading rows stay the same
                new_obj_types.append(raw_type)
                new_obj_texts.append(clean_text)

        # Commit updates
        df["Object Type"] = new_obj_types
        df["Object Text"] = new_obj_texts

        # Final tidy-up
        df = self._canonicalize_columns(df)
        df = self._reorder_visible_columns(df)
        return df



    # ------------------------------------------------------------------
    @property
    def visible_columns(self) -> List[str]:
        """Columns to show in the grid, prioritized in canonical order. Extra columns follow."""
        if self.dataframe.empty:
            return [c for c in self.DEFAULT_COLUMN_ORDER if c != self.section_column_name]
        ordered = [c for c in self.DEFAULT_COLUMN_ORDER if c in self.dataframe.columns]
        tail = [c for c in self.dataframe.columns if c not in ordered]
        cols = [*ordered, *tail]
        # Hide the auto-generated section column from the Excel preview/table
        return [c for c in cols if c != self.section_column_name]


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
            index = (len(df) if insert_at is None or insert_at < 0 or insert_at > len(df) else insert_at)
            upper = df.iloc[:index]
            lower = df.iloc[index:]
            df = pd.concat([upper, pd.DataFrame([record]), lower], ignore_index=True)

        self._validate_columns(df)
        self.dataframe = self.finalize_dataframe(df)
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
            existing = existing[~existing.get("SourceFile", pd.Series(dtype=str)).isin(sources)]
            merged = pd.concat([existing, working], ignore_index=True)

        merged = self.finalize_dataframe(merged)
        self.dataframe = merged
        return self.dataframe

    # ------------------------------------------------------------------
    def configure_requirement_pattern(self, config: Optional[dict[str, str]] | None) -> None:
        self._custom_patterns.clear()
        if not config:
            return
        mode = str(config.get("mode", "prefixes")).strip().lower()
        value = str(config.get("value", "")).strip()
        if not value:
            return
        patterns: list[re.Pattern[str]] = []
        if mode == "regex":
            try:
                patterns.append(re.compile(value, re.IGNORECASE))
            except re.error as exc:
                LOGGER.warning("Invalid custom regex '%s': %s", value, exc)
                return
        else:
            prefixes = [p.strip() for p in value.split(",") if p.strip()]
            if not prefixes:
                return
            escaped = "|".join(re.escape(p) for p in prefixes)
            expressions = [
                rf"^\s*(?P<id>(?:{escaped})[A-Za-z0-9_.-]*)\s*[:\-–]\s*(?P<body>.+)$",
                rf"^\s*(?P<id>(?:{escaped})[A-Za-z0-9_.-]*)\s+(?P<body>shall.+)$",
            ]
            try:
                patterns = [re.compile(expr, re.IGNORECASE) for expr in expressions]
            except re.error as exc:
                LOGGER.warning("Failed to compile prefix patterns '%s': %s", value, exc)
                return
        self._custom_patterns = patterns

    # ------------------------------------------------------------------
    def finalize_dataframe(self, df: pd.DataFrame) -> pd.DataFrame:
        """Apply section numbering and normalize requirement text for consistency."""
        if df is None:
            return pd.DataFrame()
        if df.empty:
            return df.copy()

        numbered = self._apply_section_numbering(df)
        normalized = self._normalize_requirement_records(numbered)
        # Normalize table captions to be above the table and always present
        normalized = self._normalize_table_captions(normalized)
        # Also ensure Figure captions are renumbered; table text is preserved
        normalized = self._renumber_captions_per_source(normalized)
        return normalized

    # ------------------------------------------------------------------
    def refresh_dataframe(self) -> pd.DataFrame:
        """Refresh the in-memory dataframe after edits."""
        self.dataframe = self.finalize_dataframe(self.dataframe)
        return self.dataframe

    # ------------------------------------------------------------------
    def configure_requirement_patterns_strict(self, mode: str, value: str) -> None:
        """Strict pattern configuration.
        - mode 'regex': compile raw expressions (one per line).
        - mode 'prefixes': build full ID matchers from prefixes (one per line or comma-separated).
        """
        mode_l = str(mode or "").strip().lower()
        if mode_l == "regex":
            self._custom_patterns = compile_user_patterns(value, manual_regex=True)
        else:
            self._custom_patterns = build_requirement_patterns_from_prefixes(value)

    # ------------------------------------------------------------------
    def create_empty_dataframe(self) -> pd.DataFrame:
        """Return a canonical empty dataframe used to initialize the UI/state.

        Includes the default visible columns; downstream operations will
        add the section column and any missing optional columns as needed.
        """
        cols = list(self.DEFAULT_COLUMN_ORDER)
        return pd.DataFrame(columns=cols)

    # ------------------------------------------------------------------
    def _extract_records_from_document(self, document, source_name: str, sheet_name: str, progress_callback=None) -> List[dict[str, str]]:
        """Extract requirement data, headings, text, tables, and images (from paragraphs) from a DOCX file."""
        from docx.table import Table  # type: ignore
        from docx.text.paragraph import Paragraph  # type: ignore

        records: List[dict[str, str]] = []
        front_records: List[Dict[str, str]] = []
        body_started = False
        first_page_done = False
        seen_req_ids: set[str] = set()

        def add_record(payload: Dict[str, str], *, mark_body: bool = False) -> None:
            nonlocal body_started
            if (mark_body or first_page_done) and not body_started:
                body_started = True
            records.append(payload)
            if not body_started:
                front_records.append(payload.copy())

        blocks = list(self._iter_doc_blocks(document))
        total_blocks = len(blocks)
        skip_until = 0  # 1-based index in 'blocks' to skip ahead after aggregating requirement text
        for idx, block in enumerate(blocks, 1):
            if skip_until and idx <= skip_until:
                continue
            if progress_callback and total_blocks > 0:
                percent = int((idx / total_blocks) * 100)
                progress_callback(percent)

            if isinstance(block, Paragraph):
                style_name = (block.style.name or "").lower()
                # Detect if this paragraph carries a page/section boundary
                try:
                    paragraph_has_boundary = self._paragraph_has_page_boundary(block)
                except Exception:
                    paragraph_has_boundary = False

                try:
                    has_img = getattr(self, "_paragraph_has_image", None)
                    if callable(has_img) and has_img(block):
                        for img_record in self._collect_paragraph_images(block, source_name, sheet_name):
                            add_record(img_record)
                except Exception:
                    pass

                text = block.text.strip()
                if not text:
                    continue

                if self._is_toc_entry(style_name, text):
                    LOGGER.debug("Skipping auto-generated section: %s", text[:60])
                    continue

                # First, classify by style name
                try:
                    obj_type = self._classify_paragraph(style_name)
                except AttributeError:
                    s = (style_name or "").lower()
                    if "heading 1" in s:
                        obj_type = "Heading 1"
                    elif "heading 2" in s:
                        obj_type = "Heading 2"
                    elif "heading 3" in s:
                        obj_type = "Heading 3"
                    elif "heading 4" in s:
                        obj_type = "Heading 4"
                    elif "heading 5" in s:
                        obj_type = "Heading 5"
                    elif "heading 6" in s:
                        obj_type = "Heading 6"
                    elif "heading 7" in s:
                        obj_type = "Heading 7"
                    elif "heading 8" in s:
                        obj_type = "Heading 8"
                    elif "heading 9" in s:
                        obj_type = "Heading 9"
                    elif "heading 10" in s:
                        obj_type = "Heading 10"
                    else:
                        obj_type = "Requirement"

                # Only treat as list if it is NOT a Heading by style
                if not str(obj_type).lower().startswith("heading"):
                    try:
                        list_meta = self._paragraph_list_info(block, document)
                    except AttributeError:
                        list_meta = None
                    if list_meta:
                        try:
                            html_inline = self._paragraph_inlines_to_html(block)
                        except AttributeError:
                            html_inline = f"<p>{html.escape(text)}</p>"
                        add_record({
                            "Object Type": "Text",
                            "Requirement ID": "",
                            "Object Text": text,
                            "Up Trace": "",
                            "SourceFile": source_name,
                            "SheetName": sheet_name,
                            "SourceType": "docx",
                            "Attachment Type": "list_item",
                            "Attachment Data": json.dumps({
                                "level": int(list_meta.get("level", 1)),
                                "ordered": bool(list_meta.get("ordered", True)),
                                "fmt": str(list_meta.get("fmt", "decimal")),
                                "html": html_inline if html_inline else html.escape(text),
                            }),
                        })
                        if paragraph_has_boundary and not first_page_done:
                            first_page_done = True
                        continue
                try:
                    req_id, heading_hint, body_text = self._maybe_extract_requirement_strict(text)
                except AttributeError:
                    rid = ""; heading_hint = ""; body = text
                    try:
                        m = re.match(r"\s*([A-Za-z][A-Za-z0-9_.\-]*\d+)\s*[:\-]\s*(.+)", text)
                        if not m:
                            m = re.match(r"\s*([A-Za-z][A-Za-z0-9_.\-]*\d+)\s+(shall|must|should).+", text, flags=re.I)
                        if m:
                            rid = clean_id(m.group(1))
                            body = text[m.end(1):].lstrip(" :-") if m.lastindex and m.lastindex >= 2 else text
                    except Exception:
                        rid = ""; body = text
                    req_id, heading_hint, body_text = rid, "", body

                LOGGER.debug("Paragraph: %s", text[:100])
                if req_id:
                    LOGGER.debug("  matched Requirement ID: %s", req_id)

                    # Determine which ID to link from heading: prefer next LLR id if present
                    link_req_id = str(req_id)
                    j = idx + 1
                    while j <= total_blocks:
                        nxt = blocks[j - 1]
                        from docx.text.paragraph import Paragraph as _P
                        from docx.table import Table as _T
                        if isinstance(nxt, _T):
                            break
                        if isinstance(nxt, _P):
                            nxt_text = nxt.text.strip() if nxt.text else ""
                            if not nxt_text:
                                j += 1
                                continue
                            # Stop at next heading-like or next requirement ID
                            nxt_style = (nxt.style.name or "").lower()
                            if nxt_style.startswith("heading"):
                                break
                            # Next requirement ID?
                            rid2, _, _ = self._maybe_extract_requirement_strict(nxt_text)
                            if rid2:
                                if "LLR" in str(rid2).upper():
                                    link_req_id = str(rid2)
                                break
                            # Add this paragraph as its own Text row (split by internal newlines)
                            for line in [ln.strip() for ln in nxt_text.splitlines() if ln.strip()]:
                                add_record({
                                    "Object Type": "Text",
                                    "Requirement ID": "",
                                    "Object Text": line,
                                    "Up Trace": "",
                                    "SourceFile": source_name,
                                    "SheetName": sheet_name,
                                    "SourceType": "docx",
                                    "Attachment Type": "",
                                    "Attachment Data": "",
                                }, mark_body=True)
                            j += 1
                            continue
                        break
                    # Insert heading row (above requirement) with deduped numeric and linked ID
                    if heading_hint or str(obj_type).lower().startswith("heading"):
                        heading_text = heading_hint.strip() if heading_hint else str(req_id).strip()
                        heading_text = heading_text.replace("\n", " ").strip()
                        heading_text = _dedup_section_label(heading_text)
                        m_head = re.match(r"^\s*(\d+(?:\.\d+)*)\b", heading_text)
                        if m_head:
                            depth = 1 + m_head.group(1).count(".")
                            depth = max(1, min(depth, 10))
                            add_record({
                                "Object Type": f"Heading {depth}",
                                "Requirement ID": "",
                                "Object Text": (f"{heading_text} {link_req_id}".strip() if link_req_id else heading_text),
                                "Up Trace": "",
                                "SourceFile": source_name,
                                "SheetName": sheet_name,
                                "SourceType": "docx",
                                "Attachment Type": "",
                                "Attachment Data": "",
                                "Linked ID / Description": link_req_id,
                            }, mark_body=True)
                        else:
                            add_record({
                                "Object Type": "Heading",
                                "Requirement ID": "",
                                "Object Text": (f"{heading_text} {link_req_id}".strip() if link_req_id else heading_text),
                                "Up Trace": "",
                                "SourceFile": source_name,
                                "SheetName": sheet_name,
                                "SourceType": "docx",
                                "Attachment Type": "",
                                "Attachment Data": "",
                                "Linked ID / Description": link_req_id,
                            }, mark_body=True)
                    # Classify requirement type based on ID token (optional)
                    obj_req_type = "Requirement"
                    up_id = str(req_id).upper()
                    if "LLR" in up_id:
                        obj_req_type = "Requirement LLR"
                    elif "HLR" in up_id:
                        obj_req_type = "Requirement HLR"
                    elif "-SRS-" in up_id or up_id.endswith("-SRS") or up_id.startswith("SRS-") or " SRS-" in up_id:
                        obj_req_type = "Requirement System"

                    if str(req_id) in seen_req_ids:
                        if paragraph_has_boundary and not first_page_done:
                            first_page_done = True
                        # Skip duplicate requirement IDs
                        continue
                    seen_req_ids.add(str(req_id))

                    add_record({
                        "Object Type": obj_req_type,
                        "Requirement ID": req_id,
                        "Object Text": "",  # keep body in separate Text rows
                        "Up Trace": "",
                        "SourceFile": source_name,
                        "SheetName": sheet_name,
                        "SourceType": "docx",
                        "Attachment Type": "",
                        "Attachment Data": "",
                    }, mark_body=True)
                    # Add the immediate body (right-of-ID) as first Text row if present
                    if body_text and body_text.strip():
                        for line in [ln.strip() for ln in body_text.splitlines() if ln.strip()]:
                            add_record({
                                "Object Type": "Text",
                                "Requirement ID": "",
                                "Object Text": line,
                                "Up Trace": "",
                                "SourceFile": source_name,
                                "SheetName": sheet_name,
                                "SourceType": "docx",
                                "Attachment Type": "",
                                "Attachment Data": "",
                            }, mark_body=True)
                    if j > idx + 1:
                        skip_until = j - 1
                    if paragraph_has_boundary and not first_page_done:
                        first_page_done = True
                    continue

                # Fallback: detect numbered headings like "1 Introduction" or "2.3 Scope"
                # Only treat as heading if it does not look like a TOC entry line
                if not obj_type.startswith("Heading"):
                    if not self._is_toc_entry_line(text):
                        # Avoid list/bullet paragraphs being misclassified as headings
                        if ("list" in style_name) or ("bullet" in style_name) or re.match(r"^\s*([a-zA-Z]|[ivxlcdmIVXLCDM]+)\s*[\).]", text):
                            m = None
                        else:
                            m = re.match(r"^\s*(\d+(?:\.\d+)*)\s+\S", text)
                        if m:
                            depth = 1 + m.group(1).count(".")
                            depth = max(1, min(depth, 10))
                            add_record(
                                {
                                    "Object Type": f"Heading {depth}",
                                    "Requirement ID": "",
                                    "Object Text": text,
                                    "Up Trace": "",
                                    "SourceFile": source_name,
                                    "SheetName": sheet_name,
                                    "SourceType": "docx",
                                    "Attachment Type": "",
                                    "Attachment Data": "",
                                },
                                mark_body=True,
                            )
                            if paragraph_has_boundary and not first_page_done:
                                first_page_done = True
                            continue

                if req_id and obj_type == "Requirement":
                    add_record({
                        "Object Type": "Requirement",
                        "Requirement ID": req_id,
                        "Object Text": body_text,
                        "Up Trace": "",
                        "SourceFile": source_name,
                        "SheetName": sheet_name,
                        "SourceType": "docx",
                        "Attachment Type": "",
                        "Attachment Data": "",
                    })
                    if paragraph_has_boundary and not first_page_done:
                        first_page_done = True
                    continue

                if obj_type.startswith("Heading"):
                    add_record({
                        "Object Type": obj_type,
                        "Requirement ID": "",
                        "Object Text": text,
                        "Up Trace": "",
                        "SourceFile": source_name,
                        "SheetName": sheet_name,
                        "SourceType": "docx",
                        "Attachment Type": "",
                        "Attachment Data": "",
                    }, mark_body=True)
                    if paragraph_has_boundary and not first_page_done:
                        first_page_done = True
                    continue

                try:
                    html_inline = self._paragraph_inlines_to_html(block)
                except AttributeError:
                    html_inline = f"<p>{html.escape(body_text or text)}</p>"
                add_record({
                    "Object Type": "Text",
                    "Requirement ID": req_id or "",
                    "Object Text": body_text or text,
                    "Up Trace": "",
                    "SourceFile": source_name,
                    "SheetName": sheet_name,
                    "SourceType": "docx",
                    "Attachment Type": "html",
                    "Attachment Data": html_inline,
                })
                if paragraph_has_boundary and not first_page_done:
                    first_page_done = True
                continue

            if isinstance(block, Table):
                try:
                    html_table = self._table_to_html(block)
                except AttributeError:
                    # Fallback simple table converter
                    try:
                        rows_html: List[str] = []
                        for row in block.rows:
                            cells_html: List[str] = []
                            for cell in row.cells:
                                txt = html.escape(cell.text.strip()) if cell.text else "&nbsp;"
                                cells_html.append(f"<td>{txt}</td>")
                            rows_html.append("<tr>" + "".join(cells_html) + "</tr>")
                        html_table = (
                            '<table border="1" style="border-collapse: collapse; width: 100%;">'
                            + "".join(rows_html)
                            + "</table>"
                        )
                    except Exception:
                        html_table = ""
                if not html_table.strip():
                    continue
                add_record({
                    "Object Type": "Table",
                    "Requirement ID": "",
                    "Object Text": "",
                    "Up Trace": "",
                    "SourceFile": source_name,
                    "SheetName": sheet_name,
                    "SourceType": "docx",
                    "Attachment Type": "table",
                    "Attachment Data": html_table,
                })

        self.front_matter_records[source_name] = [copy.deepcopy(r) for r in front_records]
        try:
            self.front_matter_count[source_name] = int(len(front_records))
        except Exception:
            self.front_matter_count[source_name] = 0

        if not records:
            LOGGER.warning("No paragraphs or tables parsed from %s", source_name)
            return records

        return records

    def _clean_word_dataframe(self, df: pd.DataFrame) -> pd.DataFrame:
        if df is None or df.empty:
            return pd.DataFrame() if df is None else df.copy()

        work = df.reset_index(drop=True)

        # Remove auto-generated Table of Contents / List of Figures / List of Tables sections
        drop_idx: List[int] = []
        skip_mode: Optional[str] = None
        for idx, row in work.iterrows():
            obj_type = str(row.get("Object Type", "") or "").lower()
            text = str(row.get("Object Text", "") or "").strip()
            if obj_type.startswith("heading"):
                skip_mode = None
                mode = self._is_auto_list_heading(text)
                if mode:
                    skip_mode = mode
                    drop_idx.append(idx)
                continue
            if skip_mode:
                # While in skip_mode drop everything until the next heading,
                # regardless of exact entry formatting — this suppresses
                # imported TOC/LOF/LOT blocks robustly.
                drop_idx.append(idx)
                continue

        if drop_idx:
            work = work.drop(drop_idx).reset_index(drop=True)

        if getattr(self, "skip_front_matter_for_word", False) and {"SourceFile", "Object Type"}.issubset(work.columns):
            drop_idx = []
            have_counts = bool(getattr(self, 'front_matter_count', {}))
            for src, sub in work.groupby(work["SourceFile"].astype(str), sort=False):
                if have_counts and str(src) in self.front_matter_count:
                    n = int(self.front_matter_count.get(str(src), 0))
                    if n > 0:
                        drop_idx.extend(list(sub.index[:n]))
                        continue
                heading_mask = sub["Object Type"].astype(str).str.lower().str.startswith("heading")
                if not heading_mask.any():
                    # Try numeric heading detection: first row whose Object Text starts with 1 / 1.2 etc.
                    try:
                        nt = sub["Object Text"].astype(str).str.match(r"^\s*\d+(?:\.\d+)*\s+\S")
                        heading_mask = nt
                    except Exception:
                        pass
                if heading_mask.any():
                    first_idx = heading_mask[heading_mask].index[0]
                    drop_idx.extend([i for i in sub.index if i < first_idx])
            if drop_idx:
                work = work.drop(drop_idx).reset_index(drop=True)

        return work

    def get_front_matter_records(self, source_name: str) -> List[Dict[str, str]]:
        records = self.front_matter_records.get(source_name, [])
        return [copy.deepcopy(r) for r in records]

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
    def _classify_paragraph(self, style_name: str) -> str:
        style_lower = style_name.lower()
        if "heading 1" in style_lower:
            return "Heading 1"
        if "heading 2" in style_lower:
            return "Heading 2"
        if "heading 3" in style_lower:
            return "Heading 3"
        if "heading 4" in style_lower:
            return "Heading 4"
        if "heading 5" in style_lower:
            return "Heading 5"
        if "heading 6" in style_lower:
            return "Heading 6"
        if "heading 7" in style_lower:
            return "Heading 7"
        if "heading 8" in style_lower:
            return "Heading 8"
        if "heading 9" in style_lower:
            return "Heading 9"
        if "heading 10" in style_lower:
            return "Heading 10"
        if "list" in style_lower or "bullet" in style_lower:
            return "Text"
        return "Requirement"

    # ------------------------------------------------------------------
    def _paragraph_has_image(self, paragraph) -> bool:
        """Detect inline/floating images, including legacy VML, inside a paragraph."""
        a_ns = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
        vml_ns = "{urn:schemas-microsoft-com:vml}"
        for run in paragraph.runs:
            el = run._element
            if el.findall(f".//{a_ns}blip") or el.findall(f".//{a_ns}drawing"):
                return True
            # Legacy VML image data
            if el.findall(f".//{vml_ns}imagedata"):
                return True
        return False

    # ------------------------------------------------------------------
    def _paragraph_has_page_boundary(self, paragraph) -> bool:
        """Detect explicit page/section boundaries in a paragraph.

        Looks for:\n- w:br w:type="page" (manual page break)\n- w:pPr/w:pageBreakBefore\n- w:pPr/w:sectPr (section break that typically starts a new page)
        """
        try:
            # search runs for explicit page breaks
            for run in paragraph.runs:
                el = run._element
                for br in el.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}br'):
                    t = (br.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type') or '').lower()
                    if t == 'page':
                        return True
        except Exception:
            pass

        try:
            # paragraph properties: pageBreakBefore or sectPr
            p_el = paragraph._element
            if p_el.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pageBreakBefore'):
                return True
            if p_el.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sectPr'):
                return True
        except Exception:
            pass
        return False

    # ------------------------------------------------------------------
    def _collect_paragraph_images(self, paragraph, source_name: str, sheet_name: str) -> List[dict[str, str]]:
        """Extract images (inline, floating, or legacy VML) from a paragraph."""
        from docx.oxml.ns import qn  # type: ignore

        records: List[dict[str, str]] = []
        seen: set[str] = set()
        el = None

        for run in paragraph.runs:
            el = run._element
            # DrawingML images
            for blip in el.findall(".//{http://schemas.openxmlformats.org/drawingml/2006/main}blip"):
                rel_id = blip.get(qn("r:embed"))
                if not rel_id or rel_id in seen:
                    continue
                seen.add(rel_id)
                part = paragraph.part.related_parts.get(rel_id)
                if not part:
                    continue
                mime = getattr(part, "content_type", "image/png")
                filename = Path(str(part.partname)).name
                payload = json.dumps({
                    "mime": mime,
                    "data": base64.b64encode(part.blob).decode("ascii"),
                    "filename": filename,
                })
                records.append({
                    "Object Type": "Image",
                    "Requirement ID": "",
                    "Object Text": filename,
                    "Up Trace": "",
                    "SourceFile": source_name,
                    "SheetName": sheet_name,
                    "SourceType": "docx",
                    "Attachment Type": "image",
                    "Attachment Data": payload,
                })

            # Legacy VML images (v:imagedata r:id="..." )
            for imdata in el.findall(".//{urn:schemas-microsoft-com:vml}imagedata"):
                rel_id = imdata.get(qn("r:id")) or imdata.get(qn("r:embed"))
                if not rel_id or rel_id in seen:
                    continue
                seen.add(rel_id)
                part = paragraph.part.related_parts.get(rel_id)
                if not part:
                    continue
                mime = getattr(part, "content_type", "image/png")
                filename = Path(str(part.partname)).name
                payload = json.dumps({
                    "mime": mime,
                    "data": base64.b64encode(part.blob).decode("ascii"),
                    "filename": filename,
                })
                records.append({
                    "Object Type": "Image",
                    "Requirement ID": "",
                    "Object Text": filename,
                    "Up Trace": "",
                    "SourceFile": source_name,
                    "SheetName": sheet_name,
                    "SourceType": "docx",
                    "Attachment Type": "image",
                    "Attachment Data": payload,
                })

        return records

    # ------------------------------------------------------------------
    def extract_all_images_from_docx(self, document, *, source_name: str, sheet_name: str) -> List[dict[str, str]]:
        """
        Scan the entire DOCX package (body, headers, footers, etc.) and extract
        all embedded images using relationship traversal (no external deps).
        """
        image_records: List[dict[str, str]] = []
        try:
            from docx.opc.package import Package  # type: ignore
        except Exception:
            # Fallback: minimal traversal from document.part
            return self._bfs_collect_images_from_part(document.part, source_name, sheet_name)

        pkg = document.part.package if hasattr(document.part, "package") else None
        if pkg is None or not hasattr(pkg, "parts"):
            # Fallback to BFS from document.part
            return self._bfs_collect_images_from_part(document.part, source_name, sheet_name)

        # Traverse every part reachable from the document root
        seen_parts: set[str] = set()
        pending: List[object] = [document.part]

        while pending:
            part = pending.pop()
            partname = str(getattr(part, "partname", ""))
            if partname in seen_parts:
                continue
            seen_parts.add(partname)

            # collect images on this part (ignore header/footer parts to avoid front-page bleed)
            for rel in getattr(part, "rels", {}).values():
                try:
                    reltype = str(rel.reltype).lower()
                except Exception:
                    reltype = ""
                if "image" in reltype:
                    try:
                        img_part = rel.target_part
                    except Exception:
                        img_part = None
                    if not img_part:
                        continue
                    # Skip images sourced from header/footer parts
                    source_partname = str(getattr(part, "partname", ""))
                    if source_partname:
                        src_lower = source_partname.lower()
                        if ("/word/header" in src_lower) or ("/word/footer" in src_lower):
                            continue
                    filename = Path(str(getattr(img_part, "partname", "image"))).name
                    mime = getattr(img_part, "content_type", "image/png")
                    blob = getattr(img_part, "blob", None)
                    if not blob:
                        continue

                    payload = json.dumps({
                        "mime": mime,
                        "data": base64.b64encode(blob).decode("ascii"),
                        "filename": filename,
                    })

                    image_records.append({
                        "Object Type": "Image",
                        "Requirement ID": "",
                        "Object Text": filename,
                        "Up Trace": "",
                        "SourceFile": source_name,
                        "SheetName": sheet_name,
                        "SourceType": "docx",
                        "Attachment Type": "image",
                        "Attachment Data": payload,
                    })

                # traverse deeper
                try:
                    tgt = rel.target_part
                    if tgt is not None:
                        tname = str(getattr(tgt, "partname", ""))
                        if tname and tname not in seen_parts:
                            pending.append(tgt)
                except Exception:
                    pass

        # De-dup by filename (package usually ensures unique names)
        if image_records:
            uniq = {}
            for r in image_records:
                try:
                    name = self._safe_payload_name(r.get("Attachment Data", ""))
                except Exception:
                    name = ""
                if name and name not in uniq:
                    uniq[name] = r
            image_records = list(uniq.values())

        return image_records

    # ------------------------------------------------------------------
    def _bfs_collect_images_from_part(self, root_part, source_name: str, sheet_name: str) -> List[dict[str, str]]:
        """Fallback: BFS starting from a given part to collect image relationships."""
        image_records: List[dict[str, str]] = []
        seen_parts: set[str] = set()
        pending: List[object] = [root_part]

        while pending:
            part = pending.pop()
            partname = str(getattr(part, "partname", ""))
            if partname in seen_parts:
                continue
            seen_parts.add(partname)

            for rel in getattr(part, "rels", {}).values():
                try:
                    reltype = str(rel.reltype).lower()
                except Exception:
                    reltype = ""
                if "image" in reltype:
                    try:
                        img_part = rel.target_part
                    except Exception:
                        img_part = None
                    if not img_part:
                        continue
                    filename = Path(str(getattr(img_part, "partname", "image"))).name
                    mime = getattr(img_part, "content_type", "image/png")
                    blob = getattr(img_part, "blob", None)
                    if not blob:
                        continue

                    payload = json.dumps({
                        "mime": mime,
                        "data": base64.b64encode(blob).decode("ascii"),
                        "filename": filename,
                    })

                    image_records.append({
                        "Object Type": "Image",
                        "Requirement ID": "",
                        "Object Text": filename,
                        "Up Trace": "",
                        "SourceFile": source_name,
                        "SheetName": sheet_name,
                        "SourceType": "docx",
                        "Attachment Type": "image",
                        "Attachment Data": payload,
                    })

                # traverse deeper
                try:
                    tgt = rel.target_part
                    if tgt is not None:
                        tname = str(getattr(tgt, "partname", ""))
                        if tname and tname not in seen_parts:
                            pending.append(tgt)
                except Exception:
                    pass

        # De-dup by filename
        if image_records:
            uniq = {}
            for r in image_records:
                try:
                    name = self._safe_payload_name(r.get("Attachment Data", ""))
                except Exception:
                    name = ""
                if name and name not in uniq:
                    uniq[name] = r
            image_records = list(uniq.values())

        return image_records

    # ------------------------------------------------------------------
    def _safe_payload_name(self, payload: str) -> str:
        try:
            data = json.loads(payload or "{}")
            return str(data.get("filename", "")).strip()
        except Exception:
            return ""

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
        """Split a paragraph into requirement ID and body text using robust pattern search."""
        text = text.strip()
        if not text:
            return "", ""

        rid = ""
        body = text

        # 1) Custom patterns (prefix/regex)
        for pattern in self._custom_patterns:
            match = pattern.search(text)
            if match:
                rid = clean_id(match.group("id") if "id" in match.groupdict() else match.group(0))
                body = match.group("body") if "body" in match.groupdict() else text
                LOGGER.debug("Matched custom pattern → ReqID='%s'", rid)
                return rid, body.strip()

        # 2) Built-in fallback patterns
        base_patterns = [
            r"(?:Requirement|Req)\s*(?:ID)?\s*[:\-–]\s*(?P<id>[A-Za-z0-9_.\-]+)\s*(?P<body>.*)",
            r"(?P<id>[A-Za-z]{1,}[A-Za-z0-9_.\-]*\d+)\s*[:\-–]\s*(?P<body>.+)",
            r"(?P<id>[A-Za-z]{1,}[A-Za-z0-9_.\-]*\d+)\s+(?P<body>(?:shall|must|should).+)",
        ]
        for expr in base_patterns:
            match = re.match(expr, text, flags=re.IGNORECASE)
            if match:
                rid = clean_id(match.group("id"))
                body = match.group("body") or text
                LOGGER.debug("Matched fallback pattern → ReqID='%s'", rid)
                return rid, body.strip()

        # 3) Prefix-based fallback
        if hasattr(self, "prefix_list") and self.prefix_list:
            for pref in self.prefix_list:
                regexes = build_regexes_from_input(pref)
                for regex in regexes:
                    found = regex.findall(text)
                    if found:
                        rid = clean_id(found[0])
                        body = text.split(found[0], 1)[-1].strip()
                        LOGGER.debug("Matched prefix pattern '%s' → ReqID='%s'", pref, rid)
                        return rid, body

        # 4) No match
        LOGGER.debug("No ReqID found in paragraph: %s", text[:80])
        return "", text

    # ------------------------------------------------------------------
    def _maybe_extract_requirement_strict(self, text: str) -> tuple[str, str, str]:
        """Split a paragraph using only user-defined patterns. No default detection.
        Returns (req_id, heading_hint_left_of_id, body_right_of_id).
        """
        text = (text or "").strip()
        if not text:
            return "", "", ""

        for pattern in self._custom_patterns:
            match = pattern.search(text)
            if match:
                rid = match.group("id") if "id" in match.groupdict() else match.group(0)
                # If body not captured, derive it by slicing the text around the ID
                if "body" in match.groupdict() and match.group("body") is not None:
                    # Capture any preceding heading hint from the remaining left side
                    try:
                        start = match.start("id") if "id" in match.groupdict() else match.start(0)
                    except Exception:
                        start = text.find(str(rid)) if str(rid) in text else 0
                    left = text[:start]
                    left = re.sub(r"(?i)Requirement\s*(ID|No)\s*[:\-–]?\s*$", "", left).strip()
                    body = match.group("body")
                else:
                    try:
                        start = match.start("id") if "id" in match.groupdict() else match.start(0)
                        end = match.end("id") if "id" in match.groupdict() else match.end(0)
                    except Exception:
                        start = text.find(str(rid)) if str(rid) in text else 0
                        end = start + len(str(rid))
                    left = text[:start]
                    right = text[end:]
                    # Strip any 'Requirement ID' label near the ID (left side)
                    left = re.sub(r"(?i)Requirement\s*(ID|No)\s*[:\-–]?\s*$", "", left).strip()
                    body = right.strip()
                LOGGER.debug("Matched custom pattern -> ReqID='%s'", rid)
                return str(rid).strip(), str(left).strip(), str(body).strip()

        LOGGER.debug("No ReqID match (custom-only) for: %s", text[:80])
        return "", "", text
    # ------------------------------------------------------------------
    def _normalize_requirement_records(self, df: pd.DataFrame) -> pd.DataFrame:
        if df.empty or "Object Text" not in df.columns or "Requirement ID" not in df.columns:
            return df

        cleaned = df.copy()

        def _normalize_row(row: pd.Series) -> str:
            req_id = str(row.get("Requirement ID", "") or "").strip()
            text_value = row.get("Object Text", "")
            if pd.isna(text_value):
                text_value = ""
            text_str = str(text_value).strip()
            if not req_id or not text_str:
                return text_str

            patterns = [
                rf"^\s*Requirement\s*ID\s*[:\-–]\s*{re.escape(req_id)}\s*",
                rf"^\s*Req(?:uirement)?\s*ID\s*[:\-–]\s*{re.escape(req_id)}\s*",
                rf"^\s*{re.escape(req_id)}\s*[:\-–]?\s*",
            ]
            for pattern in patterns:
                updated = re.sub(pattern, "", text_str, flags=re.IGNORECASE)
                if updated != text_str:
                    text_str = updated.strip()
                    break
            return text_str

        cleaned["Object Text"] = cleaned.apply(_normalize_row, axis=1)
        return cleaned

    # ------------------------------------------------------------------
    def _normalize_table_captions(self, df: pd.DataFrame) -> pd.DataFrame:
        """Ensure each table has a caption row immediately above.

        - If a caption is below the table, move it above.
        - If missing, insert a blank caption row ("Table :") above the table.
        """
        if df is None or df.empty:
            return df
        work = df.reset_index(drop=True).copy()
        out_rows: list[dict] = []
        i = 0
        n = len(work)
        def _is_caption(text: str) -> bool:
            s = str(text or "")
            return bool(self._TABLE_PREFIX.match(s) or TABLE_NO_ANY_RE.match(s) or re.match(r"^\s*Table\s*:\s*$", s, flags=re.I))

        while i < n:
            row = work.iloc[i].to_dict()
            att = str(row.get("Attachment Type", "") or "").lower()
            if att != "table":
                out_rows.append(row)
                i += 1
                continue

            # Found a table row; determine caption location
            prev_is_caption = False
            next_is_caption = False
            prev_row = out_rows[-1] if out_rows else None
            if prev_row is not None:
                if str(prev_row.get("Object Type", "")).lower() == "text" and _is_caption(prev_row.get("Object Text", "")):
                    prev_is_caption = True
            if i + 1 < n:
                nxt = work.iloc[i + 1].to_dict()
                if str(nxt.get("Object Type", "")).lower() == "text" and _is_caption(nxt.get("Object Text", "")):
                    next_is_caption = True

            if prev_is_caption:
                # Caption already above: keep it, append only the table.
                out_rows.append(row)
                # If there is also a caption below, skip it to avoid duplicates.
                i += 2 if next_is_caption else 1
                continue

            if next_is_caption:
                # Move the below-caption above the table
                out_rows.append(nxt)
                out_rows.append(row)
                i += 2
                continue

            # No caption found: insert a minimal placeholder above
            caption_row = {c: row.get(c, "") for c in work.columns}
            caption_row["Object Type"] = "Text"
            caption_row["Object Text"] = "Table :"
            caption_row["Attachment Type"] = ""
            caption_row["Attachment Data"] = ""
            out_rows.append(caption_row)
            out_rows.append(row)
            i += 1

        return pd.DataFrame(out_rows, columns=work.columns)

    # ------------------------------------------------------------------
    def update_cell(self, row: int, column: str, value: str) -> None:
        if self.dataframe.empty:
            raise RequirementDataError("No data loaded")
        if column not in self.dataframe.columns:
            raise RequirementDataError(f"Unknown column: {column}")
        if row < 0 or row >= len(self.dataframe.index):
            raise RequirementDataError(f"Row {row} is outside of the dataframe range")
        self.dataframe.at[row, column] = value
        # If Object Type or Object Text changed, rebuild numbering → merged into Object Text
        if column in ("Object Type", "Object Text"):
            self.dataframe = self.finalize_dataframe(self.dataframe)

    # ------------------------------------------------------------------
    def to_html_preview(self) -> str:
        if self.dataframe.empty:
            return ""
        parts: List[str] = ["<div>"]

        # skip_mode: None | 'toc' | 'lof' | 'lot'
        skip_mode: Optional[str] = None
        # after we detect a heading for LOF/LOT/TOC, keep skipping lines that look like entries;
        # exit skip when we hit a heading or a line that clearly isn't an entry.
        for _, row in self.dataframe.iterrows():
            obj_type = str(row.get("Object Type", "")).strip()
            obj_type_l = obj_type.lower()
            section = str(row.get(self.section_column_name, "")).strip()
            text = str(row.get("Object Text", "") or "").strip()
            req_id = str(row.get("Requirement ID", "") or "").strip()
            attachment_type = str(row.get("Attachment Type", "") or "").strip().lower()
            attachment_data = str(row.get("Attachment Data", "") or "").strip()

            # 1) Any line that itself says "List of Figures/Tables" or "Table of Contents"
            #    turns skip_mode on (even if it’s not styled as a heading).
            heading_type = self._is_auto_list_heading(text)
            if heading_type:
                skip_mode = heading_type
                # do NOT render the heading itself
                continue

            # 2) A "real" heading (Heading 1/2/3) always ends skip_mode
            if obj_type_l.startswith("heading"):
                skip_mode = None

            # 3) While in skip_mode, skip typical entry lines (regardless of style/hyperlinking).
            if skip_mode:
                is_entry = (
                    (skip_mode == "toc" and self._is_toc_entry_line(text)) or
                    (skip_mode == "lof" and self._is_lof_entry_line(text)) or
                    (skip_mode == "lot" and self._is_lot_entry_line(text))
                )
                # Also swallow blank spacers between entries
                if is_entry or not text:
                    continue
                else:
                    # we encountered non-entry content; exit skip mode and render this line normally
                    skip_mode = None

            # ---- Normal rendering below (unchanged) -------------------------------
            heading_prefix = f"{section} " if section else ""

            if obj_type_l == "heading 1":
                parts.append(f"<h1>{heading_prefix}{html.escape(text)}</h1>")
            elif obj_type_l == "heading 2":
                parts.append(f"<h2>{heading_prefix}{html.escape(text)}</h2>")
            elif obj_type_l == "heading 3":
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
                    mime = "image/png"; data = attachment_data; filename = "image"
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

    # --- NEW: internal util to iterate requirement “blocks” in df order
    def _iter_requirement_blocks(self, df: pd.DataFrame):
        """
        Yield tuples of (req_id, rows_list) where rows_list covers the
        requirement row and the following rows (Text/Table/Image) until
        the next requirement id or a heading.
        """
        if df.empty:
            return
        current_id = None
        bucket = []
        for _, row in df.iterrows():
            obj_type = str(row.get("Object Type", "")).strip().lower()
            rid = str(row.get("Requirement ID", "")).strip()
            is_heading = obj_type.startswith("heading")
            is_req_row = (obj_type == "requirement") and bool(rid)

            if is_req_row or is_heading:
                if current_id and bucket:
                    yield current_id, bucket
                bucket = []
                current_id = rid if is_req_row else None
                if is_req_row:
                    bucket.append(row)
                continue

            if current_id:
                bucket.append(row)

        if current_id and bucket:
            yield current_id, bucket

    def ensure_trace_column(self, direction: str) -> str:
        """
        Ensure that the trace column ('Up Trace' or 'Down Trace') exists in the dataframe.
        """
        col = "Down Trace" if str(direction).strip().lower().startswith("down") else "Up Trace"
        if col not in self.dataframe.columns:
            self.dataframe[col] = ""
            self.dataframe = self._reorder_visible_columns(self.dataframe)
        return col

    # ------------------------------------------------------------------
    def renumber_headings_from(self, start_index: int, base_label: str) -> pd.DataFrame:
        """
        Renumber headings from a given global row index forward, forcing a new
        numeric base (e.g., "2" or "2.1" or "2.1.1"). This rewrites the
        Object Text prefixes for Heading 1..6 rows from that index onward and
        cascades numbering across levels.
        """
        df = self.dataframe.copy()
        if df.empty:
            return df

        # Validate base label
        m = re.match(r"^\s*(\d+(?:\.\d+){0,5})\s*$", str(base_label or ""))
        if not m:
            raise RequirementDataError("Invalid base label. Use forms like '2', '2.1', '2.1.1'.")
        parts = [int(p) for p in m.group(1).split('.')]

        # Helpers
        def level_from_type(ot: str) -> int | None:
            mm = re.search(r"heading\s*(\d+)", str(ot or ""), re.IGNORECASE)
            return int(mm.group(1)) if mm else None

        def strip_leading_number(s: str) -> str:
            s = str(s or '')
            mm = re.match(r"^\s*\d+(?:\s*\.\s*\d+)*\.?\s*(.*)$", s)
            return (mm.group(1) if mm else s).strip()

        # Counters for Heading 1..6
        h = [0, 0, 0, 0, 0, 0]
        assigned_first = False

        for idx in range(max(0, int(start_index)), len(df)):
            obj_type = str(df.at[idx, "Object Type"] if "Object Type" in df.columns else "")
            lvl = level_from_type(obj_type)
            if not lvl or not (1 <= lvl <= 6):
                continue

            txt = str(df.at[idx, "Object Text"] if "Object Text" in df.columns else "")
            clean_txt = strip_leading_number(txt)

            if not assigned_first:
                # Seed counters from base label; if base has fewer parts than current level,
                # fill missing deeper parts with 1s so the first occurrence is valid.
                h = [0, 0, 0, 0, 0, 0]
                for i, num in enumerate(parts[:6]):
                    h[i] = num
                # Ensure length up to current level
                for j in range(len(parts), lvl):
                    h[j] = 1
                label = ".".join(str(h[i]) for i in range(lvl))
                df.at[idx, "Object Text"] = f"{label} {clean_txt}".strip()
                assigned_first = True
                continue

            # Cascade increments depending on the current heading level
            if lvl == 1:
                h[0] += 1
                for j in range(1, 6):
                    h[j] = 0
            else:
                if h[0] == 0:
                    h[0] = 1
                # Ensure parents are at least 1
                for j in range(1, lvl - 1):
                    if h[j] == 0:
                        h[j] = 1
                h[lvl - 1] += 1
                for j in range(lvl, 6):
                    h[j] = 0
            label = ".".join(str(h[i]) for i in range(lvl))
            df.at[idx, "Object Text"] = f"{label} {clean_txt}".strip()

        # Commit and run finalize to rebuild section column and tidy
        prev = getattr(self, "ignore_explicit_heading_numbers", False)
        try:
            # Ensure we honor the explicit labels we just wrote
            self.ignore_explicit_heading_numbers = False
            self.dataframe = self.finalize_dataframe(df)
        finally:
            self.ignore_explicit_heading_numbers = prev
        return self.dataframe
    # ------------------------------------------------------------------
    # --- NEW: build a flattened/exportable dataframe (single row per requirement)
    def build_grouped_export(self) -> pd.DataFrame:
        """
        Returns a DataFrame with 1 row per Requirement ID:
          - Requirement ID
          - Requirement Content (merged into ONE cell)
          - Referenced Tables
          - Referenced Figures
          - Image Path(s) (comma separated temp files written from base64)
        """
        df = self.dataframe.copy()
        if df.empty:
            return pd.DataFrame(columns=[
                "Requirement ID", "Requirement Content",
                "Referenced Tables", "Referenced Figures", "Image Path(s)"
            ])

        out_rows = []
        for req_id, rows in self._iter_requirement_blocks(df):
            content_chunks = []
            tables = []
            figures = []
            image_paths = []

            for r in rows:
                att_type = str(r.get("Attachment Type", "")).strip().lower()
                obj_text = str(r.get("Object Text", "") or "").strip()
                if att_type == "table" and obj_text:
                    tables.append(obj_text)
                elif att_type == "image":
                    payload = str(r.get("Attachment Data", "") or "").strip()
                    if payload:
                        try:
                            data = json.loads(payload)
                            raw_b64 = data.get("data") or ""
                            if raw_b64:
                                import tempfile, os
                                img_bytes = base64.b64decode(raw_b64)
                                suffix = ".png"
                                fname = (data.get("filename") or "image").strip()
                                if "." in fname:
                                    suffix = "." + fname.split(".")[-1]
                                tmp = tempfile.NamedTemporaryFile(prefix="reqimg_", suffix=suffix, delete=False)
                                tmp.write(img_bytes)
                                tmp.flush(); tmp.close()
                                image_paths.append(tmp.name)
                        except Exception:
                            pass
                    if obj_text:
                        figures.append(obj_text)
                else:
                    if obj_text:
                        content_chunks.append(obj_text)

            merged_content = " ".join(ch for ch in content_chunks if ch).strip()
            ref_tables = ", ".join(sorted(set(tables))) if tables else ""
            ref_figs = ", ".join(sorted(set(figures))) if figures else ""
            img_str = ", ".join(image_paths) if image_paths else ""

            out_rows.append({
                "Requirement ID": req_id,
                "Requirement Content": merged_content,
                "Referenced Tables": ref_tables,
                "Referenced Figures": ref_figs,
                "Image Path(s)": img_str
            })

        # Ensure orphan requirement rows still appear
        seen = {r["Requirement ID"] for r in out_rows}
        for _, row in df.iterrows():
            rid = str(row.get("Requirement ID", "")).strip()
            if rid and rid not in seen and str(row.get("Object Type", "")).strip().lower() == "requirement":
                out_rows.append({
                    "Requirement ID": rid,
                    "Requirement Content": str(row.get("Object Text", "") or "").strip(),
                    "Referenced Tables": "",
                    "Referenced Figures": "",
                    "Image Path(s)": ""
                })

        return pd.DataFrame(out_rows, columns=[
            "Requirement ID", "Requirement Content",
            "Referenced Tables", "Referenced Figures", "Image Path(s)"
        ])


# ------------------------------------------------------------
# Cross-artifact trace building (for Trace Matrix tab selection)
# ------------------------------------------------------------
from typing import Dict, Tuple, Set

_SPLIT_TRACE_TOKENS = re.compile(r"[,\n;/]+")

def _extract_id_token(cell: str) -> str:
    """Return the likely ID token from a mixed 'ID - description' cell."""
    if cell is None:
        return ""
    s = str(cell).strip()
    if not s:
        return ""
    # Split on dash variants (– — -) with spaces around
    parts = re.split(r"\s+[–—-]\s+", s, maxsplit=1)
    return parts[0].strip()

def _parse_trace_list(cell) -> list[str]:
    """Split a cell into distinct ID-like tokens, preserving order."""
    if cell is None:
        return []
    if not isinstance(cell, str):
        cell = str(cell) if pd.notna(cell) else ""
    tokens = []
    for raw in _SPLIT_TRACE_TOKENS.split(cell):
        t = _extract_id_token(raw.strip())
        if t:
            tokens.append(t)
    # De-dup while keeping first occurrence
    seen = set()
    out = []
    for t in tokens:
        if t not in seen:
            seen.add(t); out.append(t)
    return out

def _normalize_for_trace(df: pd.DataFrame) -> pd.DataFrame:
    """Map incoming DataFrame to a stable schema used by the matrix builder."""
    if df is None or df.empty:
        return pd.DataFrame(columns=["Requirement ID","Object Text","Up Trace","Down Trace","Linked ID / Description","SourceFile","SourceType"])
    colmap = {c.lower(): c for c in df.columns}
    def get(*names):
        for n in names:
            if n.lower() in colmap:
                return colmap[n.lower()]
        return None
    rid = get("Requirement ID","Req ID","ReqID","ID")
    up  = get("Up Trace","UpTrace")
    dn  = get("Down Trace","DownTrace")
    bi  = get("Linked ID / Description","Bi Trace","Bi-Directional Trace","BiTrace","Bidirectional Trace")
    src = get("SourceFile")
    st  = get("SourceType")
    out = pd.DataFrame()
    if rid: out["Requirement ID"] = df[rid].astype(str).str.strip()
    else:   out["Requirement ID"] = ""
    out["Object Text"] = df[get("Object Text","Title","Description")] if get("Object Text","Title","Description") else ""
    out["Up Trace"] = df[up] if up else ""
    out["Down Trace"] = df[dn] if dn else ""
    out["Linked ID / Description"] = df[bi] if bi else ""
    out["SourceFile"] = df[src] if src else ""
    out["SourceType"] = df[st] if st else ""
    return out

def build_cross_pairs(self, tab_frames: Dict[str, pd.DataFrame]) -> Tuple[list[str], list[tuple], list[str]]:
    """
    Build cross-artifact trace pairs between the selected tabs.

    Parameters
    ----------
    tab_frames : Dict[str, DataFrame]
        Mapping of 'tab label' -> DataFrame subset for that tab (raw, not reset index).

    Returns
    -------
    columns : list[str]
        ["From Artifact","From ID","↔","To Artifact","To ID"]
    rows : list[tuple]
        List of 5-tuples corresponding to the columns.
    row_sources : list[str]
        Preferred source tab label for navigation (usually From Artifact).
    """
    # 1) Normalize frames and index IDs per tab
    norm: Dict[str, pd.DataFrame] = {k: _normalize_for_trace(v) for k,v in tab_frames.items()}
    id_index: Dict[str, Set[str]] = {k: set(norm[k]["Requirement ID"].dropna().astype(str).str.strip()) for k in norm}

    # 2) Build directed edges, then collapse to undirected marks
    directed_edges: Set[tuple] = set()  # (src_tab, src_id, dst_tab, dst_id, "→")
    for tab_name, ndf in norm.items():
        if ndf.empty: continue
        for _, row in ndf.iterrows():
            rid = str(row.get("Requirement ID","")).strip()
            if not rid: continue

            # Gather link lists
            bi_list = _parse_trace_list(row.get("Linked ID / Description",""))
            up_list = _parse_trace_list(row.get("Up Trace",""))
            dn_list = _parse_trace_list(row.get("Down Trace",""))

            # Helper: add links to any *other* selected tab where ID exists
            def add_links(ids: list[str], mark: str):
                for tid in ids:
                    for other_tab, other_ids in id_index.items():
                        if other_tab == tab_name:
                            continue
                        if tid in other_ids:
                            if mark == "↔":
                                directed_edges.add((tab_name, rid, other_tab, tid, "→"))
                                directed_edges.add((other_tab, tid, tab_name, rid, "→"))
                            elif mark == "→":
                                directed_edges.add((tab_name, rid, other_tab, tid, "→"))
                            elif mark == "←":
                                directed_edges.add((other_tab, tid, tab_name, rid, "→"))
            add_links(bi_list, "↔")
            add_links(up_list, "→")
            add_links(dn_list, "←")

    # 3) Collapse to undirected symbols
    pairs: Set[tuple] = set()  # (A, aid, B, bid, mark)
    edge_set = set((a,aid,b,bid) for (a,aid,b,bid,_) in directed_edges)
    for (a,aid,b,bid,_) in directed_edges:
        reverse = (b,bid,a,aid) in edge_set
        mark = "↔" if reverse else "→"
        pairs.add((a,aid,b,bid,mark))

    # 4) Emit rows
    rows = sorted(list(pairs), key=lambda t: (t[0].lower(), t[1].lower(), t[2].lower(), t[3].lower()))
    columns = ["From Artifact","From ID","↔","To Artifact","To ID"]
    # navigation prefers From Artifact tab
    row_sources = [r[0] for r in rows]
    return columns, [(a, aid, m, b, bid) for a,aid,b,bid,m in rows], row_sources
TABLE_NO_ANY_RE = re.compile(r"^\s*Table\s*(?:No\.?|Number)?\s*[:\-]?\s*(\d+)\b", re.IGNORECASE)
