#!/usr/bin/env python3
"""
Core LaTeX-to-Word Converter Engine
====================================
Converts Markdown with LaTeX equations ($..$ inline, $$...$$ display) into
professional Microsoft Word (.docx) documents with native, editable OMML
equations and optional automatic equation numbering.

Pipeline: LaTeX -> MathML (latex2mathml) -> OMML (MML2OMML.XSL via lxml XSLT)
"""

import sys
import os
import re
from typing import List, Optional

from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from lxml import etree
from latex2mathml.converter import convert as latex2mathml_convert

from core.math_tokens import (
    split_inline_math, split_table_row, is_fence_line, fence_marker,
)


# ---------------------------------------------------------------------------
# XSL resolution
# ---------------------------------------------------------------------------
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_PROJECT_DIR = os.path.dirname(_SCRIPT_DIR)

# Explicit override first (useful for CI boxes and non-standard installs).
_XSL_ENV_VAR = "LWX_MML2OMML_XSL"

_XSL_CANDIDATES = [
    os.path.join(_SCRIPT_DIR, "MML2OMML.XSL"),
    os.path.join(_PROJECT_DIR, "MML2OMML.XSL"),
    # System Office installation
    r"C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL",
    r"C:\Program Files (x86)\Microsoft Office\root\Office16\MML2OMML.XSL",
]


def _resolve_xsl() -> Optional[str]:
    """Find the first available MML2OMML.XSL on the system."""
    env_path = os.environ.get(_XSL_ENV_VAR, "").strip()
    if env_path and os.path.isfile(env_path):
        return env_path
    for path in _XSL_CANDIDATES:
        if os.path.isfile(path):
            return path
    return None


# ---------------------------------------------------------------------------
# Converter class
# ---------------------------------------------------------------------------
class LatexToWordConverter:
    """Converts Markdown+LaTeX input into a .docx with native OMML equations.

    Usage::

        conv = LatexToWordConverter()
        stats = conv.convert_markdown_to_docx(md_text, "output.docx")
        print(stats)
    """

    def __init__(self):
        self.equation_counter: int = 0
        self._stats = self._fresh_stats()
        self._cur_line: int = 0

        self.xsl_path = _resolve_xsl()
        self.transform: Optional[etree.XSLT] = None

        if self.xsl_path:
            try:
                xslt_tree = etree.parse(self.xsl_path)
                self.transform = etree.XSLT(xslt_tree)
            except Exception as exc:
                sys.stderr.write(
                    f"Warning: Could not load MML2OMML.XSL ({self.xsl_path}): {exc}\n")

    @staticmethod
    def _fresh_stats() -> dict:
        return {
            "equations_total": 0,
            "equations_converted": 0,
            "equations_failed": 0,
            "failed_formulas": [],   # [{"line": int, "latex": str}, ...]
            "output_path": "",
        }

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------
    def reset(self):
        """Reset counters for a fresh document."""
        self.equation_counter = 0
        self._stats = self._fresh_stats()
        self._cur_line = 0

    def convert_markdown_to_docx(
        self,
        markdown_content: str,
        output_path: str,
        number_equations: bool = True,
        eq_format: str = "({n})",
    ) -> dict:
        """Convert *markdown_content* to a Word document at *output_path*.

        Returns a stats dict with conversion counts.
        """
        self.reset()
        doc = Document()
        self.setup_document_styles(doc)

        lines = markdown_content.splitlines()
        i = 0
        in_display_math = False
        display_close = ""          # "$$" or "\\]" depending on opener
        display_math_buf: List[str] = []
        in_table = False
        table_buf: List[str] = []
        in_code_fence = False
        code_fence_marker = ""
        code_buf: List[str] = []

        while i < len(lines):
            line = lines[i]
            stripped = line.strip()
            self._cur_line = i + 1

            # ── fenced code blocks (``` / ~~~) ─────────────────────
            if is_fence_line(line):
                marker = fence_marker(line)[0] * 3
                if not in_code_fence:
                    in_code_fence = True
                    code_fence_marker = marker
                    code_buf = []
                elif marker[0] == code_fence_marker[0]:
                    in_code_fence = False
                    self._add_code_block(doc, code_buf)
                    code_buf = []
                i += 1; continue
            if in_code_fence:
                code_buf.append(line)
                i += 1; continue

            # ── inside display math ($$…$$ or \[…\]) ──────────────
            if in_display_math:
                if stripped.endswith(display_close):
                    in_display_math = False
                    display_math_buf.append(stripped[:-len(display_close)].strip())
                    formula = " ".join(b for b in display_math_buf if b)
                    self._add_display_equation(
                        doc, formula, number_equations, eq_format)
                    display_math_buf = []
                else:
                    display_math_buf.append(stripped)
                i += 1; continue

            # ── display math opener: $$...$$ ──────────────────────
            if stripped.startswith("$$"):
                rest = stripped[2:].strip()
                if rest.endswith("$$") and len(rest) > 2:
                    # single-line $$...$$
                    self._add_display_equation(
                        doc, rest[:-2].strip(), number_equations, eq_format)
                elif not rest.endswith("$$"):
                    # multi-line: accumulate until closing $$
                    in_display_math = True
                    display_close = "$$"
                    display_math_buf = [rest] if rest else []
                # else: bare "$$  $$" empty pair — skip
                i += 1; continue

            # ── display math opener: \[ ... \] ────────────────────
            if stripped.startswith("\\["):
                rest = stripped[2:].strip()
                if rest.endswith("\\]") and len(rest) >= 2:
                    formula = rest[:-2].strip()
                    if formula:
                        self._add_display_equation(
                            doc, formula, number_equations, eq_format)
                else:
                    in_display_math = True
                    display_close = "\\]"
                    display_math_buf = [rest] if rest else []
                i += 1; continue

            # ── markdown table ────────────────────────────────────
            if stripped.startswith("|") and stripped.endswith("|"):
                in_table = True
                table_buf.append(stripped)
                i += 1; continue
            elif in_table:
                in_table = False
                self._add_markdown_table(doc, table_buf)
                table_buf = []
                # fall through to evaluate current line

            # ── blank ─────────────────────────────────────────────
            if not stripped:
                i += 1; continue

            # ── horizontal rule ───────────────────────────────────
            if re.match(r'^-{3,}$', stripped):
                i += 1; continue

            # ── headings ──────────────────────────────────────────
            if stripped.startswith("#### "):
                self._add_heading(doc, stripped[5:].strip(), level=4)
            elif stripped.startswith("### "):
                self._add_heading(doc, stripped[4:].strip(), level=3)
            elif stripped.startswith("## "):
                self._add_heading(doc, stripped[3:].strip(), level=2)
            elif stripped.startswith("# "):
                self._add_heading(doc, stripped[2:].strip(), level=1)

            # ── images ────────────────────────────────────────────
            elif stripped.startswith("![") and "](" in stripped and stripped.endswith(")"):
                m = re.match(r"!\[(.*?)\]\((.*?)\)", stripped)
                if m:
                    self._add_image(doc, m.group(2), m.group(1), output_path)

            # ── blockquotes ───────────────────────────────────────
            elif stripped.startswith("> "):
                self._add_blockquote(doc, stripped[2:].strip())

            # ── bullet list items (-, *, +) ───────────────────────
            elif re.match(r'^[-*+] ', stripped):
                self._add_list_item(doc, stripped[2:].strip(), numbered=False)

            # ── numbered list items (1. / 1) ) ────────────────────
            elif re.match(r'^\d+[.)] ', stripped):
                self._add_list_item(
                    doc, re.sub(r'^\d+[.)] ', '', stripped), numbered=True)

            # ── normal paragraph ──────────────────────────────────
            else:
                self.add_formatted_paragraph(doc, stripped)

            i += 1

        # flush remaining table
        if in_table and table_buf:
            self._add_markdown_table(doc, table_buf)

        doc.save(output_path)
        self._stats["output_path"] = os.path.abspath(output_path)
        return dict(self._stats)

    # ------------------------------------------------------------------
    # LaTeX → OMML conversion
    # ------------------------------------------------------------------
    def latex_to_omml(self, latex_code: str, is_display: bool = True) -> Optional[OxmlElement]:
        """Convert a LaTeX string to a native Word OMML XML element.

        Two-pass strategy: convert the minimally-cleaned formula first so
        fidelity-preserving constructs (``\\mathbb``, ``\\text``) survive;
        if that fails, retry with the legacy compatibility rewrites.
        """
        if not self.transform:
            return None
        clean = self._clean_latex(latex_code)
        if not clean:
            return None

        attempts = [clean]
        legacy = self._legacy_compat_latex(clean)
        if legacy != clean:
            attempts.append(legacy)

        last_exc: Optional[Exception] = None
        for attempt in attempts:
            try:
                mathml_str = latex2mathml_convert(attempt)
                mathml_tree = etree.fromstring(mathml_str.encode("utf-8"))
                omml_tree = self.transform(mathml_tree)
                omml_root = omml_tree.getroot()

                xml_bytes = etree.tostring(omml_root, encoding="unicode")

                if is_display and "oMathPara" not in xml_bytes:
                    xml_bytes = (
                        '<m:oMathPara xmlns:m='
                        '"http://schemas.openxmlformats.org/officeDocument/2006/math">'
                        f'{xml_bytes}</m:oMathPara>'
                    )

                return parse_xml(xml_bytes)
            except Exception as exc:
                last_exc = exc

        preview = latex_code[:40].replace("\n", " ")
        sys.stderr.write(
            f"OMML conversion failed for '{preview}…': {last_exc}\n")
        return None

    # ------------------------------------------------------------------
    # Document styles
    # ------------------------------------------------------------------
    def setup_document_styles(self, doc: Document):
        """Academic formatting: 1″ margins, TNR 12 pt, 1.15 spacing."""
        for section in doc.sections:
            section.top_margin = Inches(1.0)
            section.bottom_margin = Inches(1.0)
            section.left_margin = Inches(1.0)
            section.right_margin = Inches(1.0)

        style = doc.styles["Normal"]
        style.font.name = "Times New Roman"
        style.font.size = Pt(12)
        style.font.color.rgb = RGBColor(33, 33, 33)
        pf = style.paragraph_format
        pf.line_spacing = 1.15
        pf.space_after = Pt(6)
        pf.space_before = Pt(0)

    # ------------------------------------------------------------------
    # Inline paragraph with math
    # ------------------------------------------------------------------
    def add_formatted_paragraph(self, doc: Document, text: str):
        """Add paragraph with mixed text and inline $…$ / \\(…\\) OMML math."""
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_after = Pt(6)
        self._emit_inline_tokens(p, text)

    def _emit_inline_tokens(self, paragraph, text: str):
        """Append text/math/code runs to *paragraph* via the shared tokenizer."""
        for token in split_inline_math(text):
            if token.kind == "math":
                self._stats["equations_total"] += 1
                omml = self.latex_to_omml(token.value, is_display=False)
                if omml is not None:
                    paragraph._p.append(omml)
                    self._stats["equations_converted"] += 1
                else:
                    run = paragraph.add_run(token.value)
                    run.italic = True
                    run.font.name = "Cambria Math"
                    self._stats["equations_failed"] += 1
                    self._record_failure(token.value)
            elif token.kind == "code":
                run = paragraph.add_run(token.value.strip("`"))
                run.font.name = "Consolas"
                run.font.size = Pt(10.5)
            else:
                self._add_rich_text_runs(paragraph, token.value)

    def _record_failure(self, latex_code: str):
        self._stats["failed_formulas"].append(
            {"line": self._cur_line, "latex": latex_code[:80]})

    # ------------------------------------------------------------------
    # Display equations with auto-numbering
    # ------------------------------------------------------------------
    def _add_display_equation(
        self, doc: Document, formula: str,
        number_equations: bool = True, eq_format: str = "({n})"
    ):
        """Add a display equation, optionally with a right-aligned number."""
        self._stats["equations_total"] += 1
        omml_elem = self.latex_to_omml(formula, is_display=True)

        if number_equations:
            self.equation_counter += 1
            eq_label = eq_format.replace("{n}", str(self.equation_counter))
            self._add_numbered_equation(doc, omml_elem, formula, eq_label)
        else:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(6)
            p.paragraph_format.space_after = Pt(6)
            if omml_elem is not None:
                p._p.append(omml_elem)
                self._stats["equations_converted"] += 1
            else:
                self._fallback_equation_text(p, formula)
                self._stats["equations_failed"] += 1
                self._record_failure(formula)

    def _add_numbered_equation(
        self, doc: Document, omml_elem, formula: str, eq_label: str
    ):
        """Insert equation inside a 1×3 invisible table for numbering."""
        tbl = doc.add_table(rows=1, cols=3)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl.autofit = False

        # page width minus margins ≈ 6.5 inches
        page_w = Inches(6.5)
        tbl.columns[0].width = Inches(0.6)
        tbl.columns[1].width = Inches(4.8)
        tbl.columns[2].width = Inches(1.1)

        # remove all borders
        self._remove_table_borders(tbl)

        # ── centre cell: equation ──
        eq_cell = tbl.cell(0, 1)
        eq_p = eq_cell.paragraphs[0]
        eq_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        eq_p.paragraph_format.space_before = Pt(4)
        eq_p.paragraph_format.space_after = Pt(4)
        if omml_elem is not None:
            # For display math inside a table cell, unwrap oMathPara → oMath
            inner = self._unwrap_omathpara(omml_elem)
            try:
                eq_p._p.append(inner)
            except Exception:
                eq_p._p.append(omml_elem)
            self._stats["equations_converted"] += 1
        else:
            self._fallback_equation_text(eq_p, formula)
            self._stats["equations_failed"] += 1
            self._record_failure(formula)

        # ── right cell: equation number ──
        num_cell = tbl.cell(0, 2)
        num_p = num_cell.paragraphs[0]
        num_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        num_p.paragraph_format.space_before = Pt(4)
        num_p.paragraph_format.space_after = Pt(4)
        run = num_p.add_run(eq_label)
        run.font.name = "Cambria Math"
        run.font.size = Pt(12)

        # left cell stays empty
        left_p = tbl.cell(0, 0).paragraphs[0]
        left_p.paragraph_format.space_before = Pt(4)
        left_p.paragraph_format.space_after = Pt(4)

    @staticmethod
    def _unwrap_omathpara(omml_elem):
        """Return the inner m:oMath of an m:oMathPara (table-cell safe).

        Unlike a regex over serialized XML, extracting the child element
        keeps all namespace declarations intact when the child is
        re-serialized or appended elsewhere.
        """
        try:
            if etree.QName(omml_elem).localname == "oMathPara":
                for child in omml_elem:
                    if etree.QName(child).localname == "oMath":
                        return child
        except Exception:
            pass
        return omml_elem

    # ------------------------------------------------------------------
    # Markdown table → Word table
    # ------------------------------------------------------------------
    def _add_markdown_table(self, doc: Document, table_lines: List[str]):
        """Convert markdown pipe-table lines to a Word table."""
        rows_data = []
        for line in table_lines:
            line = line.strip()
            if not line:
                continue
            if re.match(r'^\|?[\s\:\-]+(?:\|[\s\:\-]+)*\|?$', line):
                continue
            # math-aware split: $|x|$ keeps its pipes, \| stays literal
            cells = [c.strip() for c in split_table_row(line)]
            rows_data.append(cells)

        if not rows_data:
            return

        num_cols = max(len(r) for r in rows_data)
        tbl = doc.add_table(rows=len(rows_data), cols=num_cols)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl.autofit = True

        for r_idx, row in enumerate(rows_data):
            for c_idx, val in enumerate(row):
                if c_idx >= num_cols:
                    continue
                cell = tbl.cell(r_idx, c_idx)
                cell_p = cell.paragraphs[0]
                cell_p.text = ""
                cell_p.paragraph_format.space_after = Pt(2)
                cell_p.paragraph_format.space_before = Pt(2)

                for token in split_inline_math(val):
                    if token.kind == "math":
                        self._stats["equations_total"] += 1
                        omml = self.latex_to_omml(token.value, is_display=False)
                        if omml is not None:
                            cell_p._p.append(omml)
                            self._stats["equations_converted"] += 1
                        else:
                            r = cell_p.add_run(token.value)
                            r.italic = True
                            self._stats["equations_failed"] += 1
                            self._record_failure(token.value)
                    elif token.kind == "code":
                        r = cell_p.add_run(token.value.strip("`"))
                        r.font.name = "Consolas"
                        r.font.size = Pt(9.5)
                        if r_idx == 0:
                            r.font.bold = True
                    else:
                        # text: \| is a literal pipe in table cells
                        r = cell_p.add_run(token.value.replace("\\|", "|"))
                        if r_idx == 0:
                            r.font.bold = True

                # header shading
                if r_idx == 0:
                    shading = parse_xml(
                        f'<w:shd {nsdecls("w")} w:fill="F2F2F2"/>')
                    cell._tc.get_or_add_tcPr().append(shading)
                    for run in cell_p.runs:
                        run.font.size = Pt(10.5)
                else:
                    for run in cell_p.runs:
                        run.font.size = Pt(10)

        # spacer
        sp = doc.add_paragraph()
        sp.paragraph_format.space_before = Pt(0)
        sp.paragraph_format.space_after = Pt(4)

    # ------------------------------------------------------------------
    # Heading helpers
    # ------------------------------------------------------------------
    _HEADING_CFG = {
        1: {"size": 16, "bold": True, "italic": False, "space_before": 14, "space_after": 6, "rgb": (0, 0, 0)},
        2: {"size": 14, "bold": True, "italic": False, "space_before": 12, "space_after": 4, "rgb": (30, 30, 30)},
        3: {"size": 12, "bold": True, "italic": True,  "space_before": 8,  "space_after": 2, "rgb": (50, 50, 50)},
        4: {"size": 12, "bold": True, "italic": False, "space_before": 6,  "space_after": 2, "rgb": (60, 60, 60)},
    }

    def _add_heading(self, doc: Document, text: str, level: int):
        cfg = self._HEADING_CFG.get(level, self._HEADING_CFG[4])
        h = doc.add_heading(text, level=level)
        h.paragraph_format.space_before = Pt(cfg["space_before"])
        h.paragraph_format.space_after = Pt(cfg["space_after"])
        for r in h.runs:
            r.font.name = "Times New Roman"
            r.font.size = Pt(cfg["size"])
            r.font.bold = cfg["bold"]
            r.font.italic = cfg["italic"]
            r.font.color.rgb = RGBColor(*cfg["rgb"])

    # ------------------------------------------------------------------
    # Minor element helpers
    # ------------------------------------------------------------------
    def _add_image(self, doc, img_path, caption, output_path):
        if not os.path.isabs(img_path):
            base = os.path.dirname(os.path.abspath(output_path))
            resolved = os.path.normpath(os.path.join(base, img_path))
            if not os.path.exists(resolved):
                resolved = os.path.abspath(img_path)
        else:
            resolved = img_path
        if os.path.exists(resolved):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.add_run().add_picture(resolved, width=Inches(5.5))
            if caption:
                cp = doc.add_paragraph()
                cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r = cp.add_run(caption)
                r.font.size = Pt(10); r.font.italic = True

    def _add_blockquote(self, doc, text):
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.5)
        p.paragraph_format.right_indent = Inches(0.5)
        self._emit_inline_tokens(p, text)
        for r in p.runs:
            r.italic = True
            if r.font.size is None:
                r.font.size = Pt(11)

    def _add_list_item(self, doc, text, numbered=False):
        style = "List Number" if numbered else "List Bullet"
        p = doc.add_paragraph(style=style)
        # clear auto-generated run and use formatted paragraph logic
        p.text = ""
        self._emit_inline_tokens(p, text)

    def _add_code_block(self, doc, code_lines):
        """Render a fenced code block as monospaced, lightly-shaded paragraphs."""
        for raw in code_lines:
            p = doc.add_paragraph()
            p.paragraph_format.line_spacing = 1.0
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.left_indent = Inches(0.25)
            run = p.add_run(raw if raw else " ")
            run.font.name = "Consolas"
            run.font.size = Pt(9.5)
            shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="F5F5F5"/>')
            p._p.get_or_add_pPr().append(shading)
        # small spacer after the block
        sp = doc.add_paragraph()
        sp.paragraph_format.space_after = Pt(4)

    # ------------------------------------------------------------------
    # Text formatting helpers
    # ------------------------------------------------------------------
    @staticmethod
    def _add_rich_text_runs(paragraph, text: str):
        """Parse **bold** and *italic* markdown within plain text."""
        # bold
        parts = re.split(r'(\*\*[^\*]+?\*\*)', text)
        for part in parts:
            if part.startswith("**") and part.endswith("**") and len(part) > 4:
                r = paragraph.add_run(part[2:-2])
                r.bold = True
            else:
                # italic
                subs = re.split(r'(\*[^\*\n]+?\*)', part)
                for s in subs:
                    if s.startswith("*") and s.endswith("*") and len(s) > 2:
                        r = paragraph.add_run(s[1:-1])
                        r.italic = True
                    else:
                        if s:
                            paragraph.add_run(s)

    @staticmethod
    def _fallback_equation_text(paragraph, formula: str):
        """Plain-text fallback when OMML conversion fails."""
        clean = formula.strip().strip("$").strip()
        run = paragraph.add_run(clean)
        run.italic = True
        run.font.name = "Cambria Math"

    # ------------------------------------------------------------------
    # LaTeX cleaning
    # ------------------------------------------------------------------
    @staticmethod
    def _clean_latex(latex_code: str) -> str:
        """Strip delimiters and apply only *safe* normalisations.

        Deliberately conservative: modern latex2mathml handles ``\\mathbb``,
        ``\\text`` and ``\\operatorname`` natively, so we must NOT rewrite
        them here (rewriting ``\\mathbb{R}`` to ``\\mathbf{R}`` loses the
        blackboard-bold glyph). Word-boundary regexes keep ``\\middle`` safe
        from the ``\\mid`` → ``|`` normalisation.
        """
        s = latex_code.strip()
        # strip delimiters
        if s.startswith("$$") and s.endswith("$$"):
            s = s[2:-2].strip()
        elif s.startswith("$") and s.endswith("$"):
            s = s[1:-1].strip()

        # \mid → | but never \middle (word-boundary aware)
        s = re.sub(r'\\mid(?![a-zA-Z])', '|', s)
        return s

    @staticmethod
    def _legacy_compat_latex(latex_code: str) -> str:
        """Second-chance compatibility pass for older latex2mathml builds.

        Only applied when the native conversion of ``_clean_latex`` output
        fails. Keeps documents convertible at the cost of some fidelity
        (e.g. \\mathbb is demoted to \\mathbf).
        """
        s = latex_code
        s = re.sub(r'\\mathbb(?![a-zA-Z])', r'\\mathbf', s)
        s = re.sub(r'\\text\{([^}]*)\}', r'\\mathrm{\1}', s)
        known_ops = ["sin", "cos", "tan", "log", "ln", "exp", "det",
                     "max", "min", "sup", "inf", "lim", "arg", "deg"]
        for op in known_ops:
            s = s.replace(f"\\operatorname{{{op}}}", f"\\{op}")
        s = re.sub(r'\\operatorname\{([^}]*)\}', r'\\mathrm{\1}', s)
        s = s.replace(r"\,", " ")
        s = s.replace(r"\;", " ")
        s = s.replace(r"\!", "")
        s = s.replace(r"\quad", "  ")
        s = s.replace(r"\qquad", "    ")
        return s

    # ------------------------------------------------------------------
    # Table border removal
    # ------------------------------------------------------------------
    @staticmethod
    def _remove_table_borders(tbl):
        """Make all table borders invisible."""
        tbl_pr = tbl._tbl.tblPr if tbl._tbl.tblPr is not None else OxmlElement('w:tblPr')
        borders = OxmlElement('w:tblBorders')
        for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            el = OxmlElement(f'w:{edge}')
            el.set(qn('w:val'), 'none')
            el.set(qn('w:sz'), '0')
            el.set(qn('w:space'), '0')
            el.set(qn('w:color'), 'auto')
            borders.append(el)
        tbl_pr.append(borders)


# ---------------------------------------------------------------------------
# CLI interface
# ---------------------------------------------------------------------------
def main():
    """Simple CLI: python -m core.converter input.md output.docx [--no-number]"""
    import argparse
    parser = argparse.ArgumentParser(
        description="Convert Markdown+LaTeX to Word with native OMML equations")
    parser.add_argument("input", help="Input .md or .tex file")
    parser.add_argument("output", help="Output .docx file")
    parser.add_argument("--no-number", action="store_true",
                        help="Disable automatic equation numbering")
    parser.add_argument("--eq-format", default="({n})",
                        help="Equation number format (default: '({n})')")
    args = parser.parse_args()

    if not os.path.isfile(args.input):
        print(f"Error: Input file not found: {args.input}")
        sys.exit(1)

    with open(args.input, "r", encoding="utf-8") as f:
        content = f.read()

    conv = LatexToWordConverter()
    stats = conv.convert_markdown_to_docx(
        content, args.output,
        number_equations=not args.no_number,
        eq_format=args.eq_format,
    )
    print(f"Done! {stats['equations_converted']}/{stats['equations_total']} "
          f"equations converted ({stats['equations_failed']} failed)")
    print(f"Output: {stats['output_path']}")


if __name__ == "__main__":
    main()
