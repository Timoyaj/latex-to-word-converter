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


# ---------------------------------------------------------------------------
# XSL resolution
# ---------------------------------------------------------------------------
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_PROJECT_DIR = os.path.dirname(_SCRIPT_DIR)

_XSL_CANDIDATES = [
    os.path.join(_SCRIPT_DIR, "MML2OMML.XSL"),
    os.path.join(_PROJECT_DIR, "MML2OMML.XSL"),
    # Bundled with the academic-doc-extractor-docx skill
    os.path.normpath(os.path.join(
        os.path.expanduser("~"),
        ".gemini", "config", "skills",
        "academic-doc-extractor-docx", "scripts", "MML2OMML.XSL",
    )),
    # System Office installation
    r"C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL",
    r"C:\Program Files (x86)\Microsoft Office\root\Office16\MML2OMML.XSL",
]


def _resolve_xsl() -> Optional[str]:
    """Find the first available MML2OMML.XSL on the system."""
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
        self._stats = {"equations_total": 0, "equations_converted": 0,
                        "equations_failed": 0, "output_path": ""}

        self.xsl_path = _resolve_xsl()
        self.transform: Optional[etree.XSLT] = None

        if self.xsl_path:
            try:
                xslt_tree = etree.parse(self.xsl_path)
                self.transform = etree.XSLT(xslt_tree)
            except Exception as exc:
                sys.stderr.write(
                    f"Warning: Could not load MML2OMML.XSL ({self.xsl_path}): {exc}\n")

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------
    def reset(self):
        """Reset counters for a fresh document."""
        self.equation_counter = 0
        self._stats = {"equations_total": 0, "equations_converted": 0,
                        "equations_failed": 0, "output_path": ""}

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
        display_math_buf: List[str] = []
        in_table = False
        table_buf: List[str] = []

        while i < len(lines):
            line = lines[i]
            stripped = line.strip()

            # ── display math $$...$$ ──────────────────────────────
            if stripped.startswith("$$"):
                if not in_display_math:
                    rest = stripped[2:].strip()
                    if rest.endswith("$$") and len(rest) > 2:
                        # single-line $$...$$
                        self._add_display_equation(
                            doc, rest[:-2].strip(), number_equations, eq_format)
                        i += 1; continue
                    elif rest.endswith("$$"):
                        # $$  $$ on same line (empty)
                        i += 1; continue
                    else:
                        in_display_math = True
                        display_math_buf = [rest] if rest else []
                        i += 1; continue
                else:
                    # closing $$
                    in_display_math = False
                    before = stripped
                    if before.endswith("$$"):
                        before = before[:-2]
                    if before.startswith("$$"):
                        before = before[2:]
                    display_math_buf.append(before.strip())
                    formula = " ".join(b for b in display_math_buf if b)
                    self._add_display_equation(
                        doc, formula, number_equations, eq_format)
                    display_math_buf = []
                    i += 1; continue

            if in_display_math:
                if stripped.endswith("$$"):
                    in_display_math = False
                    display_math_buf.append(stripped[:-2].strip())
                    formula = " ".join(b for b in display_math_buf if b)
                    self._add_display_equation(
                        doc, formula, number_equations, eq_format)
                    display_math_buf = []
                else:
                    display_math_buf.append(stripped)
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

            # ── bullet / numbered list items ──────────────────────
            elif stripped.startswith("- ") or stripped.startswith("* "):
                self._add_list_item(doc, stripped[2:].strip())

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
        """Convert a LaTeX string to a native Word OMML XML element."""
        if not self.transform:
            return None
        try:
            clean = self._clean_latex(latex_code)
            if not clean:
                return None

            mathml_str = latex2mathml_convert(clean)
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
            preview = latex_code[:40].replace("\n", " ")
            sys.stderr.write(f"OMML conversion failed for '{preview}…': {exc}\n")
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
        """Add paragraph with mixed text and inline $…$ OMML math."""
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_after = Pt(6)

        pattern = re.compile(r'(\$[^\$\n]+?\$)')
        tokens = pattern.split(text)

        for token in tokens:
            if not token:
                continue
            if token.startswith("$") and token.endswith("$") and len(token) > 2:
                self._stats["equations_total"] += 1
                omml = self.latex_to_omml(token, is_display=False)
                if omml is not None:
                    p._p.append(omml)
                    self._stats["equations_converted"] += 1
                else:
                    run = p.add_run(token[1:-1])
                    run.italic = True
                    run.font.name = "Cambria Math"
                    self._stats["equations_failed"] += 1
            else:
                self._add_rich_text_runs(p, token)

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
            # For display math inside a table cell, use oMath (not oMathPara)
            xml_str = etree.tostring(
                omml_elem if hasattr(omml_elem, 'tag') else omml_elem,
                encoding="unicode"
            )
            # Unwrap oMathPara if present — inside a table cell we just need oMath
            xml_str = re.sub(
                r'<m:oMathPara[^>]*>(.*?)</m:oMathPara>',
                r'\1', xml_str, flags=re.DOTALL
            )
            try:
                inner = parse_xml(xml_str)
                eq_p._p.append(inner)
            except Exception:
                eq_p._p.append(omml_elem)
            self._stats["equations_converted"] += 1
        else:
            self._fallback_equation_text(eq_p, formula)
            self._stats["equations_failed"] += 1

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
            cells = [c.strip() for c in line.strip().strip("|").split("|")]
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

                tokens = re.split(r'(\$[^\$\n]+?\$)', val)
                for token in tokens:
                    if not token:
                        continue
                    if token.startswith("$") and token.endswith("$") and len(token) > 2:
                        self._stats["equations_total"] += 1
                        omml = self.latex_to_omml(token, is_display=False)
                        if omml is not None:
                            cell_p._p.append(omml)
                            self._stats["equations_converted"] += 1
                        else:
                            r = cell_p.add_run(token[1:-1])
                            r.italic = True
                            self._stats["equations_failed"] += 1
                    else:
                        r = cell_p.add_run(token)
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
        r = p.add_run(text)
        r.italic = True; r.font.size = Pt(11)

    def _add_list_item(self, doc, text):
        p = doc.add_paragraph(style="List Bullet")
        # clear auto-generated run and use formatted paragraph logic
        p.text = ""
        pattern = re.compile(r'(\$[^\$\n]+?\$)')
        tokens = pattern.split(text)
        for token in tokens:
            if not token:
                continue
            if token.startswith("$") and token.endswith("$") and len(token) > 2:
                self._stats["equations_total"] += 1
                omml = self.latex_to_omml(token, is_display=False)
                if omml is not None:
                    p._p.append(omml)
                    self._stats["equations_converted"] += 1
                else:
                    r = p.add_run(token[1:-1])
                    r.italic = True
                    self._stats["equations_failed"] += 1
            else:
                self._add_rich_text_runs(p, token)

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
        """Sanitise LaTeX for latex2mathml compatibility."""
        s = latex_code.strip()
        # strip delimiters
        if s.startswith("$$") and s.endswith("$$"):
            s = s[2:-2].strip()
        elif s.startswith("$") and s.endswith("$"):
            s = s[1:-1].strip()

        # common substitutions for latex2mathml compatibility
        s = s.replace(r"\mid", "|")
        s = s.replace(r"\mathbb", r"\mathbf")
        # \text{...} → \mathrm{...}
        s = re.sub(r'\\text\{([^}]*)\}', r'\\mathrm{\1}', s)
        # \operatorname{sin} → \sin   (for known operators)
        known_ops = ["sin", "cos", "tan", "log", "ln", "exp", "det",
                     "max", "min", "sup", "inf", "lim", "arg", "deg"]
        for op in known_ops:
            s = s.replace(f"\\operatorname{{{op}}}", f"\\{op}")
        # remaining \operatorname → \mathrm
        s = re.sub(r'\\operatorname\{([^}]*)\}', r'\\mathrm{\1}', s)
        # \, spacing → thin space (just remove, latex2mathml handles it)
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
