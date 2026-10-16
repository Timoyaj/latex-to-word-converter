"""Integration tests: run fixtures through the converter and inspect the docx."""

from __future__ import annotations

import pytest
from lxml import etree

from core.converter import LatexToWordConverter
from tests.conftest import (
    md_text, docx_xml, omath_count, eq_labels, table_count, load_expected,
)


def convert(converter, name: str, out_docx, **kwargs):
    kwargs.setdefault("number_equations", True)
    kwargs.setdefault("eq_format", "({n})")
    return converter.convert_markdown_to_docx(md_text(name), str(out_docx), **kwargs)


# ── fixture-driven structural goldens ──────────────────────────────────────

@pytest.mark.parametrize(
    "fixture,expected_file",
    [
        ("inline_basics.md", "inline_basics.json"),
        ("display_multiline.md", "display_multiline.json"),
        ("display_singleline.md", "display_singleline.json"),
        ("numbering_formats.md", "numbering_formats.json"),
        ("tables_math.md", "tables_math.json"),
        ("markdown_mix.md", "markdown_mix.json"),
        ("edge_cases.md", "edge_cases.json"),
        ("failures.md", "failures.json"),
    ],
)
def test_fixture_structure(converter, out_docx, fixture, expected_file):
    exp = load_expected(expected_file)
    stats = convert(converter, fixture, out_docx)
    xml = docx_xml(out_docx)

    assert stats["equations_total"] == exp["equations_total"]
    assert stats["equations_converted"] == exp["equations_converted"]
    assert stats["equations_failed"] == exp["equations_failed"]
    assert omath_count(xml) == exp["omath_count"]
    assert table_count(xml) == exp["table_count"]
    if "eq_labels" in exp:
        assert eq_labels(xml) == exp["eq_labels"]


# ── numbering behaviour ───────────────────────────────────────────────────

def test_numbering_default_format(converter, out_docx):
    convert(converter, "numbering_formats.md", out_docx)
    assert eq_labels(docx_xml(out_docx)) == ["1", "2", "3"]


def test_numbering_custom_format(converter, out_docx):
    convert(converter, "numbering_formats.md", out_docx, eq_format="[Eq. {n}]")
    xml = docx_xml(out_docx)
    for n in ("1", "2", "3"):
        assert f"[Eq. {n}]" in xml


def test_numbering_disabled(converter, out_docx):
    stats = convert(converter, "numbering_formats.md", out_docx, number_equations=False)
    xml = docx_xml(out_docx)
    assert eq_labels(xml) == []
    assert table_count(xml) == 0
    assert stats["equations_converted"] == 3


def test_counter_resets_between_runs(converter, out_docx, tmp_path):
    convert(converter, "numbering_formats.md", out_docx)
    second = tmp_path / "second.docx"
    stats = convert(converter, "numbering_formats.md", second)
    assert eq_labels(docx_xml(second)) == ["1", "2", "3"]


# ── OMML structure ────────────────────────────────────────────────────────

def test_display_math_in_tables_is_omath_not_omathpara(converter, out_docx):
    convert(converter, "numbering_formats.md", out_docx)
    xml = docx_xml(out_docx)
    assert "<m:oMathPara" not in xml
    assert omath_count(xml) == 3


def test_display_math_without_numbering_is_omathpara(converter, out_docx):
    convert(converter, "display_multiline.md", out_docx, number_equations=False)
    xml = docx_xml(out_docx)
    assert xml.count("<m:oMathPara") == 3


def test_mathbb_keeps_blackboard_glyph(converter, out_docx):
    convert(converter, "inline_basics.md", out_docx)
    xml = docx_xml(out_docx)
    assert "0211D" in xml or "\u211d" in xml


def test_middle_not_corrupted(converter, out_docx):
    convert(converter, "edge_cases.md", out_docx)
    xml = docx_xml(out_docx)
    assert "|dle|" not in xml  # the classic \middle corruption artifact


# ── markdown features ─────────────────────────────────────────────────────

def test_markdown_tables(converter, out_docx):
    convert(converter, "tables_math.md", out_docx)
    xml = docx_xml(out_docx)
    # one markdown table only (no numbered equations in this fixture)
    assert table_count(xml) == 1
    assert omath_count(xml) == 8  # every cell's math converts


def test_abs_value_cell_not_split(converter, out_docx):
    convert(converter, "tables_math.md", out_docx)
    xml = docx_xml(out_docx)
    # $|x|$ and $\|x\|$ keep their pipes inside math:
    # exactly 5 rows x 3 columns
    assert xml.count("<w:tc>") == 15


def test_headings_and_lists(converter, out_docx):
    convert(converter, "markdown_mix.md", out_docx)
    xml = docx_xml(out_docx)
    for style in ("Heading1", "Heading2", "Heading3", "Heading4",
                  "ListBullet", "ListNumber"):
        assert style in xml, f"missing style {style}"


def test_code_fence_rendered_monospace(converter, out_docx):
    convert(converter, "edge_cases.md", out_docx)
    xml = docx_xml(out_docx)
    assert "Consolas" in xml
    assert "costs $5 and $10" in xml  # preserved verbatim


def test_currency_not_converted(converter, out_docx):
    stats = convert(converter, "edge_cases.md", out_docx)
    # edge_cases.md has 5 real equations (2 display + bracket + 2 paren-style)
    assert stats["equations_converted"] == 5
    assert stats["equations_failed"] == 0
    xml = docx_xml(out_docx)
    assert "The price is $5.00 and the other is $10.00 today" in xml


# ── failure reporting ─────────────────────────────────────────────────────

def test_failed_formulas_reported_with_lines(converter, out_docx):
    stats = convert(converter, "failures.md", out_docx)
    assert stats["equations_failed"] == 2
    assert stats["equations_converted"] == 1
    failed = stats["failed_formulas"]
    assert len(failed) == 2
    assert all(isinstance(item["line"], int) and item["line"] > 0 for item in failed)
    assert any("notacommand" in item["latex"] for item in failed)
    # fallback text is present in the document
    xml = docx_xml(out_docx)
    assert "notacommand" in xml


# ── golden OMML fragment ──────────────────────────────────────────────────

def test_golden_omml_x_squared(converter):
    """Structural golden: OMML for x^2 (regenerate with the snippet in the
    conftest docstring if the XSLT or latex2mathml output legitimately
    changes)."""
    omml = converter.latex_to_omml("x^2", is_display=False)
    assert omml is not None
    xml = etree.tostring(omml, encoding="unicode")
    assert "oMath" in xml
    # exponent structure: x with a superscript containing 2
    assert "<m:t>x</m:t>" in xml.replace(" ", "")
    assert "<m:t>2</m:t>" in xml.replace(" ", "")
