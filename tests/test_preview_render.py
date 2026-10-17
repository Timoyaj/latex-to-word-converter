"""Tests for the headless live-preview renderer (gui/preview_render.py).

Covers the prettifier, block segmentation, click-to-source spans and the
table-cell run builder used by the GUI preview pane and the browser mirror.
"""

from __future__ import annotations

import pytest

from gui.preview_render import (
    latex_to_pretty, render_document, runs_for_line, _cell_runs,
)


# --------------------------------------------------------------------- #
# latex_to_pretty
# --------------------------------------------------------------------- #
class TestPrettify:
    def test_subscripts_simple(self):
        assert latex_to_pretty("x_{i}^{2} + y_{j}^{3}") == "xᵢ² + yⱼ³"

    def test_single_letter_subscript(self):
        assert latex_to_pretty("N_h") == "Nₕ"

    def test_unmapped_subscript_falls_back(self):
        # 'b' has no subscript glyph → raw form preserved
        assert latex_to_pretty("N_{bh}") == "N_{bh}"

    def test_frac_sqrt(self):
        out = latex_to_pretty(r"\frac{-b \pm \sqrt{b^2-4ac}}{2a}")
        assert out == "(-b ± √(b²-4ac))/(2a)"

    def test_hats_and_in(self):
        out = latex_to_pretty(r"f_{1}\in\{0.05, 0.15, 0.35\}")
        assert out == "f₁∈{0.05, 0.15, 0.35}"

    def test_text_and_quad(self):
        out = latex_to_pretty(r"\text{for } \Delta\hat{\theta}_{DML} \quad \mathrm{and}")
        assert out == "for Δθ̂_{DML}   and"

    def test_greek_letters(self):
        assert latex_to_pretty(r"\sigma_h") == "σₕ"
        assert latex_to_pretty(r"\theta_i") == "θᵢ"

    def test_comparators_and_mid(self):
        out = latex_to_pretty(r"a \le b \ge c \ne d \approx e \leq 3")
        assert out == "a ≤ b ≥ c ≠ d ≈ e ≤ 3"

    def test_right_brackets_dropped(self):
        out = latex_to_pretty(r"y_i = \begin{cases} 1 \\ 0 \end{cases}")
        assert "\\begin" not in out and "\\end" not in out


# --------------------------------------------------------------------- #
# render_document
# --------------------------------------------------------------------- #
class TestRenderDocument:
    def test_inline_math_runs(self):
        blocks = render_document("Einstein said $E=mc^2$ here.", number_equations=False)
        (p,) = blocks
        kinds = [r["kind"] for r in p["runs"]]
        assert "math" in kinds
        # click-to-source span must cover the exact $…$ source
        math_run = next(r for r in p["runs"] if r["kind"] == "math")
        assert math_run["src"] == (1, 14, 22)
        assert math_run["pretty"] == "E=mc²"

    def test_bare_math_run(self):
        blocks = render_document("where N_h is fixed.", number_equations=False)
        (p,) = blocks
        bare = [r for r in p["runs"] if r["kind"] in ("bare", "bare_math")]
        assert len(bare) == 1
        assert bare[0]["text"] == "N_h"
        assert bare[0]["pretty"] == "Nₕ"

    def test_display_block_with_label(self):
        blocks = render_document("$$\nx + 1 = 2\n$$", number_equations=True,
                                 eq_format="({n})")
        d = blocks[0]
        assert d["kind"] == "display"
        assert d["pretty"] == "x + 1 = 2"
        assert d["label"] == "(1)"
        assert d["src"] == (1, 0, 2)

    def test_display_numbering_off(self):
        blocks = render_document("$$\nx + 1 = 2\n$$", number_equations=False)
        assert blocks[0]["label"] is None

    def test_display_latex_form(self):
        blocks = render_document(r"\[ a^2 + b^2 \]", number_equations=True)
        d = blocks[0]
        assert d["kind"] == "display"
        assert d["pretty"] == "a² + b²"
        assert d["label"] == "(1)"

    def test_headings(self):
        blocks = render_document("# Title\n\n## Sub", number_equations=False)
        kinds = [b["kind"] for b in blocks]
        assert kinds == ["heading", "heading"]
        assert blocks[0]["level"] == 1 and blocks[1]["level"] == 2

    def test_code_fence(self):
        blocks = render_document("```python\nprint('$$x$$')\n```",
                                 number_equations=False)
        (c,) = blocks
        assert c["kind"] == "code"
        assert c["lines"] == ["print('$$x$$')"]

    def test_quote_and_list(self):
        blocks = render_document("> quoted\n\n- item one", number_equations=False)
        kinds = [b["kind"] for b in blocks]
        assert "quote" in kinds and "list" in kinds

    def test_table_block(self):
        blocks = render_document("| a_1 | $b$ |\n|---|---|\n| 1 | 2 |",
                                 number_equations=False)
        t = blocks[0]
        assert t["kind"] == "table"
        assert len(t["rows"]) == 2
        header = t["rows"][0]
        assert header[0][0]["pretty"] == "a₁"
        assert header[1][0]["pretty"] == "b"

    def test_table_skips_math_conversion(self):
        blocks = render_document("| $x_i$ | n |\n|---|---|",
                                 number_equations=False)
        t = blocks[0]
        assert t["kind"] == "table"


# --------------------------------------------------------------------- #
# runs_for_line / _cell_runs
# --------------------------------------------------------------------- #
class TestRunsForLine:
    def test_split_keeps_clickable_spans(self):
        runs = runs_for_line("take $x_1$ and N_h", 7)
        kinds = [r["kind"] for r in runs]
        assert kinds == ["text", "math", "text", "bare"]
        math = next(r for r in runs if r["kind"] == "math")
        assert math["src"] == (7, 5, 10)
        bare = next(r for r in runs if r["kind"] in ("bare", "bare_math"))
        assert bare["src"] == (7, 15, 18)

    def test_escaped_dollar_is_text(self):
        runs = runs_for_line("price \\\\$5 and done", 3)
        assert all(r["kind"] == "text" for r in runs)

    def test_unbalanced_dollar_stays_text(self):
        runs = runs_for_line("cost $5 today", 2)
        assert all(r["kind"] == "text" for r in runs)

    def test_code_span_markdown(self):
        runs = runs_for_line("use `x_1` here", 2)
        kinds = [r["kind"] for r in runs]
        assert "code" in kinds

    def test_line_no_optional(self):
        runs = runs_for_line("plain text")
        assert [r["text"] for r in runs] == ["plain text"]
        assert runs[0]["src"] is None

    def test_cell_runs(self):
        runs = _cell_runs("a_1")
        assert runs[0]["pretty"] == "a₁"
        assert runs[0]["kind"] in ("bare", "bare_math")

    def test_cell_runs_math(self):
        runs = _cell_runs("$\\frac{1}{2}$")
        assert runs[0]["kind"] == "math"
        assert runs[0]["pretty"] == "(1)/(2)"
