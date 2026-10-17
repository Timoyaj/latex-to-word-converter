"""Regression tests for the bare-math detection bugs reported 2026-10-17:

1. ``N_h`` (single letter + subscript) was never detected.
2. Long formulas containing double-space typos were detected only partially
   (seed ``_{``/``^(`` lost bracket depth; stray `` . `` broke the span).
3. Formulas wrapped across two lines were split into two partial spans.
"""

from __future__ import annotations

import pytest

from core.math_tokens import find_bare_math_in_line, normalize_bare_math


def _bare(text: str) -> list[str]:
    """Return detected bare spans (line may contain \\n)."""
    return [text[s:e] for s, e in find_bare_math_in_line(text)]


class TestSingleLetterSubscripts:
    def test_N_h_detected(self):
        spans = _bare("The sample size N_h for stratum h is fixed.")
        assert "N_h" in spans

    def test_W_ij_detected(self):
        spans = _bare("The weight W_ij and the mean theta_i are used.")
        assert "W_ij" in spans
        assert "theta_i" in spans

    def test_word_only_line_stays_prose(self):
        assert _bare("The sample size for stratum is fixed.") == []
        assert _bare("max_size and snake_case_name") == []


class TestDoubleSpaceTypos:
    def test_full_span_despite_double_spaces(self):
        text = "sum_{k in s_h} w_hk  *  (y_hk  -  y_hbar)  =  0"
        spans = _bare(text)
        assert spans == [text]
        norm = normalize_bare_math(text)
        assert "  " not in norm

    def test_seed_brackets_keep_depth(self):
        # the seed ends at '_{' / '^(' — expansion must not stop at the
        # matching '}' / ')' that closes the group the seed opened
        text = "s^{(1)} + s^{(2)} = t"
        spans = _bare(text)
        assert spans == [text]

    def test_stray_period_tolerated(self):
        text = "T_h = n_h^(-1/2) .  + sigma_h"
        spans = _bare(text)
        assert spans == [text]
        norm = normalize_bare_math(text)
        assert "." not in norm
        assert "sigma" in norm

    def test_normal_period_still_breaks(self):
        text = "We use T_h = n_h. The sample is large."
        spans = _bare(text)
        assert spans == ["T_h = n_h"]

    def test_glued_period_no_space_continues(self):
        # decimal point inside a number
        spans = _bare("rate is 0.05 today")
        assert spans == []


class TestMultilineFormulas:
    def test_wrap_crosses_newline(self):
        text = "y_i = beta_0 + beta_1 * x_i\n+ u_i"
        spans = _bare(text)
        assert spans == [text]
        # normalize joins the wrapped lines into one formula
        assert "\n" not in normalize_bare_math(text)

    def test_inline_gap_crosses_newline(self):
        text = "y_i = beta_0 +\nbeta_1 * x_i"
        spans = _bare(text)
        assert spans == [text]


class TestEdgePreservation:
    def test_comparators_detected(self):
        text = "rate_2 <= t_1 and speed >= 3."
        spans = _bare(text)
        assert "rate_2 <= t_1" in spans
        assert "speed >= 3" in spans

    def test_subject_kept_with_operator_edge(self):
        assert "y_i = speed" in _bare("y_i = speed is recorded")

    def test_mixed_prose_and_formula(self):
        text = "The estimator theta_i + u_i is consistent."
        spans = _bare(text)
        assert spans == ["theta_i + u_i"]


class TestNormalizeBareMath:
    def test_nested_groups_roundtrip(self):
        # braces inside groups survive; bare braces are LaTeX-escaped
        assert normalize_bare_math("f_{1} in {a, b}") == "f_{1} \\in \\{a, b\\}"

    def test_double_spaces_collapsed(self):
        assert normalize_bare_math("a  +  b") == "a + b"

    def test_stray_periods_removed(self):
        assert normalize_bare_math("a = b .  + c") == "a = b + c"
        assert normalize_bare_math("a = b . + c") == "a = b + c"

    def test_newline_collapsed(self):
        assert normalize_bare_math("a =\nb") == "a = b"
