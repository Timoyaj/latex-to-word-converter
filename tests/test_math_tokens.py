"""Unit tests for core.math_tokens — the shared inline/display math scanner."""

from core.math_tokens import (
    Token,
    split_inline_math,
    display_blocks,
    count_display_math,
    is_fence_line,
    fence_marker,
    INLINE_DOLLAR_RE,
    find_bare_math,
    find_bare_math_in_line,
    normalize_bare_math,
)


def kinds(tokens):
    return [t.kind for t in tokens]


def math_values(tokens):
    return [t.value for t in tokens if t.kind == "math"]


class TestSplitInlineMath:
    def test_simple_math(self):
        toks = split_inline_math("energy $E = mc^2$ here")
        assert kinds(toks) == ["text", "math", "text"]
        assert math_values(toks) == ["E = mc^2"]

    def test_multiple_maths(self):
        toks = split_inline_math("$a$ and $b$")
        assert math_values(toks) == ["a", "b"]

    def test_currency_pair_is_not_math(self):
        toks = split_inline_math("The price is $5.00 and the other is $10.00 today")
        assert math_values(toks) == []
        assert kinds(toks) == ["text"]

    def test_currency_then_math(self):
        toks = split_inline_math("costs $5.00 but $x+1$ is math")
        assert math_values(toks) == ["x+1"]

    def test_math_ending_in_digit_still_math(self):
        # closing $ not followed by a digit => fine ("$x+1$" at end of text)
        assert math_values(split_inline_math("solve $x+1$")) == ["x+1"]
        # "$2$" is math even though it opens with a digit
        assert math_values(split_inline_math("value $2$ today")) == ["2"]

    def test_dollar_followed_by_digit_after_close_is_literal(self):
        # "$a$100" should not become math (Pandoc rule)
        assert math_values(split_inline_math("costs $a$100 now")) == []

    def test_escaped_dollar_is_literal(self):
        toks = split_inline_math(r"cost is \$5 and $x+1$ here")
        assert math_values(toks) == ["x+1"]
        assert any("$5" in t.value for t in toks if t.kind == "text")

    def test_code_span_protected(self):
        toks = split_inline_math("a `code $x$ span` and $y^2$ math")
        assert math_values(toks) == ["y^2"]
        assert any(t.kind == "code" for t in toks)

    def test_no_whitespace_after_opening_dollar(self):
        assert math_values(split_inline_math("$ x $")) == []

    def test_paren_inline_math(self):
        toks = split_inline_math(r"the slope \( m \) is steep")
        assert math_values(toks) == [" m "]

    def test_multiline_is_split_per_line_by_callers(self):
        # bodies never span lines
        assert math_values(split_inline_math("line1 $a\nb$ line2")) == []


class TestDisplayBlocks:
    def test_single_line(self):
        assert display_blocks(["$$x^2$$"]) == [(0, 0)]

    def test_multi_line(self):
        lines = ["text", "$$", "x^2", "$$", "more"]
        assert display_blocks(lines) == [(1, 3)]

    def test_empty_pair_skipped(self):
        assert display_blocks(["$$  $$"]) == []

    def test_unterminated_ignored(self):
        assert display_blocks(["$$", "x^2"]) == []


class TestCountDisplayMath:
    def test_counts_blocks(self):
        text = "$$a$$\n\n$$\nb\n$$\n"
        assert count_display_math(text) == 2

    def test_fences_protected(self):
        text = "```\n$$x$$\n```\n\n$$a$$\n"
        assert count_display_math(text) == 1


class TestFences:
    def test_backtick_fence(self):
        assert is_fence_line("```python")
        assert fence_marker("```python") == "```"

    def test_tilde_fence(self):
        assert is_fence_line("~~~")
        assert fence_marker("~~~") == "~~~"

    def test_plain_line(self):
        assert not is_fence_line("just text")


class TestPublicRegexes:
    def test_gui_regex_matches_tokenizer(self):
        # the exported regex must agree with the tokenizer on a tricky line
        line = "price $5.00 and $x+1$ end"
        assert math_values(split_inline_math(line)) == ["x+1"]
        spans = [m.group(0) for m in INLINE_DOLLAR_RE.finditer(line)]
        assert spans == ["$x+1$"]


class TestBareMath:
    """Formulas typed WITHOUT $ or \\( delimiters must still be recognized."""

    USER_EXAMPLE = (
        "f_1 in {0.05, 0.15, 0.35}; f_2 in {0.10, 0.25, 0.50}, "
        "E_2[ N_hat * theta_hat_DML | s_1 ] = sum_{k in s_1} w_1k Y_k "
        "+ o_p(N n^(-1/2))."
    )

    def test_user_example_detected(self):
        toks = split_inline_math(self.USER_EXAMPLE)
        bare = [t.value for t in toks if t.kind == "bare_math"]
        assert len(bare) == 3
        assert bare[0] == "f_1 in {0.05, 0.15, 0.35}"
        assert bare[1] == "f_2 in {0.10, 0.25, 0.50}"
        assert bare[2].startswith("E_2[ N_hat * theta_hat_DML")
        assert bare[2].endswith("o_p(N n^(-1/2))")

    def test_prose_identifiers_stay_prose(self):
        toks = split_inline_math(
            "the max_size of a snake_case_name is not math")
        assert math_values(toks) == []
        assert not any(t.kind == "bare_math" for t in toks)

    def test_comparisons_detected(self):
        toks = split_inline_math("rate_2 <= t_1 and speed >= 3.")
        bare = [t.value for t in toks if t.kind == "bare_math"]
        assert bare == ["rate_2 <= t_1", "speed >= 3"]

    def test_bare_formula_inside_prose(self):
        toks = split_inline_math("Let f_1 be the rate; it is small.")
        assert [t.value for t in toks if t.kind == "bare_math"] == ["f_1"]

    def test_code_and_delimited_win(self):
        toks = split_inline_math("with $x^2$ and `sum_1` plus sum_1 code")
        kinds = [(t.kind, t.value) for t in toks]
        assert ("math", "x^2") in kinds
        assert ("bare_math", "sum_1") in kinds  # trailing prose trimmed
        assert any(k == "code" for k, _ in kinds)

    def test_find_bare_math_in_line_offsets(self):
        line = "see f_1 and $x$ here"
        spans = find_bare_math_in_line(line)
        assert len(spans) == 1
        s, e = spans[0]
        assert line[s:e] == "f_1"


class TestNormalizeBareMath:
    def test_user_example_normalization(self):
        norm = normalize_bare_math(
            "E_2[ N_hat * theta_hat_DML | s_1 ] = sum_{k in s_1} "
            "w_1k Y_k + o_p(N n^(-1/2))")
        assert r"\hat{N}" in norm
        assert r"\hat{\theta}_{DML}" in norm
        assert r"\cdot" in norm
        assert r"\mid" in norm
        assert r"\sum_{k \in s_{1}}" in norm
        assert r"w_{1k}" in norm
        assert r"n^{-1/2}" in norm
        # subscript braces must NOT be escaped as set braces
        assert r"\sum_\{" not in norm

    def test_set_braces_escaped(self):
        norm = normalize_bare_math("f_1 in {0.05, 0.15, 0.35}")
        assert r"f_{1}" in norm
        assert r"\in" in norm
        assert r"\{0.05, 0.15, 0.35\}" in norm

    def test_comparisons_and_greek(self):
        assert r"\leq" in normalize_bare_math("rate_2 <= t_1")
        assert r"\geq" in normalize_bare_math("speed >= 3")
        norm = normalize_bare_math("alpha_1 + beta_2")
        assert r"\alpha" in norm and r"\beta" in norm
