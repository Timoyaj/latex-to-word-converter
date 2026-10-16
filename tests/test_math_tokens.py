"""Unit tests for core.math_tokens — the shared inline/display math scanner."""

from core.math_tokens import (
    Token,
    split_inline_math,
    display_blocks,
    count_display_math,
    is_fence_line,
    fence_marker,
    INLINE_DOLLAR_RE,
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
