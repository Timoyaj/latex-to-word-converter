"""Unit tests for LaTeX cleaning — the minimal cleaner + legacy fallback."""

from core.converter import LatexToWordConverter as C


class TestCleanLatex:
    def test_strips_dollar_delimiters(self):
        assert C._clean_latex("$x^2$") == "x^2"
        assert C._clean_latex("$$x^2$$") == "x^2"

    def test_mid_becomes_pipe(self):
        assert C._clean_latex(r"x \mid y") == "x | y"

    def test_middle_survives(self):
        # regression: naive replace(r"\mid", "|") corrupted \middle -> |dle|
        assert C._clean_latex(r"\left( x \middle| y \right)") == (
            r"\left( x \middle| y \right)"
        )
        assert C._clean_latex(r"a \middle| b") == r"a \middle| b"

    def test_mathbb_preserved(self):
        # regression: \mathbb{R} used to be demoted to \mathbf{R}
        assert C._clean_latex(r"\mathbb{R}") == r"\mathbb{R}"

    def test_text_preserved(self):
        assert C._clean_latex(r"\text{J}\cdot\text{s}") == r"\text{J}\cdot\text{s}"

    def test_operatorname_preserved(self):
        assert C._clean_latex(r"\operatorname{argmax}") == r"\operatorname{argmax}"


class TestLegacyCompat:
    def test_demotes_mathbb_only_in_fallback(self):
        assert C._legacy_compat_latex(r"\mathbb{R}") == r"\mathbf{R}"

    def test_text_to_mathrm(self):
        assert C._legacy_compat_latex(r"\text{rad}") == r"\mathrm{rad}"

    def test_known_operatorname(self):
        assert C._legacy_compat_latex(r"\operatorname{sin}") == r"\sin"

    def test_unknown_operatorname(self):
        assert C._legacy_compat_latex(r"\operatorname{argmax}") == r"\mathrm{argmax}"

    def test_middle_still_survives_legacy(self):
        assert C._legacy_compat_latex(r"a \middle| b") == r"a \middle| b"


class TestTwoPassConversion:
    def test_mathbb_converts_natively(self, converter):
        omml = converter.latex_to_omml(r"\mathbb{R}", is_display=False)
        assert omml is not None
        from lxml import etree
        xml = etree.tostring(omml, encoding="unicode")
        # native path keeps the blackboard glyph (U+0211D), not a bold R
        assert "0211D" in xml or "\u211d" in xml

    def test_bad_latex_returns_none(self, converter):
        assert converter.latex_to_omml(r"\notacommand{{", is_display=False) is None
