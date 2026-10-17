#!/usr/bin/env python3
"""
Preview rendering for the GUI side-by-side conversion pane
==========================================================
Turns Markdown/LaTeX source into lightweight "document blocks" with a
Unicode-pretty rendering of every formula, so the user can vet exactly what
will be converted — including *bare* (un-delimited) detections — before
saving a .docx.

Pure logic (no tkinter) so it is unit-testable; `gui/app.py` binds the
resulting blocks into a read-only Text widget with clickable formula tags.
"""

from __future__ import annotations

import re
from typing import Dict, List, Optional, Tuple

from core.math_tokens import (
    split_inline_math, split_table_row, is_fence_line, fence_marker,
    display_blocks, find_bare_math_in_line, normalize_bare_math,
    INLINE_DOLLAR_RE, INLINE_PAREN_RE, CODE_SPAN_RE,
)


# ══════════════════════════════════════════════════════════════════════════
# LaTeX → Unicode prettifier (preview-only; the docx uses real OMML)
# ══════════════════════════════════════════════════════════════════════════
_GREEK = {
    "alpha": "α", "beta": "β", "gamma": "γ", "delta": "δ", "epsilon": "ε",
    "varepsilon": "ϵ", "zeta": "ζ", "eta": "η", "theta": "θ",
    "vartheta": "ϑ", "iota": "ι", "kappa": "κ", "lambda": "λ", "mu": "μ",
    "nu": "ν", "xi": "ξ", "pi": "π", "rho": "ρ", "sigma": "σ", "tau": "τ",
    "upsilon": "υ", "phi": "φ", "varphi": "ϕ", "chi": "χ", "psi": "ψ",
    "omega": "ω", "Gamma": "Γ", "Delta": "Δ", "Theta": "Θ", "Lambda": "Λ",
    "Xi": "Ξ", "Pi": "Π", "Sigma": "Σ", "Upsilon": "Υ", "Phi": "Φ",
    "Psi": "Ψ", "Omega": "Ω",
}
_SYMBOLS = {
    "infty": "∞", "in": "∈", "notin": "∉", "subset": "⊂", "subseteq": "⊆",
    "supset": "⊃", "supseteq": "⊇", "leq": "≤", "geq": "≥", "neq": "≠",
    "le": "≤", "ge": "≥", "ne": "≠", "lt": "<", "gt": ">",
    "approx": "≈", "equiv": "≡", "cdot": "·", "times": "×", "div": "÷",
    "pm": "±", "mp": "∓", "mid": "∣", "to": "→", "rightarrow": "→",
    "leftarrow": "←", "Rightarrow": "⇒", "ldots": "…", "cdots": "⋯",
    "dots": "…", "sum": "Σ", "prod": "∏", "int": "∫", "oint": "∮",
    "partial": "∂", "nabla": "∇", "forall": "∀", "exists": "∃",
    "neg": "¬", "wedge": "∧", "vee": "∨", "cup": "∪", "cap": "∩",
    "emptyset": "∅", "infty": "∞", "prime": "′", "star": "⋆",
}
_SUP = str.maketrans("0123456789+-=()ni xabdehjklmnoprstuvw",
                     "⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼⁽⁾ⁿⁱ·ˣᵃᵇᵈᵉʰʲᵏˡᵐⁿᵒᵖʳˢᵗᵘᵛʷ")
_SUB = str.maketrans("0123456789+-=()aehijklmnoprstuvx",
                     "₀₁₂₃₄₅₆₇₈₉₊₋₌₍₎ₐₑₕᵢⱼₖₗₘₙₒₚᵣₛₜᵤᵥₓ")
_BB = {"R": "ℝ", "N": "ℕ", "Z": "ℤ", "Q": "ℚ", "C": "ℂ", "P": "ℙ", "H": "ℍ"}

_ACCENTS = [("hat", "\u0302"), ("bar", "\u0304"), ("tilde", "\u0303"),
            ("vec", "\u20d7"), ("dot", "\u0307"), ("ddot", "\u0308"),
            ("prime", "′")]


def _sub_or_raw(content: str) -> str:
    if all(c.isspace() or ord(c) > 127 or c.translate(_SUB) != c
           for c in content):
        return content.translate(_SUB)
    return f"_{{{content}}}" if len(content) != 1 else f"_{content}"


def _sup_or_raw(content: str) -> str:
    if all(c.isspace() or ord(c) > 127 or c.translate(_SUP) != c
           for c in content):
        return content.translate(_SUP)
    return f"^{{{content}}}" if len(content) != 1 else f"^{content}"


def latex_to_pretty(latex: str) -> str:
    """Approximate visual rendering of LaTeX for the preview pane."""
    s = latex.strip()
    s = s.replace(r"\left", "").replace(r"\right", "")
    s = s.replace(r"\!", "").replace(r"\,", " ").replace(r"\;", " ")
    s = s.replace(r"\{", "\x01").replace(r"\}", "\x02")   # literal braces
    s = s.replace(r"\qquad", "\x04").replace(r"\quad", "\x03")  # wide spaces

    # roots first (their braces nest inside \frac args), then fractions
    s = re.sub(r'\\sqrt\[([^\]]*)\]\{([^{}]*)\}', r'\2^(1/\1)', s)
    for _ in range(8):
        new = re.sub(r'\\sqrt\{([^{}]*)\}', r'√(\1)', s)
        if new == s:
            break
        s = new
    for _ in range(8):
        new = re.sub(r'\\frac\{([^{}]*)\}\{([^{}]*)\}', r'(\1)/(\2)', s)
        if new == s:
            break
        s = new

    # accents: \hat{N} → N̂
    for name, mark in _ACCENTS:
        for _ in range(6):
            new = re.sub(rf'\\{name}\{{([^{{}}]*)\}}',
                         lambda m, mk=mark: m.group(1) + mk, s)
            if new == s:
                break
            s = new

    s = re.sub(r'\\mathbb\{([^{}]*)\}',
               lambda m: "".join(_BB.get(c, c) for c in m.group(1)), s)
    s = re.sub(r'\\(?:mathrm|text|operatorname|mathbf|mathit)\{([^{}]*)\}',
               r'\1', s)

    # word commands
    for w in sorted(_GREEK, key=len, reverse=True):
        s = re.sub(rf'\\{w}(?![a-zA-Z])', _GREEK[w], s)
    for w in sorted(_SYMBOLS, key=len, reverse=True):
        s = re.sub(rf'\\{w}(?![a-zA-Z])', _SYMBOLS[w], s)
    s = re.sub(r'\\[a-zA-Z]+', '', s)          # drop unknown commands

    # sub/superscripts, innermost first
    for _ in range(8):
        new = re.sub(r'_\{([^{}]*)\}',
                     lambda m: _sub_or_raw(m.group(1)), s)
        new = re.sub(r'\^\{([^{}]*)\}',
                     lambda m: _sup_or_raw(m.group(1)), new)
        if new == s:
            break
        s = new
    s = re.sub(r'_([A-Za-z0-9])', lambda m: _sub_or_raw(m.group(1)), s)
    s = re.sub(r'\^([A-Za-z0-9])', lambda m: _sup_or_raw(m.group(1)), s)

    s = s.replace("\x01", "{").replace("\x02", "}")
    s = re.sub(r'\s+', ' ', s).strip()
    s = re.sub(r' *\x04 *', "     ", s)      # \qquad → 5 spaces
    s = re.sub(r' *\x03 *', "   ", s)        # \quad  → 3 spaces
    return s


# ══════════════════════════════════════════════════════════════════════════
# Document blocks
# ══════════════════════════════════════════════════════════════════════════
# Each block is a dict:
#   {"kind": "heading"|"para"|"quote"|"list"|"display"|"code"|"table",
#    "line": first source line (1-based), "runs": [...], ...}
# Each run is a dict:
#   {"kind": "text"|"math"|"bare"|"code", "text": str,          # raw source
#    "pretty": str,                                             # rendered
#    "src": (line, col0, col1) | None}                          # click target


def _run(kind: str, text: str, src: Optional[Tuple[int, int, int]]) -> dict:
    if kind in ("math", "bare", "bare_math"):
        pretty = latex_to_pretty(
            text if kind == "math" else normalize_bare_math(text))
    elif kind == "code":
        pretty = text.strip("`")
    else:
        pretty = text
    return {"kind": kind, "text": text, "pretty": pretty, "src": src}


def _cell_runs(cell: str) -> List[dict]:
    """Tokenize a table cell into text/math/bare/code runs (no click src)."""
    out: List[dict] = []
    for tok in split_inline_math(cell):
        out.append(_run(tok.kind, tok.value, None))
    return out or [_run("text", cell, None)]


def runs_for_line(line: str, line_no: int = None) -> List[dict]:
    """Tokenize one source line into runs carrying click offsets."""
    spans: List[Tuple[int, int, str]] = []
    for m in CODE_SPAN_RE.finditer(line):
        spans.append((m.start(), m.end(), "code"))
    for m in INLINE_DOLLAR_RE.finditer(line):
        spans.append((m.start(), m.end(), "math"))
    for m in INLINE_PAREN_RE.finditer(line):
        spans.append((m.start(), m.end(), "math"))
    for s, e in find_bare_math_in_line(line):
        spans.append((s, e, "bare"))
    spans.sort()

    runs: List[dict] = []
    pos = 0

    def _sp(a: int, b: int):
        return (line_no, a, b) if line_no is not None else None

    for s, e, kind in spans:
        if s < pos:
            continue
        if s > pos:
            runs.append(_run("text", line[pos:s], _sp(pos, s)))
        raw = line[s:e]
        if kind == "math":
            # strip delimiters for display; click span keeps them
            if raw.startswith("$$") and raw.endswith("$$"):
                raw = raw[2:-2]
            elif raw.startswith("$") and raw.endswith("$"):
                raw = raw[1:-1]
            elif raw.startswith("\\(") and raw.endswith("\\)"):
                raw = raw[2:-2]
        runs.append(_run(kind, raw, _sp(s, e)))
        pos = e
    if pos < len(line):
        runs.append(_run("text", line[pos:], _sp(pos, len(line))))
    return runs


def render_document(text: str,
                    number_equations: bool = True,
                    eq_format: str = "({n})") -> List[dict]:
    """Segment *text* into preview blocks mirroring the converter's parse."""
    lines = text.split("\n")
    blocks: List[dict] = []
    disp = set()
    for a, b in display_blocks(lines):
        disp.update(range(a, b + 1))

    i = 0
    eq_no = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        # ── code fences ──────────────────────────────────────────
        if is_fence_line(line):
            marker = fence_marker(line)[0] * 3
            body = []
            j = i + 1
            while j < len(lines) and not (
                    is_fence_line(lines[j])
                    and fence_marker(lines[j])[0] == marker[0]):
                body.append(lines[j])
                j += 1
            blocks.append({"kind": "code", "line": i + 1,
                           "lines": body, "runs": []})
            i = j + 1
            continue

        # ── display equations ────────────────────────────────────
        if stripped.startswith("\\[") and stripped.endswith("\\]") and len(stripped) > 3:
            raw = stripped[2:-2].strip()
            eq_no += 1
            label = (eq_format.replace("{n}", str(eq_no))
                     if number_equations else None)
            blocks.append({"kind": "display", "line": i + 1,
                           "text": raw, "pretty": latex_to_pretty(raw),
                           "label": label,
                           "src": (i + 1, 0, len(line)),
                           "runs": []})
            i += 1
            continue

        if i in disp:
            j = i
            while j + 1 < len(lines) and (j + 1) in disp:
                j += 1
            raw = " ".join(
                ln.strip().strip("$").strip()
                for ln in lines[i:j + 1] if ln.strip())
            eq_no += 1
            label = (eq_format.replace("{n}", str(eq_no))
                     if number_equations else None)
            blocks.append({"kind": "display", "line": i + 1,
                           "text": raw, "pretty": latex_to_pretty(raw),
                           "label": label,
                           "src": (i + 1, 0, len(line)),
                           "runs": []})
            i = j + 1
            continue

        # ── blank ────────────────────────────────────────────────
        if not stripped:
            i += 1
            continue

        # ── headings ─────────────────────────────────────────────
        m = re.match(r'^(#{1,4})\s+(.*)$', stripped)
        if m:
            level = len(m.group(1))
            blocks.append({"kind": "heading", "level": level, "line": i + 1,
                           "runs": [_run("text", m.group(2),
                                         (i + 1, 0, len(line)))]})
            i += 1
            continue

        # ── horizontal rule ──────────────────────────────────────
        if re.match(r'^-{3,}$', stripped):
            i += 1
            continue

        # ── tables ───────────────────────────────────────────────
        if stripped.startswith("|") and stripped.endswith("|"):
            rows = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                row = lines[i].strip()
                if not re.match(r'^\|?[\s\:\-]+(?:\|[\s\:\-]+)*\|?$', row):
                    cells = split_table_row(row)
                    rows.append([_cell_runs(c.strip()) for c in cells])
                i += 1
            blocks.append({"kind": "table", "line": i, "rows": rows, "runs": []})
            continue

        # ── quote / list ─────────────────────────────────────────
        if stripped.startswith("> "):
            blocks.append({"kind": "quote", "line": i + 1,
                           "runs": runs_for_line(stripped[2:], i + 1)})
            i += 1
            continue
        m = re.match(r'^([-*+]|\d+[.)])\s+(.*)$', stripped)
        if m:
            blocks.append({"kind": "list", "line": i + 1,
                           "bullet": m.group(1),
                           "runs": runs_for_line(m.group(2), i + 1)})
            i += 1
            continue

        # ── paragraph (consecutive prose lines) ──────────────────
        runs: List[dict] = []
        start = i
        while i < len(lines):
            ln = lines[i]
            s = ln.strip()
            if (not s or is_fence_line(ln) or i in disp
                    or re.match(r'^#{1,4}\s', s) or re.match(r'^-{3,}$', s)
                    or (s.startswith("|") and s.endswith("|"))
                    or s.startswith("> ") or re.match(r'^([-*+]|\d+[.)])\s', s)):
                break
            runs.extend(runs_for_line(ln, i + 1))
            if i > start:
                runs.append(_run("text", "\n", None))
            i += 1
        blocks.append({"kind": "para", "line": start + 1, "runs": runs})

    return blocks
