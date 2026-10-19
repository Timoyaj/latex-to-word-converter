#!/usr/bin/env python3
"""
Shared Markdown / math tokenization
====================================
Single source of truth for locating inline math (`$...$`, `\\(...\\)`) and
display math (`$$...$$`, `\\[...\\]`) in Markdown text. Used by both the
conversion engine (`core/converter.py`) and the GUI editor's syntax
highlighting / equation counter so the two can never drift apart.

Inline-math discovery follows Pandoc's dollar-math rules:

* an escaped ``\\$`` is a literal dollar sign, never a math delimiter;
* the opening ``$`` must not be followed by whitespace or another ``$``;
* the closing ``$`` must not be preceded by whitespace and must not be
  followed by a digit (so ``$5.00 and $10.00`` stays prose, while
  ``$2$`` and ``$x+1$`` remain math).

Inline code spans (`` `...` ``) and fenced code blocks are never scanned
for math.
"""

from __future__ import annotations

import re
from typing import List, NamedTuple, Tuple


class Token(NamedTuple):
    """A chunk of an inline stream: plain text, math, or code."""

    kind: str  # 'text' | 'math' | 'code'
    value: str


# ── code fences ───────────────────────────────────────────────────────────
FENCE_RE = re.compile(r"^\s*(`{3,}|~{3,})")


def is_fence_line(line: str) -> bool:
    """True for ``` or ~~~ fence delimiter lines."""
    return bool(FENCE_RE.match(line))


def fence_marker(line: str) -> str:
    """Return the fence marker (e.g. '```') of a fence line, or ''."""
    m = FENCE_RE.match(line)
    return m.group(1) if m else ""


# ── inline math ───────────────────────────────────────────────────────────
# $...$ with Pandoc guard rules; body is single-line and cannot contain '$'.
_INLINE_DOLLAR_RE = re.compile(
    r"(?<![\\$])"          # not escaped, not part of $$
    r"\$(?!\$)(?!\s)"      # opening $ + non-space, non-$ right side
    r"(?P<body>[^$\n]+?)"  # formula body (lazy)
    r"(?<!\s)\$"           # closing $ with non-space to its left
    r"(?![0-9])"           # ...and not immediately followed by a digit
)

# \( ... \) inline math (also single-line).
_INLINE_PAREN_RE = re.compile(
    r"\\\((?P<body>[^$\n]+?)\\\)"
)

# `code` span (single backticks; double backticks handled by splitting).
_CODE_SPAN_RE = re.compile(r"(`+)(.+?)\1")

# Public aliases so the GUI editor can highlight with the exact same rules.
INLINE_DOLLAR_RE = _INLINE_DOLLAR_RE
INLINE_PAREN_RE = _INLINE_PAREN_RE
CODE_SPAN_RE = _CODE_SPAN_RE


# ── bare math (formulas written WITHOUT $ or \( delimiters) ────────────────
# Detection is seed-based and deliberately conservative: a span is only
# treated as math when it contains a *strong* signal (digit/brace/paren
# sub- or superscripts, operator words like sum_, accent suffixes like
# _hat, comparison operators …). Plain identifiers such as max_size or
# snake_case_name never trigger conversion.
_BARE_SEED_RE = re.compile(
    r"""
      _\{ |                          # _{
      _[0-9(] |                      # _1  _(
      \^\{ |                         # ^{
      \^[0-9(] |                     # ^2  ^(
      # operator word + short subscript (max_x ✓, max_size ✗ — identifiers)
      \b(?:sum|prod|int|lim|inf|sup|det|max|min|log|exp|ln)_
        (?:\{|[0-9]|[A-Za-z](?![A-Za-z0-9])) |
      # single-letter or greek base + short subscript: N_h, W_ij, theta_i
      \b(?:[A-Za-z]|alpha|beta|gamma|delta|epsilon|varepsilon|zeta|eta|theta|
          vartheta|iota|kappa|lambda|mu|nu|xi|pi|rho|sigma|tau|upsilon|phi|
          varphi|chi|psi|omega|Gamma|Delta|Theta|Lambda|Xi|Pi|Sigma|
          Upsilon|Phi|Psi|Omega)_
        (?:\{|\d|[A-Za-z]{1,2}(?![A-Za-z0-9])) |
      \w+_(?:hat|bar|tilde|dot|ddot|star|prime)(?![A-Za-z]) |
      <= | >= | != | ->
    """,
    re.VERBOSE,
)

# Words that terminate bare-math expansion at bracket depth 0 ("in" is
# intentionally allowed — set membership: "k in s_1").
_BARE_STOPWORDS = frozenset("""
a an the and or but nor so yet as at by to for with from on about into over
after before between out up down not no yes can will just also than too very
here there such each both all any some more most other only own same is are
was were be been being am it its this that these those then thus hence so
which who whom whose where when while why how if else do does did done have
has had let we you they he she i our your their of off example examples etc
vs versus per among during without within upon above below under again
further once twice because since until unless therefore whereas respectively
given suppose assume note recall define defined defines means denote denotes
denoting get got make makes made use used using value values variable
variables function functions equation equations formula formulas term terms
set sets word words left right where plus minus equals equal divided
multiply times result results example. code text data input output
see show shows shown give gives yield yields follow follows produce produces
""".split())

_ALLOWED_FORMULA_WORDS = frozenset("""
in notin subset subseteq supset supseteq approx equiv neq to implies iff
cdot times div sum prod int lim inf sup det max min log ln exp arg deg
hat bar tilde dot dddot star prime
alpha beta gamma delta epsilon varepsilon zeta eta theta vartheta iota
kappa lambda mu nu xi pi rho sigma tau upsilon phi varphi chi psi omega
Gamma Delta Theta Lambda Xi Pi Sigma Upsilon Phi Psi Omega
""".split())


def find_bare_math(text: str) -> List[Tuple[int, int]]:
    """Locate un-delimited formula spans as (start, end) character offsets.

    Runs on a code-free, delimited-math-free text chunk. Each strong-signal
    seed is expanded left/right across math-plausible characters while
    tracking bracket depth; expansion stops at newlines, sentence-ending
    periods, ';' ',' at depth 0, and prose stopwords at depth 0.
    """
    spans: List[Tuple[int, int]] = []
    n = len(text)
    for seed in _BARE_SEED_RE.finditer(text):
        s, e = seed.start(), seed.end()
        # skip seeds already inside an accepted span
        if any(s >= ts and e <= te for ts, te in spans):
            continue

        # ── expand left ─────────────────────────────────────────
        # seed may end with an opener (e.g. "_{") that the walk already passed
        depth = max(0, _close_count(text[s:e]) - _open_count(text[s:e]))
        i = s
        while i > 0:
            ch = text[i - 1]
            if ch in ")]}":
                depth += 1
            elif ch in "([{":
                if depth == 0:
                    break
                depth -= 1
            if depth == 0:
                if ch in ";,":
                    break
                if ch == ".":
                    # sentence period = glued to the left, space/end on the
                    # right; a stray " . " inside a formula (typo) continues
                    glued = (i - 1 > 0 and text[i - 2] not in " \t")
                    followed = text[i] if i < n else " "
                    if glued and followed in " \t":
                        break
                if ch == "$":
                    break
            # stopword check at word start (\n and space both act as gaps)
            if ch in " \n" and depth == 0:
                word = _word_before(text, i - 1)
                if word and word.lower() in _BARE_STOPWORDS:
                    break
            i -= 1
        s = i

        # ── expand right ────────────────────────────────────────
        depth = max(0, _open_count(text[s:e]) - _close_count(text[s:e]))
        j = e
        while j < n:
            ch = text[j]
            if ch in "([{":
                depth += 1
            elif ch in ")]}":
                if depth == 0:
                    break
                depth -= 1
            if depth == 0:
                if ch in ";,":
                    break
                if ch == ".":
                    glued = (j > 0 and text[j - 1] not in " \t")
                    followed = text[j + 1] if j + 1 < n else " "
                    if glued and followed in " \t":
                        break
                if ch == "$":
                    break
            if ch in " \n" and depth == 0:
                word = _word_after(text, j + 1)
                if word and word.lower() in _BARE_STOPWORDS:
                    break
            j += 1
        e = j

        span = text[s:e].strip()
        # trim dangling punctuation and connectives at the edges
        span = span.strip(" \t\n.,;:")
        for wd in ("in", "and", "or", "for", "with", "of", "to", "on", "at",
                   "by", "is", "are", "as", "but"):
            if span.lower().endswith(" " + wd):
                span = span[: -(len(wd) + 1)].rstrip()
            if span.lower().startswith(wd + " "):
                span = span[len(wd) + 1:].lstrip()
        span = _trim_prose_edges(span)
        if not span or not _BARE_SEED_RE.search(span):
            continue
        # recompute exact offsets within the expanded region
        rel = text[s:e].find(span)
        if rel == -1:
            continue
        spans.append((s + rel, s + rel + len(span)))
    return _merge_spans(spans)


# trailing word with optional sub/superscript part, e.g. "Y_k", "w_1k", "code"
_EDGE_TOKEN_RE = re.compile(
    r'[A-Za-z][A-Za-z0-9]*(?:[_^](?:\{[^}]*\}|[A-Za-z0-9]+))?')
_EDGE_KEEP_OPS = re.compile(r'(<=|>=|!=|->|[=<>+*|])')


def _trim_prose_edges(span: str) -> str:
    """Shrink a span until both edges look like math, not prose.

    An edge token survives when it carries a sub/superscript (``N_h``,
    ``Y_k``), is a single-letter identifier or allowed vocabulary
    (greek/operators/relations), is numeric, or sits directly against a
    relation/assignment operator (``speed >= 3``, ``y_i = speed`` keep their
    subjects). Everything else is prose creeping in from expansion
    (``sample size N_h`` → ``N_h``, ``sum_1 code`` → ``sum_1``).
    """
    def keep(tok_start: int, tok_end: int, side: str) -> bool:
        tok = span[tok_start:tok_end]
        m = _EDGE_TOKEN_RE.match(tok)
        base = m.group(0).split("_")[0].split("^")[0] if m else tok
        has_sub = bool(m and m.group(0) != base)
        if has_sub or base in _ALLOWED_FORMULA_WORDS or len(base) == 1:
            return True
        if base.isdigit():
            return True
        if side == "left":
            if _EDGE_KEEP_OPS.match((span[tok_end:tok_end + 8] or "").lstrip()):
                return True
        else:
            if _EDGE_KEEP_OPS.search((span[:tok_start] or "")[-6:]):
                return True
        return False

    while span:
        # left edge
        m = re.match(
            r'([A-Za-z][A-Za-z0-9]*:?(?:[_^](?:\{[^}]*\}|[A-Za-z0-9]+))?)\s+',
            span)
        if m and not keep(0, m.end(1), "left"):
            span = span[m.end():]
            continue
        # right edge
        m = re.search(r'\s+([A-Za-z][A-Za-z0-9]*(?:[_^](?:\{[^}]*\}|[A-Za-z0-9]+))?)$', span)
        if m and not keep(m.start(1), m.end(1), "right"):
            span = span[:m.start()].rstrip()
            continue
        # single-token span: keep as-is
        return span
    return span


def _word_before(text: str, idx: int) -> str:
    j = idx
    while j > 0 and (text[j - 1].isalpha() or text[j - 1] == "_"):
        j -= 1
    return text[j:idx]


def _word_after(text: str, idx: int) -> str:
    j = idx
    while j < len(text) and (text[j].isalpha() or text[j] == "_"):
        j += 1
    return text[idx:j]


def _open_count(s: str) -> int:
    return sum(s.count(c) for c in "([{")


def _close_count(s: str) -> int:
    return sum(s.count(c) for c in ")]}")


def _merge_spans(spans: List[Tuple[int, int]]) -> List[Tuple[int, int]]:
    merged: List[Tuple[int, int]] = []
    for s, e in sorted(spans):
        if merged and s < merged[-1][1]:
            merged[-1] = (merged[-1][0], max(merged[-1][1], e))
        else:
            merged.append((s, e))
    return merged


def find_bare_math_in_line(line: str) -> List[Tuple[int, int]]:
    """Bare-math spans of a single line, with offsets into the raw line.

    Delimited math ($…$, \\(…\\)) and code spans are excluded first, so the
    GUI highlighter can tag exactly what the converter will treat as math.
    """
    covered: List[Tuple[int, int]] = []
    for m in _CODE_SPAN_RE.finditer(line):
        covered.append(m.span())
    for m in _INLINE_DOLLAR_RE.finditer(line):
        covered.append(m.span())
    for m in _INLINE_PAREN_RE.finditer(line):
        covered.append(m.span())
    covered = _merge_spans(covered)

    spans: List[Tuple[int, int]] = []
    pos = 0
    for cs, ce in covered + [(len(line), len(line))]:
        if cs > pos:
            for s, e in find_bare_math(line[pos:cs]):
                spans.append((pos + s, pos + e))
        pos = max(pos, ce)
    return _merge_spans(spans)


# ── ASCII-math → LaTeX normalisation ──────────────────────────────────────
_GREEK_WORDS = (
    "alpha beta gamma delta epsilon varepsilon zeta eta theta vartheta iota "
    "kappa lambda mu nu xi pi rho sigma tau upsilon phi varphi chi psi omega "
    "Gamma Delta Theta Lambda Xi Pi Sigma Upsilon Phi Psi Omega"
).split()
_OPERATOR_WORDS = (
    "sum prod int lim inf sup det max min log ln exp arg deg"
).split()
_RELATION_WORDS = {
    "in": "in", "notin": "notin", "subset": "subset", "subseteq": "subseteq",
    "supset": "supset", "supseteq": "supseteq", "approx": "approx",
    "equiv": "equiv", "neq": "neq", "implies": "implies", "to": "to",
}


def normalize_bare_math(s: str) -> str:
    """Translate informal ASCII math into LaTeX that latex2mathml accepts.

    Handles spelled greek/operators, ``_hat`` accents (N_hat → \\hat{N}),
    set braces, ``*`` → ``\\cdot``, comparisons (<= → \\leq), multi-char
    subscripts (w_1k → w_{1k}) and parenthesised exponents (n^(-1/2)).
    """
    # 0. drop stray typographical periods ("n^(-1/2) .  +" → "n^(-1/2) +")
    s = re.sub(r'\s+\.(?=\s|$)', '', s)

    # 1. spelled greek letters → commands (theta_hat later becomes \hat{\theta})
    for w in _GREEK_WORDS:
        s = re.sub(rf'(?<![A-Za-z\\])({w})(?![A-Za-z])', r'\\\1', s)

    # 2. accent suffixes: N_hat → \hat{N}, \theta_hat_DML → \hat{\theta}_DML
    s = re.sub(
        r'((?:\\[a-zA-Z]+|[A-Za-z]))_(hat|bar|tilde|dot|ddot|star|prime)(?![A-Za-z])',
        r'\\\2{\1}', s)

    # 3. spelled operators → commands (sum_ → \sum_)
    for w in _OPERATOR_WORDS:
        s = re.sub(rf'(?<![A-Za-z\\])({w})(?![A-Za-z])', r'\\\1', s)

    # 4. spelled relations → commands (in → \in)
    for w in sorted(_RELATION_WORDS, key=len, reverse=True):
        s = re.sub(rf'(?<![A-Za-z\\])({w})(?![A-Za-z])', r'\\\1', s)

    # 5. comparison / arrow operators
    s = s.replace("<=", r"\leq ").replace(">=", r"\geq ")
    s = s.replace("!=", r"\neq ").replace("==", "=")
    s = s.replace("->", r"\to ")

    # 6. multiplication sign and spaced conditioning bar
    s = s.replace("*", r" \cdot ")
    s = re.sub(r'\s+\|\s+', r' \\mid ', s)

    # 7. parenthesised sub/superscripts: ^(-1/2) → ^{-1/2}, _(a b) → _{a b}
    s = re.sub(r'\^\(([^()]*)\)', r'^{\1}', s)
    s = re.sub(r'_\(([^()]*)\)', r'_{\1}', s)

    # 8. brace multi-char sub/superscripts: w_1k → w_{1k}, x^2 stays fine
    s = re.sub(r'(?<![_\\^])([_^])(?!\{)([A-Za-z0-9]+)', r'\1{\2}', s)

    # 9. literal set braces {…} → \{…\} (protect command args & _{}/^{} groups)
    placeholders: List[str] = []
    def _stash(m):
        placeholders.append(m.group(0))
        return f"\x00{len(placeholders) - 1}\x00"

    # stash repeatedly: nested groups (_{k in s_{1}}) resolve innermost-first
    for _ in range(8):
        new = re.sub(r'\\[a-zA-Z]+\{[^{}]*\}', _stash, s)
        new = re.sub(r'[_^]\{[^{}]*\}', _stash, new)
        if new == s:
            break
        s = new
    s = s.replace("{", r"\{").replace("}", r"\}")
    # restore in reverse so outer groups re-embed their inner placeholders
    for i in range(len(placeholders) - 1, -1, -1):
        s = s.replace(f"\x00{i}\x00", placeholders[i])

    # 10. tidy whitespace
    s = re.sub(r'\s+', ' ', s).strip()
    return s


def split_inline_math(text: str) -> List[Token]:
    """Split an inline text run into text/math/code tokens.

    Math tokens carry the *unwrapped* formula body (no ``$`` delimiters).
    Text tokens have ``\\$`` unescaped to a literal ``$``. Code tokens are
    returned verbatim and are never scanned for math.
    """
    tokens: List[Token] = []

    # 1. peel off code spans verbatim
    segments: List[Tuple[str, bool]] = []  # (chunk, is_code)
    pos = 0
    for m in _CODE_SPAN_RE.finditer(text):
        if m.start() > pos:
            segments.append((text[pos:m.start()], False))
        segments.append((m.group(0), True))
        pos = m.end()
    if pos < len(text):
        segments.append((text[pos:], False))

    # 2. scan non-code segments for math
    for chunk, is_code in segments:
        if is_code:
            tokens.append(Token("code", chunk))
            continue
        tokens.extend(_scan_math(chunk))

    return tokens


def _scan_math(chunk: str) -> List[Token]:
    """Find math spans in a code-free chunk and build tokens.

    Handles delimited math ($…$, \\(…\\)) first, then scans the remaining
    text for *bare* un-delimited formulas.
    """
    out: List[Token] = []
    # Merge candidates from both patterns, earliest match first.
    matches: List[re.Match] = []
    for m in _INLINE_DOLLAR_RE.finditer(chunk):
        matches.append(m)
    for m in _INLINE_PAREN_RE.finditer(chunk):
        matches.append(m)
    matches.sort(key=lambda m: (m.start(), m.end()))

    # Drop overlaps (keep the earlier match).
    pruned: List[re.Match] = []
    last_end = -1
    for m in matches:
        if m.start() >= last_end:
            pruned.append(m)
            last_end = m.end()

    pos = 0
    for m in pruned:
        if m.start() > pos:
            out.extend(_scan_bare(chunk[pos:m.start()], pos))
        out.append(Token("math", m.group("body")))
        pos = m.end()
    if pos < len(chunk):
        out.extend(_scan_bare(chunk[pos:], pos))
    return out


def _scan_bare(text: str, offset: int) -> List[Token]:
    """Split a plain-text run into text/bare-math tokens."""
    out: List[Token] = []
    pos = 0
    for s, e in find_bare_math(text):
        if s > pos:
            out.append(Token("text", _unescape(text[pos:s])))
        out.append(Token("bare_math", text[s:e]))
        pos = e
    if pos < len(text):
        out.append(Token("text", _unescape(text[pos:])))
    if not out and text:
        out.append(Token("text", _unescape(text)))
    return out


def _unescape(text: str) -> str:
    """Turn escaped literal dollars into plain dollars in text runs."""
    return text.replace(r"\$", "$")


# ── markdown table rows ───────────────────────────────────────────────────
def split_table_row(line: str) -> List[str]:
    """Split a pipe-table row into raw cells.

    Pipes *inside* math spans (``$|x|$``) are part of the formula; escaped
    pipes (``\\|``) are literal text. Only bare pipes outside math/code split
    cells. Cell content keeps its ``\\|`` escapes so math can distinguish
    ``\\|`` (norm glyph) from a literal pipe — callers unescape text tokens.
    """
    inner = line.strip()
    if inner.startswith("|"):
        inner = inner[1:]
    if inner.endswith("|") and not inner.endswith("\\|"):
        inner = inner[:-1]

    cells: List[str] = []
    buf: List[str] = []
    i = 0
    n = len(inner)
    while i < n:
        ch = inner[i]
        # escaped pipe: keep, don't split
        if ch == "\\" and i + 1 < n and inner[i + 1] == "|":
            buf.append("\\|")
            i += 2
            continue
        # math / code spans: consume verbatim so inner pipes never split
        if ch == "$":
            m = _INLINE_DOLLAR_RE.match(inner, i)
            if m:
                buf.append(m.group(0))
                i = m.end()
                continue
        if ch == "\\" and i + 1 < n and inner[i + 1] == "(":
            m = _INLINE_PAREN_RE.match(inner, i)
            if m:
                buf.append(m.group(0))
                i = m.end()
                continue
        if ch == "`":
            m = _CODE_SPAN_RE.match(inner, i)
            if m:
                buf.append(m.group(0))
                i = m.end()
                continue
        # bare pipe outside protected spans: cell boundary
        if ch == "|":
            cells.append("".join(buf))
            buf = []
            i += 1
            continue
        buf.append(ch)
        i += 1
    cells.append("".join(buf))
    return cells


# ── display math ──────────────────────────────────────────────────────────
def display_blocks(lines: List[str]) -> List[Tuple[int, int]]:
    """Locate display-math blocks as (start_idx, end_idx) line ranges.

    Handles ``$$...$$`` on one line and spanning many lines. A closing
    ``$$`` on the same line as content ends a single-line block. Returns
    inclusive line-index pairs; ``start == end`` for one-liners.
    """
    blocks: List[Tuple[int, int]] = []
    i = 0
    n = len(lines)
    while i < n:
        stripped = lines[i].strip()
        if stripped.startswith("$$"):
            rest = stripped[2:].strip()
            if rest.endswith("$$") and len(rest) >= 2:
                if rest[:-2].strip():
                    blocks.append((i, i))
                i += 1
                continue
            # open a multiline block
            j = i + 1
            closed = False
            while j < n:
                s = lines[j].strip()
                if s.endswith("$$"):
                    blocks.append((i, j))
                    closed = True
                    break
                j += 1
            if closed:
                i = j + 1
            else:
                i += 1  # unterminated: ignore
            continue
        if not stripped.startswith(("|", "#", "-", "*", ">")):
            # display opener after prose on the same line ("and $$ … $$")
            idx = stripped.find("$$")
            if idx > 0 and (not stripped[idx - 1].isalnum()
                            or stripped[idx - 1].isspace()):
                rest = stripped[idx + 2:].strip()
                if rest.endswith("$$") and "$$" not in rest[:-2] \
                        and rest[:-2].strip():
                    blocks.append((i, i))
                    i += 1
                    continue
                if not rest:
                    j = i + 1
                    closed = False
                    while j < n:
                        s = lines[j].strip()
                        if s.endswith("$$"):
                            blocks.append((i, j))
                            closed = True
                            break
                        j += 1
                    if closed:
                        i = j + 1
                    else:
                        i += 1  # unterminated: ignore
                    continue
        i += 1
    return blocks


def count_display_math(text: str) -> int:
    """Count display equations in raw text, ignoring fenced code blocks."""
    lines = text.splitlines()
    kept: List[str] = []
    in_fence = False
    marker = ""
    for line in lines:
        if is_fence_line(line):
            if not in_fence:
                in_fence = True
                marker = fence_marker(line)[0] * 3  # ` or ~
                continue
            elif fence_marker(line).startswith(marker[0]):
                in_fence = False
                continue
        if not in_fence:
            kept.append(line)
    return len(display_blocks(kept))
