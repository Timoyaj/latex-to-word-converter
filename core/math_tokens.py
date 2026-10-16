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
    """Find math spans in a code-free chunk and build tokens."""
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
            out.append(Token("text", _unescape(chunk[pos:m.start()])))
        out.append(Token("math", m.group("body")))
        pos = m.end()
    if pos < len(chunk):
        out.append(Token("text", _unescape(chunk[pos:])))
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
