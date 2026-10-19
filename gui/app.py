#!/usr/bin/env python3
"""
LaTeX → Word Converter — GUI Application
==========================================
A polished tkinter interface for converting Markdown/LaTeX documents
into professional Microsoft Word (.docx) files with native OMML equations
and automatic equation numbering.

UI/UX highlights:
  * menu bar with keyboard shortcuts (Ctrl+O/S/F, F3, Ctrl+Return …)
  * code editor with line-number gutter, current-line highlight and
    Catppuccin syntax highlighting driven by the shared math tokenizer
  * live preview on a white "paper" panel with clickable formulas
  * find & replace bar (Ctrl+F) with match count
  * flat icon toolbar with hover states and tooltips
  * zoom (Ctrl+= / Ctrl+- / Ctrl+0), word-wrap and preview toggles
  * placeholder hints, empty-state guidance and collapsible status log
  * live status bar: caret position, word count, display/inline equation counts
  * options persisted to ~/.latex-to-word-converter.json (geometry + settings)
  * non-blocking conversion with a status log and optional auto-open
"""

import sys
import os
import re
import json
import subprocess
import threading
import queue
from pathlib import Path
from tkinter import (
    Tk, Toplevel, Text, StringVar, BooleanVar,
    filedialog, messagebox,
    END, DISABLED, NORMAL, WORD, HORIZONTAL, BOTH, LEFT, RIGHT, X, Y, TOP, BOTTOM,
    NW, SE, E, W, SUNKEN, FLAT, VERTICAL,
)
import tkinter
import tkinter as tk
import tkinter.ttk as ttk

# Ensure project root is importable
_HERE = os.path.dirname(os.path.abspath(__file__))
_PROJECT_ROOT = os.path.dirname(_HERE)
if _PROJECT_ROOT not in sys.path:
    sys.path.insert(0, _PROJECT_ROOT)

from core.converter import LatexToWordConverter
from core.math_tokens import (
    count_display_math, split_inline_math, is_fence_line, fence_marker,
    INLINE_DOLLAR_RE, INLINE_PAREN_RE, find_bare_math_in_line,
    normalize_bare_math,
)
from gui.preview_render import render_document

APP_NAME = "LaTeX → Word Converter"
APP_VERSION = "1.1"
_SETTINGS_PATH = Path.home() / ".latex-to-word-converter.json"


def open_path(path: str):
    """Cross-platform 'open in file manager / default app'."""
    if sys.platform == "win32":
        os.startfile(path)  # noqa: S606 — documented Windows shell API
    elif sys.platform == "darwin":
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])


# ── design tokens ─────────────────────────────────────────────────
# Light chrome (app panels / cards)
_BG          = "#f3f4f8"   # app background
_SURFACE     = "#ffffff"   # cards / paper
_BORDER      = "#e2e6ef"   # hairlines
_INK         = "#1e293b"   # primary text
_MUTED       = "#64748b"   # secondary text
_FAINT       = "#94a3b8"   # tertiary text
_ACCENT      = "#4f46e5"   # indigo accent
_ACCENT_HOV  = "#6366f1"
_ACCENT_SOFT = "#eef2ff"   # tinted hover / chips
_OK          = "#059669"
_WARN        = "#b45309"
_ERR         = "#dc2626"
_LOG_BG      = "#f1f5f9"

# Dark editor (Catppuccin Mocha)
_EDITOR_BG   = "#1e1e2e"
_EDITOR_FG   = "#cdd6f4"
_GUTTER_BG   = "#181825"
_GUTTER_FG   = "#6c7086"
_LINE_HL     = "#2a2a3d"
_MATH_FG     = "#f9e2af"   # inline $…$
_DISPLAY_FG  = "#a6e3a1"   # display $$…$$

_UI_FONT     = ("Segoe UI", 10)
_UI_BOLD     = ("Segoe UI Semibold", 10)
_CAPTION     = ("Segoe UI Semibold", 8)
_MONO        = "Consolas"
_SERIF       = "Georgia"
_BASE_SIZE   = 11          # editor / preview base font size (zoomable)


class Tooltip:
    """Small delayed tooltip for any widget."""

    def __init__(self, widget, text: str, delay: int = 550):
        self.widget = widget
        self.text = text
        self.delay = delay
        self._job = None
        self._tip = None
        widget.bind("<Enter>", self._schedule, add="+")
        widget.bind("<Leave>", self._hide, add="+")
        widget.bind("<ButtonPress>", self._hide, add="+")
        widget.bind("<Destroy>", lambda e: self._hide(), add="+")

    def _schedule(self, _e=None):
        self._cancel()
        self._job = self.widget.after(self.delay, self._show)

    def _cancel(self):
        if self._job:
            try:
                self.widget.after_cancel(self._job)
            except Exception:
                pass
            self._job = None

    def _show(self):
        if self._tip or not self.text:
            return
        x = self.widget.winfo_rootx() + 8
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 6
        self._tip = tw = Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        tk.Label(
            tw, text=self.text, justify=LEFT, anchor=W,
            background="#111827", foreground="#f9fafb",
            font=("Segoe UI", 9), padx=8, pady=4,
        ).pack()
        tw.lift()

    def _hide(self, _e=None):
        self._cancel()
        if self._tip:
            try:
                self._tip.destroy()
            except Exception:
                pass
            self._tip = None


# ══════════════════════════════════════════════════════════════════
# Main application class
# ══════════════════════════════════════════════════════════════════
class ConverterApp:
    """Tkinter GUI wrapping the LatexToWordConverter engine."""

    # ── templates ─────────────────────────────────────────────────
    TEMPLATES = {
        "Fraction": r"$$\frac{a}{b}$$",
        "Summation": r"$$\sum_{i=1}^{n} x_i$$",
        "Integral": r"$$\int_{a}^{b} f(x)\, dx$$",
        "Matrix (2×2)": r"$$\begin{pmatrix} a & b \\ c & d \end{pmatrix}$$",
        "Square root": r"$$\sqrt{a^2 + b^2}$$",
        "Greek inline": r"$\alpha, \beta, \gamma, \theta$",
        "Superscript / subscript": r"$x_{i}^{2} + y_{j}^{3}$",
    }

    def __init__(self, root: Tk):
        self.root = root
        self.converter = LatexToWordConverter()
        self.msg_queue: queue.Queue = queue.Queue()
        self._last_output: str = ""
        self._current_file: str = ""
        self._busy: bool = False

        self.settings = self._load_settings()
        try:
            zoom = int(self.settings.get("zoom", 0) or 0)
        except (TypeError, ValueError):
            zoom = 0
        self._zoom = max(-3, min(10, zoom))
        self._wrap_enabled = bool(self.settings.get("wrap", False))

        self._configure_root()
        self._build_ui()
        self._bind_shortcuts()
        self._restore_settings()
        self._poll_queue()

    # ------------------------------------------------------------------
    # Settings persistence
    # ------------------------------------------------------------------
    @staticmethod
    def _load_settings() -> dict:
        try:
            with open(_SETTINGS_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
                return data if isinstance(data, dict) else {}
        except Exception:
            return {}

    def _save_settings(self):
        try:
            self.settings.update({
                "geometry": self.root.winfo_geometry(),
                "autonumber": bool(self.autonumber_var.get()),
                "eq_format": self.eq_fmt_var.get(),
                "autopen": bool(self.autopen_var.get()),
                "zoom": int(self._zoom),
                "wrap": bool(self._wrap_enabled),
                "last_open_dir": self._last_open_dir,
                "last_output_dir": self._last_output_dir,
            })
            with open(_SETTINGS_PATH, "w", encoding="utf-8") as f:
                json.dump(self.settings, f, indent=2)
        except Exception:
            pass  # never block shutdown on settings I/O

    # ------------------------------------------------------------------
    # Root window configuration
    # ------------------------------------------------------------------
    def _configure_root(self):
        self.root.title(APP_NAME)
        self.root.geometry("1060x800")
        self.root.minsize(860, 640)
        self.root.configure(bg=_BG)
        self.root.protocol("WM_DELETE_WINDOW", self._quit)

        self._apply_windows_dark_titlebar()

        style = ttk.Style()
        try:
            style.theme_use("vista")
        except Exception:
            try:
                style.theme_use("clam")
            except Exception:
                pass

        # custom styles
        style.configure("TLabel", font=_UI_FONT, background=_BG,
                        foreground=_INK)
        style.configure("TCheckbutton", font=_UI_FONT, background=_BG,
                        foreground=_INK)
        style.map("TCheckbutton", background=[("active", _BG)])
        style.configure("Subtitle.TLabel", font=_UI_BOLD,
                        background=_BG, foreground=_INK)
        style.configure("Chip.TLabel", font=("Segoe UI Semibold", 9),
                        background=_ACCENT_SOFT, foreground=_ACCENT,
                        padding=(10, 3))
        style.configure("TLabelframe", background=_BG, relief=FLAT,
                        borderwidth=0)
        style.configure("TLabelframe.Label", font=_CAPTION,
                        background=_BG, foreground=_MUTED)
        style.configure("Status.TLabel", font=("Segoe UI", 9),
                        background=_BG, foreground=_MUTED)
        style.configure("Horizontal.TProgressbar", troughcolor=_BORDER,
                        background=_ACCENT, borderwidth=0)

    def _apply_windows_dark_titlebar(self):
        """Best-effort dark title bar on Windows 10 1809+ / 11 (ignores failures)."""
        try:
            import ctypes
            hwnd = ctypes.windll.user32.GetParent(self.root.winfo_id())
            for attr in (20, 19):  # DWMWA_USE_IMMERSIVE_DARK_MODE (11/10)
                res = ctypes.windll.dwmapi.DwmSetWindowAttribute(
                    hwnd, attr, ctypes.byref(ctypes.c_int(1)), 4)
                if res == 0:
                    break
        except Exception:
            pass

    # ------------------------------------------------------------------
    # Build the complete UI
    # ------------------------------------------------------------------
    def _build_ui(self):
        self._build_menubar()

        outer = ttk.Frame(self.root, padding=(14, 10, 14, 8))
        outer.pack(fill=BOTH, expand=True)

        # ── compact header ────────────────────────────────────────
        header = tk.Frame(outer, bg=_BG)
        header.pack(fill=X, pady=(0, 8))
        tk.Label(header, text="◆", font=("Segoe UI", 15, "bold"),
                 fg=_ACCENT, bg=_BG).pack(side=LEFT, padx=(0, 8))
        tk.Label(header, text=APP_NAME, font=("Segoe UI Semibold", 15),
                 fg=_INK, bg=_BG).pack(side=LEFT)
        ttk.Label(header, text=f"v{APP_VERSION}",
                  style="Chip.TLabel").pack(side=LEFT, padx=(10, 0))
        tk.Label(header, text="Native editable OMML equations for Word",
                 font=("Segoe UI", 9), fg=_MUTED,
                 bg=_BG).pack(side=LEFT, padx=(12, 0), pady=(3, 0))

        # ── toolbar (flat icon buttons on a card) ─────────────────
        tb = tk.Frame(outer, bg=_SURFACE, highlightbackground=_BORDER,
                      highlightthickness=1, bd=0)
        tb.pack(fill=X, pady=(0, 8))
        self._toolbar_button(tb, "📂", "Open", self._open_file,
                             "Open a Markdown / LaTeX file  (Ctrl+O)")
        self._toolbar_button(tb, "✨", "Sample", self._load_sample,
                             "Load the bundled sample document")
        self._toolbar_button(tb, "📋", "Paste", self._paste,
                             "Paste from the clipboard")
        self._toolbar_button(tb, "🔍", "Find", lambda: self.toggle_find_bar(True),
                             "Find & replace in the editor  (Ctrl+F)")
        self._toolbar_button(tb, "📐", "Wrap", self._wrap_formula,
                             "Wrap the selection as $…$ or $$…$$  (Ctrl+M)")
        self._toolbar_button(tb, "🩺", "Inspect", self._open_inspector,
                             "Inspect a formula and preview its conversion")
        self._toolbar_button(tb, "🗑", "Clear", self._clear,
                             "Clear the editor")

        tk.Frame(tb, width=1, bg=_BORDER).pack(side=LEFT, fill=Y, padx=8, pady=6)

        # template menu
        self._tpl_var = StringVar(value="Insert Template…")
        tpl_combo = ttk.Combobox(tb, textvariable=self._tpl_var,
                                 values=list(self.TEMPLATES.keys()),
                                 state="readonly", width=19)
        tpl_combo.pack(side=LEFT, padx=(0, 4), pady=6)
        tpl_combo.bind("<<ComboboxSelected>>", self._insert_template)
        Tooltip(tpl_combo, "Insert a ready-made equation template at the caret")

        # equation counter chip (right side)
        self._eq_count_var = StringVar(value="0 display equations detected")
        ttk.Label(tb, textvariable=self._eq_count_var,
                  style="Chip.TLabel").pack(side=RIGHT, padx=10)

        # ── main split: editor (left) | live preview (right) ─────
        paned = ttk.PanedWindow(outer, orient=HORIZONTAL)
        paned.pack(fill=BOTH, expand=True, pady=(0, 8))
        self._paned = paned

        left = tk.Frame(paned, bg=_BG)
        right = tk.Frame(paned, bg=_BG)
        paned.add(left, weight=3)
        paned.add(right, weight=2)
        self._right_pane = right
        self._preview_visible = True

        # ── find & replace bar (hidden until Ctrl+F) ──────────────
        self.find_frame = tk.Frame(left, bg=_SURFACE,
                                   highlightbackground=_BORDER,
                                   highlightthickness=1, bd=0)
        self._find_visible = False

        tk.Label(self.find_frame, text="Find", font=("Segoe UI", 9),
                 fg=_MUTED, bg=_SURFACE).pack(side=LEFT, padx=(10, 4), pady=6)
        self._find_var = StringVar()
        self._find_entry = ttk.Entry(self.find_frame, textvariable=self._find_var, width=24)
        self._find_entry.pack(side=LEFT, padx=2, pady=6)
        self._find_entry.bind("<Return>", lambda e: self._find_next())
        self._find_entry.bind("<Escape>", lambda e: self.toggle_find_bar(False))

        tk.Label(self.find_frame, text="Replace", font=("Segoe UI", 9),
                 fg=_MUTED, bg=_SURFACE).pack(side=LEFT, padx=(10, 4))
        self._replace_var = StringVar()
        self._replace_entry = ttk.Entry(self.find_frame, textvariable=self._replace_var, width=18)
        self._replace_entry.pack(side=LEFT, padx=2, pady=6)
        self._replace_entry.bind("<Return>", lambda e: self._replace_current())
        self._replace_entry.bind("<Escape>", lambda e: self.toggle_find_bar(False))

        self._flat_button(self.find_frame, "▲",
                          lambda: self._find_next(backwards=True),
                          "Previous match").pack(side=LEFT, padx=1)
        self._flat_button(self.find_frame, "▼", self._find_next,
                          "Next match").pack(side=LEFT, padx=1)
        self._flat_button(self.find_frame, "Replace",
                          self._replace_current, "Replace current match").pack(side=LEFT, padx=2)
        self._flat_button(self.find_frame, "All",
                          self._replace_all, "Replace all matches").pack(side=LEFT, padx=1)

        self._find_status_var = StringVar(value="")
        tk.Label(self.find_frame, textvariable=self._find_status_var,
                 font=("Segoe UI", 9), fg=_MUTED,
                 bg=_SURFACE).pack(side=LEFT, padx=10)
        self._flat_button(self.find_frame, "✕",
                          lambda: self.toggle_find_bar(False),
                          "Close find bar  (Esc)").pack(side=RIGHT, padx=6)

        # ── editor card (gutter + text + scrollbars) ──────────────
        ed_card = self._card(left)
        ed_card.pack(fill=BOTH, expand=True)
        self._ed_card = ed_card

        ed_head = tk.Frame(ed_card, bg=_SURFACE)
        ed_head.pack(fill=X)
        tk.Label(ed_head, text="EDITOR", font=_CAPTION, fg=_MUTED,
                 bg=_SURFACE).pack(side=LEFT, padx=(12, 8), pady=6)
        self._editor_hint = tk.Label(
            ed_head, text="Ctrl+M wrap  ·  Ctrl+F find  ·  F5 refresh",
            font=("Segoe UI", 8), fg=_FAINT, bg=_SURFACE)
        self._editor_hint.pack(side=RIGHT, padx=12)
        self._file_chip = tk.Label(
            ed_head, text="untitled.md", font=("Segoe UI", 9),
            fg=_MUTED, bg=_SURFACE)
        self._file_chip.pack(side=RIGHT, padx=8)
        tk.Frame(ed_card, height=1, bg=_BORDER).pack(fill=X)

        ed_body = tk.Frame(ed_card, bg=_EDITOR_BG)
        ed_body.pack(fill=BOTH, expand=True)

        self.gutter = Text(
            ed_body, width=4, padx=6, pady=6, takefocus=0,
            borderwidth=0, relief=FLAT, state=DISABLED, wrap="none",
            font=(_MONO, _BASE_SIZE), cursor="arrow",
            bg=_GUTTER_BG, fg=_GUTTER_FG,
            selectbackground=_GUTTER_BG, selectforeground=_GUTTER_FG,
        )
        self.gutter.pack(side=LEFT, fill=Y)

        v_scroll = ttk.Scrollbar(ed_body, orient=VERTICAL)
        h_scroll = ttk.Scrollbar(ed_card, orient=HORIZONTAL)
        self.v_scroll = v_scroll
        self._h_scroll = h_scroll

        self.editor = Text(
            ed_body, wrap="none", undo=True,
            font=(_MONO, _BASE_SIZE),
            bg=_EDITOR_BG, fg=_EDITOR_FG,
            insertbackground="#f5f5f5",
            selectbackground="#45475a",
            relief=FLAT, borderwidth=0,
            padx=10, pady=8,
            yscrollcommand=self._on_editor_scroll,
            xscrollcommand=h_scroll.set,
        )
        self.editor.pack(side=LEFT, fill=BOTH, expand=True)
        v_scroll.pack(side=RIGHT, fill=Y)
        v_scroll.config(command=self._on_scrollbar)
        h_scroll.pack(fill=X)
        h_scroll.config(command=self.editor.xview)

        # empty-state hint overlaid on the dark editor
        self._placeholder = tk.Label(
            self.editor,
            text="Paste Markdown or LaTeX here…\n\n"
                 "Tip: $inline$ and $$display$$ math become editable\n"
                 "Word equations. Select a formula and press Ctrl+M to\n"
                 "wrap it, or open a file with Ctrl+O.",
            font=(_MONO, 10), fg="#585b70", bg=_EDITOR_BG,
            justify=LEFT, anchor=NW, cursor="xterm")
        self._placeholder.bind("<Button-1>", lambda e: self.editor.focus_set())

        # syntax-highlight + decoration tags
        self.editor.tag_configure("inline_math",  foreground=_MATH_FG)
        self.editor.tag_configure("display_math", foreground=_DISPLAY_FG)
        self.editor.tag_configure("bare_math", foreground="#fab387",
                                  underline=True)
        self.editor.tag_configure("code_block", foreground="#9399b2",
                                  background="#181825")
        self.editor.tag_configure("current_line", background=_LINE_HL)
        self.editor.tag_configure("found", background="#45475a",
                                  foreground="#f9e2af")

        self.editor.bind("<<Modified>>", self._on_modified)
        self.editor.bind("<KeyRelease>", self._on_editor_change)
        self.editor.bind("<ButtonRelease-1>", self._on_editor_change)
        self.editor.bind("<Configure>", self._on_editor_change)

        # ── live conversion preview (paper card) ──────────────────
        pv_card = self._card(right)
        pv_card.pack(fill=BOTH, expand=True)
        self._pv_card = pv_card

        pv_head = tk.Frame(pv_card, bg=_SURFACE)
        pv_head.pack(fill=X)
        tk.Label(pv_head, text="PREVIEW", font=_CAPTION, fg=_MUTED,
                 bg=_SURFACE).pack(side=LEFT, padx=(12, 8), pady=6)
        tk.Label(pv_head, text="click a formula to jump to source",
                 font=("Segoe UI", 8), fg=_FAINT,
                 bg=_SURFACE).pack(side=RIGHT, padx=12)
        tk.Frame(pv_card, height=1, bg=_BORDER).pack(fill=X)

        pv_body = tk.Frame(pv_card, bg=_SURFACE)
        pv_body.pack(fill=BOTH, expand=True)

        self._pv_scroll = ttk.Scrollbar(pv_body, orient=VERTICAL)

        self.preview = Text(
            pv_body, wrap=WORD, state=DISABLED,
            font=(_SERIF, _BASE_SIZE), bg=_SURFACE, fg=_INK,
            relief=FLAT, borderwidth=0, padx=16, pady=12,
            cursor="arrow", spacing1=2, spacing3=6,
        )
        self.preview.pack(side=LEFT, fill=BOTH, expand=True)
        self._pv_scroll.pack(side=RIGHT, fill=Y)
        self._pv_scroll.config(command=self.preview.yview)
        self.preview.configure(yscrollcommand=self._on_preview_scroll)

        self._configure_preview_tags()

        self.preview.bind("<Button-1>", self._on_preview_click)
        self.preview.bind("<Double-Button-1>", lambda e: self._open_inspector())

        self._preview_srcmap: dict = {}
        self._preview_line_index: list = []
        self._preview_job = None
        self._sync_job = None
        self._pv_scroll_job = None
        self._scroll_syncing = False

        # ── options card ──────────────────────────────────────────
        opt = self._card(outer)
        opt.pack(fill=X, pady=(0, 8))

        row1 = tk.Frame(opt, bg=_SURFACE)
        row1.pack(fill=X, padx=12, pady=(8, 4))
        tk.Label(row1, text="OPTIONS", font=_CAPTION, fg=_MUTED,
                 bg=_SURFACE).pack(side=LEFT, padx=(0, 12))

        self.autonumber_var = BooleanVar(value=bool(self.settings.get("autonumber", True)))
        cb = tk.Checkbutton(row1, text="Auto-number equations",
                            variable=self.autonumber_var,
                            command=self._toggle_eq_format,
                            font=("Segoe UI", 9), fg=_INK, bg=_SURFACE,
                            activebackground=_SURFACE, activeforeground=_INK,
                            selectcolor=_SURFACE, bd=0, highlightthickness=0,
                            cursor="hand2")
        cb.pack(side=LEFT)
        Tooltip(cb, "Number display equations in an invisible 3-column table")

        tk.Label(row1, text="format", font=("Segoe UI", 9), fg=_MUTED,
                 bg=_SURFACE).pack(side=LEFT, padx=(12, 4))
        self.eq_fmt_var = StringVar(value=str(self.settings.get("eq_format", "({n})")))
        self.eq_fmt_entry = ttk.Entry(row1, textvariable=self.eq_fmt_var, width=9)
        self.eq_fmt_entry.pack(side=LEFT)
        Tooltip(self.eq_fmt_entry, "Numbering template — {n} is replaced by the equation number")

        self.autopen_var = BooleanVar(value=bool(self.settings.get("autopen", False)))
        cb2 = tk.Checkbutton(row1, text="Open document when done",
                             variable=self.autopen_var,
                             font=("Segoe UI", 9), fg=_INK, bg=_SURFACE,
                             activebackground=_SURFACE, activeforeground=_INK,
                             selectcolor=_SURFACE, bd=0, highlightthickness=0,
                             cursor="hand2")
        cb2.pack(side=LEFT, padx=(18, 0))
        Tooltip(cb2, "Launch the generated .docx in Word after conversion")

        row2 = tk.Frame(opt, bg=_SURFACE)
        row2.pack(fill=X, padx=12, pady=(0, 9))
        tk.Label(row2, text="Save to", font=("Segoe UI", 9), fg=_MUTED,
                 bg=_SURFACE).pack(side=LEFT, padx=(0, 8))
        self.output_var = StringVar()
        self.output_entry = ttk.Entry(row2, textvariable=self.output_var)
        self.output_entry.pack(side=LEFT, padx=(0, 6), fill=X, expand=True)
        self._flat_button(row2, "Browse…", self._browse_output,
                          "Choose the destination .docx file").pack(side=LEFT)

        # ── action bar ────────────────────────────────────────────
        act = tk.Frame(outer, bg=_BG)
        act.pack(fill=X, pady=(0, 6))

        self.convert_btn = tk.Button(
            act, text="⚡  Convert & Save    Ctrl+↵",
            font=("Segoe UI Semibold", 11), fg="white",
            bg=_ACCENT, activebackground=_ACCENT_HOV,
            activeforeground="white", disabledforeground="#e0e7ff",
            relief=FLAT, bd=0, padx=24, pady=9, cursor="hand2",
            highlightthickness=0, command=self._start_conversion,
        )
        self.convert_btn.pack(side=LEFT)
        self.convert_btn.bind(
            "<Enter>", lambda e: self.convert_btn.configure(bg=_ACCENT_HOV)
            if str(self.convert_btn["state"]) != DISABLED else None)
        self.convert_btn.bind(
            "<Leave>", lambda e: self.convert_btn.configure(bg=_ACCENT))
        Tooltip(self.convert_btn, "Convert the editor contents and write the .docx  (Ctrl+Return)")

        self.progress = ttk.Progressbar(act, mode="indeterminate", length=160)
        self.progress.pack(side=LEFT, padx=(16, 0), pady=8)

        # post-conversion buttons (hidden initially)
        self._post_frame = tk.Frame(act, bg=_BG)
        self._flat_button(self._post_frame, "📄  Open File",
                          self._open_output_file,
                          "Open the generated document").pack(side=LEFT, padx=(0, 4))
        self._flat_button(self._post_frame, "📁  Open Folder",
                          self._open_output_folder,
                          "Show the output folder in Explorer").pack(side=LEFT)

        # ── collapsible log card ──────────────────────────────────
        self._log_wrap = tk.Frame(outer, bg=_BG)
        self._log_wrap.pack(fill=X)

        log_head = tk.Frame(self._log_wrap, bg=_BG)
        log_head.pack(fill=X)
        self._log_toggle = tk.Button(
            log_head, text="▾  Status Log", font=("Segoe UI Semibold", 9),
            fg=_MUTED, bg=_BG, activebackground=_BG, activeforeground=_INK,
            relief=FLAT, bd=0, padx=2, pady=2, cursor="hand2",
            highlightthickness=0, command=self._toggle_log)
        self._log_toggle.pack(side=LEFT)
        self._log_visible = True

        self._log_frame = tk.Frame(self._log_wrap, bg=_LOG_BG,
                                   highlightbackground=_BORDER,
                                   highlightthickness=1, bd=0)
        self._log_frame.pack(fill=X, pady=(3, 0))

        self.log = Text(
            self._log_frame, height=4, wrap=WORD, state=DISABLED,
            font=(_MONO, 9),
            bg=_LOG_BG, fg="#334155",
            relief=FLAT, borderwidth=0,
            padx=10, pady=6,
        )
        self.log.pack(fill=X)
        self.log.tag_configure("ok", foreground=_OK)
        self.log.tag_configure("warn", foreground=_WARN)
        self.log.tag_configure("err", foreground=_ERR)
        self.log.tag_configure("dim", foreground=_MUTED)

        # ── status bar ────────────────────────────────────────────
        status = tk.Frame(outer, bg=_BG)
        status.pack(fill=X, pady=(6, 0))
        self._status_var = StringVar(value="Ready")
        ttk.Label(status, textvariable=self._status_var,
                  style="Status.TLabel").pack(side=LEFT)

        # zoom controls (right)
        self._zoom_var = StringVar(value="100%")
        self._flat_button(status, "+", self._zoom_in, "Zoom in  (Ctrl+=)").pack(side=RIGHT)
        ttk.Label(status, textvariable=self._zoom_var,
                  style="Status.TLabel", width=5,
                  anchor="center").pack(side=RIGHT)
        self._flat_button(status, "−", self._zoom_out, "Zoom out  (Ctrl+-)").pack(side=RIGHT)

        tk.Frame(status, width=1, height=14, bg=_BORDER).pack(
            side=RIGHT, fill=Y, padx=10, pady=2)
        self._math_stat_var = StringVar(value="0 display · 0 inline")
        ttk.Label(status, textvariable=self._math_stat_var,
                  style="Status.TLabel").pack(side=RIGHT)

        self._log("Ready. Paste or open a LaTeX/Markdown file to begin.")
        if self.converter.xsl_path:
            self._log(f"OMML stylesheet loaded: {self.converter.xsl_path}", "dim")
        else:
            self._log("⚠ MML2OMML.XSL not found — equations will render as plain text.",
                      "warn")
        self._update_placeholder()

    # ------------------------------------------------------------------
    # Small UI building blocks
    # ------------------------------------------------------------------
    @staticmethod
    def _card(parent) -> tk.Frame:
        """White surface card with a hairline border."""
        return tk.Frame(parent, bg=_SURFACE, highlightbackground=_BORDER,
                        highlightthickness=1, bd=0)

    def _toolbar_button(self, parent, icon, label, command, tip):
        btn = tk.Button(
            parent, text=f"{icon}  {label}", command=command,
            font=("Segoe UI", 9), fg=_INK, bg=_SURFACE,
            activebackground=_ACCENT_SOFT, activeforeground=_ACCENT,
            relief=FLAT, bd=0, padx=10, pady=6, cursor="hand2",
            highlightthickness=0, takefocus=0,
        )
        btn.pack(side=LEFT, padx=1, pady=4)
        btn.bind("<Enter>", lambda e: btn.configure(bg=_ACCENT_SOFT, fg=_ACCENT))
        btn.bind("<Leave>", lambda e: btn.configure(bg=_SURFACE, fg=_INK))
        Tooltip(btn, tip)
        return btn

    def _flat_button(self, parent, text, command, tip=None):
        btn = tk.Button(
            parent, text=text, command=command,
            font=("Segoe UI", 9), fg=_MUTED, bg=_SURFACE,
            activebackground=_ACCENT_SOFT, activeforeground=_ACCENT,
            relief=FLAT, bd=0, padx=8, pady=4, cursor="hand2",
            highlightthickness=0, takefocus=0,
        )
        btn.bind("<Enter>", lambda e: btn.configure(bg=_ACCENT_SOFT, fg=_ACCENT))
        btn.bind("<Leave>", lambda e: btn.configure(bg=_SURFACE, fg=_MUTED))
        if tip:
            Tooltip(btn, tip)
        return btn

    def _configure_preview_tags(self):
        z = getattr(self, "_zoom", 0)
        base = _BASE_SIZE + z
        self.preview.tag_configure("h1", font=(_SERIF, base + 4, "bold"),
                                   spacing1=12, spacing3=5)
        self.preview.tag_configure("h2", font=(_SERIF, base + 2, "bold"),
                                   spacing1=9, spacing3=3)
        self.preview.tag_configure("h3", font=(_SERIF, base + 1, "bold italic"),
                                   spacing1=7, spacing3=2)
        self.preview.tag_configure("h4", font=(_SERIF, base, "bold"),
                                   spacing1=6, spacing3=2)
        self.preview.tag_configure("math_run", foreground=_ACCENT)
        self.preview.tag_configure("bare_run", foreground="#c2410c",
                                   underline=True)
        self.preview.tag_configure("display_run", foreground=_ACCENT,
                                   justify="center", font=(_SERIF, base + 1),
                                   spacing1=9, spacing3=3)
        self.preview.tag_configure("eq_label", foreground=_FAINT,
                                   justify="right")
        self.preview.tag_configure("code_run", font=(_MONO, max(7, base - 1)),
                                   background=_LOG_BG, foreground="#475569")
        self.preview.tag_configure("quote_run", font=(_SERIF, base, "italic"),
                                   lmargin1=18, lmargin2=18,
                                   foreground="#475569")
        self.preview.tag_configure("table_run", font=(_SERIF, base),
                                   lmargin1=6, foreground=_INK)
        self.preview.tag_configure("rule_run", foreground="#cbd5e1")
        self.preview.tag_configure("formula_hot", background=_ACCENT_SOFT)

    def _update_placeholder(self, text: str = None):
        """Show the empty-state hint only while the editor is empty."""
        try:
            if text is None:
                text = self.editor.get("1.0", END)
            empty = not text.strip()
        except Exception:
            empty = False
        if empty:
            self._placeholder.place(in_=self.editor, x=12, y=10, anchor=NW)
            self._placeholder.lift()
        else:
            self._placeholder.place_forget()

    def _update_file_chip(self):
        name = Path(self._current_file).name if self._current_file else "untitled.md"
        try:
            self._file_chip.configure(text=name)
        except Exception:
            pass

    # ------------------------------------------------------------------
    # View controls: zoom, word wrap, preview pane, log
    # ------------------------------------------------------------------
    def _apply_zoom(self):
        z = self._zoom
        size = max(7, _BASE_SIZE + z)
        self.editor.configure(font=(_MONO, size))
        self.gutter.configure(font=(_MONO, size))
        self.preview.configure(font=(_SERIF, size))
        self._configure_preview_tags()
        try:
            self._zoom_var.set(f"{round(size / _BASE_SIZE * 100)}%")
        except Exception:
            pass

    def _zoom_in(self):
        if self._zoom < 10:
            self._zoom += 1
            self._apply_zoom()

    def _zoom_out(self):
        if self._zoom > -3:
            self._zoom -= 1
            self._apply_zoom()

    def _zoom_reset(self):
        self._zoom = 0
        self._apply_zoom()

    def _toggle_wrap(self, enabled=None):
        wrap = (not self._wrap_enabled) if enabled is None else bool(enabled)
        self._wrap_enabled = wrap
        if wrap:
            self.editor.configure(wrap=WORD)
            self._h_scroll.pack_forget()
        else:
            self.editor.configure(wrap="none")
            self._h_scroll.pack(fill=X)
        self._on_editor_change()

    def _toggle_preview(self, visible=None):
        show = (not self._preview_visible) if visible is None else bool(visible)
        if show == self._preview_visible:
            return
        self._preview_visible = show
        if show:
            self._paned.add(self._right_pane, weight=2)
        else:
            self._paned.forget(self._right_pane)
        self._on_editor_change()

    def _toggle_log(self):
        self._log_visible = not self._log_visible
        if self._log_visible:
            self._log_frame.pack(fill=X, pady=(3, 0))
            self._log_toggle.configure(text="▾  Status Log")
        else:
            self._log_frame.pack_forget()
            self._log_toggle.configure(text="▸  Status Log")

    def _load_sample(self):
        sample = Path(_PROJECT_ROOT) / "examples" / "sample_equations.md"
        try:
            content = sample.read_text(encoding="utf-8")
        except Exception as exc:
            messagebox.showerror("Sample Not Found",
                                 f"Could not read the sample document:\n{exc}")
            return
        self.editor.delete("1.0", END)
        self.editor.insert("1.0", content)
        self.editor.edit_reset()
        self._current_file = str(sample)
        self.output_var.set(str(Path.home() / "Documents" /
                                "sample_equations.docx"))
        self._on_editor_change()
        self._log(f"Loaded sample: {sample}", "dim")

    # ------------------------------------------------------------------
    # Menu bar
    # ------------------------------------------------------------------
    def _build_menubar(self):
        menubar = tkinter.Menu(self.root, tearoff=0, font=("Segoe UI", 10))

        file_menu = tkinter.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Open File…        Ctrl+O",
                              accelerator="Ctrl+O", command=self._open_file)
        file_menu.add_command(label="Load Sample Document",
                              command=self._load_sample)
        file_menu.add_command(label="Save Output As…   Ctrl+S",
                              accelerator="Ctrl+S", command=self._browse_output)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self._quit)
        menubar.add_cascade(label="File", menu=file_menu)

        edit_menu = tkinter.Menu(menubar, tearoff=0)
        edit_menu.add_command(label="Undo        Ctrl+Z",
                              command=lambda: self._safe_edit(self.editor.edit_undo))
        edit_menu.add_command(label="Redo        Ctrl+Y",
                              command=lambda: self._safe_edit(self.editor.edit_redo))
        edit_menu.add_separator()
        edit_menu.add_command(label="Cut", command=lambda: self._safe_edit(lambda: self.editor.event_generate("<<Cut>>")))
        edit_menu.add_command(label="Copy", command=lambda: self._safe_edit(lambda: self.editor.event_generate("<<Copy>>")))
        edit_menu.add_command(label="Paste", command=self._paste)
        edit_menu.add_separator()
        edit_menu.add_command(label="Find / Replace…   Ctrl+F",
                              accelerator="Ctrl+F",
                              command=self.toggle_find_bar)
        edit_menu.add_command(label="Find Next          F3",
                              accelerator="F3", command=self._find_next)
        edit_menu.add_command(label="Find Previous      Shift+F3",
                              accelerator="Shift+F3",
                              command=lambda: self._find_next(backwards=True))
        edit_menu.add_separator()
        edit_menu.add_command(label="Clear Editor", command=self._clear)
        menubar.add_cascade(label="Edit", menu=edit_menu)

        insert_menu = tkinter.Menu(menubar, tearoff=0)
        for name in self.TEMPLATES:
            insert_menu.add_command(
                label=name,
                command=lambda n=name: self._insert_template(name=n))
        menubar.add_cascade(label="Insert", menu=insert_menu)

        view_menu = tkinter.Menu(menubar, tearoff=0)
        view_menu.add_command(label="Word Wrap          Ctrl+Shift+W",
                              accelerator="Ctrl+Shift+W",
                              command=self._toggle_wrap)
        view_menu.add_command(label="Show Preview       F2",
                              accelerator="F2", command=self._toggle_preview)
        view_menu.add_separator()
        view_menu.add_command(label="Zoom In            Ctrl+=",
                              accelerator="Ctrl+=", command=self._zoom_in)
        view_menu.add_command(label="Zoom Out           Ctrl+-",
                              accelerator="Ctrl+-", command=self._zoom_out)
        view_menu.add_command(label="Reset Zoom         Ctrl+0",
                              accelerator="Ctrl+0", command=self._zoom_reset)
        menubar.add_cascade(label="View", menu=view_menu)

        help_menu = tkinter.Menu(menubar, tearoff=0)
        help_menu.add_command(label="User Guide         F1",
                              accelerator="F1", command=self._show_guide)
        help_menu.add_command(label="About", command=self._about)
        menubar.add_cascade(label="Help", menu=help_menu)

        self.root.config(menu=menubar)

    @staticmethod
    def _safe_edit(fn):
        try:
            fn()
        except Exception:
            pass

    def _show_guide(self, _event=None):
        """User Guide window: how every part of the app works."""
        win = Toplevel(self.root)
        win.title(f"{APP_NAME} — User Guide")
        win.transient(self.root)
        win.configure(bg=_SURFACE)
        win.geometry("720x640")
        win.grab_set()

        body = tk.Frame(win, bg=_SURFACE)
        body.pack(fill=BOTH, expand=True)

        bar = tk.Frame(body, bg=_SURFACE)
        bar.pack(fill=X, padx=20, pady=(16, 0))
        tk.Label(bar, text="◆", font=("Segoe UI", 14, "bold"),
                 fg=_ACCENT, bg=_SURFACE).pack(side=LEFT, padx=(0, 8))
        tk.Label(bar, text="User Guide", font=("Segoe UI Semibold", 15),
                 fg=_INK, bg=_SURFACE).pack(side=LEFT)

        text = Text(body, wrap=WORD, font=("Segoe UI", 10),
                    bg=_SURFACE, fg=_INK, relief=FLAT, borderwidth=0,
                    padx=22, pady=12, cursor="arrow", spacing1=1, spacing3=5)
        scroll = ttk.Scrollbar(body, orient=VERTICAL, command=text.yview)
        text.configure(yscrollcommand=scroll.set)
        scroll.pack(side=RIGHT, fill=Y)
        text.pack(fill=BOTH, expand=True)

        text.tag_configure("title", font=("Segoe UI Semibold", 13),
                           foreground=_INK, spacing1=2, spacing3=2)
        text.tag_configure("h", font=("Segoe UI Semibold", 11),
                           foreground=_ACCENT, spacing1=16, spacing3=4)
        text.tag_configure("body", font=("Segoe UI", 10), spacing3=6,
                           lmargin1=0, lmargin2=0)
        text.tag_configure("li", font=("Segoe UI", 10), spacing3=3,
                           lmargin1=18, lmargin2=18)
        text.tag_configure("key", font=(_MONO, 9), background=_ACCENT_SOFT,
                           foreground=_ACCENT)
        text.tag_configure("mono", font=(_MONO, 9), foreground=_MUTED)
        text.tag_configure("muted", font=("Segoe UI", 9), foreground=_MUTED,
                           spacing3=8)

        def h(s):
            text.insert(END, s + "\n", "h")

        def p(s):
            text.insert(END, s + "\n\n", "body")

        def li(label, s, label_tag="key"):
            text.insert(END, "•  ", "li")
            text.insert(END, label, label_tag)
            text.insert(END, "  —  " + s + "\n", "li")

        def mono(s):
            text.insert(END, s, "mono")

        text.insert(END, "LaTeX → Word Converter\n", "title")
        text.insert(END,
                    "Turn Markdown + LaTeX into a Microsoft Word (.docx) document "
                    "whose equations are native, editable OMML math — not pictures. "
                    "Equations can be auto-numbered, and bundled Word macros can "
                    "renumber and cross-reference them inside Word.\n\n", "body")

        h("1 · Getting started")
        p("Type or paste Markdown/LaTeX into the dark editor on the left, watch "
          "the live preview on the right, then press Convert & Save to write "
          "the .docx. Load the bundled sample with the Sample toolbar button "
          "to see every supported construct at once.")

        h("2 · The editor")
        li("Line-number gutter", "always in sync with the editor's scroll position.")
        li("Syntax colors", "inline $…$ math glows gold, display $$…$$ math green, "
           "bare formulas peach and underlined, code fences gray.")
        li("Current-line highlight", "the row with your caret is softly shaded.")
        li("Empty-state hint", "when the editor is empty, a dim hint shows how to "
           "start; it disappears as soon as you type.")
        li("File chip", "the header strip shows the loaded file name.")

        h("3 · Writing math")
        li("Inline math", "wrap LaTeX in single dollars: $E=mc^2$.")
        li("Display math", "use $$…$$ on one line, or open $$ on its own line and "
           "close with $$ after the formula. A $$ opener may follow prose on the "
           "same line (\"and $$\\n y \\n$$\").")
        li("Bracketed display", "\\[ … \\] works exactly like $$…$$.")
        li("Bare formulas", "informal math without any delimiters is detected — "
           "e.g. rate_2 <= t_1, E_2[ N_hat * theta_hat_DML | s_1 ] = sum_{k in "
           "s_1} w_1k Y_k. ASCII is normalized (N_hat → N̂, sum_ → Σ, <= → ≤, "
           "w_1k → w₁ₖ). Prose like max_size or snake_case_name is left alone.")
        p("Currency amounts ($5.00 … $10.00), escaped \\$ and code fences are "
          "never treated as math — detection follows Pandoc dollar-math rules.")

        h("4 · Live preview")
        p("The PREVIEW panel renders the document exactly as conversion will: "
          "headings, lists, quotes, tables, code fences, and Unicode-prettified "
          "math with numbering labels. It scrolls in sync with the editor in "
          "both directions, and scrolling the preview moves the editor too.")
        li("Click a formula", "jumps to and selects its exact source span in the editor.")
        li("Double-click a formula", "opens the Formula Inspector pre-filled with it.")
        li("F5", "forces a preview refresh (it normally refreshes itself ~120 ms after typing).")

        h("5 · Formula tools")
        li("Wrap  (Ctrl+M)", "select any text and wrap it as $…$ — or $$…$$ when the "
           "selection spans multiple lines. Ideal for turning a bare detection "
           "into a deliberate formula.")
        li("Inspect  (double-click / toolbar)", "edit a formula, see its normalized "
           "and pretty forms, and check whether it converts to a Word math zone "
           "or would fall back to plain text. Insert it as inline math, display "
           "math, or code — nothing is saved until Convert & Save.")

        h("6 · Templates")
        p("The Insert Template dropdown (and the Insert menu) inserts ready-made "
          "constructs at the caret: fractions, summations, integrals, matrices, "
          "square roots, greek sets, and sub/superscripts.")

        h("7 · Find & replace")
        li("Ctrl+F", "opens the find bar pre-filled with the current selection.")
        li("F3 / Shift+F3", "next / previous match (wraps around).")
        li("Replace / All", "replace the current match or every match (case-insensitive).")
        li("Esc", "closes the bar and clears the match highlight.")

        h("8 · Conversion options")
        li("Auto-number equations", "display equations get sequential numbers laid "
           "out in an invisible three-column table (margin, centered equation, tag).")
        li("Numbering format", "any template containing {n} — e.g. ({n}), Eq. {n}, [{n}].")
        li("Open document when done", "launches the .docx in Word automatically.")
        li("Save to", "destination .docx path; Browse… picks it. Open a .md file and "
           "the output path is suggested automatically.")

        h("9 · Converting")
        p("Convert & Save (Ctrl+Return) runs the conversion in the background — "
          "the button shows progress and the Status Log reports "
          "“Success! n/n equations converted”. Formulas that could not convert "
          "are listed with their line number and kept as plain text. After a "
          "successful run, Open File and Open Folder buttons appear next to the "
          "log. Log entries are color-coded: green success, amber warnings, "
          "red errors. The log can be collapsed via its caption.")

        h("10 · View controls")
        li("Ctrl+= / Ctrl+- / Ctrl+0", "zoom the editor and preview in, out, or reset "
           "(the current percent shows in the status bar).")
        li("Ctrl+Shift+W", "toggle word wrap in the editor (horizontal scrollbar "
           "appears when wrap is off).")
        li("F2", "show or hide the preview pane to focus on the editor.")

        h("11 · Keyboard shortcuts")
        shortcuts = [
            ("Ctrl+O", "Open file"), ("Ctrl+S", "Choose output file"),
            ("Ctrl+Return", "Convert & save"), ("Ctrl+F", "Find & replace"),
            ("F3 / Shift+F3", "Find next / previous"),
            ("Ctrl+M", "Wrap selection as formula"),
            ("Ctrl+= / Ctrl+-", "Zoom in / out"), ("Ctrl+0", "Reset zoom"),
            ("Ctrl+Shift+W", "Toggle word wrap"), ("F2", "Toggle preview pane"),
            ("F5", "Refresh preview"), ("F1", "This user guide"),
            ("Ctrl+Z / Ctrl+Y", "Undo / redo"),
        ]
        for keys, desc in shortcuts:
            li(keys, desc)

        h("12 · Settings, files & Word macros")
        p("Window geometry, numbering options, zoom, word-wrap and folder "
          "choices persist to ~/.latex-to-word-converter.json. The sample "
          "document lives at examples/sample_equations.md. Inside Word, the "
          "bundled macros/EquationAutoNumbering.bas module converts remaining "
          "$$…$$ text, renumbers equations and inserts live Eq. (n) cross-"
          "references — import it via Word's VBA editor (Alt+F11).")
        text.insert(END, "If the log warns that MML2OMML.XSL was not found, set the "
                    "LWX_MML2OMML_XSL environment variable to a stylesheet path — "
                    "otherwise equations fall back to plain text.\n", "muted")

        text.configure(state=DISABLED)
        text.focus_set()
        win.bind("<Escape>", lambda e: win.destroy())
        return win

    def _about(self):
        win = Toplevel(self.root)
        win.title(f"About {APP_NAME}")
        win.transient(self.root)
        win.configure(bg=_BG)
        win.resizable(False, False)
        win.grab_set()

        body = tk.Frame(win, bg=_BG, padx=24, pady=20)
        body.pack(fill=BOTH, expand=True)

        tk.Label(body, text="◆", font=("Segoe UI", 22, "bold"),
                 fg=_ACCENT, bg=_BG).pack(anchor=W)
        tk.Label(body, text=APP_NAME, font=("Segoe UI Semibold", 16),
                 fg=_INK, bg=_BG).pack(anchor=W, pady=(2, 0))
        tk.Label(body, text=f"Version {APP_VERSION}",
                 font=("Segoe UI", 9), fg=_MUTED, bg=_BG).pack(anchor=W)

        tk.Label(
            body,
            text=("Markdown/LaTeX → Word with native, editable OMML equations,\n"
                  "automatic equation numbering and Word VBA automation.\n\n"
                  "Math detection follows Pandoc dollar-math rules —\n"
                  "currency, \\$ escapes and code fences are never converted."),
            font=("Segoe UI", 10), fg=_INK, bg=_BG,
            justify=LEFT).pack(anchor=W, pady=(14, 12))

        shortcuts = [
            ("Ctrl+O", "Open file"), ("Ctrl+S", "Choose output"),
            ("Ctrl+Return", "Convert & save"), ("Ctrl+F", "Find & replace"),
            ("Ctrl+M", "Wrap selection as formula"),
            ("Ctrl+= / Ctrl+-", "Zoom in / out"), ("Ctrl+0", "Reset zoom"),
            ("Ctrl+Shift+W", "Toggle word wrap"), ("F2", "Toggle preview"),
            ("F5", "Refresh preview"), ("F1", "User guide"),
        ]
        grid = tk.Frame(body, bg=_BG)
        grid.pack(fill=X)
        for row, (keys, desc) in enumerate(shortcuts):
            tk.Label(grid, text=keys, font=(_MONO, 9), fg=_ACCENT,
                     bg=_BG).grid(row=row, column=0, sticky=W, padx=(0, 14), pady=1)
            tk.Label(grid, text=desc, font=("Segoe UI", 9), fg=_MUTED,
                     bg=_BG).grid(row=row, column=1, sticky=W, pady=1)

        tk.Button(body, text="Close", command=win.destroy,
                  font=("Segoe UI", 9), fg="white", bg=_ACCENT,
                  activebackground=_ACCENT_HOV, activeforeground="white",
                  relief=FLAT, bd=0, padx=18, pady=6, cursor="hand2",
                  highlightthickness=0).pack(anchor=E, pady=(16, 0))
        win.focus_set()
        return win

    def _quit(self):
        self._save_settings()
        self.root.destroy()

    # ------------------------------------------------------------------
    # Keyboard shortcuts
    # ------------------------------------------------------------------
    def _bind_shortcuts(self):
        r = self.root
        r.bind("<Control-o>", lambda e: self._open_file())
        r.bind("<Control-s>", lambda e: self._browse_output())
        r.bind("<Control-f>", lambda e: self.toggle_find_bar(True))
        r.bind("<Control-M>", self._wrap_formula)
        r.bind("<Control-m>", self._wrap_formula)
        r.bind("<F3>", lambda e: self._find_next())
        r.bind("<Shift-F3>", lambda e: self._find_next(backwards=True))
        r.bind("<Control-Return>", lambda e: self._start_conversion())
        r.bind("<F5>", lambda e: self._refresh_preview())
        r.bind("<F2>", lambda e: self._toggle_preview())
        r.bind("<Control-plus>", lambda e: self._zoom_in())
        r.bind("<Control-equal>", lambda e: self._zoom_in())
        r.bind("<Control-minus>", lambda e: self._zoom_out())
        r.bind("<Control-0>", lambda e: self._zoom_reset())
        r.bind("<Control-Shift-W>", lambda e: self._toggle_wrap())
        r.bind("<Control-Shift-w>", lambda e: self._toggle_wrap())
        r.bind("<F1>", lambda e: self._show_guide())

    # ------------------------------------------------------------------
    # Restore persisted settings
    # ------------------------------------------------------------------
    def _restore_settings(self):
        self._last_open_dir = str(self.settings.get("last_open_dir", ""))
        self._last_output_dir = str(self.settings.get("last_output_dir", ""))
        geom = self.settings.get("geometry")
        if geom:
            try:
                self.root.geometry(geom)
            except Exception:
                pass
        self._toggle_eq_format()
        self._apply_zoom()
        if self._wrap_enabled:
            self._toggle_wrap(True)
        self._update_file_chip()
        self._on_editor_change()

    # ------------------------------------------------------------------
    # Toolbar actions
    # ------------------------------------------------------------------
    def _open_file(self):
        initial = self._last_open_dir if os.path.isdir(self._last_open_dir) else None
        path = filedialog.askopenfilename(
            title="Open Markdown / LaTeX file",
            initialdir=initial,
            filetypes=[
                ("Markdown files", "*.md"),
                ("LaTeX files", "*.tex"),
                ("Text files", "*.txt"),
                ("All files", "*.*"),
            ])
        if path:
            try:
                with open(path, "r", encoding="utf-8") as f:
                    content = f.read()
                self.editor.delete("1.0", END)
                self.editor.insert("1.0", content)
                self.editor.edit_reset()
                self._current_file = path
                self._last_open_dir = str(Path(path).parent)
                self._on_editor_change()
                self._update_file_chip()
                # suggest output name
                stem = Path(path).stem
                suggested = str(Path(path).parent / f"{stem}.docx")
                self.output_var.set(suggested)
                self._log(f"Loaded: {path}", "dim")
            except Exception as exc:
                messagebox.showerror("Error", f"Could not read file:\n{exc}")

    def _paste(self):
        try:
            text = self.root.clipboard_get()
            self.editor.insert("insert", text)
            self._on_editor_change()
        except Exception:
            pass

    def _clear(self):
        content = self.editor.get("1.0", END).strip()
        if content:
            if not messagebox.askyesno(
                    "Clear Editor",
                    "Discard the current editor contents?\n"
                    "This cannot be undone from the menu, but Ctrl+Z may still work."):
                return
        self.editor.delete("1.0", END)
        self.editor.edit_reset()
        self._current_file = ""
        self._on_editor_change()
        self._update_file_chip()

    def _insert_template(self, _event=None, name=None):
        name = name or self._tpl_var.get()
        tpl = self.TEMPLATES.get(name)
        if tpl:
            self.editor.insert("insert", "\n" + tpl + "\n")
            self.editor.focus_set()
            self._on_editor_change()
        self._tpl_var.set("Insert Template…")

    def _browse_output(self):
        initial = self._last_output_dir if os.path.isdir(self._last_output_dir) else None
        path = filedialog.asksaveasfilename(
            title="Save Word Document",
            initialdir=initial,
            defaultextension=".docx",
            filetypes=[("Word Document", "*.docx")])
        if path:
            self.output_var.set(path)
            self._last_output_dir = str(Path(path).parent)

    def _toggle_eq_format(self):
        if self.autonumber_var.get():
            self.eq_fmt_entry.configure(state=NORMAL)
        else:
            self.eq_fmt_entry.configure(state=DISABLED)

    # ------------------------------------------------------------------
    # Find & replace
    # ------------------------------------------------------------------
    def toggle_find_bar(self, show=None):
        if show is None:
            show = not self._find_visible
        self._find_visible = bool(show)
        if self._find_visible:
            self.find_frame.pack(fill=X, pady=(0, 6), before=self._ed_card)
            # pre-fill with selection
            try:
                sel = self.editor.get("sel.first", "sel.last")
                if sel and "\n" not in sel:
                    self._find_var.set(sel)
            except Exception:
                pass
            self._find_entry.focus_set()
        else:
            self.find_frame.pack_forget()
            self.editor.tag_remove("found", "1.0", END)
            self.editor.focus_set()

    def _find_next(self, backwards=False):
        needle = self._find_var.get()
        if not needle:
            return
        self.editor.tag_remove("found", "1.0", END)
        start = self.editor.index("insert")
        if not backwards:
            start = self.editor.index(f"{start}+1c")
        pos = self.editor.search(needle, start, stopindex=END,
                                 backwards=backwards, nocase=True)
        if not pos:
            # wrap around
            pos = self.editor.search(needle, END if backwards else "1.0",
                                     stopindex="1.0" if backwards else END,
                                     backwards=backwards, nocase=True)
        if pos:
            end = f"{pos}+{len(needle)}c"
            self.editor.tag_add("found", pos, end)
            self.editor.mark_set("insert", end if not backwards else pos)
            self.editor.see(pos)
            self._find_status_var.set("match found")
        else:
            self._find_status_var.set("no matches")

    def _replace_current(self):
        needle = self._find_var.get()
        if not needle:
            return
        try:
            sel = self.editor.get("sel.first", "sel.last")
        except Exception:
            sel = ""
        if sel.lower() == needle.lower() and sel:
            self.editor.delete("sel.first", "sel.last")
            self.editor.insert("insert", self._replace_var.get())
            self._on_editor_change()
        self._find_next()

    def _replace_all(self):
        needle = self._find_var.get()
        if not needle:
            return
        content = self.editor.get("1.0", END)
        pattern = re.compile(re.escape(needle), re.IGNORECASE)
        new_content, n = pattern.subn(self._replace_var.get(), content)
        if n:
            self.editor.delete("1.0", END)
            self.editor.insert("1.0", new_content)
            self._on_editor_change()
        self._find_status_var.set(f"replaced {n} occurrence{'s' if n != 1 else ''}")

    # ------------------------------------------------------------------
    # Editor decoration: highlighting, gutter, counter, status
    # ------------------------------------------------------------------
    def _on_modified(self, _event=None):
        # keep the modified flag cleared so <<Modified>> keeps firing
        if self.editor.edit_modified():
            self._on_editor_change()
        self.editor.edit_modified(False)

    def _on_editor_scroll(self, *args):
        self.v_scroll.set(*args)
        try:
            self.gutter.yview_moveto(args[0])
        except Exception:
            pass
        self._schedule_scroll_sync()

    def _on_scrollbar(self, *args):
        self.editor.yview(*args)
        self.gutter.yview_moveto(self.editor.yview()[0])

    # ── two-way editor ↔ preview scroll sync ──────────────────────
    def _schedule_scroll_sync(self):
        if self._sync_job is not None:
            try:
                self.root.after_cancel(self._sync_job)
            except Exception:
                pass
        self._sync_job = self.root.after(50, self._sync_preview_to_editor)

    def _on_preview_scroll(self, first, last):
        try:
            self._pv_scroll.set(first, last)
        except Exception:
            pass
        if self._scroll_syncing:
            return
        if self._pv_scroll_job is not None:
            try:
                self.root.after_cancel(self._pv_scroll_job)
            except Exception:
                pass
        self._pv_scroll_job = self.root.after(80, self._sync_editor_to_preview)

    def _sync_editor_to_preview(self):
        """Scroll the editor to the source line nearest the preview's top."""
        self._pv_scroll_job = None
        if not self._preview_line_index:
            return
        try:
            top = float(self.preview.index("@2,2"))
        except Exception:
            return
        best = None
        for line, pidx in self._preview_line_index:
            if float(pidx) <= top:
                best = line
            else:
                break
        if best is None:
            return
        try:
            n = max(1, int(self.editor.index("end-1c").split(".")[0]))
            self.editor.yview_moveto(min(1.0, max(0.0, (best - 1) / n)))
        except Exception:
            pass

    def _on_editor_change(self, _event=None):
        text = self.editor.get("1.0", END)
        # remove old tags (keep found/current_line handling below)
        for tag in ("inline_math", "display_math", "code_block", "current_line"):
            self.editor.tag_remove(tag, "1.0", END)

        lines = text.split("\n")
        line_starts = [0]
        for ln in lines[:-1]:
            line_starts.append(line_starts[-1] + len(ln) + 1)

        def tag_range(tag, start_char, end_char):
            self.editor.tag_add(tag, f"1.0+{start_char}c", f"1.0+{end_char}c")

        # ── code fences: tag and exclude from math scanning ──────
        in_fence = False
        fence_open = ""
        fence_line_idx = set()
        for idx, ln in enumerate(lines):
            if is_fence_line(ln):
                marker = fence_marker(ln)[0] * 3
                if not in_fence:
                    in_fence = True
                    fence_open = marker
                    fence_line_idx.add(idx)
                elif marker[0] == fence_open[0]:
                    in_fence = False
                    fence_line_idx.add(idx)
                continue
            if in_fence:
                fence_line_idx.add(idx)
        for idx in sorted(fence_line_idx):
            s = line_starts[idx]
            tag_range("code_block", s, s + len(lines[idx]))

        # ── display math $$…$$ (single- and multi-line) ──────────
        in_block = False
        block_start = 0
        for idx, ln in enumerate(lines):
            if idx in fence_line_idx:
                continue
            s = ln.strip()
            if not in_block and s.startswith("$$"):
                rest = s[2:].strip()
                if rest.endswith("$$") and len(rest) >= 2:
                    if rest[:-2].strip():
                        tag_range("display_math",
                                  line_starts[idx], line_starts[idx] + len(ln))
                else:
                    in_block = True
                    block_start = idx
            elif in_block and s.endswith("$$"):
                in_block = False
                tag_range("display_math",
                          line_starts[block_start],
                          line_starts[idx] + len(ln))
        for idx, ln in enumerate(lines):
            if idx in fence_line_idx:
                continue
            s = ln.strip()
            if s.startswith("\\[") and s.endswith("\\]") and len(s) > 3:
                tag_range("display_math",
                          line_starts[idx], line_starts[idx] + len(ln))

        # ── inline math via the shared tokenizer regexes ──────────
        for idx, ln in enumerate(lines):
            if idx in fence_line_idx:
                continue
            pos = line_starts[idx]
            for m in INLINE_DOLLAR_RE.finditer(ln):
                tag_range("inline_math", pos + m.start(), pos + m.end())
            for m in INLINE_PAREN_RE.finditer(ln):
                tag_range("inline_math", pos + m.start(), pos + m.end())
            # bare (un-delimited) formulas — underlined peach
            for s, e in find_bare_math_in_line(ln):
                tag_range("bare_math", pos + s, pos + e)

        # ── current-line highlight (under the other tags) ─────────
        try:
            cur = self.editor.index("insert").split(".")[0]
            self.editor.tag_add("current_line", f"{cur}.0", f"{cur}.0+1line")
        except Exception:
            pass

        # ── gutter numbers ────────────────────────────────────────
        n_lines = max(1, len(lines) - (1 if lines and lines[-1] == "" else 0))
        gutter_text = "\n".join(str(i) for i in range(1, n_lines + 1))
        self.gutter.configure(state=NORMAL)
        self.gutter.delete("1.0", END)
        self.gutter.insert("1.0", gutter_text)
        self.gutter.configure(state=DISABLED)
        self.gutter.yview_moveto(self.editor.yview()[0])

        # ── counter + status bar ──────────────────────────────────
        count = count_display_math(text)
        self._eq_count_var.set(
            f"{count} display equation{'s' if count != 1 else ''} detected")

        inline = 0
        for idx, ln in enumerate(lines):
            if idx in fence_line_idx:
                continue
            inline += sum(1 for t in split_inline_math(ln)
                          if t.kind in ("math", "bare_math"))
        self._math_stat_var.set(f"{count} display · {inline} inline")

        words = len(re.findall(r"\S+", text.rstrip("\n")))
        chars = len(text.rstrip("\n"))
        self._update_status(words, chars)
        self._update_placeholder(text)
        self._update_file_chip()
        self._schedule_preview_refresh()

    def _update_status(self, words=None, chars=None):
        try:
            row, col = self.editor.index("insert").split(".")
        except Exception:
            row, col = "1", "0"
        name = Path(self._current_file).name if self._current_file else "untitled"
        if words is None:
            text = self.editor.get("1.0", END).rstrip("\n")
            words = len(re.findall(r"\S+", text))
            chars = len(text)
        self._status_var.set(
            f"{name}   ·   Ln {row}, Col {int(col) + 1}   ·   "
            f"{words} words, {chars} chars")

    # ------------------------------------------------------------------
    # Live preview pane
    # ------------------------------------------------------------------
    def _schedule_preview_refresh(self):
        if self._preview_job is not None:
            try:
                self.root.after_cancel(self._preview_job)
            except Exception:
                pass
        self._preview_job = self.root.after(120, self._refresh_preview)

    def _refresh_preview(self):
        self._preview_job = None
        pv = self.preview
        pv.configure(state=NORMAL)
        pv.delete("1.0", END)
        self._preview_srcmap = {}
        self._preview_line_index = []

        text = self.editor.get("1.0", END)
        try:
            blocks = render_document(
                text,
                number_equations=self.autonumber_var.get(),
                eq_format=self.eq_fmt_var.get() or "({n})")
        except Exception:
            blocks = []

        for b in blocks:
            kind = b["kind"]
            self._preview_line_index.append(
                (b.get("line", 1), pv.index(END + "-1c")))
            if kind == "heading":
                for r in b["runs"]:
                    self._insert_preview_run(r, f"h{b['level']}")
                pv.insert(END, "\n")
            elif kind == "display":
                tag = self._preview_tag(b.get("src"))
                pv.insert(END, b["pretty"] + "\n", ("display_run", tag))
                if b.get("label"):
                    pv.insert(END, b["label"] + "\n", ("eq_label",))
                pv.insert(END, "\n")
            elif kind == "code":
                for ln in b["lines"]:
                    pv.insert(END, ln + "\n", "code_run")
                pv.insert(END, "\n")
            elif kind == "table":
                for row in b["rows"]:
                    cells = [" ".join(rr["pretty"] for rr in cell)
                             for cell in row]
                    pv.insert(END, " │ ".join(cells) + "\n", "table_run")
                pv.insert(END, "\n")
            elif kind == "quote":
                for r in b["runs"]:
                    self._insert_preview_run(r, "quote_run")
                pv.insert(END, "\n")
            elif kind == "list":
                pv.insert(END, "• ", "table_run")
                for r in b["runs"]:
                    self._insert_preview_run(r, "")
                pv.insert(END, "\n")
            else:  # para
                for r in b["runs"]:
                    if r["text"] == "\n":
                        pv.insert(END, " ")
                    else:
                        self._insert_preview_run(r, "")
                pv.insert(END, "\n\n")

        pv.configure(state=DISABLED)
        self._sync_preview_to_editor()

    def _insert_preview_run(self, run: dict, base_tag: str):
        kind = run.get("kind", "text")
        extra = {"math": "math_run", "bare": "bare_run",
                 "bare_math": "bare_run", "code": "code_run"}.get(kind, "")
        tags = tuple(t for t in (base_tag, extra) if t)
        tag = self._preview_tag(run.get("src"))
        self.preview.insert(END, run.get("pretty", run.get("text", "")),
                            tags + (tag,) if tag else tags)

    def _preview_tag(self, src):
        if not src:
            return ""
        tag = f"src_{src[0]}_{src[1]}_{src[2]}"
        self._preview_srcmap[tag] = src
        self.preview.tag_bind(tag, "<Enter>",
                              lambda e: self.preview.configure(cursor="hand2"))
        self.preview.tag_bind(tag, "<Leave>",
                              lambda e: self.preview.configure(cursor="arrow"))
        return tag

    def _on_preview_click(self, event):
        idx = self.preview.index(f"@{event.x},{event.y}")
        for tag in self.preview.tag_names(idx):
            if tag.startswith("src_") and tag in self._preview_srcmap:
                self._jump_to_source(*self._preview_srcmap[tag])
                return

    def _jump_to_source(self, line: int, c0: int, c1: int):
        self.editor.tag_remove("sel", "1.0", END)
        start, end = f"{line}.{c0}", f"{line}.{c1}"
        self.editor.mark_set("insert", start)
        self.editor.tag_add("sel", start, end)
        self.editor.see(start)
        self.editor.focus_set()
        self._update_status()

    def _sync_preview_to_editor(self):
        """Scroll the preview to the block nearest the editor's top line."""
        self._sync_job = None
        if not self._preview_line_index:
            return
        try:
            top_line = int(self.editor.index("@0,0").split(".")[0])
        except Exception:
            return
        target = self._preview_line_index[0][1]
        for line, pidx in self._preview_line_index:
            if line <= top_line:
                target = pidx
            else:
                break
        self._scroll_syncing = True
        try:
            self.preview.see(target)
        except Exception:
            pass
        finally:
            self.root.after_idle(self._end_scroll_sync)

    def _end_scroll_sync(self):
        self._scroll_syncing = False

    # ------------------------------------------------------------------
    # Formula repair helpers: wrap-as-formula + Formula Inspector
    # ------------------------------------------------------------------
    def _wrap_formula(self, _event=None):
        """Wrap the current selection as $…$ (or $$…$$ when multi-line).
        Fixes partial detections without retyping.  Returns 'break' when
        bound to a key event so Text inserts nothing."""
        try:
            sel = self.editor.get("sel.first", "sel.last")
            start, end = self.editor.index("sel.first"), self.editor.index("sel.last")
        except Exception:
            sel = ""
        if not sel.strip():
            self._status_var.set("Select a formula first, then press Ctrl+M")
            return "break" if _event else None
        body = sel.strip("\n")
        if body.strip().startswith("$"):
            wrapped = body               # already delimited
        elif "\n" in body.strip():
            wrapped = f"$$\n{body.strip()}\n$$"
        else:
            wrapped = f"${body.strip()}$"
        self.editor.delete(start, end)
        self.editor.insert(start, wrapped)
        self._on_editor_change()
        self._status_var.set("Formula wrapped — preview updated")
        return "break" if _event else None

    def _open_inspector(self, _event=None):
        """Formula Inspector dialog: edit a (possibly partial) formula,
        preview its normalized form and conversion status, then insert it
        as inline math, display math, or code."""
        win = Toplevel(self.root)
        win.title("Formula Inspector")
        win.transient(self.root)
        win.configure(bg=_BG)
        win.geometry("620x460")
        win.grab_set()

        outer = ttk.Frame(win, padding=14)
        outer.pack(fill=BOTH, expand=True)

        ttk.Label(outer, text="Formula source  (edit freely — partial detections, "
                              "double spaces and stray periods are tolerated)",
                  style="Subtitle.TLabel").pack(anchor=W)
        src = Text(outer, height=4, wrap=WORD, font=(_MONO, 11),
                   bg=_SURFACE, fg=_INK, insertbackground=_INK,
                   relief=FLAT, borderwidth=0, padx=10, pady=8,
                   highlightbackground=_BORDER, highlightcolor=_ACCENT,
                   highlightthickness=1)
        src.pack(fill=X, pady=(6, 12))

        # pre-fill: selection → clicked span → current line
        initial = ""
        try:
            initial = self.editor.get("sel.first", "sel.last")
        except Exception:
            initial = ""
        if not initial.strip():
            line = self.editor.get("insert linestart", "insert lineend")
            initial = line.strip()
        src.insert("1.0", initial.strip())

        ttk.Label(outer, text="Normalized / pretty form:",
                  style="Subtitle.TLabel").pack(anchor=W)
        norm_var = StringVar()
        ttk.Label(outer, textvariable=norm_var, font=(_SERIF, 12, "bold"),
                  foreground=_ACCENT, wraplength=570,
                  background=_SURFACE, padding=10,
                  relief=FLAT, borderwidth=1,
                  anchor=W).pack(fill=X, pady=(6, 10))

        status_var = StringVar()
        status_lbl = ttk.Label(outer, textvariable=status_var,
                               font=("Segoe UI", 10, "bold"))
        status_lbl.pack(anchor=W)

        def refresh(*_a):
            body = src.get("1.0", END).strip()
            if not body:
                norm_var.set("(empty)")
                status_var.set("")
                return
            norm = normalize_bare_math(body)
            from gui.preview_render import latex_to_pretty
            norm_var.set(latex_to_pretty(norm))
            try:
                ok = self.converter.latex_to_omml(norm) is not None
            except Exception:
                ok = False
            if ok:
                status_var.set("✓ converts to Word math (math zone)")
            else:
                status_var.set("✗ will fall back to plain text — try fixing the source above")

        src.bind("<KeyRelease>", refresh)
        refresh()

        btns = ttk.Frame(outer)
        btns.pack(fill=X, pady=16)

        def finish(mode):
            body = src.get("1.0", END).strip()
            if not body:
                win.destroy()
                return
            if mode == "inline":
                ins = f"${body}$"
            elif mode == "display":
                ins = f"$$\n{body}\n$$"
            else:
                ins = f"`{body}`"
            try:
                sel_start = self.editor.index("sel.first")
                sel_end = self.editor.index("sel.last")
                self.editor.delete(sel_start, sel_end)
            except Exception:
                sel_start = self.editor.index("insert")
            self.editor.insert(sel_start, ins)
            self._on_editor_change()
            win.destroy()

        def make_accent_btn(text, command):
            b = tk.Button(btns, text=text, command=command,
                          font=("Segoe UI Semibold", 9), fg="white", bg=_ACCENT,
                          activebackground=_ACCENT_HOV, activeforeground="white",
                          relief=FLAT, bd=0, padx=14, pady=6, cursor="hand2",
                          highlightthickness=0)
            b.pack(side=LEFT, padx=4)
            return b

        make_accent_btn("Insert as  $formula$", lambda: finish("inline"))
        ttk.Button(btns, text="Insert as display  $$…$$",
                   command=lambda: finish("display")).pack(side=LEFT, padx=4)
        ttk.Button(btns, text="Mark as code (skip math)",
                   command=lambda: finish("code")).pack(side=LEFT, padx=4)
        ttk.Button(btns, text="Cancel",
                   command=win.destroy).pack(side=RIGHT, padx=4)

        src.focus_set()
        return win

    # ------------------------------------------------------------------
    # Conversion
    # ------------------------------------------------------------------
    def _start_conversion(self):
        if self._busy:
            return
        content = self.editor.get("1.0", END).strip()
        if not content:
            messagebox.showwarning("Empty Input",
                                   "Please enter or open some LaTeX/Markdown content.")
            return

        output = self.output_var.get().strip()
        if not output:
            output = filedialog.asksaveasfilename(
                title="Save Word Document",
                defaultextension=".docx",
                filetypes=[("Word Document", "*.docx")])
            if not output:
                return
            self.output_var.set(output)
        if not output.lower().endswith(".docx"):
            output += ".docx"
            self.output_var.set(output)

        try:
            os.makedirs(os.path.dirname(os.path.abspath(output)) or ".", exist_ok=True)
        except Exception as exc:
            messagebox.showerror("Output Error", f"Cannot create output folder:\n{exc}")
            return

        self._last_output_dir = str(Path(os.path.abspath(output)).parent)

        # disable button, start progress
        self._busy = True
        self.convert_btn.configure(state=DISABLED, text="⏳  Converting…")
        self._post_frame.pack_forget()
        self.progress.start(12)
        self._log("Converting…", "dim")

        t = threading.Thread(
            target=self._conversion_worker,
            args=(content, output,
                  self.autonumber_var.get(),
                  self.eq_fmt_var.get()),
            daemon=True,
        )
        t.start()

    def _conversion_worker(self, content, output, autonumber, eq_fmt):
        try:
            stats = self.converter.convert_markdown_to_docx(
                content, output,
                number_equations=autonumber,
                eq_format=eq_fmt,
            )
            self.msg_queue.put(("ok", stats))
        except Exception as exc:
            self.msg_queue.put(("error", str(exc)))

    def _poll_queue(self):
        """Check message queue from worker thread."""
        try:
            while True:
                kind, payload = self.msg_queue.get_nowait()
                self.progress.stop()
                self._busy = False
                self.convert_btn.configure(
                    state=NORMAL, text="⚡  Convert & Save    Ctrl+↵")

                if kind == "ok":
                    stats = payload
                    self._last_output = stats["output_path"]
                    total = stats["equations_total"]
                    ok = stats["equations_converted"]
                    fail = stats["equations_failed"]
                    self._log(
                        f"✅ Success! {ok}/{total} equations converted"
                        f"{f' ({fail} failed)' if fail else ''}. "
                        f"Saved to: {stats['output_path']}",
                        "ok" if not fail else "warn")
                    for item in stats.get("failed_formulas", []):
                        self._log(
                            f"   ⚠ line {item['line']}: could not convert "
                            f"{item['latex']!r} — kept as plain text",
                            "warn")
                    self._post_frame.pack(side=RIGHT)
                    if self.autopen_var.get():
                        self._open_output_file()
                else:
                    self._log(f"❌ Error: {payload}", "err")
                    messagebox.showerror("Conversion Error", str(payload))
        except queue.Empty:
            pass
        self.root.after(100, self._poll_queue)

    # ------------------------------------------------------------------
    # Post-conversion actions
    # ------------------------------------------------------------------
    def _open_output_file(self):
        if self._last_output and os.path.isfile(self._last_output):
            open_path(self._last_output)

    def _open_output_folder(self):
        if self._last_output:
            folder = os.path.dirname(os.path.abspath(self._last_output))
            if os.path.isdir(folder):
                open_path(folder)

    # ------------------------------------------------------------------
    # Logging
    # ------------------------------------------------------------------
    def _log(self, msg: str, kind: str = None):
        self.log.configure(state=NORMAL)
        tags = (kind,) if kind else ()
        self.log.insert(END, msg + "\n", tags)
        self.log.see(END)
        self.log.configure(state=DISABLED)


# ══════════════════════════════════════════════════════════════════
# Entry point
# ══════════════════════════════════════════════════════════════════
def main():
    root = Tk()
    app = ConverterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
