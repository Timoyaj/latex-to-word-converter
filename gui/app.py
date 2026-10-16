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
  * find & replace bar (Ctrl+F) with match count
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
    Tk, Text, StringVar, BooleanVar,
    filedialog, messagebox,
    END, DISABLED, NORMAL, WORD, HORIZONTAL, BOTH, LEFT, RIGHT, X, Y, TOP, BOTTOM,
    NW, SE, E, W, SUNKEN, FLAT, VERTICAL,
)
import tkinter
import tkinter.ttk as ttk

# Ensure project root is importable
_HERE = os.path.dirname(os.path.abspath(__file__))
_PROJECT_ROOT = os.path.dirname(_HERE)
if _PROJECT_ROOT not in sys.path:
    sys.path.insert(0, _PROJECT_ROOT)

from core.converter import LatexToWordConverter
from core.math_tokens import (
    count_display_math, split_inline_math, is_fence_line, fence_marker,
    INLINE_DOLLAR_RE, INLINE_PAREN_RE,
)

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


# ── colour palette ────────────────────────────────────────────────
_BG          = "#f7f8fa"
_PANEL_FG    = "#334155"
_MUTED       = "#64748b"
_EDITOR_BG   = "#1e1e2e"
_EDITOR_FG   = "#cdd6f4"
_GUTTER_BG   = "#181825"
_GUTTER_FG   = "#6c7086"
_LINE_HL     = "#2a2a3d"
_MATH_FG     = "#f9e2af"      # yellow-ish for inline $…$
_DISPLAY_FG  = "#a6e3a1"      # green for $$…$$
_ACCENT      = "#1d4ed8"      # blue button
_ACCENT_HOV  = "#2563eb"
_LOG_BG      = "#f1f5f9"
_FIND_BG     = "#eef2ff"


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
        self.root.geometry("980x780")
        self.root.minsize(800, 620)
        self.root.configure(bg=_BG)
        self.root.protocol("WM_DELETE_WINDOW", self._quit)

        style = ttk.Style()
        try:
            style.theme_use("vista")
        except Exception:
            try:
                style.theme_use("clam")
            except Exception:
                pass

        # custom styles
        style.configure("Accent.TButton",
                        font=("Segoe UI Semibold", 12),
                        foreground="white", background=_ACCENT)
        style.map("Accent.TButton",
                  background=[("active", _ACCENT_HOV), ("disabled", "#93c5fd")])
        style.configure("Tool.TButton", font=("Segoe UI", 10))
        style.configure("TLabel", font=("Segoe UI", 10), background=_BG)
        style.configure("TCheckbutton", font=("Segoe UI", 10), background=_BG)
        style.configure("TLabelframe", background=_BG)
        style.configure("TLabelframe.Label",
                        font=("Segoe UI Semibold", 10), background=_BG,
                        foreground=_PANEL_FG)
        style.configure("Status.TLabel",
                        font=("Segoe UI", 9), background=_BG, foreground=_MUTED)
        style.configure("Horizontal.TProgressbar", troughcolor="#e2e8f0",
                        background=_ACCENT)

    # ------------------------------------------------------------------
    # Build the complete UI
    # ------------------------------------------------------------------
    def _build_ui(self):
        self._build_menubar()

        outer = ttk.Frame(self.root, padding=(10, 8, 10, 6))
        outer.pack(fill=BOTH, expand=True)

        # ── title bar ─────────────────────────────────────────────
        title_frame = ttk.Frame(outer)
        title_frame.pack(fill=X, pady=(0, 4))
        ttk.Label(title_frame,
                  text=APP_NAME,
                  font=("Segoe UI", 17, "bold")).pack(side=LEFT)
        ttk.Label(title_frame,
                  text=f"Native editable OMML equations · v{APP_VERSION}",
                  font=("Segoe UI", 9),
                  foreground=_MUTED).pack(side=LEFT, padx=(12, 0))

        # ── toolbar ───────────────────────────────────────────────
        tb = ttk.Frame(outer)
        tb.pack(fill=X, pady=(0, 4))

        ttk.Button(tb, text="📂 Open", style="Tool.TButton",
                   command=self._open_file).pack(side=LEFT, padx=2)
        ttk.Button(tb, text="📋 Paste", style="Tool.TButton",
                   command=self._paste).pack(side=LEFT, padx=2)
        ttk.Button(tb, text="🔍 Find", style="Tool.TButton",
                   command=self.toggle_find_bar).pack(side=LEFT, padx=2)
        ttk.Button(tb, text="🗑 Clear", style="Tool.TButton",
                   command=self._clear).pack(side=LEFT, padx=2)

        ttk.Separator(tb, orient=VERTICAL).pack(side=LEFT, padx=8, fill=Y, pady=2)

        # template menu
        self._tpl_var = StringVar(value="Insert Template…")
        tpl_combo = ttk.Combobox(tb, textvariable=self._tpl_var,
                                 values=list(self.TEMPLATES.keys()),
                                 state="readonly", width=20)
        tpl_combo.pack(side=LEFT, padx=2)
        tpl_combo.bind("<<ComboboxSelected>>", self._insert_template)

        # equation counter label (right side)
        self._eq_count_var = StringVar(value="0 display equations detected")
        ttk.Label(tb, textvariable=self._eq_count_var,
                  font=("Segoe UI", 9), foreground="#6366f1").pack(side=RIGHT, padx=6)

        # ── find & replace bar (hidden until Ctrl+F) ──────────────
        self.find_frame = ttk.Frame(outer, padding=(8, 6))
        self._find_visible = False

        ttk.Label(self.find_frame, text="Find:").pack(side=LEFT)
        self._find_var = StringVar()
        self._find_entry = ttk.Entry(self.find_frame, textvariable=self._find_var, width=28)
        self._find_entry.pack(side=LEFT, padx=4)
        self._find_entry.bind("<Return>", lambda e: self._find_next())
        self._find_entry.bind("<Escape>", lambda e: self.toggle_find_bar(False))

        ttk.Label(self.find_frame, text="Replace:").pack(side=LEFT, padx=(10, 0))
        self._replace_var = StringVar()
        self._replace_entry = ttk.Entry(self.find_frame, textvariable=self._replace_var, width=22)
        self._replace_entry.pack(side=LEFT, padx=4)
        self._replace_entry.bind("<Return>", lambda e: self._replace_current())
        self._replace_entry.bind("<Escape>", lambda e: self.toggle_find_bar(False))

        ttk.Button(self.find_frame, text="▲", width=3,
                   command=lambda: self._find_next(backwards=True)).pack(side=LEFT, padx=1)
        ttk.Button(self.find_frame, text="▼", width=3,
                   command=self._find_next).pack(side=LEFT, padx=1)
        ttk.Button(self.find_frame, text="Replace",
                   command=self._replace_current).pack(side=LEFT, padx=1)
        ttk.Button(self.find_frame, text="All",
                   command=self._replace_all).pack(side=LEFT, padx=1)

        self._find_status_var = StringVar(value="")
        ttk.Label(self.find_frame, textvariable=self._find_status_var,
                  foreground=_MUTED).pack(side=LEFT, padx=8)
        ttk.Button(self.find_frame, text="✕",
                   command=lambda: self.toggle_find_bar(False)).pack(side=RIGHT)

        # ── editor (gutter + text + scrollbars) ───────────────────
        ed_frame = ttk.LabelFrame(outer, text="LaTeX / Markdown Input", padding=4)
        ed_frame.pack(fill=BOTH, expand=True, pady=(0, 6))

        self.gutter = Text(
            ed_frame, width=4, padx=6, pady=6, takefocus=0,
            borderwidth=0, relief=FLAT, state=DISABLED, wrap="none",
            font=("Consolas", 11), cursor="arrow",
            bg=_GUTTER_BG, fg=_GUTTER_FG,
            selectbackground=_GUTTER_BG, selectforeground=_GUTTER_FG,
        )
        self.gutter.pack(side=LEFT, fill=Y)

        v_scroll = ttk.Scrollbar(ed_frame, orient=VERTICAL)
        h_scroll = ttk.Scrollbar(outer, orient=HORIZONTAL)
        self.v_scroll = v_scroll

        self.editor = Text(
            ed_frame, wrap="none", undo=True,
            font=("Consolas", 11),
            bg=_EDITOR_BG, fg=_EDITOR_FG,
            insertbackground="#f5f5f5",
            selectbackground="#45475a",
            relief=FLAT, borderwidth=0,
            padx=8, pady=6,
            yscrollcommand=self._on_editor_scroll,
            xscrollcommand=h_scroll.set,
        )
        self.editor.pack(side=LEFT, fill=BOTH, expand=True)
        v_scroll.pack(side=RIGHT, fill=Y)
        v_scroll.config(command=self._on_scrollbar)
        h_scroll.pack(fill=X, pady=(4, 0))
        h_scroll.config(command=self.editor.xview)

        # syntax-highlight + decoration tags
        self.editor.tag_configure("inline_math",  foreground=_MATH_FG)
        self.editor.tag_configure("display_math", foreground=_DISPLAY_FG)
        self.editor.tag_configure("code_block", foreground="#9399b2",
                                  background="#181825")
        self.editor.tag_configure("current_line", background=_LINE_HL)
        self.editor.tag_configure("found", background="#45475a",
                                  foreground="#f9e2af")

        self.editor.bind("<<Modified>>", self._on_modified)
        self.editor.bind("<KeyRelease>", self._on_editor_change)
        self.editor.bind("<ButtonRelease-1>", self._on_editor_change)
        self.editor.bind("<Configure>", self._on_editor_change)

        # ── options panel ─────────────────────────────────────────
        opt = ttk.LabelFrame(outer, text="Conversion Options", padding=(8, 4))
        opt.pack(fill=X, pady=(0, 6))

        row1 = ttk.Frame(opt)
        row1.pack(fill=X, pady=2)

        self.autonumber_var = BooleanVar(value=bool(self.settings.get("autonumber", True)))
        cb = ttk.Checkbutton(row1, text="Auto-number equations",
                             variable=self.autonumber_var,
                             command=self._toggle_eq_format)
        cb.pack(side=LEFT)

        ttk.Label(row1, text="   Format:").pack(side=LEFT)
        self.eq_fmt_var = StringVar(value=str(self.settings.get("eq_format", "({n})")))
        self.eq_fmt_entry = ttk.Entry(row1, textvariable=self.eq_fmt_var, width=10)
        self.eq_fmt_entry.pack(side=LEFT, padx=4)

        self.autopen_var = BooleanVar(value=bool(self.settings.get("autopen", False)))
        ttk.Checkbutton(row1, text="Open file after conversion",
                        variable=self.autopen_var).pack(side=LEFT, padx=(18, 0))

        row2 = ttk.Frame(opt)
        row2.pack(fill=X, pady=2)

        ttk.Label(row2, text="Output file:").pack(side=LEFT)
        self.output_var = StringVar()
        self.output_entry = ttk.Entry(row2, textvariable=self.output_var, width=52)
        self.output_entry.pack(side=LEFT, padx=4, fill=X, expand=True)
        ttk.Button(row2, text="Browse…", command=self._browse_output).pack(side=LEFT)

        # ── action bar ────────────────────────────────────────────
        act = ttk.Frame(outer)
        act.pack(fill=X, pady=(0, 4))

        self.convert_btn = ttk.Button(
            act, text="⚡  Convert to Word  (Ctrl+Return)",
            style="Accent.TButton",
            command=self._start_conversion,
        )
        self.convert_btn.pack(side=LEFT, ipady=6, ipadx=14)

        self.progress = ttk.Progressbar(act, mode="indeterminate", length=180)
        self.progress.pack(side=LEFT, padx=12)

        # post-conversion buttons (hidden initially)
        self._post_frame = ttk.Frame(act)
        ttk.Button(self._post_frame, text="📄 Open File",
                   command=self._open_output_file).pack(side=LEFT, padx=2)
        ttk.Button(self._post_frame, text="📁 Open Folder",
                   command=self._open_output_folder).pack(side=LEFT, padx=2)

        # ── log area ──────────────────────────────────────────────
        log_frame = ttk.LabelFrame(outer, text="Status Log", padding=4)
        log_frame.pack(fill=X)

        self.log = Text(
            log_frame, height=5, wrap=WORD, state=DISABLED,
            font=("Consolas", 9),
            bg=_LOG_BG, fg="#334155",
            relief=FLAT, borderwidth=0,
            padx=6, pady=4,
        )
        self.log.pack(fill=X)

        # ── status bar ────────────────────────────────────────────
        status = ttk.Frame(outer)
        status.pack(fill=X, pady=(4, 0))
        self._status_var = StringVar(value="Ready")
        ttk.Label(status, textvariable=self._status_var,
                  style="Status.TLabel").pack(side=LEFT)
        self._math_stat_var = StringVar(value="0 display · 0 inline")
        ttk.Label(status, textvariable=self._math_stat_var,
                  style="Status.TLabel").pack(side=RIGHT)

        self._log("Ready. Paste or open a LaTeX/Markdown file to begin.")
        if self.converter.xsl_path:
            self._log(f"OMML stylesheet loaded: {self.converter.xsl_path}")
        else:
            self._log("⚠ MML2OMML.XSL not found — equations will render as plain text.")

    # ------------------------------------------------------------------
    # Menu bar
    # ------------------------------------------------------------------
    def _build_menubar(self):
        menubar = tkinter.Menu(self.root, tearoff=0, font=("Segoe UI", 10))

        file_menu = tkinter.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Open File…        Ctrl+O",
                              accelerator="Ctrl+O", command=self._open_file)
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

        help_menu = tkinter.Menu(menubar, tearoff=0)
        help_menu.add_command(label="About", command=self._about)
        menubar.add_cascade(label="Help", menu=help_menu)

        self.root.config(menu=menubar)

    @staticmethod
    def _safe_edit(fn):
        try:
            fn()
        except Exception:
            pass

    def _about(self):
        messagebox.showinfo(
            "About",
            f"{APP_NAME}\nVersion {APP_VERSION}\n\n"
            "Markdown/LaTeX → Word with native OMML equations,\n"
            "automatic numbering, and Word VBA automation.\n\n"
            "Math detection: Pandoc dollar-math rules\n"
            "(currency, \\$ escapes and code fences are safe).",
        )

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
        r.bind("<F3>", lambda e: self._find_next())
        r.bind("<Shift-F3>", lambda e: self._find_next(backwards=True))
        r.bind("<Control-Return>", lambda e: self._start_conversion())

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
                # suggest output name
                stem = Path(path).stem
                suggested = str(Path(path).parent / f"{stem}.docx")
                self.output_var.set(suggested)
                self._log(f"Loaded: {path}")
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
        self.editor.delete("1.0", END)
        self.editor.edit_reset()
        self._current_file = ""
        self._on_editor_change()

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
            self.find_frame.pack(fill=X, pady=(0, 4), before=self.editor.master)
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

    def _on_scrollbar(self, *args):
        self.editor.yview(*args)
        self.gutter.yview_moveto(self.editor.yview()[0])

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
            inline += sum(1 for t in split_inline_math(ln) if t.kind == "math")
        self._math_stat_var.set(f"{count} display · {inline} inline")

        words = len(re.findall(r"\S+", text.rstrip("\n")))
        chars = len(text.rstrip("\n"))
        self._update_status(words, chars)

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
        self.convert_btn.configure(state=DISABLED)
        self._post_frame.pack_forget()
        self.progress.start(12)
        self._log("Converting…")

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
                self.convert_btn.configure(state=NORMAL)

                if kind == "ok":
                    stats = payload
                    self._last_output = stats["output_path"]
                    total = stats["equations_total"]
                    ok = stats["equations_converted"]
                    fail = stats["equations_failed"]
                    self._log(
                        f"✅ Success! {ok}/{total} equations converted"
                        f"{f' ({fail} failed)' if fail else ''}. "
                        f"Saved to: {stats['output_path']}"
                    )
                    for item in stats.get("failed_formulas", []):
                        self._log(
                            f"   ⚠ line {item['line']}: could not convert "
                            f"{item['latex']!r} — kept as plain text"
                        )
                    self._post_frame.pack(side=RIGHT)
                    if self.autopen_var.get():
                        self._open_output_file()
                else:
                    self._log(f"❌ Error: {payload}")
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
    def _log(self, msg: str):
        self.log.configure(state=NORMAL)
        self.log.insert(END, msg + "\n")
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
