#!/usr/bin/env python3
"""
LaTeX → Word Converter — GUI Application
==========================================
A polished tkinter interface for converting Markdown/LaTeX documents
into professional Microsoft Word (.docx) files with native OMML equations
and automatic equation numbering.
"""

import sys
import os
import re
import subprocess
import threading
import queue
from pathlib import Path
from tkinter import (
    Tk, Text, StringVar, BooleanVar, IntVar,
    filedialog, messagebox, scrolledtext,
    END, DISABLED, NORMAL, WORD, HORIZONTAL,
)
import tkinter.ttk as ttk

# Ensure project root is importable
_HERE = os.path.dirname(os.path.abspath(__file__))
_PROJECT_ROOT = os.path.dirname(_HERE)
if _PROJECT_ROOT not in sys.path:
    sys.path.insert(0, _PROJECT_ROOT)

from core.converter import LatexToWordConverter
from core.math_tokens import (
    count_display_math, is_fence_line, fence_marker,
    INLINE_DOLLAR_RE, INLINE_PAREN_RE,
)


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
_EDITOR_BG   = "#1e1e2e"
_EDITOR_FG   = "#cdd6f4"
_MATH_FG     = "#f9e2af"      # yellow-ish for inline $…$
_DISPLAY_FG  = "#a6e3a1"      # green for $$…$$
_ACCENT      = "#1d4ed8"      # blue button
_ACCENT_HOV  = "#2563eb"
_LOG_BG      = "#f1f5f9"


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

        self._configure_root()
        self._build_ui()
        self._poll_queue()

    # ------------------------------------------------------------------
    # Root window configuration
    # ------------------------------------------------------------------
    def _configure_root(self):
        self.root.title("LaTeX → Word Converter")
        self.root.geometry("960x760")
        self.root.minsize(800, 620)
        self.root.configure(bg=_BG)
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
                  background=[("active", _ACCENT_HOV)])

    # ------------------------------------------------------------------
    # Build the complete UI
    # ------------------------------------------------------------------
    def _build_ui(self):
        outer = ttk.Frame(self.root, padding=10)
        outer.pack(fill="both", expand=True)

        # ── title bar ─────────────────────────────────────────────
        title_frame = ttk.Frame(outer)
        title_frame.pack(fill="x", pady=(0, 6))
        ttk.Label(title_frame,
                  text="LaTeX → Word Converter",
                  font=("Segoe UI", 18, "bold")).pack(side="left")
        ttk.Label(title_frame,
                  text="Native editable OMML equations",
                  font=("Segoe UI", 10),
                  foreground="#64748b").pack(side="left", padx=(12, 0))

        # ── toolbar ───────────────────────────────────────────────
        tb = ttk.Frame(outer)
        tb.pack(fill="x", pady=(0, 4))

        ttk.Button(tb, text="📂 Open File", command=self._open_file).pack(side="left", padx=2)
        ttk.Button(tb, text="📋 Paste", command=self._paste).pack(side="left", padx=2)
        ttk.Button(tb, text="🗑 Clear", command=self._clear).pack(side="left", padx=2)

        ttk.Separator(tb, orient="vertical").pack(side="left", padx=8, fill="y", pady=2)

        # template menu
        self._tpl_var = StringVar(value="Insert Template…")
        tpl_combo = ttk.Combobox(tb, textvariable=self._tpl_var,
                                 values=list(self.TEMPLATES.keys()),
                                 state="readonly", width=20)
        tpl_combo.pack(side="left", padx=2)
        tpl_combo.bind("<<ComboboxSelected>>", self._insert_template)

        # equation counter label (right side)
        self._eq_count_var = StringVar(value="0 equations detected")
        ttk.Label(tb, textvariable=self._eq_count_var,
                  font=("Segoe UI", 9), foreground="#6366f1").pack(side="right", padx=6)

        # ── editor ────────────────────────────────────────────────
        ed_frame = ttk.LabelFrame(outer, text="LaTeX / Markdown Input", padding=4)
        ed_frame.pack(fill="both", expand=True, pady=(0, 8))

        self.editor = Text(
            ed_frame, wrap=WORD, undo=True,
            font=("Consolas", 11),
            bg=_EDITOR_BG, fg=_EDITOR_FG,
            insertbackground="#f5f5f5",
            selectbackground="#45475a",
            relief="flat", borderwidth=0,
            padx=8, pady=6,
        )
        self.editor.pack(fill="both", expand=True)

        # syntax-highlight tags
        self.editor.tag_configure("inline_math",  foreground=_MATH_FG)
        self.editor.tag_configure("display_math", foreground=_DISPLAY_FG)
        self.editor.tag_configure("code_block", foreground="#9399b2",
                                  background="#181825")

        self.editor.bind("<KeyRelease>", self._on_editor_change)
        self.editor.bind("<ButtonRelease-1>", self._on_editor_change)

        # ── options panel ─────────────────────────────────────────
        opt = ttk.LabelFrame(outer, text="Conversion Options", padding=8)
        opt.pack(fill="x", pady=(0, 8))

        row1 = ttk.Frame(opt)
        row1.pack(fill="x", pady=2)

        self.autonumber_var = BooleanVar(value=True)
        cb = ttk.Checkbutton(row1, text="Auto-number equations",
                             variable=self.autonumber_var,
                             command=self._toggle_eq_format)
        cb.pack(side="left")

        ttk.Label(row1, text="   Format:").pack(side="left")
        self.eq_fmt_var = StringVar(value="({n})")
        self.eq_fmt_entry = ttk.Entry(row1, textvariable=self.eq_fmt_var, width=10)
        self.eq_fmt_entry.pack(side="left", padx=4)

        row2 = ttk.Frame(opt)
        row2.pack(fill="x", pady=2)

        ttk.Label(row2, text="Output file:").pack(side="left")
        self.output_var = StringVar()
        self.output_entry = ttk.Entry(row2, textvariable=self.output_var, width=50)
        self.output_entry.pack(side="left", padx=4, fill="x", expand=True)
        ttk.Button(row2, text="Browse…", command=self._browse_output).pack(side="left")

        # ── action bar ────────────────────────────────────────────
        act = ttk.Frame(outer)
        act.pack(fill="x", pady=(0, 4))

        self.convert_btn = ttk.Button(
            act, text="⚡  Convert to Word",
            style="Accent.TButton",
            command=self._start_conversion,
        )
        self.convert_btn.pack(side="left", ipady=6, ipadx=16)

        self.progress = ttk.Progressbar(act, mode="indeterminate", length=200)
        self.progress.pack(side="left", padx=12)

        # post-conversion buttons (hidden initially)
        self._post_frame = ttk.Frame(act)
        ttk.Button(self._post_frame, text="📄 Open File",
                   command=self._open_output_file).pack(side="left", padx=2)
        ttk.Button(self._post_frame, text="📁 Open Folder",
                   command=self._open_output_folder).pack(side="left", padx=2)

        # ── log area ──────────────────────────────────────────────
        log_frame = ttk.LabelFrame(outer, text="Status Log", padding=4)
        log_frame.pack(fill="x")

        self.log = Text(
            log_frame, height=5, wrap=WORD, state=DISABLED,
            font=("Consolas", 9),
            bg=_LOG_BG, fg="#334155",
            relief="flat", borderwidth=0,
            padx=6, pady=4,
        )
        self.log.pack(fill="x")

        self._log("Ready. Paste or open a LaTeX/Markdown file to begin.")
        if self.converter.xsl_path:
            self._log(f"OMML stylesheet loaded: {self.converter.xsl_path}")
        else:
            self._log("⚠ MML2OMML.XSL not found — equations will render as plain text.")

    # ------------------------------------------------------------------
    # Toolbar actions
    # ------------------------------------------------------------------
    def _open_file(self):
        path = filedialog.askopenfilename(
            title="Open Markdown / LaTeX file",
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
        self._on_editor_change()

    def _insert_template(self, _event=None):
        name = self._tpl_var.get()
        tpl = self.TEMPLATES.get(name)
        if tpl:
            self.editor.insert("insert", "\n" + tpl + "\n")
            self._on_editor_change()
        self._tpl_var.set("Insert Template…")

    def _browse_output(self):
        path = filedialog.asksaveasfilename(
            title="Save Word Document",
            defaultextension=".docx",
            filetypes=[("Word Document", "*.docx")])
        if path:
            self.output_var.set(path)

    def _toggle_eq_format(self):
        if self.autonumber_var.get():
            self.eq_fmt_entry.configure(state=NORMAL)
        else:
            self.eq_fmt_entry.configure(state=DISABLED)

    # ------------------------------------------------------------------
    # Editor syntax highlighting + equation counting
    # ------------------------------------------------------------------
    def _on_editor_change(self, _event=None):
        text = self.editor.get("1.0", END)
        # remove old tags
        self.editor.tag_remove("inline_math", "1.0", END)
        self.editor.tag_remove("display_math", "1.0", END)
        self.editor.tag_remove("code_block", "1.0", END)

        lines = text.split("\n")
        # absolute char offset of each line start (tk counts '\n' as 1 char)
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
            e = s + len(lines[idx])
            tag_range("code_block", s, e)

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
                        # tag the whole line
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
        # \[ ... \] display blocks (single-line only for tagging simplicity)
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

        count = count_display_math(text)
        self._eq_count_var.set(
            f"{count} display equation{'s' if count != 1 else ''} detected")

    # ------------------------------------------------------------------
    # Conversion
    # ------------------------------------------------------------------
    def _start_conversion(self):
        content = self.editor.get("1.0", END).strip()
        if not content:
            messagebox.showwarning("Empty Input", "Please enter or open some LaTeX/Markdown content.")
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

        # disable button, start progress
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
                    self._post_frame.pack(side="right")
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
