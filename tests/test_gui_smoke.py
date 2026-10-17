"""GUI smoke tests — skipped automatically when no display/tkinter is present."""

from __future__ import annotations

import pytest

tk = pytest.importorskip("tkinter")


@pytest.fixture()
def app():
    try:
        root = tk.Tk()
    except tk.TclError:
        pytest.skip("no display available for Tk")
    root.withdraw()
    try:
        from gui.app import ConverterApp
        yield ConverterApp(root)
    finally:
        root.destroy()


class TestGuiSmoke:
    def test_app_constructs(self, app):
        assert app.converter is not None

    def test_equation_counter(self, app):
        app.editor.insert("1.0", "$$x^2$$\n\nand $$\ny\n$$\n")
        app._on_editor_change()
        assert app._eq_count_var.get().startswith("2 display")

    def test_counter_skips_code_fences(self, app):
        app.editor.insert("1.0", "```\n$$x$$\n```\n")
        app._on_editor_change()
        assert app._eq_count_var.get().startswith("0 display")

    def test_template_insertion(self, app):
        app._tpl_var.set("Fraction")
        app._insert_template()
        content = app.editor.get("1.0", "end")
        assert r"\frac{a}{b}" in content

    def test_open_path_helper_exists(self):
        from gui.app import open_path
        assert callable(open_path)
