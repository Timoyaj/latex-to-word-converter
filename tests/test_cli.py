"""CLI tests: python -m core.converter input.md output.docx"""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

from tests.conftest import MD_FIXTURES, docx_xml, eq_labels, omath_count

PROJECT_ROOT = Path(__file__).resolve().parents[1]


def run_cli(*args):
    return subprocess.run(
        [sys.executable, "-m", "core.converter", *args],
        capture_output=True, text=True, cwd=PROJECT_ROOT,
    )


def test_cli_basic(tmp_path):
    out = tmp_path / "out.docx"
    result = run_cli(str(MD_FIXTURES / "numbering_formats.md"), str(out))
    assert result.returncode == 0, result.stderr
    assert "3/3 equations converted" in result.stdout
    assert out.is_file()
    assert eq_labels(docx_xml(out)) == ["1", "2", "3"]


def test_cli_no_number(tmp_path):
    out = tmp_path / "out.docx"
    result = run_cli(str(MD_FIXTURES / "numbering_formats.md"), str(out), "--no-number")
    assert result.returncode == 0, result.stderr
    assert eq_labels(docx_xml(out)) == []
    assert omath_count(docx_xml(out)) == 3


def test_cli_eq_format(tmp_path):
    out = tmp_path / "out.docx"
    result = run_cli(
        str(MD_FIXTURES / "numbering_formats.md"), str(out),
        "--eq-format", "[Eq. {n}]",
    )
    assert result.returncode == 0, result.stderr
    xml = docx_xml(out)
    assert "[Eq. 1]" in xml and "[Eq. 3]" in xml


def test_cli_reports_failures(tmp_path):
    out = tmp_path / "out.docx"
    result = run_cli(str(MD_FIXTURES / "failures.md"), str(out))
    assert result.returncode == 0, result.stderr
    assert "(2 failed)" in result.stdout


def test_cli_missing_input(tmp_path):
    result = run_cli(str(tmp_path / "nope.md"), str(tmp_path / "out.docx"))
    assert result.returncode == 1
    assert "not found" in result.stdout.lower() + result.stderr.lower()
