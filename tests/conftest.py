"""Shared pytest fixtures and docx-inspection helpers for the test suite."""

from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path

import pytest

from core.converter import LatexToWordConverter

FIXTURES = Path(__file__).parent / "fixtures"
MD_FIXTURES = FIXTURES / "md"
EXPECTED_DIR = FIXTURES / "expected"
VBA_FIXTURES = FIXTURES / "vba"


@pytest.fixture()
def converter():
    return LatexToWordConverter()


@pytest.fixture()
def out_docx(tmp_path):
    """Path for a scratch output document inside the pytest tmp dir."""
    return tmp_path / "out.docx"


def md_text(name: str) -> str:
    return (MD_FIXTURES / name).read_text(encoding="utf-8")


def docx_xml(path) -> str:
    """word/document.xml of a generated .docx, as text."""
    with zipfile.ZipFile(path) as z:
        return z.read("word/document.xml").decode("utf-8")


def omath_count(xml: str) -> int:
    """Count math zones: oMath minus the oMathPara wrapper's own tag."""
    return len(re.findall(r"<m:oMath[ >]", xml))


def eq_labels(xml: str) -> list:
    """Equation number labels like (1), [Eq. 2] found in plain runs."""
    return re.findall(r"<w:t[^>]*>\((\d+)\)</w:t>", xml)


def table_count(xml: str) -> int:
    return xml.count("<w:tbl>")


def load_expected(name: str) -> dict:
    """Structural golden facts for a fixture (equations, tables, labels...)."""
    with open(EXPECTED_DIR / name, encoding="utf-8") as f:
        return json.load(f)
