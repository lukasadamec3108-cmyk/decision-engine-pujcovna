"""Testy českého formátování čísel (modul formatting, etapa M1 – přesunuto z ui_theme beze změny)."""

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import formatting as f  # noqa: E402

NB = " "


def test_numbers_and_units():
    assert f.fmt_num(1234567) == f"1{NB}234{NB}567"
    assert f.fmt_num(12.346, 2) == "12,35"
    assert f.fmt_czk(-1500) == f"−1{NB}500{NB}Kč"
    assert f.fmt_pct(24.6) == f"24,6{NB}%"
    assert f.fmt_days(12.5) == f"12,5{NB}dní"
    assert f.fmt_num(None) == "–"


def test_negative_zero_has_no_minus():
    assert f.fmt_num(-0.0) == "0"
    assert f.fmt_num(-0.004, 2) == "0,00"
    assert f.fmt_czk(-0.4) == f"0{NB}Kč"


def test_ui_theme_reexports_the_same_functions():
    import ui_theme
    for name in ("fmt_num", "fmt_czk", "fmt_pct", "fmt_days", "NBSP"):
        assert getattr(ui_theme, name) is getattr(f, name)


def test_module_has_no_streamlit_dependency():
    assert "import streamlit" not in (ROOT / "formatting.py").read_text(encoding="utf-8")
