"""Testy záložky „Kapacita strojů“ v běžící aplikaci (Streamlit AppTest). Heslo je náhodné, jen pro test.

Spuštění: python -m pytest tests   nebo   python tests/test_capacity_tab.py
"""

import os
import secrets
import sys
from datetime import date
from pathlib import Path

from streamlit.testing.v1 import AppTest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
PASSWORD = secrets.token_urlsafe(12)


def logged_in_app():
    os.chdir(ROOT)
    at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=300)
    at.secrets["password"] = PASSWORD
    at.run()
    at.text_input[0].set_value(PASSWORD).run()
    assert not at.exception, [e.value for e in at.exception]
    return at


def tab(at):
    return at.tabs[6]


def grid_html(at):
    return [m.value for m in tab(at).markdown if "cc-table" in m.value][0]


def chip(at):
    return [m.value for m in tab(at).markdown if "de-status" in m.value]


def test_capacity_tab_is_seventh_and_original_tabs_unchanged():
    at = logged_in_app()
    assert [t.label for t in at.tabs] == ["🏆 Doporučení", "📈 Možné výsledky", "⚠️ Riziko",
                                          "🎯 Co ovlivňuje zisk", "🛒 Nákup strojů", "🗓️ Plán zakázek",
                                          "📅 Kapacita strojů"]
    html_text = grid_html(at)
    for name in ("Mini bagr", "Pásový dumper", "Silniční válec", "Nákladní auto"):
        assert name in html_text
    assert html_text.count("cc over") == 5                         # 4 dny válec + 1 den kolový bagr
    conflicts = [d.value for d in tab(at).dataframe if "Konflikt" in d.value.columns][0]
    assert conflicts["Stroj"].tolist() == ["Kolový bagr", "Silniční válec"]
    assert any("po termínu" in m.value for m in tab(at).markdown)  # zpožděné zakázky z plánu záložky 6


def test_filters_type_period_and_dates():
    at = logged_in_app()
    at.multiselect(key="cap_types").set_value(["SILNICNI_VALEC"]).run()
    html_text = grid_html(at)
    assert "Silniční válec" in html_text and "Mini bagr" not in html_text
    at.slider(key="cap_period").set_value((14, 19)).run()
    html_text = grid_html(at)
    assert "<th>14" in html_text and "<th>19" in html_text and "<th>13" not in html_text
    assert at.number_input(key="cap_detail_day").value == 14        # detail dne drží rozsah období
    at.date_input(key="cap_plan_start").set_value(date(2026, 12, 20)).run()
    assert not at.exception, [e.value for e in at.exception]
    html_text = grid_html(at)
    assert "Ne 3. 1." in html_text                                   # den 15 = 3. 1. 2027 (přes konec roku)


def test_day_detail_and_request_check():
    at = logged_in_app()
    at.selectbox(key="cap_detail_type").set_value("MINI_BAGR")
    at.number_input(key="cap_detail_day").set_value(18).run()
    tables = [d.value for d in tab(at).dataframe]
    assert any("Poznámka" in t.columns and t["Poznámka"].str.startswith("po termínu").any() for t in tables)
    # výchozí požadavek: mini bagr dny 13–16 – drží ho fiktivní pronájem FIKT-P01
    assert "nerealizovatelné" in chip(at)[-1]
    at.number_input(key="cap_req_from").set_value(21)
    at.number_input(key="cap_req_to").set_value(24).run()
    assert "realizovatelné" in chip(at)[-1] and "nerealizovatelné" not in chip(at)[-1]
    at.number_input(key="cap_req_from").set_value(30)
    at.number_input(key="cap_req_to").set_value(29).run()
    assert "nelze určit" in chip(at)[-1] and "Nelze vyhodnotit" in chip(at)[-1]
    assert not at.exception


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_"):
            fn()
            print("OK", name)
