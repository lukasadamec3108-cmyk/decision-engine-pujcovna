"""Zapamatování vstupů stránky Plánování kapacity (filtry, detail dne, ověření požadavku). Hesla jsou náhodná.

Streamlit stav nevykresleného widgetu smaže; stránka proto hodnoty ukládá i pod INV_KEEP + klíč (app.kept_input)
a po návratu je obnoví. Testy ověřují zachování hodnot, stejné výsledky, validaci, reset a odhlášení.

Spuštění: python -m pytest tests/test_capacity_filters.py
"""

import logging
import os
import secrets
import sys
from datetime import date
from pathlib import Path

from streamlit.testing.v1 import AppTest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))
from app_pages import open_page  # noqa: E402

KEY = "kapacita"
PASSWORD = secrets.token_urlsafe(12)

KEPT = ("cap_types", "cap_plan_start", "cap_period", "cap_view", "cap_detail_type", "cap_detail_day",
        "cap_req_type", "cap_req_from", "cap_req_to", "cap_req_qty", "cap_req_tent")
DEFAULTS = {"cap_types": [], "cap_plan_start": None, "cap_period": (1, 25), "cap_view": "Kalendář dostupnosti",
            "cap_detail_type": "MINI_BAGR", "cap_detail_day": 1, "cap_req_type": "MINI_BAGR",
            "cap_req_from": 13, "cap_req_to": 16, "cap_req_qty": 1, "cap_req_tent": False}
CHANGED = {"cap_types": ["SILNICNI_VALEC", "MINI_BAGR"], "cap_plan_start": date(2026, 3, 2), "cap_period": (2, 20),
           "cap_view": "Kalendář dostupnosti", "cap_detail_type": "SILNICNI_VALEC", "cap_detail_day": 15,
           "cap_req_type": "PASOVY_DUMPER", "cap_req_from": 21, "cap_req_to": 24, "cap_req_qty": 2,
           "cap_req_tent": True}


def app():
    os.chdir(ROOT)
    at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=300)
    at.secrets["password"] = PASSWORD
    at.run()
    return login(at)


def login(at):
    at.text_input[0].set_value(PASSWORD).run()
    assert not at.exception, [e.value for e in at.exception]
    return at


def page(at=None):
    at = open_page(at or app(), KEY)
    assert not at.exception, [e.value for e in at.exception]
    return at


def values(at):
    """Zobrazené hodnoty vstupů stránky (jen vykreslené widgety)."""
    out = {}
    for kind in ("multiselect", "date_input", "slider", "radio", "selectbox", "number_input", "checkbox"):
        for w in at.get(kind):
            if w.key in KEPT:
                out[w.key] = w.value
    return out


def set_inputs(at, changed=CHANGED):
    at.multiselect(key="cap_types").set_value(changed["cap_types"])
    at.date_input(key="cap_plan_start").set_value(changed["cap_plan_start"])
    at.slider(key="cap_period").set_value(changed["cap_period"]).run()
    at.selectbox(key="cap_detail_type").set_value(changed["cap_detail_type"])
    at.number_input(key="cap_detail_day").set_value(changed["cap_detail_day"])
    at.selectbox(key="cap_req_type").set_value(changed["cap_req_type"])
    at.number_input(key="cap_req_from").set_value(changed["cap_req_from"])
    at.number_input(key="cap_req_to").set_value(changed["cap_req_to"])
    at.number_input(key="cap_req_qty").set_value(changed["cap_req_qty"])
    at.checkbox(key="cap_req_tent").set_value(changed["cap_req_tent"]).run()
    assert not at.exception, [e.value for e in at.exception]
    assert values(at) == changed
    return at


def content(at):
    """Výsledky stránky: texty (kalendář, souhrn, ověření požadavku) a tabulky."""
    texts = [m.value for m in at.main.markdown if not m.value.lstrip().startswith("<style")]
    tables = [d.value.to_json() for d in at.main.dataframe]
    return texts, tables


def test_inputs_survive_dashboard_and_economics_pages():
    at = set_inputs(page())
    before = content(at)
    for other in ("dashboard", "ekonomika", "investice", "rozhodovaci_centrum"):
        open_page(at, other)
        assert not at.exception, other
        page(at)
        assert values(at) == CHANGED, other                              # nic se nevrátilo na výchozí
        assert content(at) == before, other                              # stejné výsledky jako před odchodem


def test_results_equal_direct_input_in_fresh_session():
    at = set_inputs(page())
    open_page(at, "dashboard")
    page(at)
    fresh = set_inputs(page())                                           # stejné vstupy zadané napřímo
    assert content(at) == content(fresh)


def test_default_results_identical_on_first_visit_and_after_return():
    at = page()
    first = content(at)
    open_page(at, "dashboard")
    page(at)
    assert values(at) == DEFAULTS and content(at) == first


def test_view_toggle_keeps_day_detail_and_request_inputs():
    at = set_inputs(page())
    at.radio(key="cap_view").set_value("Přehled využití").run()
    assert not at.exception and "cap_req_from" not in values(at)
    at.selectbox(key="use_detail_type").set_value("MINI_BAGR").run()
    open_page(at, "dashboard")                                           # i přes jinou stránku
    page(at)
    assert at.radio(key="cap_view").value == "Přehled využití"
    assert at.selectbox(key="use_detail_type").value == "MINI_BAGR"
    at.radio(key="cap_view").set_value("Kalendář dostupnosti").run()
    assert values(at) == CHANGED


def test_show_in_calendar_button_still_overrides_filter():
    at = set_inputs(page())
    at.radio(key="cap_view").set_value("Přehled využití").run()
    at.selectbox(key="use_detail_type").set_value("SILNICNI_VALEC").run()
    at.button(key="use_to_calendar").click().run()
    assert not at.exception
    assert at.multiselect(key="cap_types").value == ["SILNICNI_VALEC"]
    assert at.selectbox(key="cap_detail_type").value == "SILNICNI_VALEC"
    assert at.radio(key="cap_view").value == "Kalendář dostupnosti"
    assert at.slider(key="cap_period").value == (2, 20)                  # ostatní vstupy zůstaly
    open_page(at, "dashboard")
    page(at)
    assert at.multiselect(key="cap_types").value == ["SILNICNI_VALEC"]


def test_links_from_fleet_economics_override_kept_values():
    at = set_inputs(page())
    open_page(at, "ekonomika")
    at.slider(key="fe_period").set_value((3, 18)).run()
    at.button(key="fe_usage_detail").click().run()                       # odkaz „Podrobné využití…“
    assert not at.exception and not at.warning
    assert at.slider(key="cap_period").value == (3, 18)
    assert at.multiselect(key="cap_types").value == []
    assert at.radio(key="cap_view").value == "Přehled využití"
    open_page(at, "ekonomika")
    assert at.slider(key="fe_period").value == (3, 18)                   # Ekonomika si filtr pamatuje dál
    at.button(key="fe_cal_K2").click().run()                             # odkaz do kalendáře na stroj
    assert not at.exception
    assert at.multiselect(key="cap_types").value == ["SILNICNI_VALEC"]
    assert at.selectbox(key="cap_detail_type").value == "SILNICNI_VALEC"
    assert at.radio(key="cap_view").value == "Kalendář dostupnosti"
    assert at.slider(key="cap_period").value == (3, 18)
    assert at.number_input(key="cap_req_from").value == 21               # ověření požadavku zůstalo


def test_reset_restores_defaults_results_and_only_this_page():
    at = page()
    default_content = content(at)
    open_page(at, "ekonomika")
    at.slider(key="fe_period").set_value((3, 18)).run()
    open_page(at, "investice")
    at.number_input(key="inv_fixed_cost").set_value(120000).run()
    page(at)
    set_inputs(at)
    at.radio(key="cap_view").set_value("Přehled využití").run()
    [button] = [b for b in at.main.button if b.key == "cap_reset"]
    assert button.label == "Obnovit výchozí nastavení"
    button.click().run()
    assert not at.exception and not at.warning
    assert values(at) == DEFAULTS
    assert content(at) == default_content                               # i výsledky jako při prvním otevření
    assert not [k for k in at.session_state if str(k).startswith("_keep_cap_")
                and at.session_state[k] != DEFAULTS.get(k[len("_keep_"):])]
    open_page(at, "dashboard")
    page(at)
    assert values(at) == DEFAULTS
    assert at.session_state["_keep_inv_fixed_cost"] == 120000            # ostatní stránky beze změny
    open_page(at, "ekonomika")
    assert at.slider(key="fe_period").value == (3, 18)


def test_invalid_kept_values_fall_back_to_defaults():
    at = set_inputs(page())
    bad = {"cap_types": ["NEEXISTUJICI"], "cap_plan_start": "2. 3. 2026", "cap_period": (5, 999),
           "cap_view": "Neznámé zobrazení", "cap_detail_type": "NEEXISTUJICI", "cap_detail_day": "15",
           "cap_req_type": "NEEXISTUJICI", "cap_req_from": 0, "cap_req_to": 400, "cap_req_qty": 99,
           "cap_req_tent": "ano", "use_detail_type": "NEEXISTUJICI"}
    open_page(at, "dashboard")
    for key, value in bad.items():
        at.session_state["_keep_" + key] = value
    page(at)
    assert not at.exception, [e.value for e in at.exception]
    assert values(at) == DEFAULTS
    at.radio(key="cap_view").set_value("Přehled využití").run()
    assert not at.exception and at.selectbox(key="use_detail_type").value == "MINI_BAGR"     # první volba


def test_valid_part_is_kept_when_other_part_is_invalid():
    at = set_inputs(page())
    open_page(at, "dashboard")
    at.session_state["_keep_cap_period"] = (0, 20)
    page(at)
    assert not at.exception
    assert values(at) == {**CHANGED, "cap_period": (1, 25), "cap_detail_day": 15}


def test_detail_type_outside_new_type_filter_falls_back():
    at = set_inputs(page())
    at.multiselect(key="cap_types").set_value(["PASOVY_DUMPER"]).run()     # detail stroje už není ve filtru
    assert not at.exception
    assert at.selectbox(key="cap_detail_type").value == "PASOVY_DUMPER"


def test_logout_clears_kept_inputs_and_new_login_starts_with_defaults():
    at = set_inputs(page())
    assert all("_keep_" + k in at.session_state for k in KEPT)
    [b for b in at.sidebar.button if b.key == "logout"][0].click().run()
    assert not at.exception
    assert not [k for k in at.session_state if str(k).startswith(("cap_", "_keep_cap_", "use_", "_keep_use_"))]
    page(login(at))
    assert values(at) == DEFAULTS


def test_no_streamlit_default_value_warning(caplog):
    logger = logging.getLogger("streamlit")
    logger.addHandler(caplog.handler)
    try:
        with caplog.at_level(logging.WARNING):
            at = set_inputs(page())
            at.slider(key="cap_period").set_value((5, 10)).run()             # den detailu se drží v rozsahu
            open_page(at, "dashboard")
            page(at)
            assert at.number_input(key="cap_detail_day").value == 10
    finally:
        logger.removeHandler(caplog.handler)
    assert not [r for r in caplog.records if "default value" in r.getMessage()]
