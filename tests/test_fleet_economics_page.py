"""Testy stránky „Ekonomika flotily“ (E2.2 – K3.4b kostra, K3.4c napojení výpočtu). Hesla jsou náhodná.

Stránka zobrazuje výsledek fleet_economics.period_result převedený fleet_view.build: upozornění, čtyři hlavní
ukazatele, rozpad výsledku, chybějící údaje a konflikty s variantami. Nic nevybírá ani nemění.

Spuštění: python -m pytest tests/test_fleet_economics_page.py
"""

import os
import re
import secrets
import sys
from datetime import date
from pathlib import Path

from streamlit.testing.v1 import AppTest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))
import auth  # noqa: E402
import capacity_view as cv  # noqa: E402
import datamodel as dm  # noqa: E402
import fleet_economics as fe  # noqa: E402
import fleet_view as fv  # noqa: E402
import navigace  # noqa: E402
import scheduling as sch  # noqa: E402
from app_pages import open_page  # noqa: E402

KEY = "ekonomika"
LEGACY = secrets.token_urlsafe(12)
TESTER = secrets.token_urlsafe(12)
USERS = {
    "lukas": {"name": "Lukáš", "role": "admin", "active": True, "legacy_password": True},
    "honza": {"name": "Honza Test", "role": "tester", "active": True,
              "password_hash": auth.hash_password(TESTER, iterations=100_000)},
}


def app(users=False):
    os.chdir(ROOT)
    at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=300)
    at.secrets["password"] = LEGACY
    if users:
        at.secrets["users"] = USERS
    at.run()
    return at


def page(at=None):
    at = at or app()
    if not any(b.key == "logout" for b in at.sidebar.button):
        at.text_input[0].set_value(LEGACY).run()
    open_page(at, KEY)
    assert not at.exception, [e.value for e in at.exception]
    return at


def login_user(at, user, password):
    [u] = [t for t in at.text_input if t.label == "Uživatelské jméno"]
    [p] = [t for t in at.text_input if t.label == "Heslo"]
    u.set_value(user)
    p.set_value(password).run()
    return at


def body(at):
    """Viditelný obsah stránky bez stylů a bez společné hlavičky aplikace."""
    return [m.value for m in at.main.markdown
            if not m.value.lstrip().startswith("<style") and 'class="de-header' not in m.value]


def plain(values):
    text = re.sub(r"<[^>]+>", " ", " ".join(values)).replace("\xa0", " ")
    return re.sub(r"\s+", " ", text)


def kpi_values(at):
    """{popisek ukazatele: hodnota} z KPI karet stránky."""
    html_kpis = next(m for m in body(at) if 'class="dd-kpis"' in m)
    pairs = re.findall(r'<span class="dd-lbl">([^<]+)</span></div><div class="dd-val[^"]*">([^<]*)</div>', html_kpis)
    return {label: value.replace("\xa0", " ") for label, value in pairs}


def method_label(method):
    return next(label for label, m in methods().items() if m == method)


def methods():
    at_src = (ROOT / "app.py").read_text(encoding="utf-8")
    block = at_src[at_src.index("FLEET_ECON_METHODS = {"):]
    block = block[:block.index("}") + 1]
    labels = re.findall(r'"([^"]+)": (None|"months"|"calendar_days")', block)
    return {label: (None if m == "None" else m.strip('"')) for label, m in labels}


def sample_view(**kw):
    model = dm.load_model()
    jobs = dm.legacy_jobs(model)
    plan = sch.evaluate(jobs, sch.optimize_schedule(jobs, dm.fleet_capacity(model))["starts"])
    starts = cv.job_starts_from_plan(jobs, plan)
    return fv.build(fe.period_result(model, 1, 25, starts, **kw), cv.type_names(model), kw.get("plan_start"))


# ---------------------------------------------------------------------------
# Navigace a přístup
# ---------------------------------------------------------------------------

def test_page_is_in_overview_group_directly_below_dashboard():
    assert navigace.KEYS[:2] == ["dashboard", KEY]
    assert navigace.groups()["Přehled"] == ["dashboard", KEY]
    entry = next(p for p in navigace.PAGES if p[0] == KEY)
    assert entry[2] == "Ekonomika flotily" and entry[5] == "ekonomika-flotily"
    assert navigace.page_file(KEY).is_file()
    assert navigace.DEFAULT_PAGE == "dashboard"
    assert KEY not in {k for k, _ in navigace.ORIGINAL_TABS.values()}


def test_not_reachable_without_login():
    at = open_page(app(users=True), KEY)
    assert not at.exception
    assert [t.label for t in at.text_input] == ["Uživatelské jméno", "Heslo"]
    assert "Ekonomika flotily" not in plain(m.value for m in at.markdown)
    assert not any('class="dd-kpis"' in m.value for m in at.markdown)


def test_reachable_for_logged_in_roles_like_other_overviews():
    for user, password in (("lukas", LEGACY), ("honza", TESTER)):
        at = login_user(app(users=True), user, password)
        assert any(b.key == "logout" for b in at.sidebar.button), user
        open_page(at, KEY)
        assert not at.exception and "Ekonomika flotily" in plain(body(at)), user
        assert "Celkový výsledek flotily" in plain(body(at)), user
        open_page(at, "dashboard")
        assert not at.exception and "Souhrn výsledků" in plain(body(at)), user


# ---------------------------------------------------------------------------
# Napojení na výpočet
# ---------------------------------------------------------------------------

def test_page_shows_the_same_values_as_the_computation():
    at = page()
    view = sample_view()
    shown = kpi_values(at)
    assert shown == {h["label"]: h["text"].replace("\xa0", " ") for h in view["headline"]}
    text = plain(body(at))
    for group in view["components"]:
        for row in group["rows"]:
            assert row["item"] in text and row["text"].replace("\xa0", " ") in text, row["item"]


def test_default_view_on_sample_data():
    at = page()
    text = plain(body(at))
    assert "Ukázková data / neúplné" in text
    assert fv.NOT_NET_PROFIT in text and fv.FICTITIOUS in text
    assert list(kpi_values(at)) == ["Výsledek z pronájmu strojů", "Výsledek z vlastních zakázek",
                                    "Fixní náklady flotily", "Celkový výsledek flotily"]
    assert kpi_values(at) == {"Výsledek z pronájmu strojů": "nejvýše 194 500 Kč",
                              "Výsledek z vlastních zakázek": "nejvýše 635 000 Kč",
                              "Fixní náklady flotily": "nezadáno",
                              "Celkový výsledek flotily": "nejvýše 829 500 Kč"}
    assert "Období: dny 1–25 (25 dní)" in text
    for section in ("Rozpad výsledku", "Co chybí k přesnému výsledku", "Kapacitní konflikty a možné varianty"):
        assert section in text
    assert "Mini bagr: variabilní náklad pronájmu/den" in text and "řešení konfliktů (FIKT-P05, J005, J006)" in text
    for status in ("horní mez", "chybí údaje", "nezapočítáno", "podmíněný", "úplný"):
        assert status in text, status


def test_default_period_matches_capacity_planning():
    at = page()
    economics = [s for s in at.slider if s.key == "fe_period"][0].value
    open_page(at, "kapacita")
    capacity = [s for s in at.slider if s.key == "cap_period"][0].value
    assert economics == capacity == (1, 25)


def test_default_fixed_cost_method_is_not_chosen():
    at = page()
    radio = [r for r in at.radio if r.key == "fe_method"][0]
    assert methods()[radio.value] is None
    assert not [n for n in at.number_input if n.key == "fe_months"]
    assert "metoda přepočtu na období není zvolena" in plain(body(at))


def test_explicit_months_method():
    at = page()
    [r for r in at.radio if r.key == "fe_method"][0].set_value(method_label("months")).run()
    [n for n in at.number_input if n.key == "fe_months"][0].set_value(25 / 30).run()
    assert not at.exception
    shown = kpi_values(at)
    view = sample_view(fixed_cost_method="months", months=25 / 30)             # přesně zadaná hodnota
    assert shown["Fixní náklady flotily"] == {h["key"]: h for h in view["headline"]}["fixni"]["text"].replace("\xa0", " ")
    assert shown["Celkový výsledek flotily"].startswith("nejvýše")
    assert "režie není známá" in plain(body(at))


def test_calendar_method_needs_plan_start():
    at = page()
    [r for r in at.radio if r.key == "fe_method"][0].set_value(method_label("calendar_days")).run()
    text = plain(body(at))
    assert "Výsledek nelze spočítat" in text and "plan_start" in text
    assert not any('class="dd-kpis"' in m for m in body(at))                  # žádná vymyšlená čísla
    [d for d in at.date_input if d.key == "fe_plan_start"][0].set_value(date(2026, 1, 30)).run()
    assert not at.exception
    shown = kpi_values(at)
    view = sample_view(fixed_cost_method="calendar_days", plan_start=date(2026, 1, 30))
    assert shown["Fixní náklady flotily"] == {h["key"]: h for h in view["headline"]}["fixni"]["text"].replace("\xa0", " ")
    assert "30. 1. 2026" in plain(body(at))


def test_conflicts_and_variants_are_shown_not_decided():
    at = page()
    text = plain(body(at))
    assert "2 případy" in text
    assert "Silniční válec · dny 15–18 · FIKT-P05 × J006" in text and "Kolový bagr · den 11 · J005 × servis" in text
    for title in ("Posun zakázky", "Zrušení pronájmu", "Zkrácení pronájmu", "Posun servisu",
                  "Externí zapůjčení / náhradní stroj", "Ponechat k ručnímu rozhodnutí"):
        assert title in text, title
    assert "Rozhodnutí je na majiteli půjčovny" in text and "nic se neruší ani nepřesouvá" in text
    assert "doporuč" not in text.lower()                                       # nic se nedoporučuje


def test_incomplete_result_is_never_called_profit():
    at = page()
    text = plain(body(at)).lower()
    assert text.count("zisk") == text.count("nejde o čistý zisk") >= 1
