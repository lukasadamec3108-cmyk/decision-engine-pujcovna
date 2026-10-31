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
    # nic se nedoporučuje – slovo je jen ve výslovném „Nejde o doporučení…“ u porovnání variant
    assert "doporuč" not in text.replace(fv.COMPARISON_DISCLAIMER, "").lower()


def test_incomplete_result_is_never_called_profit():
    at = page()
    text = plain(body(at)).lower()
    assert text.count("zisk") == text.count("nejde o čistý zisk") >= 1


# ---------------------------------------------------------------------------
# K3.4d – porovnání variant, blokující závazky, odkazy
# ---------------------------------------------------------------------------

def expander(at, prefix, index):
    return [e for e in at.main.expander if e.label.startswith(prefix)][index]


def test_comparison_is_information_only_and_says_when_not_possible():
    at = page()
    labels = [e.label for e in at.main.expander if e.label.startswith("Porovnání variant")]
    assert labels == ["Porovnání variant – jen informace (0)", "Porovnání variant – jen informace (3)"]
    j005 = plain(m.value for m in expander(at, "Porovnání variant", 0).markdown)
    assert "Varianty nelze porovnat – vyčíslit lze jen jednu variantu" in j005
    p05 = plain(m.value for m in expander(at, "Porovnání variant", 1).markdown)
    assert "Varianty nelze spolehlivě porovnat" in p05
    assert p05.count("nelze porovnat") == 3 and "storno / smluvní náhrada" in p05
    assert "posun J006 × zrušení FIKT-P05" in p05
    for text in (j005, p05):
        assert fv.COMPARISON_DISCLAIMER in text
        assert "lepší" not in text                                    # na ukázkových datech nic „jistě lepší“
    assert "0 Kč" not in p05                                          # chybějící údaj není nula


def test_blocking_commitments_are_shown_for_job_shift():
    text = plain(body(page()))
    assert ("Bližší termíny blokují: pronájem FIKT-P03 (dny 14–20), servis / údržba KOLOVY_BAGR-01 od dne 11 "
            "(dny 11–12), vlastní zakázka J001 (dny 1–5), vlastní zakázka J007 (dny 1–5), vlastní zakázka J009 "
            "(dny 13–16), vlastní zakázka J011 (dny 17–20)") in text
    assert "Bližší termíny blokují: pronájem FIKT-P05 (dny 15–19)" in text


def test_calendar_link_opens_capacity_planning_filtered_to_the_machine():
    at = page()
    assert [b.key for b in at.main.button if b.key.startswith("fe_cal_")] == ["fe_cal_K1", "fe_cal_K2"]
    [b for b in at.main.button if b.key == "fe_cal_K2"][0].click().run()
    assert not at.exception and not at.error and not at.warning
    assert "Kapacita strojů po dnech" in plain(m.value for m in at.main.markdown)
    assert [m.value for m in at.multiselect if m.key == "cap_types"] == [["SILNICNI_VALEC"]]
    assert [s.value for s in at.selectbox if s.key == "cap_detail_type"] == ["SILNICNI_VALEC"]
    assert [r.value for r in at.radio if r.key == "cap_view"] == ["Kalendář dostupnosti"]
    open_page(at, KEY)                                                # zpět – stránka beze změny výsledků
    assert kpi_values(at)["Celkový výsledek flotily"] == "nejvýše 829 500 Kč"


def test_jobs_link_points_to_own_jobs_page():
    at = page()
    links = [p for p in at.main.get("page_link")]
    assert len(links) == 2                                            # oba konflikty obsahují zakázku
    assert {p.proto.label for p in links} == {"Vlastní zakázky"}
    url = next(page[5] for page in navigace.PAGES if page[0] == "zakazky")
    assert all(p.proto.page == url for p in links)                    # adresa stránky Vlastní zakázky


def test_links_do_not_change_data_or_plan():
    before = dm.data_fingerprint()
    at = page()
    [b for b in at.main.button if b.key == "fe_cal_K1"][0].click().run()
    assert dm.data_fingerprint() == before
    open_page(at, KEY)
    assert kpi_values(at) == {h["label"]: h["text"].replace("\xa0", " ") for h in sample_view()["headline"]}


# ---------------------------------------------------------------------------
# K3.4e1 – předběžné a nejisté závazky, evidence závazků
# ---------------------------------------------------------------------------

def section(at, title_start):
    """Text bloků stránky od nadpisu sekce po další nadpis sekce."""
    out, inside = [], False
    for m in body(at):
        if "Každý pronájem a každá vlastní zakázka" in m:                # začátek evidence – už jiná část
            break
        if 'class="dd-section-title"' in m:
            inside = title_start in m
            continue
        if inside:
            out.append(m)
    return plain(out)


def ledger_rows(at):
    exp = [e for e in at.main.expander if e.label.startswith("Evidence závazků")][0]
    html_rows = " ".join(m.value for m in exp.markdown)
    return exp, re.findall(r'<span class="l">([A-Z0-9-]+) – ', html_rows)


def test_pending_commitments_are_shown_as_not_counted():
    at = page()
    text = section(at, "Předběžné a nejisté závazky")
    assert "nezapočítáno – 2 závazky" in plain(body(at))
    assert fv.PENDING_NOTE in text
    for ref, period, machine, amount in (("FIKT-P06", "dny 21–25", "Pásový dumper", "nejvýše 40 000 Kč"),
                                         ("FIKT-P07", "dny 19–23", "Smykový nakladač", "nejvýše 37 500 Kč")):
        assert f"{ref} – pronájem" in text and period in text and machine in text and amount in text
    assert text.count("předběžný – jen potenciál") == 2 and text.count("nezapočítáno") >= 2
    # výsledek se nemění – předběžné částky nejsou ve výsledku
    assert kpi_values(at)["Celkový výsledek flotily"] == "nejvýše 829 500 Kč"


def test_crossing_commitments_in_shorter_period():
    at = page()
    [s for s in at.slider if s.key == "fe_period"][0].set_value((1, 14)).run()
    assert not at.exception
    text = section(at, "Předběžné a nejisté závazky")
    for ref in ("FIKT-P01", "FIKT-P03", "J009"):
        assert ref in text, ref
    assert text.count("přesahuje období – chybí pravidlo časového rozlišení") == 3
    assert "FIKT-P06" not in text                                       # předběžná rezervace je mimo období
    view = fv.build(fe.period_result(dm.load_model(), 1, 14, sample_starts()), cv.type_names(dm.load_model()))
    assert kpi_values(at) == {h["label"]: h["text"].replace("\xa0", " ") for h in view["headline"]}


def test_ledger_lists_all_commitments_once_grouped_by_state():
    at = page()
    exp, refs = ledger_rows(at)
    assert exp.label == "Evidence závazků (19)"
    assert len(refs) == 19 and len(set(refs)) == 19
    model = dm.load_model()
    assert sorted(refs) == sorted(list(model.rental_bookings["booking_id"]) + list(model.jobs["job_id"]))
    groups = re.findall(r'class="dd-group">([^<]+)</div>', " ".join(m.value for m in exp.markdown))
    assert groups == ["Jisté využití – započteno do výsledku (14)", "V konfliktu – nezapočítáno, rozhodne majitel (3)",
                      "Předběžné – nezapočítáno, jen potenciál (2)"]
    text = plain(m.value for m in exp.markdown)
    assert "termín: návrh plánu" in text and "termín: rezervace" in text and "kus: nejednoznačné přiřazení" in text
    assert "zisk" not in text.lower()


def test_ledger_in_shorter_period_keeps_every_commitment():
    at = page()
    [s for s in at.slider if s.key == "fe_period"][0].set_value((1, 14)).run()
    exp, refs = ledger_rows(at)
    assert len(refs) == len(set(refs)) == 19
    groups = re.findall(r'class="dd-group">([^<]+)</div>', " ".join(m.value for m in exp.markdown))
    assert [g.rsplit(" (", 1)[0] for g in groups] == [
        "Jisté využití – započteno do výsledku", "V konfliktu – nezapočítáno, rozhodne majitel",
        "Přesah mimo období – nezapočítáno, chybí pravidlo časového rozlišení", "Mimo zvolené období – nezapočítáno"]


def sample_starts():
    model = dm.load_model()
    jobs = dm.legacy_jobs(model)
    plan = sch.evaluate(jobs, sch.optimize_schedule(jobs, dm.fleet_capacity(model))["starts"])
    return cv.job_starts_from_plan(jobs, plan)



# ---------------------------------------------------------------------------
# K3.4e2 – souhrn využití strojů, odkaz do Plánování kapacity, metodika
# ---------------------------------------------------------------------------

def test_usage_summary_section_matches_view():
    at = page()
    text = section(at, "Využití strojů za období")
    usage = sample_view()["usage_summary"]
    assert "Využití strojů za období dny 1–25 (25 dní) · kapacita flotily 225 strojodnů" in plain(body(at))
    for row in usage["rows"]:
        assert f"{row['label']} {row['meaning']} {row['text']}" in text, row["label"]
    assert usage["composition"] in text and all(n in text for n in usage["notes"])


def test_usage_link_opens_capacity_overview_with_same_period():
    at = page()
    [s for s in at.slider if s.key == "fe_period"][0].set_value((3, 20)).run()
    [b for b in at.main.button if b.key == "fe_usage_detail"][0].click().run()
    assert not at.exception and not at.error and not at.warning
    assert [s.value for s in at.slider if s.key == "cap_period"] == [(3, 20)]
    assert [r.value for r in at.radio if r.key == "cap_view"] == ["Přehled využití"]
    assert [m.value for m in at.multiselect if m.key == "cap_types"] == [[]]
    open_page(at, KEY)                                                   # zpět – výsledky beze změny
    assert not at.exception


def test_methodology_expander():
    at = page()
    [exp] = [e for e in at.main.expander if e.label == "Jak se počítá"]
    text = plain(m.value for m in exp.markdown)
    for title, desc in fv.METHODOLOGY:
        assert title in text and desc in text, title
    assert kpi_values(at)["Celkový výsledek flotily"] == "nejvýše 829 500 Kč"     # výsledek beze změny
