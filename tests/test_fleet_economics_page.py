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


def expected_kpis(view):
    """V5 (schváleno): celkový výsledek → pronájmy → zakázky → kapacitní konflikty (fleet_view.kpis „konflikty“).
    Fixní náklady flotily jsou v rozpadu výsledku (row_value)."""
    head = {h["key"]: h for h in view["headline"]}
    out = {head[k]["label"]: head[k]["text"].replace("\xa0", " ") for k in ("celkem", "pronajem", "zakazky")}
    out["Kapacitní konflikty"] = next(k for k in view["kpis"] if k["key"] == "konflikty")["text"].replace("\xa0", " ")
    return out


def fixed_row(view):
    """Řádek „Fixní náklady flotily“ v rozpadu výsledku (fleet_view.components). Ověří, že jde o stejnou částku
    jako dřívější KPI karta (headline „fixni“) – v rozpadu je jen se znaménkem odečtu."""
    row = next(r for g in view["components"] for r in g["rows"] if r["item"] == "Fixní náklady flotily")
    head = {h["key"]: h for h in view["headline"]}["fixni"]
    assert row["text"].lstrip("−") == head["text"] and row["status"] == head["status"]
    return row["text"].replace("\xa0", " ")


def row_value(at, label):
    """Hodnota řádku „popisek – hodnota“ (dash_rows) podle popisku – první výskyt na stránce."""
    raw = " ".join(body(at))
    m = re.search(r'<span class="l">' + re.escape(label) + r'(?:<small>.*?</small>)?</span><span class="v">([^<]*)</span>', raw)
    return m.group(1).replace("\xa0", " ") if m else None


def details(at, summary_start):
    """Texty rozbalovacích detailů HTML (<details>), jejichž popisek začíná summary_start – v pořadí na stránce."""
    raw = " ".join(body(at))
    blocks = re.findall(r'<details class="de-more"><summary>(.*?)</summary>(.*?)</details>', raw, flags=re.S)
    return [(plain([summ]).strip(), plain([content])) for summ, content in blocks
            if plain([summ]).strip().startswith(summary_start)]


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
    at = app(users=True)
    try:                                                   # bez přihlášení stránka vůbec není registrovaná
        open_page(at, KEY)
        raise AssertionError("stránka je dostupná bez přihlášení")
    except ValueError as exc:
        assert "Known pages: ['Přihlášení']" in str(exc)
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
    assert shown == expected_kpis(view)
    assert list(shown) == ["Celkový výsledek flotily", "Výsledek z pronájmu strojů", "Výsledek z vlastních zakázek",
                           "Kapacitní konflikty"]                                     # V5: schválené pořadí
    fixed = {h["key"]: h for h in view["headline"]}["fixni"]                         # dřívější KPI karta
    assert row_value(at, fixed["label"]) == fixed_row(view)
    text = plain(body(at))
    assert fixed["meaning"] in text
    for group in view["components"]:
        for row in group["rows"]:
            assert row["item"] in text and row["text"].replace("\xa0", " ") in text, row["item"]


def test_default_view_on_sample_data():
    at = page()
    text = plain(body(at))
    assert "Ukázková data / neúplné" in text
    assert fv.NOT_NET_PROFIT in text and fv.FICTITIOUS in text
    assert list(kpi_values(at)) == ["Celkový výsledek flotily", "Výsledek z pronájmu strojů",
                                    "Výsledek z vlastních zakázek", "Kapacitní konflikty"]
    assert kpi_values(at) == {"Celkový výsledek flotily": "nejvýše 829 500 Kč",
                              "Výsledek z pronájmu strojů": "nejvýše 194 500 Kč",
                              "Výsledek z vlastních zakázek": "nejvýše 635 000 Kč",
                              "Kapacitní konflikty": "2 případy"}
    assert row_value(at, "Fixní náklady flotily") == "nezadáno"                    # V5: v rozpadu výsledku
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
    assert row_value(at, "Fixní náklady flotily") == fixed_row(view)
    assert shown == expected_kpis(view)
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
    assert row_value(at, "Fixní náklady flotily") == fixed_row(view)
    assert shown == expected_kpis(view)
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

def test_comparison_is_information_only_and_says_when_not_possible():
    at = page()
    blocks = details(at, "Porovnání variant")         # V5: detail HTML (uvnitř rozbalovací sekce nesmí být st.expander)
    assert [label for label, _ in blocks] == ["Porovnání variant – jen informace (0)",
                                              "Porovnání variant – jen informace (3)"]
    j005 = blocks[0][1]
    assert "Varianty nelze porovnat – vyčíslit lze jen jednu variantu" in j005
    p05 = blocks[1][1]
    assert "Varianty nelze spolehlivě porovnat" in p05
    assert p05.count("nelze porovnat") == 3 and "storno / smluvní náhrada" in p05
    assert "Posun zakázky J006 × Zrušení pronájmu FIKT-P05" in p05
    for text in (j005, p05):
        assert fv.COMPARISON_DISCLAIMER in text
        assert "lepší" not in text                                    # na ukázkových datech nic „jistě lepší“
    assert "0 Kč" not in p05                                          # chybějící údaj není nula


def test_blocking_commitments_are_shown_for_job_shift():
    text = plain(body(page()))
    assert ("Bližší termíny blokují: pronájem FIKT-P03 (dny 14–20), servis – Kolový bagr, kus 01 "
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
    links = [p for p in at.main.get("page_link") if p.proto.label != "Zpět na Dashboard"]   # V5: návrat zvlášť
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
    assert kpi_values(at) == expected_kpis(sample_view())


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
    exp = [e for e in at.main.expander if e.label.startswith("Evidence zakázek a pronájmů")][0]
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
    assert kpi_values(at) == expected_kpis(view)


def test_ledger_lists_all_commitments_once_grouped_by_state():
    at = page()
    exp, refs = ledger_rows(at)
    assert exp.label == "Evidence zakázek a pronájmů (19)"
    assert len(refs) == 19 and len(set(refs)) == 19
    model = dm.load_model()
    assert sorted(refs) == sorted(list(model.rental_bookings["booking_id"]) + list(model.jobs["job_id"]))
    groups = re.findall(r'class="dd-group">([^<]+)</div>', " ".join(m.value for m in exp.markdown))
    assert groups == ["Jisté využití – započteno do výsledku (14)", "V konfliktu – nezapočítáno, rozhodne majitel (3)",
                      "Předběžné – nezapočítáno, jen potenciál (2)"]
    text = plain(m.value for m in exp.markdown)
    assert "termín z návrhu plánu zakázek" in text and "termín z rezervace" in text
    assert "konkrétní stroj: nelze určit (data uvádějí jen typ stroje a kusů tohoto typu je víc)" in text
    assert "konkrétní stroj: určen jednoznačně (jediný kus tohoto typu)" in text
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
    [exp] = [e for e in at.main.expander if e.label == "Metodika výpočtů"]
    text = plain(m.value for m in exp.markdown)
    for title, points in fv.METHODOLOGY:
        assert title in text and all(p in text for p in points), title
    assert len([m for m in exp.markdown if 'class="fe-points"' in m.value]) == 1      # body jako seznamy
    assert kpi_values(at)["Celkový výsledek flotily"] == "nejvýše 829 500 Kč"     # výsledek beze změny


# ---------------------------------------------------------------------------
# K3.4e3a – vzhled a mobilní zobrazení (jen prezentace)
# ---------------------------------------------------------------------------

def test_nowrap_ids_keeps_identifiers_together_and_escapes():
    import ui_theme as ui
    out = ui.nowrap_ids("Silniční válec · dny 15–18 · FIKT-P05 × J006 <b>")
    assert '<span class="nw">FIKT-P05</span>' in out and "&lt;b&gt;" in out
    assert "J006" in out and '<span class="nw">J006</span>' not in out          # bez spojovníku není co lámat
    assert '<span class="nw">KOLOVY_BAGR-01</span>' in ui.nowrap_ids("servis / údržba KOLOVY_BAGR-01 od dne 11")
    assert ui.nowrap_ids("dny 15–18 a před-běžné") == "dny 15–18 a před-běžné"      # běžná slova beze změny


def test_status_is_brief_and_details_keep_every_warning():
    """V5: na první pohled jen podstata (neúplný výsledek, počet chybějících údajů); ostatní upozornění a celý
    seznam chybějících údajů v detailu „Co chybí k přesnému výsledku“. Žádný text se neztratil."""
    at = page()
    view = sample_view()
    raw = " ".join(body(at))
    outside = plain([re.sub(r"<details.*?</details>", " ", raw, flags=re.S)])         # viditelné bez rozbalení
    [(label, inside)] = details(at, "Co chybí k přesnému výsledku")
    assert label == "Co chybí k přesnému výsledku 7 položek"
    for b in view["banners"]:
        if b["tone"] in ("bad", "warn") or b["text"] == fv.FICTITIOUS:
            assert b["text"] in outside, b["text"]                      # podstata a ukázková data vždy vidět
        else:
            assert b["text"] in inside and b["text"] not in outside, b["text"]
    assert fv.NOT_NET_PROFIT in outside and "Chybí 7 údajů nebo rozhodnutí" in outside
    assert "Ukázková data / neúplné" in outside
    assert outside.count("horní mez") >= 3                              # horní meze u hlavních ukazatelů
    for grp in view["missing"]:
        assert grp["group"] in inside and all(item in inside for item in grp["items"])
    assert 'class="de-insight' not in raw                               # V5: bez svislých proužků


def test_conflict_titles_and_details_do_not_break_identifiers():
    html_body = " ".join(body(at := page()))
    assert '<span class="nw">FIKT-P05</span> × J006' in html_body
    assert "Kolový bagr, kus 01" in html_body and "KOLOVY_BAGR-01" not in html_body   # K3.4e3b: bez kódu kusu
    assert "Silniční válec · dny 15–18 · FIKT-P05 × J006" in plain(body(at))  # viditelný text beze změny


def test_navigation_links_have_unified_labels_and_icons():
    at = page()
    buttons = {b.key: b for b in at.main.button if b.key.startswith("fe_")}
    assert buttons["fe_cal_K1"].label == "Zobrazit v kalendáři – Kolový bagr"
    assert buttons["fe_cal_K2"].label == "Zobrazit v kalendáři – Silniční válec"
    assert buttons["fe_usage_detail"].label == "Podrobné využití v Plánování kapacity"
    assert buttons["fe_usage_detail"].proto.icon == ":material/bar_chart:"
    assert buttons["fe_cal_K2"].proto.icon == ":material/calendar_month:"
    import ui_theme as ui
    for rule in ('[class*="st-key-fe_cal_"] button p', ".st-key-fe_usage_detail button p", "font-weight: 600"):
        assert rule in ui.ECON_CSS, rule


def test_mobile_rows_stack_label_above_amount():
    import ui_theme as ui
    mobile = ui.ECON_CSS.split("@media (max-width: 640px)")[-1]
    assert "flex-direction: column" in mobile and "align-items: flex-start" in mobile
    assert "flex-wrap: wrap" not in mobile                               # dřívější nejednotné zalamování
    assert ui.ECON_KEY == "de_tab_ekonomika" and ".st-key-de_tab_ekonomika .dd-row" in mobile   # jen tato stránka



# ---------------------------------------------------------------------------
# K3.4e3b – srozumitelné popisky na stránce
# ---------------------------------------------------------------------------

def test_page_has_no_technical_codes_and_merged_method_item():
    at = page()
    text = plain(body(at))
    for code in ("direct_cost_includes_machines", "servis / údržba", "kus odvozen", "termín: ", "kus: ",
                 "vyřeší konflikt: –"):
        assert code not in text, code
    [(_, missing)] = details(at, "Co chybí k přesnému výsledku")
    assert missing.count(fv.METHOD_MISSING) == 1 and fv.MACHINES_SETTING_MISSING in missing
    assert "nelze vyhodnotit · vyřešení konfliktu nelze posoudit" in text
    assert "Servis servis – Kolový bagr, kus 01, od dne 11" in text
    assert kpi_values(at) == expected_kpis(sample_view())


# ---------------------------------------------------------------------------
# Zapamatování filtrů (období, datum prvního dne, metoda fixních nákladů, počet měsíců)
# ---------------------------------------------------------------------------

FE_KEYS = ("fe_period", "fe_plan_start", "fe_method", "fe_months")


def filters(at):
    """Aktuálně zobrazené hodnoty filtrů stránky (počet měsíců jen je-li pole vidět)."""
    values = {"fe_period": [s.value for s in at.slider if s.key == "fe_period"][0],
              "fe_plan_start": [d.value for d in at.date_input if d.key == "fe_plan_start"][0],
              "fe_method": methods()[[r.value for r in at.radio if r.key == "fe_method"][0]]}
    months = [n.value for n in at.number_input if n.key == "fe_months"]
    values["fe_months"] = months[0] if months else None
    return values


def set_filters(at, period=(3, 20), plan_start=date(2026, 1, 30), method="months", months=0.75):
    [s for s in at.slider if s.key == "fe_period"][0].set_value(period)
    [d for d in at.date_input if d.key == "fe_plan_start"][0].set_value(plan_start)
    [r for r in at.radio if r.key == "fe_method"][0].set_value(method_label(method)).run()
    if method == "months":
        [n for n in at.number_input if n.key == "fe_months"][0].set_value(months).run()
    assert not at.exception, [e.value for e in at.exception]
    return at


def period_view(start, end, **kw):
    model = dm.load_model()
    return fv.build(fe.period_result(model, start, end, sample_starts(), **kw), cv.type_names(model),
                    kw.get("plan_start"))


def headline_texts(view):
    """Všechny hlavní ukazatele výpočtu (včetně fixních nákladů) + kapacitní konflikty (V5)."""
    out = {h["label"]: h["text"].replace("\xa0", " ") for h in view["headline"]}
    return {**out, **expected_kpis(view), "Fixní náklady flotily": fixed_row(view)}


def shown_values(at):
    """KPI karty + fixní náklady flotily z rozpadu výsledku (V5: dřívější KPI karta)."""
    return {**kpi_values(at), "Fixní náklady flotily": row_value(at, "Fixní náklady flotily")}


CHANGED = {"fe_period": (3, 20), "fe_plan_start": date(2026, 1, 30), "fe_method": "months", "fe_months": 0.75}
DEFAULTS = {"fe_period": (1, 25), "fe_plan_start": None, "fe_method": None, "fe_months": None}


def test_filters_survive_switching_to_other_pages_and_back():
    at = set_filters(page())
    assert filters(at) == CHANGED
    before_kpis, before_text = kpi_values(at), plain(body(at))
    for key in ("dashboard", "kapacita", "investice", "rozhodovaci_centrum"):
        open_page(at, key)
        assert not at.exception, key
        open_page(at, KEY)
        assert not at.exception and not at.warning, key
        assert filters(at) == CHANGED, key                              # nic se nevrátilo na výchozí
        assert kpi_values(at) == before_kpis and plain(body(at)) == before_text, key
    # stejné vstupy → stejný výsledek jako přímý výpočet (výpočty se nezměnily)
    view = period_view(3, 20, fixed_cost_method="months", months=0.75, plan_start=date(2026, 1, 30))
    assert shown_values(at) == headline_texts(view)


def test_calendar_method_and_date_survive_page_switch():
    at = set_filters(page(), period=(2, 18), method="calendar_days")
    open_page(at, "dashboard")
    open_page(at, KEY)
    assert not at.exception
    assert filters(at) == {"fe_period": (2, 18), "fe_plan_start": date(2026, 1, 30),
                           "fe_method": "calendar_days", "fe_months": None}
    view = period_view(2, 18, fixed_cost_method="calendar_days", plan_start=date(2026, 1, 30))
    assert shown_values(at) == headline_texts(view)


def test_months_value_is_kept_when_switching_methods():
    at = set_filters(page(), months=0.75)
    radio = lambda: [r for r in at.radio if r.key == "fe_method"][0]     # noqa: E731
    for method in (None, "calendar_days", None):
        radio().set_value(method_label(method)).run()
        assert not at.exception and not [n for n in at.number_input if n.key == "fe_months"]
    radio().set_value(method_label("months")).run()
    assert filters(at)["fe_months"] == 0.75                               # ne výchozí 1,00
    assert shown_values(at) == headline_texts(period_view(3, 20, fixed_cost_method="months", months=0.75,
                                                        plan_start=date(2026, 1, 30)))
    radio().set_value(method_label(None)).run()                          # skryté pole přežije i přechod stránky
    open_page(at, "kapacita")
    open_page(at, KEY)
    radio().set_value(method_label("months")).run()
    assert filters(at)["fe_months"] == 0.75


def test_invalid_stored_values_fall_back_to_defaults_without_error():
    at = page()
    at.session_state["_keep_fe_period"] = (5, 999)                       # mimo dostupné období
    at.session_state["_keep_fe_method"] = "Neexistující metoda"
    at.session_state["_keep_fe_plan_start"] = "30.01.2026"               # není datum
    at.session_state["_keep_fe_months"] = 0                              # mimo povolený rozsah
    open_page(at, "dashboard")
    open_page(at, KEY)
    assert not at.exception, [e.value for e in at.exception]
    assert filters(at) == DEFAULTS
    [r for r in at.radio if r.key == "fe_method"][0].set_value(method_label("months")).run()
    assert not at.exception and filters(at)["fe_months"] == 1.0
    assert shown_values(at) == headline_texts(sample_view(fixed_cost_method="months", months=1.0))


def test_valid_part_of_stored_values_is_kept_when_other_part_is_invalid():
    at = set_filters(page())
    at.session_state["_keep_fe_period"] = (0, 20)                        # neplatné jen období
    open_page(at, "dashboard")
    open_page(at, KEY)
    assert not at.exception
    assert filters(at) == {**CHANGED, "fe_period": (1, 25)}


def test_reset_button_restores_defaults_and_only_page_filters():
    at = page()
    open_page(at, "investice")
    at.number_input(key="inv_fixed_cost").set_value(120000).run()
    open_page(at, "kapacita")
    [s for s in at.slider if s.key == "cap_period"][0].set_value((2, 10)).run()
    open_page(at, KEY)
    set_filters(at)
    cap = at.session_state["cap_period"] if "cap_period" in at.session_state else None
    [button] = [b for b in at.main.button if b.key == "fe_reset"]
    assert button.label == "Obnovit výchozí nastavení"
    button.click().run()
    assert not at.exception and not at.warning
    assert filters(at) == DEFAULTS
    assert shown_values(at) == headline_texts(sample_view())               # stejné jako při prvním otevření
    [r for r in at.radio if r.key == "fe_method"][0].set_value(method_label("months")).run()
    assert filters(at)["fe_months"] == 1.0                               # i uložený počet měsíců je pryč
    # ostatní stránky beze změny
    assert at.session_state["_keep_inv_fixed_cost"] == 120000
    assert (at.session_state["cap_period"] if "cap_period" in at.session_state else None) == cap
    open_page(at, "investice")
    assert at.number_input(key="inv_fixed_cost").value == 120000


def test_fresh_session_starts_with_defaults():
    set_filters(page())
    at = page()                                                          # nová relace
    assert filters(at) == DEFAULTS
    assert shown_values(at) == headline_texts(sample_view())


def test_logout_clears_stored_filters():
    at = login_user(app(users=True), "lukas", LEGACY)
    open_page(at, KEY)
    set_filters(at)
    assert all("_keep_" + k in at.session_state for k in FE_KEYS)
    [b for b in at.sidebar.button if b.key == "logout"][0].click().run()
    assert not at.exception
    assert not [k for k in at.session_state if str(k).startswith(("fe_", "_keep_fe_"))]
    login_user(at, "honza", TESTER)
    open_page(at, KEY)
    assert not at.exception
    assert filters(at) == DEFAULTS


def test_filter_widgets_do_not_trigger_streamlit_value_warning():
    at = set_filters(page())
    open_page(at, "dashboard")
    open_page(at, KEY)
    messages = [e.value for e in at.exception] + [w.value for w in at.warning]
    assert not [m for m in messages if "default value" in str(m) or "Session State API" in str(m)]


# ---------------------------------------------------------------------------
# Redesign MALOSTROJE V5 – manažerský přehled ve stylu Dashboardu (jen vzhled)
# ---------------------------------------------------------------------------

def test_v5_back_to_dashboard_link():
    at = page()
    [back] = [p for p in at.main.get("page_link") if p.proto.label == "Zpět na Dashboard"]
    assert navigace.DEFAULT_PAGE == "dashboard" and back.proto.page == ""             # výchozí stránka = kořen adresy
    assert back.proto.icon == next(pg[3] for pg in navigace.PAGES if pg[0] == "dashboard")
    raw = " ".join(body(at))
    assert raw.index("Ukázková data / neúplné") < raw.index('class="dd-kpis"')        # označení dat nad KPI
    assert raw.index('class="dd-kpis"') < raw.index('class="fe-st"')                   # stav hned pod KPI
    assert plain(body(at)).count(fv.FICTITIOUS) == 1                                   # celá věta jen jednou (dole)
    assert plain(body(at)).count("Ukázková data / neúplné") == 1


def test_v5_details_sections_are_collapsed_and_in_approved_order():
    at = page()
    assert [(e.label, e.proto.expanded) for e in at.main.expander] == [
        ("Změnit období a fixní náklady", False), ("Využití strojů a kapacit", False),
        ("Kapacitní konflikty a varianty řešení (2)", False), ("Zakázka, nebo pronájem?", False),
        ("Evidence zakázek a pronájmů (19)", False), ("Metodika výpočtů", False)]
    keys = {w.key for w in [*at.slider, *at.date_input, *at.radio]}
    assert {"fe_period", "fe_plan_start", "fe_method"} <= keys                       # filtry beze změny


def test_v5_conflict_kpi_is_highlighted_only_when_conflicts_exist():
    at = page()
    raw = " ".join(body(at))
    assert 'class="dd-kpi alert"' in raw and raw.count('class="dd-kpi alert"') == 1
    at = page()
    [s for s in at.slider if s.key == "fe_period"][0].set_value((1, 5)).run()       # období bez konfliktu
    assert not at.exception
    assert kpi_values(at)["Kapacitní konflikty"] == "0 případů"
    assert 'class="dd-kpi alert"' not in " ".join(body(at))


def test_v5_header_uses_dashboard_illustration():
    at = page()
    [hdr] = [m.value for m in at.main.markdown if 'class="de-header' in m.value]
    assert "app/static/zahlavi_dashboard.webp" in hdr and ">Ekonomika flotily<" in hdr


def test_v5_cards_are_tinted_without_vertical_stripes():
    import ui_theme as ui
    css = ui.ECON_LAYOUT_CSS
    for tone in ("bad", "good", "neutral", "accent", "warn"):
        assert f'[class*="st-key-de_dash_"][class*="_{tone}"] {{ border-left: 1px solid' in css, tone
    assert ui.BRAND_RED_SOFT in css and "#EEF7F0" in css                            # jemně červená / zelená
    assert ".st-key-de_tab_ekonomika" in css and ".st-key-de_tab_dashboard" not in css
    assert ui.ECON_CSS.replace("</style>", "") in ui.GAPS_CSS.replace(ui.GAPS_KEY, ui.ECON_KEY)   # Chybějící údaje
    assert "ECON_LAYOUT_CSS" not in (ROOT / "app.py").read_text(encoding="utf-8").split('if PAGE == "chybejici_udaje"')[-1]

