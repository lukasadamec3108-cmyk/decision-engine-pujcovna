"""Testy vícestránkové navigace (etapa M2): přepínání sekcí, zachování vstupů, sdílené požadavky.

Heslo je náhodné, jen pro test. Stránky a jejich pořadí bere test z navigace.py.
"""

import os
import secrets
import sys
from pathlib import Path

from streamlit.testing.v1 import AppTest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))
import navigace  # noqa: E402
from app_pages import open_page  # noqa: E402

PASSWORD = secrets.token_urlsafe(12)
INV_KEYS = ("inv_machine_1", "inv_machine_2", "inv_fixed_cost", "inv_scenario", "inv_expected_days", "inv_n",
            "inv_budget")


def app():
    os.chdir(ROOT)
    at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=300)
    at.secrets["password"] = PASSWORD
    at.run()
    at.text_input[0].set_value(PASSWORD).run()
    assert not at.exception, [e.value for e in at.exception]
    return at


def ok(at):
    assert not at.exception, [e.value for e in at.exception]
    assert not at.error, [e.value for e in at.error]


def text(at):
    return " ".join(m.value for m in at.main.markdown).replace(" ", " ")


def inv_values(at):
    out = {}
    for key in INV_KEYS:
        for kind in (at.selectbox, at.number_input, at.slider):
            found = [w for w in kind if w.key == key]
            if found:
                out[key] = found[0].value
    return out



def investment_panel(at):
    return [n for n in at.main if getattr(n, "key", None) == "de_inv_inputs"][0]


def test_registry_has_all_sections_in_required_order():
    assert [p[2] for p in navigace.PAGES] == [
        "Dashboard", "Ekonomika flotily", "Pronájem strojů", "Vlastní zakázky", "Strojový park",
        "Plánování kapacity", "Rozhodovací centrum", "Investice", "Nastavení", "Chybějící údaje"]   # E3.2
    assert len(navigace.PAGES) == 10
    assert navigace.DEFAULT_PAGE == "dashboard"
    for key in navigace.KEYS:
        assert navigace.page_file(key).is_file(), key
    assert sorted(navigace.ORIGINAL_TABS) == list(range(1, 9))          # každá původní záložka má místo


def test_every_section_renders_and_has_expected_content():
    at = app()
    assert any('class="de-header' in m.value for m in at.main.markdown)        # výchozí stránka = Dashboard
    assert "Souhrn výsledků" in text(at)
    expect = {                                   # co musí na stránce být (původní obsah záložek a nové sekce)
        "ekonomika": ("Ekonomika flotily", "Ukázková data / neúplné", "Neúplný výsledek – nejde o čistý zisk."),
        "dashboard": ("Vytížení strojového parku", "Ekonomický výsledek", "Pronájmy a zakázky", "Kapacitní konflikty",
                      "Provoz půjčovny a vlastních zakázek", "Ekonomika a hospodaření", "Doporučení a upozornění",
                      "Doporučená sestava"),
        "pronajem": ("Pronájem strojů", "Nový požadavek na pronájem", "Jak plánovač rozhoduje"),
        "zakazky": ("Plán zakázek a vytížení strojů", "Harmonogram", "Detail zakázek"),
        "strojovy_park": ("Počet kusů ve flotile", "Ekonomika jednoho kusu"),
        "kapacita": ("Kapacita strojů po dnech",),
        "rozhodovaci_centrum": ("Rozhodovací plánovač",),
        "investice": (),
        "nastaveni": ("Správa účtů", "Data"),
        "chybejici_udaje": ("Chybějící údaje", "Co doplnit jako první", "Zatím není evidováno"),   # E3.2
    }
    for _ in range(2):                                                   # tam i zpět
        for key in navigace.KEYS:
            open_page(at, key)
            ok(at)
            for snippet in expect[key]:
                assert snippet in text(at), (key, snippet)
            assert any(b.key == "logout" for b in at.sidebar.button), key     # odhlášení na každé stránce
            sidebar_inputs = [w for kind in (at.sidebar.selectbox, at.sidebar.number_input, at.sidebar.slider)
                              for w in kind]
            assert not sidebar_inputs, key                                     # v panelu jen navigace a uživatel
            has_inputs = bool([w for w in at.main.number_input if w.key == "inv_fixed_cost"])
            assert has_inputs == (key == "investice"), key                    # vstupy jen v Investicích
            has_panel = any(getattr(n, "key", None) == "de_inv_inputs" for n in at.main)
            assert has_panel == (key == "investice"), key
            if key == "investice":
                assert len(at.tabs) == 0
            else:
                assert len(at.tabs) == 0, key


def test_original_functions_are_on_their_pages():
    at = app()
    assert len(at.get("page_link")) >= 4                                # odkazy z dashboardu do sekcí
    open_page(at, "investice")
    assert [d.proto.label for d in at.get("download_button")] == ["Stáhnout report (DOCX)"]   # export DOCX
    assert len(at.get("plotly_chart")) >= 4
    open_page(at, "zakazky")
    assert any("Konec (den)" in d.value.columns for d in at.dataframe)  # harmonogram / plán
    # Zachovat původní akci a schválený přímý přechod ke konfliktům pronájmů a servisu.
    assert [p.proto.label for p in at.main.get("page_link")] == [
        "Zkontrolovat pronájmy a servis v Plánování kapacity",
        "Ověřit novou zakázku v Rozhodovacím centru",
    ]
    open_page(at, "kapacita")
    assert at.radio(key="cap_view").value == "Kalendář dostupnosti"
    assert at.multiselect(key="cap_types") and at.slider(key="cap_period")
    open_page(at, "rozhodovaci_centrum")
    assert at.radio(key="dp_kind") is not None
    assert [b for b in at.main.button if b.label == "Přidat k porovnání"]


def test_investment_inputs_survive_page_switching():
    at = app()
    open_page(at, "investice")
    defaults = inv_values(at)
    assert defaults["inv_fixed_cost"] == 100000 and defaults["inv_scenario"] == "Realistický"
    assert defaults["inv_expected_days"] == 18 and defaults["inv_n"] == 1000 and defaults["inv_budget"] == 3000000
    options = at.selectbox(key="inv_machine_1").options
    changed = dict(inv_machine_1=options[2], inv_machine_2=options[3], inv_fixed_cost=120000,
                   inv_scenario="Optimistický", inv_expected_days=24, inv_n=2000, inv_budget=2500000)
    for key, value in changed.items():
        at.selectbox(key=key).set_value(value) if key in ("inv_machine_1", "inv_machine_2", "inv_scenario") else (
            at.number_input(key=key).set_value(value) if key in ("inv_fixed_cost", "inv_budget")
            else at.slider(key=key).set_value(value))
    at.run()
    ok(at)
    assert inv_values(at) == changed
    result = text(at)
    for key in navigace.KEYS:                                            # všechny ostatní stránky
        if key != "investice":
            open_page(at, key)
            ok(at)
            if key == "dashboard":                                       # dashboard počítá ze stejných vstupů
                assert f"porovnání {changed['inv_machine_1']} × {changed['inv_machine_2']}" in text(at)
    open_page(at, "investice")
    ok(at)
    assert inv_values(at) == changed                                     # nic se nevrátilo na výchozí
    assert text(at) == result                                           # stejné výsledky jako před odchodem
    fresh = app()                                                        # a stejné jako při přímém zadání
    open_page(fresh, "investice")
    for key, value in changed.items():
        fresh.selectbox(key=key).set_value(value) if key in ("inv_machine_1", "inv_machine_2", "inv_scenario") else (
            fresh.number_input(key=key).set_value(value) if key in ("inv_fixed_cost", "inv_budget")
            else fresh.slider(key=key).set_value(value))
    fresh.run()
    assert text(fresh) == result


def test_default_investment_results_identical_on_first_visit_and_after_return():
    at = app()
    open_page(at, "investice")
    first = text(at)
    open_page(at, "kapacita")
    open_page(at, "investice")
    assert text(at) == first and inv_values(at)["inv_fixed_cost"] == 100000


def test_rental_page_shares_requests_with_decision_centre():
    at = app()
    open_page(at, "pronajem")
    ok(at)
    assert not [r for r in at.radio if r.key == "dp_kind"]                 # pronájem = typ požadavku pevně
    at.selectbox(key="dp_r_type").set_value("PASOVY_DUMPER")
    at.number_input(key="dp_r_from").set_value(22)
    at.number_input(key="dp_r_to").set_value(23)
    [b for b in at.main.button if b.label == "Přidat k porovnání"][0].click().run()
    ok(at)
    assert len(at.session_state["dp_requests"]) == 1
    open_page(at, "rozhodovaci_centrum")
    ok(at)
    assert len(at.session_state["dp_requests"]) == 1 and 'class="dp-card"' in text(at)
    centre = text(at)
    open_page(at, "dashboard")
    ok(at)
    box = [m.value for m in at.main.markdown if "rozpracovaný požadavek" in m.value]
    assert box and "1 rozpracovaný požadavek" in box[0]
    ref = at.session_state["dp_requests"][0]["ref"]
    assert f"<b>{ref}</b>" in box[0] and "Provoz: " in box[0]
    op = box[0].split("Provoz: ")[1].split("<")[0].split(" · ")[0]
    assert op in centre                                          # stejný provozní štítek jako v Rozhodovacím centru


def all_widget_keys(at):
    kinds = ("selectbox", "number_input", "slider", "multiselect", "radio", "text_input", "checkbox", "button",
             "date_input", "toggle")
    return [w.key for kind in kinds for w in at.get(kind) if w.key]


def test_investment_panel_holds_each_input_exactly_once():
    at = app()
    for key in navigace.KEYS:
        open_page(at, key)
        keys = all_widget_keys(at)
        assert len(keys) == len(set(keys)), (key, keys)                  # žádný duplicitní ovládací prvek
        found = sorted(k for k in keys if k in INV_KEYS)        # pořadí v panelu hlídá test_ui_regression
        assert found == (sorted(INV_KEYS) if key == "investice" else []), (key, found)
    open_page(at, "investice")
    panel = investment_panel(at)
    assert sorted(w.key for kind in (panel.selectbox, panel.number_input, panel.slider) for w in kind) == \
        sorted(INV_KEYS)                                                  # všech 7 vstupů je v panelu
    groups = [m.value for m in panel.markdown if m.value.startswith("### ")]
    assert groups == ["### 1 · Porovnávané stroje", "### 2 · Trh", "### 3 · Investiční rozpočet", "### 4 · Náklady"]
    [editor] = [e for e in at.main.expander if e.label == "Upravit investiční scénář"]
    assert not editor.proto.expanded
    assert sorted(w.key for kind in (editor.selectbox, editor.number_input, editor.slider) for w in kind) == sorted(INV_KEYS)
    assert not panel.expander
    assert [b.key for b in at.sidebar.button] == ["logout"]



def test_investment_summary_follows_inputs():
    at = app()
    open_page(at, "investice")
    summary = [m.value for m in at.main.markdown if 'class="de-inv-summary"' in m.value]
    assert len(summary) == 1
    plain = summary[0].replace("\u00a0", " ")
    assert "Mini bagr" in plain and "Pásový dumper" in plain and "Realistický" in plain
    assert "100 000 Kč měsíčně" in plain and "Investiční rozpočet" in plain and "3 000 000 Kč" in plain
    at.selectbox(key="inv_scenario").set_value("Pesimistický")
    at.number_input(key="inv_budget").set_value(2000000).run()
    plain = [m.value for m in at.main.markdown if 'class="de-inv-summary"' in m.value][0].replace("\u00a0", " ")
    assert "Pesimistický" in plain and "2 000 000 Kč" in plain


def test_equal_machines_warning_points_to_panel():
    at = app()
    open_page(at, "investice")
    at.selectbox(key="inv_machine_2").set_value(at.selectbox(key="inv_machine_1").value).run()
    assert any("Nastavení investičního scénáře" in w.value and "různé" in w.value for w in at.warning)
    assert len(at.tabs) == 0                                          # výsledky se jako dřív nezobrazí
    assert at.selectbox(key="inv_machine_1") and at.selectbox(key="inv_machine_2")   # stroje jdou změnit


def _investice_layout(at):
    """Typy a klíče prvků hlavní plochy stránky Investice v pořadí (bez stylů a hlavičky)."""
    out = []
    for k in sorted(at.main.children):
        n = at.main.children[k]
        if n.type == "markdown" and ("<style>" in n.value or 'class="de-header' in n.value):
            continue
        out.append((n.type, getattr(n, "key", None), getattr(n, "label", None)))
    return out


def test_investment_sidebar_replaces_tabs_and_keeps_panel_before_content():
    at = app()
    open_page(at, "investice")
    layout = _investice_layout(at)
    kinds = [(t, k) for t, k, _ in layout]
    panel_i = next(i for i, (kind, _, label) in enumerate(layout) if kind == "expander" and label == "Upravit investiční scénář")
    tiles_i = kinds.index(("flex_container", "de_inv_view_tiles"))
    content_i = kinds.index(("flex_container", "de_inv_obsah"))
    assert panel_i < tiles_i < content_i
    assert not at.tabs
    assert [b.label for b in at.main.button if str(b.key).startswith("inv_view_")] == [
        label.split(" ", 1)[1] for label in navigace.INVESTICE_TABS]
    assert [e.label for e in at.sidebar.expander] == ["Přehled", "Provoz"]
    assert not at.sidebar.button(key="logout").disabled


def _content_keys(at):
    area = [n for n in at.main.children.values() if getattr(n, "key", None) == "de_inv_obsah"][0]
    return [area.children[k].key for k in sorted(area.children)]


def test_switching_investment_tab_keeps_inputs_and_shows_one_content():
    at = app()
    open_page(at, "investice")
    assert _content_keys(at) == ["de_inv_obsah_1_on"] + [f"de_inv_obsah_{i}_off" for i in range(2, 6)]
    at.number_input(key="inv_fixed_cost").set_value(130000).run()
    first = text(at)
    for i, label in enumerate(navigace.INVESTICE_TABS, start=1):
        at.query_params["pohled"] = list(navigace.INVESTICE_VIEWS)[i - 1]
        at.run()
        ok(at)
        keys = _content_keys(at)
        assert [k for k in keys if k.endswith("_on")] == [f"de_inv_obsah_{i}_on"], (label, keys)
        assert len(keys) == 5                                 # obsah všech záložek se dál vykresluje
        assert at.number_input(key="inv_fixed_cost").value == 130000
        assert len(at.tabs) == 0
    at.query_params["pohled"] = "doporuceni"
    at.run()
    assert text(at) == first                                  # stejné výsledky jako před přepínáním
    assert [d.proto.label for d in at.get("download_button")] == ["Stáhnout report (DOCX)"]


def test_investment_sidebar_callbacks_preserve_scenario_and_unknown_url_is_safe():
    at = app()
    open_page(at, "investice")
    at.button(key="inv_view_riziko").click().run()
    ok(at)
    assert at.query_params["pohled"] == "riziko"
    assert "de_inv_obsah_3_on" in _content_keys(at)
    at.number_input(key="inv_fixed_cost").set_value(130000).run()
    for i, slug in enumerate(navigace.INVESTICE_VIEWS, 1):
        at.button(key="inv_view_" + slug).click().run()
        ok(at)
        assert "de_inv_obsah_" + str(i) + "_on" in _content_keys(at)
        assert at.number_input(key="inv_fixed_cost").value == 130000
        active = [m.value for m in at.main.markdown if 'id="de-inv-active-tile-style"' in m.value]
        assert len(active) == 1 and f".st-key-inv_view_{slug} button" in active[0]
        status = [m.value for m in at.main.markdown if 'class="inv-current-view"' in m.value]
        assert len(status) == 1 and navigace.INVESTICE_VIEWS[slug].split(" ", 1)[1] in status[0]

    at.query_params["pohled"] = "<script>unknown</script>"
    at.run()
    ok(at)
    assert "de_inv_obsah_1_on" in _content_keys(at)
    assert "<script>unknown</script>" not in text(at)
