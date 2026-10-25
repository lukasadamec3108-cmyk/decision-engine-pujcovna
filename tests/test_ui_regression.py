"""Regresní test rozhraní: výchozí stav všech původních záložek, postranního panelu a stránek proti základu.

Chrání hlavně záložky 1–5, které jinak nemají testy rozhraní. Kontroluje:
  - popisky záložek (Investice mají stejné záložky 1–5 jako dřív),
  - všechna viditelná čísla v každé oblasti (žádné nesmí zmizet, změnit se ani přibýt),
  - widgety (typ, klíč, popisek, výchozí hodnota, volby),
  - počet grafů, tabulek a tlačítek exportu DOCX v každé oblasti.
Viditelný text a rozbalovací bloky se neporovnávají (při vizuálních úpravách se smějí měnit);
v základu jsou jen pro informaci.

Etapa M2 (vícestránková aplikace): původní záložky 1–8 se hledají tam, kam se přesunuly
(navigace.ORIGINAL_TABS) a porovnávají se s PŮVODNÍM základem z verze se záložkami
(tests/fixtures/ui_baseline.json – v M2 se nemění). Jediné povolené rozdíly jsou vyjmenované níže:
  - investiční vstupy v panelu dostaly klíče (inv_*), aby se jejich hodnoty zachovaly mezi stránkami,
  - správa účtů (admin_pw1, admin_pw2) se přesunula z panelu na stránku Nastavení – beze změny.
Nové stránky (Dashboard, Pronájem strojů, Strojový park, Nastavení) a panel každé stránky hlídá
základ tests/fixtures/ui_pages_baseline.json. Po ZÁMĚRNÉ a schválené změně ho obnovte:
    python tests/test_ui_regression.py --update
Heslo je náhodné, jen pro test.
"""

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tools" / "regression"))
import extract  # noqa: E402
import navigace  # noqa: E402

BASELINE = ROOT / "tests" / "fixtures" / "ui_baseline.json"            # původní aplikace se záložkami
PAGES_BASELINE = ROOT / "tests" / "fixtures" / "ui_pages_baseline.json"  # stránky M2
_STATE = {}

# Povolené a schválené rozdíly proti původnímu panelu (etapa M2) – definice v tools/regression/extract.py
MOVED_TO_SETTINGS = extract.M2_MOVED_TO_SETTINGS


def current_state() -> dict:
    if "ui" not in _STATE:
        _STATE["ui"] = extract.ui_state(extract.logged_in_app(ROOT))
    return _STATE["ui"]


def baseline() -> dict:
    return json.loads(BASELINE.read_text(encoding="utf-8"))


def expected_original() -> dict:
    """Původní základ upravený jen o povolené rozdíly M2 (viz docstring a extract.m2_expected)."""
    return extract.m2_expected(baseline())


def _area(state, name):
    return state["sidebar"] if name == "sidebar" else state["tabs"][name]


AREAS = [str(i) for i in range(1, extract.TAB_COUNT + 1)] + ["sidebar"]


def _numbers_diff(got, want):
    missing = {k: v - got.get(k, 0) for k, v in want.items() if got.get(k, 0) < v}
    extra = {k: v - want.get(k, 0) for k, v in got.items() if want.get(k, 0) < v}
    return missing, extra


def test_tab_labels_unchanged():
    labels = baseline()["tab_labels"]
    assert current_state()["tab_labels"] == labels[:5] == navigace.INVESTICE_TABS      # záložky Investic
    assert len(labels) == len(navigace.ORIGINAL_TABS)                                  # 6–8 → vlastní stránky
    for key in navigace.KEYS:
        want = navigace.INVESTICE_TABS if key == "investice" else []
        assert current_state()["pages"][key]["tab_labels"] == want, key


def test_visible_numbers_unchanged_in_every_tab():
    now, base = current_state(), expected_original()
    for name in AREAS:
        missing, extra = _numbers_diff(_area(now, name)["numbers"], _area(base, name)["numbers"])
        assert not missing and not extra, f"oblast {name}: chybí {missing}, navíc {extra}"


def test_widgets_unchanged_in_every_tab():
    now, base = current_state(), expected_original()
    for name in AREAS:
        assert _area(now, name)["widgets"] == _area(base, name)["widgets"], f"oblast {name}: liší se widgety"


def test_moved_admin_widgets_unchanged_on_settings_page():
    moved = [w for w in baseline()["sidebar"]["widgets"] if w["key"] in MOVED_TO_SETTINGS]
    assert len(moved) == 2
    on_page = [w for w in current_state()["pages"]["nastaveni"]["main"]["widgets"] if w["key"] in MOVED_TO_SETTINGS]
    assert on_page == moved
    assert "Správa účtů (jen administrátor)" in current_state()["pages"]["nastaveni"]["main"]["expanders"]


def test_charts_tables_and_export_present_in_every_tab():
    now, base = current_state(), expected_original()
    for name in AREAS:
        assert _area(now, name)["elements"] == _area(base, name)["elements"], f"oblast {name}: chybí nebo přibyl prvek"


def test_no_widget_lost_anywhere():
    """Každý původní widget (záložky i panel) je v nové aplikaci na nějaké stránce – nic nezmizelo."""
    now = current_state()
    found = [w for p in now["pages"].values() for part in ("main", "sidebar") for w in p[part]["widgets"]]
    base = expected_original()
    for name in AREAS:
        for w in _area(base, name)["widgets"]:
            assert w in found, f"oblast {name}: chybí widget {w['label']}"
    for w in baseline()["sidebar"]["widgets"]:
        if w["key"] in MOVED_TO_SETTINGS:
            assert w in found


def test_pages_match_pages_baseline():
    base = json.loads(PAGES_BASELINE.read_text(encoding="utf-8"))
    now = current_state()["pages"]
    assert sorted(now) == sorted(base) == sorted(navigace.KEYS)
    for key in navigace.KEYS:
        for part in ("main", "sidebar"):
            got, want = now[key][part], base[key][part]
            missing, extra = _numbers_diff(got["numbers"], want["numbers"])
            assert not missing and not extra, f"stránka {key}/{part}: chybí {missing}, navíc {extra}"
            assert got["widgets"] == want["widgets"], f"stránka {key}/{part}: liší se widgety"
            assert got["elements"] == want["elements"], f"stránka {key}/{part}: liší se prvky"


def test_baseline_covers_all_tabs():
    base = baseline()
    assert len(base["tab_labels"]) == extract.TAB_COUNT
    for i in range(1, 6):                                          # záložky 1–5 musí mít co hlídat
        assert sum(base["tabs"][str(i)]["numbers"].values()) > 10, f"záložka {i}: základ bez čísel"


if __name__ == "__main__":
    if "--update" in sys.argv:
        PAGES_BASELINE.write_text(json.dumps(current_state()["pages"], ensure_ascii=False, indent=1, sort_keys=True),
                                  encoding="utf-8")
        print(f"Základ stránek uložen: {PAGES_BASELINE} (původní základ {BASELINE.name} se nemění)")
    else:
        print(__doc__)
