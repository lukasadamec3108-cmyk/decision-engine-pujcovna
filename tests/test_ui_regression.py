"""Regresní test rozhraní: výchozí stav všech 8 záložek a postranního panelu proti uloženému základu.

Chrání hlavně záložky 1–5, které jinak nemají testy rozhraní. Kontroluje:
  - popisky záložek,
  - všechna viditelná čísla v každé záložce (žádné nesmí zmizet, změnit se ani přibýt),
  - widgety (typ, klíč, popisek, výchozí hodnota, volby),
  - počet grafů, tabulek a tlačítek exportu DOCX v každé záložce.
Viditelný text a rozbalovací bloky se neporovnávají (při vizuálních úpravách se smějí měnit);
v základu jsou jen pro informaci.

Základ: tests/fixtures/ui_baseline.json. Po ZÁMĚRNÉ a schválené změně čísel nebo widgetů ho obnovte:
    python tests/test_ui_regression.py --update
Heslo je náhodné, jen pro test.
"""

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools" / "regression"))
import extract  # noqa: E402

BASELINE = ROOT / "tests" / "fixtures" / "ui_baseline.json"
_STATE = {}


def current_state() -> dict:
    if "ui" not in _STATE:
        _STATE["ui"] = extract.ui_state(extract.logged_in_app(ROOT))
    return _STATE["ui"]


def baseline() -> dict:
    return json.loads(BASELINE.read_text(encoding="utf-8"))


def _area(state, name):
    return state["sidebar"] if name == "sidebar" else state["tabs"][name]


AREAS = [str(i) for i in range(1, extract.TAB_COUNT + 1)] + ["sidebar"]


def test_tab_labels_unchanged():
    assert current_state()["tab_labels"] == baseline()["tab_labels"]


def test_visible_numbers_unchanged_in_every_tab():
    now, base = current_state(), baseline()
    for name in AREAS:
        got, want = _area(now, name)["numbers"], _area(base, name)["numbers"]
        missing = {k: v - got.get(k, 0) for k, v in want.items() if got.get(k, 0) < v}
        extra = {k: v - want.get(k, 0) for k, v in got.items() if want.get(k, 0) < v}
        assert not missing and not extra, f"oblast {name}: chybí {missing}, navíc {extra}"


def test_widgets_unchanged_in_every_tab():
    now, base = current_state(), baseline()
    for name in AREAS:
        assert _area(now, name)["widgets"] == _area(base, name)["widgets"], f"oblast {name}: liší se widgety"


def test_charts_tables_and_export_present_in_every_tab():
    now, base = current_state(), baseline()
    for name in AREAS:
        assert _area(now, name)["elements"] == _area(base, name)["elements"], f"oblast {name}: chybí nebo přibyl prvek"


def test_baseline_covers_all_tabs():
    base = baseline()
    assert len(base["tab_labels"]) == extract.TAB_COUNT
    for i in range(1, 6):                                          # záložky 1–5 musí mít co hlídat
        assert sum(base["tabs"][str(i)]["numbers"].values()) > 10, f"záložka {i}: základ bez čísel"


if __name__ == "__main__":
    if "--update" in sys.argv:
        BASELINE.write_text(json.dumps(current_state(), ensure_ascii=False, indent=1, sort_keys=True), encoding="utf-8")
        print(f"Základ uložen: {BASELINE}")
    else:
        print(__doc__)
