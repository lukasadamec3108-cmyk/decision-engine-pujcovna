"""Dashboard (etapa M3 – první fáze): jen zobrazuje existující výsledky – čísla musí odpovídat sekcím.

Heslo je náhodné, jen pro test.
"""

import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tests"))
from app_pages import open_page  # noqa: E402
from test_navigation import app, ok  # noqa: E402



def plain(value: str) -> str:
    return re.sub(r"<[^>]+>", "", value).replace("\u00a0", " ").strip()


def dash_kpi(at, label: str) -> str:
    """Hodnota hlavní KPI karty Dashboardu podle popisku (bez HTML)."""
    html_all = " ".join(m.value for m in at.main.markdown if "dd-kpis" in m.value and "<style>" not in m.value)
    found = re.search(r'dd-lbl">' + re.escape(label) + r'</span></div><div class="dd-val[^"]*">(.*?)</div>', html_all)
    assert found, label
    return plain(found.group(1))


def section_kpi(at, label: str) -> str:
    """Hodnota KPI karty v sekci (ui.kpi_cards) podle popisku (bez HTML)."""
    html_all = " ".join(m.value for m in at.main.markdown if "de-kpi" in m.value and "<style>" not in m.value)
    found = re.search(r'class="lbl">' + re.escape(label) + r'</div><div class="val[^"]*">(.*?)</div>', html_all)
    assert found, label
    return plain(found.group(1))


def test_dashboard_kpis_match_detail_sections():
    at = app()                                              # výchozí stránka = Dashboard
    ok(at)
    util = dash_kpi(at, "Vytížení strojového parku")
    result = dash_kpi(at, "Ekonomický výsledek")
    conflicts = dash_kpi(at, "Kapacitní konflikty")
    volume = dash_kpi(at, "Pronájmy a zakázky")
    assert util == "61,8–62,2 %" and result == "892 000 Kč" and conflicts == "5dní"
    assert volume == "5pronájmů · 12zakázek"
    open_page(at, "kapacita")
    assert section_kpi(at, "Dny s konfliktem") == "5"                 # kalendář ve výchozím nastavení
    at.radio(key="cap_view").set_value("Přehled využití").run()
    assert section_kpi(at, "Využití flotily") == util                  # přehled využití ve výchozím nastavení
    open_page(at, "zakazky")
    assert dash_kpi(at, "Zisk zakázek po penále") == result             # karta stránky Vlastní zakázky
    assert "12 naplánovaných zakázek" in " ".join(m.value for m in at.main.markdown)


def test_dashboard_is_transparent_about_partial_data():
    at = app()
    shown = plain(" ".join(m.value for m in at.main.markdown if "<style>" not in m.value))
    assert "dílčí údaj – jen vlastní zakázky" in shown
    assert "celkový výsledek půjčovny zatím není k dispozici" in shown
    assert "stav k dnešku nelze určit" in shown
    assert "Ukázková data" in " ".join(m.value for m in at.sidebar.markdown)


def test_dashboard_recommendation_and_conflicts_match_sections():
    at = app()
    shown = plain(" ".join(m.value for m in at.main.markdown if "<style>" not in m.value))
    kv = " ".join(m.value for m in at.main.markdown if 'class="dd-kv"' in m.value)
    assert '<div class="k">Doporučený stroj</div><div class="val">Pásový dumper</div>' in kv
    assert "Kolový bagr – den 11" in shown and "Silniční válec – den 15 – den 18" in shown
    assert "2 konflikty k řešení" in shown
    open_page(at, "kapacita")
    conflicts = [d.value for d in at.main.dataframe if "Konflikt" in d.value.columns][0]
    assert conflicts["Stroj"].tolist() == ["Kolový bagr", "Silniční válec"]
    open_page(at, "investice")
    assert any("Pásový dumper" in m.value and "de-hero" in m.value for m in at.main.markdown)


def test_dashboard_has_no_inputs_only_links():
    at = app()
    kinds = ("selectbox", "number_input", "slider", "multiselect", "radio", "text_input", "checkbox",
             "date_input", "toggle")
    assert not [w for kind in kinds for w in at.main.get(kind)]          # žádné nové ovládací prvky
    assert not at.main.button                                            # ani tlačítka
    links = [p.proto.label for p in at.get("page_link")]
    for name in ("Plánování kapacity", "Vlastní zakázky", "Pronájem strojů", "Investice", "Strojový park",
                 "Rozhodovací centrum", "Otevřít Investice", "Otevřít kalendář kapacity", "Nový pronájem"):
        assert name in links, name
