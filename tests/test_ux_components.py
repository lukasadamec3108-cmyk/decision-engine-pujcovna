"""UX optimalizace: společné komponenty (rozbalovací detail, nadpis sekce, KPI a varianty) a to, že
informace přesunuté do detailů zůstaly na stránce (jen nejsou trvale rozbalené). Hesla jsou náhodná.

Spuštění: python -m pytest tests/test_ux_components.py
"""

import os
import re
import secrets
import sys
from pathlib import Path

from streamlit.testing.v1 import AppTest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))
import fleet_view as fv  # noqa: E402
import ui_theme as ui  # noqa: E402
from app_pages import open_page  # noqa: E402
from test_fleet_economics_page import sample_view  # noqa: E402

PASSWORD = secrets.token_urlsafe(12)


def test_more_html_is_native_details_and_escapes_texts():
    out = ui.more_html("Proč <b>", "<ul><li>obsah</li></ul>", "2 < 3")
    assert out.startswith('<details class="de-more"><summary>') and out.endswith("</details>")
    assert "Proč &lt;b&gt;" in out and "2 &lt; 3" in out                   # pevné texty se escapují
    assert "<ul><li>obsah</li></ul>" in out                                 # tělo je hotové HTML
    assert " open" not in out                                               # výchozí stav: sbaleno


def test_more_html_without_meta_and_section_title_has_no_digits_in_markup():
    assert 'class="meta"' not in ui.more_html("Detail", "x")
    for markup in (ui.more_html("Detail", "x"),):
        assert not re.findall(r"\d", re.sub(r">[^<]*<", "><", markup))      # žádné číslice v atributech


def test_global_css_contains_details_and_section_title():
    import inspect
    src = inspect.getsource(ui.inject_css)
    assert "details.de-more" in src and ".de-section-title" in src and "focus-visible" in src


def test_kpi_cards_do_not_stretch_and_sub_is_not_pinned_to_bottom():
    css = ui.dash_css("x")
    kpis = re.search(r"\.dd-kpis \{[^}]*\}", css).group(0)
    assert "align-items: start" in kpis
    sub = re.search(r"\.dd-sub \{[^}]*\}", css).group(0)
    assert "margin-top: auto" not in sub


def app():
    os.chdir(ROOT)
    at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=300)
    at.secrets["password"] = PASSWORD
    at.run()
    at.text_input[0].set_value(PASSWORD).run()
    assert not at.exception
    return at


def markdown(at):
    return " ".join(m.value for m in at.main.markdown if "<style>" not in m.value)


def plain(text):
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", text).replace("\xa0", " "))


def test_economics_details_keep_all_information():
    at = open_page(app(), "ekonomika")
    raw = markdown(at)
    text = plain(raw)
    view = sample_view()
    for h in view["headline"]:                                              # význam ukazatele v detailu karty
        assert h["meaning"] in text and h["period"] in text
    assert raw.count("Co ukazatel zahrnuje") == len(view["headline"])
    usage = view["usage_summary"]
    assert usage["composition"] in text and all(n in text for n in usage["notes"])
    flat = re.sub(r"\s+", "", plain(raw))
    for case in view["conflicts"]:                                          # varianty: změny a varování zůstaly
        for v in case["variants"]:
            for x in [v["change"], *[f"Varování: {w}" for w in v["warnings"]]]:
                if x:
                    assert re.sub(r"\s+", "", x) in flat, x
    assert "Podrobnosti varianty" in raw and "varování: 1" in raw          # počet varování vidět bez rozbalení
    assert fv.JR_INTRO in text and "Chybí k přesnějšímu srovnání" in raw
    assert "nejde o čistý zisk ani o rozhodnutí" in text                    # upozornění zůstává viditelné


def test_rozhodovaci_centrum_has_standard_header():
    at = open_page(app(), "rozhodovaci_centrum")
    raw = markdown(at)
    assert 'class="de-tab-title"' in raw and "Rozhodovací centrum" in raw
    assert "Jak plánovač odpovídá" in raw and "Provoz" in raw and "Ekonomika" in raw
    assert "záložk" not in plain(raw)                                       # žádný zastaralý název


def test_no_outdated_tab_wording_on_pages():
    at = app()
    for key in ("kapacita", "zakazky", "pronajem", "rozhodovaci_centrum", "dashboard", "ekonomika"):
        open_page(at, key)
        exps = " ".join(e.label for e in at.main.expander)
        assert not re.search(r"záložk[ayu] [0-9]", plain(markdown(at)) + exps), key
