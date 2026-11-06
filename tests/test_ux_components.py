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
    assert 'class="de-tab-title' in raw and "Rozhodovací centrum" in raw         # V3: název je i v záhlaví stránky
    assert "Jak plánovač odpovídá" in raw and "Provoz" in raw and "Ekonomika" in raw
    assert "záložk" not in plain(raw)                                       # žádný zastaralý název


def test_no_outdated_tab_wording_on_pages():
    at = app()
    for key in ("kapacita", "zakazky", "pronajem", "rozhodovaci_centrum", "dashboard", "ekonomika"):
        open_page(at, key)
        exps = " ".join(e.label for e in at.main.expander)
        assert not re.search(r"záložk[ayu] [0-9]", plain(markdown(at)) + exps), key


# ---------------------------------------------------------------------------
# Redesign MALOSTROJE V3 – záhlaví stránky
# ---------------------------------------------------------------------------

def test_header_html_has_title_period_demo_and_keeps_de_header_class():
    out = ui.header_html("Dashboard", "Podtitulek <x>", "Ukázková data", "Rozsah dat: dny plánu 1–25")
    assert 'class="de-header de-ph"' in out                                 # podle třídy ho poznají regresní nástroje
    assert ">Dashboard<" in out and "Podtitulek &lt;x&gt;" in out and "Rozsah dat: dny plánu 1–25" in out
    assert "Ukázková data" in out and "de-ph-art" in out
    assert "Rozsah dat" not in ui.header_html("X", "Y", "Ukázková data")      # bez rozsahu se štítek nevypíše


def test_header_illustration_is_own_svg_without_external_references():
    svg = (ROOT / "assets" / "zahlavi_stroj.svg").read_text(encoding="utf-8")
    assert svg.lstrip().startswith("<svg") and "<image" not in svg and "href" not in svg   # žádná vložená grafika
    assert "http" not in svg.replace("http://www.w3.org/2000/svg", "")


def test_every_page_has_header_with_page_name_and_real_data_period():
    import navigace
    at = app()
    for key, _, title, *_ in navigace.PAGES:
        open_page(at, key)
        assert not at.exception, key
        [hdr] = [m.value for m in at.main.markdown if 'class="de-header' in m.value]
        assert f">{title}<" in hdr and navigace.HEADER_SUBTITLES[key] in hdr, key
        assert "Rozsah dat: dny plánu 1–25" in hdr and "Ukázková data" in hdr, key   # skutečný rozsah ukázkových dat
        assert "Decision Engine" not in hdr, key                                      # viditelně jen MALOSTROJE


def test_duplicate_content_title_is_hidden_but_intro_stays():
    at = open_page(app(), "ekonomika")
    raw = markdown(at)
    assert 'class="de-tab-title de-dup"' in raw and "Společný ekonomický výsledek flotily" in raw
    at = open_page(at, "dashboard")
    raw = markdown(at)
    assert 'de-dup' in raw and "Souhrn výsledků sekcí" in raw                   # úvod Dashboardu zůstává
    at = open_page(at, "kapacita")
    assert 'class="de-tab-title" ' in markdown(at)                              # jiný nadpis obsahu zůstává vidět


def test_login_screen_shows_malostroje():
    os.chdir(ROOT)
    at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=300)
    at.secrets["password"] = PASSWORD
    at.run()
    raw = " ".join(m.value for m in at.main.markdown)
    assert "MALOSTROJE" in raw and "Decision Engine" not in raw and 'class="de-header' not in raw


# ---------------------------------------------------------------------------
# Redesign MALOSTROJE V4 – Dashboard podle předlohy (jen vzhled)
# ---------------------------------------------------------------------------

def test_dash_bar_has_no_digits_and_clamps_share():
    bar = ui.dash_bar(0.49, "accent")
    assert bar.count("<i></i>") == 49 and 'class="dd-bar accent"' in bar
    assert not re.findall(r"\d", bar)                                       # regrese čísel se nemění
    assert ui.dash_bar(None, "dark") == "" and ui.dash_bar(1.7, "dark").count("<i></i>") == 100


def test_original_header_art_is_served_as_static_file_only_on_dashboard():
    from PIL import Image
    assert ui.header_art_src().startswith("data:image/svg+xml")             # ostatní stránky: zástupná ilustrace
    assert ui.header_art_src(original=True) == "app/static/zahlavi_dashboard.webp"
    orig, web = Image.open(ROOT / "assets" / "zahlavi_original.png"), Image.open(ui.HEADER_ART_STATIC)
    assert web.format == "WEBP" and web.mode == "RGBA" and orig.mode == "LA"           # průhledné pozadí
    assert abs(web.width / web.height - orig.width / orig.height) < 0.01                # stejné proporce
    assert sorted(p.suffix for p in (ROOT / "static").iterdir()) == [".webp"]           # ve static/ jen obrázek
    assert "enableStaticServing = true" in (ROOT / ".streamlit" / "config.toml").read_text(encoding="utf-8")


def test_dashboard_header_uses_original_art_other_pages_do_not():
    at = open_page(app(), "dashboard")
    [hdr] = [m.value for m in at.main.markdown if 'class="de-header' in m.value]
    assert "app/static/zahlavi_dashboard.webp" in hdr and len(hdr) < 5000            # obrázek není vložený do HTML
    at = open_page(at, "kapacita")
    [hdr] = [m.value for m in at.main.markdown if 'class="de-header' in m.value]
    assert "zahlavi_dashboard" not in hdr and "data:image/svg+xml" in hdr


def test_dashboard_v4_layout_keeps_values_rows_and_links():
    at = open_page(app(), "dashboard")
    raw = markdown(at)
    assert 'class="dd-kpi alert"' in raw                                     # existující konflikt → růžová karta
    assert raw.count('class="dd-bar ') == 4                                  # pruhy: pronájmy, zakázky, servis, volná
    for text in ("Nejvytíženější stroj", "Zakázky po termínu", "Zisk zakázek po penále", "Doporučená sestava",
                 "předběžné rezervace", "stav k dnešku nelze určit", "nejde o čistý zisk", "Souhrn výsledků sekcí",
                 "Doporučení a upozornění"):
        assert text in raw, text
    assert ".dd-kpi.alert" in " ".join(m.value for m in at.main.markdown)
