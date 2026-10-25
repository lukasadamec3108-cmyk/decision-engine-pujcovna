"""Společné funkce regresních nástrojů: co je v aplikaci vidět (čísla, widgety, texty).

Používá je `tools/regression/snapshot.py` i test `tests/test_ui_regression.py`.
Nic se tu nepočítá ani nemění – jen se čte, co aplikace zobrazila v AppTest.

Heslo pro přihlášení se vždy vytvoří náhodně (secrets.token_urlsafe) a nikam se neukládá.
"""

import collections
import hashlib
import os
import re
import secrets
from pathlib import Path

# Čísla tak, jak je aplikace píše: „1 234 567“, „12,5“, „−3“ (znaménko se neukládá, stačí hodnota).
NUM = re.compile(r"\d+(?:[  ]\d{3})*(?:,\d+)?")
TAB_COUNT = 8
WIDGET_TYPES = ("multiselect", "date_input", "slider", "radio", "selectbox", "number_input", "checkbox",
                "button", "download_button", "text_input", "toggle", "select_slider", "time_input", "text_area",
                "color_picker")
# Funkční prvky, které nesmějí zmizet (počet v záložce): grafy, tabulky a tlačítko exportu
COUNTED = ("plotly_chart", "dataframe", "download_button")


def numbers(text: str) -> list:
    return [n.replace(" ", " ") for n in NUM.findall(text)]


def leaves(node):
    kids = getattr(node, "children", None)
    if kids:
        for k in (kids.values() if isinstance(kids, dict) else kids):
            yield from leaves(k)
    else:
        yield node


def visible_texts(node) -> list:
    """Texty prvků (markdown, metriky, upozornění, tabulky). Bloky <style> se vynechávají – nejsou vidět."""
    out = []
    for n in leaves(node):
        if n.type == "dataframe":
            out.append(n.value.to_csv())
            continue
        for attr in ("value", "label", "body"):
            v = getattr(n, attr, None)
            if isinstance(v, str) and "<style>" not in v:
                out.append(v)
    return out


def visible_numbers(node) -> dict:
    """Četnost každého viditelného čísla (multimnožina) – zachytí změnu i ztrátu jakékoli hodnoty."""
    return dict(collections.Counter(x for t in visible_texts(node) for x in numbers(t)))


def text_fingerprint(node) -> str:
    """Otisk viditelného textu bez HTML značek a mezer – mění se při jakékoli změně textu."""
    raw = "\n".join(re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", t)).strip() for t in visible_texts(node))
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()[:16]


def widgets(node) -> list:
    """Inventář widgetů: typ, klíč, popisek, hodnota (ve výchozím stavu = výchozí hodnota) a volby."""
    out = []
    for n in leaves(node):
        if n.type in WIDGET_TYPES:
            opts = getattr(n, "options", None)
            value = getattr(n, "value", None)
            if n.type in ("text_input", "text_area"):            # obsah textových polí (i hesel) se neukládá
                value = "vyplněno" if value else "prázdné"
            out.append(dict(type=n.type, key=getattr(n, "key", None), label=getattr(n, "label", None),
                            value=str(value), options=[str(o) for o in (opts or [])]))
    return out


def element_counts(node) -> dict:
    """Počet grafů, tabulek a tlačítek exportu – ztráta funkčního prvku se pozná i bez textu."""
    c = collections.Counter(n.type for n in leaves(node))
    return {t: c.get(t, 0) for t in COUNTED}


def expander_labels(node) -> list:
    """Popisky rozbalovacích bloků (jen pro informaci – při úpravách vzhledu se smějí měnit)."""
    return [e.label for e in getattr(node, "expander", [])]


def logged_in_app(root, app_file: str = "app.py", timeout: int = 300):
    """Spustí aplikaci v AppTest a přihlásí se náhodným jednorázovým heslem (nikam se neukládá)."""
    from streamlit.testing.v1 import AppTest

    root = Path(root).resolve()
    os.chdir(root)
    password = secrets.token_urlsafe(12)
    at = AppTest.from_file(str(root / app_file), default_timeout=timeout)
    at.secrets["password"] = password
    at.run()
    at.text_input[0].set_value(password).run()
    if at.exception:
        raise RuntimeError("Aplikace skončila výjimkou: " + "; ".join(e.value[:200] for e in at.exception))
    return at


class Area:
    """Část plochy aplikace (vybrané prvky hlavní plochy) – chová se pro funkce výše jako uzel AppTest."""

    def __init__(self, nodes):
        self.children = list(nodes)
        self.expander = [e for c in self.children
                         for e in ([c] if c.type == "expander" else []) + list(getattr(c, "expander", []))]


def page_body(at) -> Area:
    """Obsah stránky bez společné hlavičky aplikace a bloků stylů (odpovídá dřívějšímu obsahu záložky)."""
    kids = at.main.children
    nodes = [kids[k] for k in sorted(kids)]
    return Area(n for n in nodes if not (n.type == "markdown" and isinstance(n.value, str) and (
        "<style>" in n.value or 'class="de-header' in n.value)))


def area_state(node) -> dict:
    return dict(numbers=visible_numbers(node), widgets=widgets(node), elements=element_counts(node),
                expanders=expander_labels(node), text=text_fingerprint(node))


def is_page_app(at) -> bool:
    """Vícestránková aplikace (etapa M2) – odhlášení v panelu a žádné hlavní záložky na výchozí stránce."""
    return len(at.tabs) == 0 and any(b.key == "logout" for b in at.sidebar.button)


def open_page(at, key: str):
    import navigace
    at.switch_page(navigace.page_relpath(key)).run()
    if at.exception:
        raise RuntimeError(f"Stránka {key} skončila výjimkou: " + "; ".join(e.value[:200] for e in at.exception))
    return at


def ui_state(at) -> dict:
    """Výchozí stav aplikace: popisky, čísla, widgety, prvky a otisk textu každé oblasti.

    Původní aplikace se záložkami: oblasti „1“–„8“ = záložky, „sidebar“ = postranní panel.
    Vícestránková aplikace (M2): oblasti „1“–„8“ = místo, kam se původní záložka přesunula
    (navigace.ORIGINAL_TABS; Investice po záložkách), „sidebar“ = panel na stránce Investice
    (tam jsou investiční vstupy) a navíc „pages“ = každá stránka (obsah a panel) zvlášť.
    """
    if not is_page_app(at):
        tabs = at.tabs
        return dict(
            tab_labels=[t.label for t in tabs],
            tabs={str(i + 1): area_state(tabs[i]) for i in range(len(tabs))},
            sidebar=area_state(at.sidebar),
        )
    import navigace
    pages, bodies = {}, {}
    for key in navigace.KEYS:
        open_page(at, key)
        bodies[key] = page_body(at)
        pages[key] = dict(main=area_state(bodies[key]), sidebar=area_state(at.sidebar),
                          tab_labels=[t.label for t in at.tabs])
        if key == "investice":
            inv_tabs = {t.label: area_state(t) for t in at.tabs}
            inv_sidebar = area_state(at.sidebar)
    tabs = {}
    for number, (key, tab_label) in navigace.ORIGINAL_TABS.items():
        tabs[str(number)] = inv_tabs[tab_label] if tab_label else pages[key]["main"]
    return dict(tab_labels=navigace.INVESTICE_TABS, tabs=tabs, sidebar=inv_sidebar, pages=pages)


# Schválené rozdíly vícestránkové aplikace (M2) proti původní verzi se záložkami – jen v postranním panelu:
#   investiční vstupy dostaly klíče (zachování hodnot mezi stránkami) a správa účtů se přesunula na Nastavení.
M2_ADDED_KEYS = {"Stroj 1": "inv_machine_1", "Stroj 2": "inv_machine_2",
                 "Měsíční fixní náklady na jeden stroj": "inv_fixed_cost", "Scénář trhu": "inv_scenario",
                 "Investiční rozpočet": "inv_budget",
                 "Odhad poptávky pro rychlý výpočet (dny/měsíc)": "inv_expected_days",
                 "Počet simulací (měsíců)": "inv_n"}
M2_MOVED_TO_SETTINGS = ("admin_pw1", "admin_pw2")


def m2_expected(state: dict) -> dict:
    """Stav původní aplikace (se záložkami) převedený o schválené rozdíly M2 – s ním se porovnává nová verze."""
    import copy
    out = copy.deepcopy(state)
    side = []
    for w in out["sidebar"]["widgets"]:
        if w["key"] in M2_MOVED_TO_SETTINGS:
            continue
        if w["key"] is None and w["label"] in M2_ADDED_KEYS:
            w = dict(w, key=M2_ADDED_KEYS[w["label"]])
        side.append(w)
    out["sidebar"]["widgets"] = side
    out["tab_labels"] = state["tab_labels"][:5]
    return out
