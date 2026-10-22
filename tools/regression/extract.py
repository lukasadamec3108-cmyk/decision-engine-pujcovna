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


def ui_state(at) -> dict:
    """Výchozí stav všech záložek a postranního panelu: popisky, čísla, widgety, prvky a otisk textu."""
    tabs = at.tabs
    return dict(
        tab_labels=[t.label for t in tabs],
        tabs={str(i + 1): dict(numbers=visible_numbers(tabs[i]), widgets=widgets(tabs[i]),
                               elements=element_counts(tabs[i]), expanders=expander_labels(tabs[i]),
                               text=text_fingerprint(tabs[i]))
              for i in range(len(tabs))},
        sidebar=dict(numbers=visible_numbers(at.sidebar), widgets=widgets(at.sidebar),
                     elements=element_counts(at.sidebar), expanders=expander_labels(at.sidebar),
                     text=text_fingerprint(at.sidebar)),
    )
