"""České formátování čísel (mezera mezi tisíci, desetinná čárka, znaménko minus).

Bez závislosti na Streamlitu – používá ho rozhraní (ui_theme ho znovu vystavuje) i export DOCX.
Funkce jsou přesunuté beze změny z ui_theme.py (etapa M1).
"""

NBSP = " "


def fmt_num(value, decimals: int = 0) -> str:
    """Číslo v českém formátu: mezera mezi tisíci, desetinná čárka, znaménko minus."""
    if value is None:
        return "–"
    text = f"{abs(float(value)):,.{decimals}f}"
    text = text.replace(",", NBSP).replace(".", ",")
    return ("−" if float(value) < 0 and text.strip("0,") else "") + text


def fmt_czk(value, decimals: int = 0) -> str:
    return f"{fmt_num(value, decimals)}{NBSP}Kč"


def fmt_pct(value_percent, decimals: int = 1) -> str:
    """Formátuje hodnotu, která už je v procentech (např. 24.6 → „24,6 %“)."""
    return f"{fmt_num(value_percent, decimals)}{NBSP}%"


def fmt_days(value, decimals: int = 1) -> str:
    return f"{fmt_num(value, decimals)}{NBSP}dní"
