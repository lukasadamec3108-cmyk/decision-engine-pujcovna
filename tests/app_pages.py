"""Pomocné funkce testů pro vícestránkovou aplikaci (etapa M2).

Stránky registruje navigace.py; AppTest na ně přepíná přes soubor stránky (AppTest.switch_page).
"""

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import navigace  # noqa: E402


def open_page(at, key: str):
    """Přepne běžící AppTest na stránku (klíč z navigace.KEYS) a překreslí ji."""
    at.switch_page(navigace.page_relpath(key)).run()
    return at


def page(at):
    """Hlavní plocha aktivní stránky (dříve obsah záložky)."""
    return at.main


ECON_PAGE_START = "# EKONOMIKA FLOTILY"
ECON_PAGE_END = "# STROJOVÝ PARK"


def outside_economics_page(module: str, alias: str) -> list:
    """Řádky app.py, které používají ekonomiku flotily MIMO stránku Ekonomika flotily (K3.4c).

    Povolené je jen: import, importlib.reload, CODE_MODULES (klíč mezipaměti) a blok stránky
    „Ekonomika flotily“. Jinde se moduly fleet_usage / fleet_economics / fleet_view použít nesmějí.
    """
    import re
    lines = (ROOT / "app.py").read_text(encoding="utf-8").splitlines()
    start = next(i for i, line in enumerate(lines) if line.startswith(ECON_PAGE_START))
    end = next(i for i, line in enumerate(lines) if line.startswith(ECON_PAGE_END))
    pattern = re.compile(rf"\b({re.escape(module)}|{re.escape(alias)}\.)")
    allowed = (f"import {module} as {alias}", f"{alias} = importlib.reload({alias})", "CODE_MODULES = (")
    return [f"{i + 1}: {line}" for i, line in enumerate(lines)
            if pattern.search(line) and not (start <= i < end) and not line.strip().startswith(allowed)]
