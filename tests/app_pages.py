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
