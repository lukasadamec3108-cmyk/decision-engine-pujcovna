"""Porovnání dvou regresních snímků (viz snapshot.py).

    python tools/regression/compare.py <před.json> <po.json> [--allow-added-numbers 7] [--strict-text]

Přísné rozdíly (návratový kód 1):
  - jakákoli změna výpočtů záložek 1–6 nebo exportu DOCX,
  - jiné popisky záložek, jiný widget (typ, klíč, popisek, výchozí hodnota, volby),
  - jiný počet grafů, tabulek nebo tlačítek exportu v záložce,
  - viditelné číslo, které zmizelo nebo se změnilo, a nové číslo (to lze povolit pro vybranou
    záložku přepínačem --allow-added-numbers, např. když nový popisek obsahuje číslo),
  - jiné souhrnné ukazatele či tabulky záložky 7 při stejných filtrech, jiný verdikt plánovače.
Informativní (návratový kód 0): změna rozbalovacích bloků a viditelného textu – při vizuálních úpravách se očekává;
přepínač --strict-text ji počítá také jako přísný rozdíl.
"""

import argparse
import collections
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import extract  # noqa: E402


def _diff_keys(a: dict, b: dict) -> list:
    return sorted(k for k in set(a) | set(b) if a.get(k) != b.get(k))


def compare(a: dict, b: dict, allow_added=(), strict_text=False) -> tuple:
    strict, info = [], []

    # výpočty a DOCX
    for scen in sorted(set(a["calc"]) | set(b["calc"])):
        ca, cb = a["calc"].get(scen, {}), b["calc"].get(scen, {})
        for k in _diff_keys(ca, cb):
            strict.append(f"výpočet {scen}: {k}" if k != "docx_parts" else
                          f"DOCX {scen}: liší se části {_diff_keys(ca.get(k, {}), cb.get(k, {}))}")

    # výchozí stav záložek
    ua, ub = a["ui"], b["ui"]
    if "pages" in ub and "pages" not in ua:          # záložky (main) × stránky (M2): jen schválené rozdíly
        info.append("M2: záložky 6–8 jsou samostatné stránky, investiční vstupy mají klíče, správa účtů je "
                    "na stránce Nastavení")
        moved = [w for w in ua["sidebar"]["widgets"] if w["key"] in extract.M2_MOVED_TO_SETTINGS]
        if moved != [w for w in ub["pages"]["nastaveni"]["main"]["widgets"]
                     if w["key"] in extract.M2_MOVED_TO_SETTINGS]:
            strict.append("správa účtů: widgety na stránce Nastavení neodpovídají původním")
        ua = extract.m2_expected(ua)
    if ua["tab_labels"] != ub["tab_labels"]:
        strict.append(f"popisky záložek: {ua['tab_labels']} → {ub['tab_labels']}")
    areas = {**{f"záložka {k}": (ua["tabs"].get(k), ub["tabs"].get(k)) for k in sorted(set(ua["tabs"]) | set(ub["tabs"]))},
             "postranní panel": (ua["sidebar"], ub["sidebar"])}
    for name, (ta, tb) in areas.items():
        if ta is None or tb is None:
            strict.append(f"{name}: chybí v jednom snímku")
            continue
        if ta["widgets"] != tb["widgets"]:
            strict.append(f"{name}: liší se widgety")
        if ta.get("elements") != tb.get("elements"):
            strict.append(f"{name}: jiný počet grafů, tabulek nebo tlačítek exportu {ta.get('elements')} → {tb.get('elements')}")
        if ta.get("expanders") != tb.get("expanders"):
            info.append(f"{name}: změnily se rozbalovací bloky {ta.get('expanders')} → {tb.get('expanders')}")
        na, nb = collections.Counter(ta["numbers"]), collections.Counter(tb["numbers"])
        removed, added = dict(na - nb), dict(nb - na)
        if removed:
            strict.append(f"{name}: čísla zmizela nebo se změnila {removed}")
        if added:
            tab_no = name.split()[-1]
            (info if tab_no in allow_added else strict).append(f"{name}: nová čísla {added}")
        if ta["text"] != tb["text"]:
            (strict if strict_text else info).append(f"{name}: změnil se viditelný text")

    # záložka 7 při filtrech
    for i, (ra, rb) in enumerate(zip(a["tab7"], b["tab7"])):
        label = f"záložka 7, filtry {ra['filters'] or 'výchozí'}"
        if ra["kpis"] != rb["kpis"]:
            strict.append(f"{label}: liší se souhrnné ukazatele")
        if ra["tables"] != rb["tables"]:
            strict.append(f"{label}: liší se tabulky")
        if rb["exceptions"]:
            strict.append(f"{label}: výjimka {rb['exceptions']}")
        na, nb = collections.Counter(ra["numbers"]), collections.Counter(rb["numbers"])
        if na - nb:
            strict.append(f"{label}: čísla zmizela nebo se změnila {dict(na - nb)}")
        if nb - na:
            (info if "7" in allow_added else strict).append(f"{label}: nová čísla {dict(nb - na)}")
    if len(a["tab7"]) != len(b["tab7"]):
        strict.append("záložka 7: jiný počet kombinací filtrů")

    if a["planner"] != b["planner"]:
        strict.append("plánovač: liší se otisky verdiktů")
    return strict, info


def main(argv=None):
    ap = argparse.ArgumentParser(description="Porovnání dvou regresních snímků Decision Engine.")
    ap.add_argument("before")
    ap.add_argument("after")
    ap.add_argument("--allow-added-numbers", nargs="*", default=[], metavar="ZÁLOŽKA",
                    help="čísla záložek, ve kterých smějí přibýt nová čísla (např. 7)")
    ap.add_argument("--strict-text", action="store_true", help="změnu viditelného textu brát jako chybu")
    args = ap.parse_args(argv)
    a = json.load(open(args.before, encoding="utf-8"))
    b = json.load(open(args.after, encoding="utf-8"))
    strict, info = compare(a, b, set(args.allow_added_numbers), args.strict_text)
    n_calc = sum(len(v) for v in a["calc"].values())
    print(f"Porovnáno: {len(a['calc'])} scénáře výpočtů ({n_calc} hodnot včetně DOCX), {len(a['ui']['tabs'])} záložek, "
          f"{len(a['tab7'])} kombinací filtrů záložky 7, {len(a['planner'])} případů plánovače.")
    for x in info:
        print("  info:", x)
    for x in strict:
        print("  ROZDÍL:", x)
    print("VÝSLEDEK:", "bez rozdílů" if not strict else f"{len(strict)} přísných rozdílů")
    return 1 if strict else 0


if __name__ == "__main__":
    sys.exit(main())
