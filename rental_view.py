"""Pronájem strojů – zobrazení výsledku plánovače v kartách (modernizace UI).

Jen převádí hotový výsledek planner_view.evaluate (karty požadavků) na KPI karty a karty variant.
Nic nepočítá: hodnoty, štítky i zdůvodnění jsou z plánovače, částky se jen formátují.
"""

import html

import capacity_view as cv
import planner_view as pv
import ui_theme as ui
from formatting import fmt_czk


def dp_economy_value(card: dict) -> tuple:
    """(hodnota HTML, vysvětlení HTML) ekonomického přínosu z hotového výsledku plánovače – nic nepočítá.

    Přesný příspěvek → částka; příspěvek s nezadaným nákladem → „nejvýše“ horní mez z plánovače;
    jinak „Nelze určit“ a důvod (požadavek se nevejde / neplatný údaj / chybí údaje).
    """
    e = card["economy"]
    if e is None:
        return "Nelze určit", html.escape(card["econ_why"])
    if e["status"] == pv.EXACT:
        return fmt_czk(e["value"]), "příspěvek = tržba − variabilní náklady (nejde o čistý zisk)"
    if e["status"] == pv.WITH_UNKNOWN:
        missing = [i["label"] for i in e.get("inputs", []) if i["state"] == pv.NEZADANO and i.get("applies")]
        why = "příspěvek závisí na nezadaném údaji" + (f": {html.escape(', '.join(missing))}" if missing else "")
        if e["hi"] is None:
            return "Nelze určit", why
        return f"nejvýše {fmt_czk(e['hi'])}", why
    return "Nelze určit", html.escape("neplatný údaj: " + ("; ".join(e["invalid"]) or "neznámý"))


DP_REC_TONE = {pv.EC_REC: "good", pv.EC_NO: "bad"}      # „Méně výhodné“ = lepší je jiná varianta → neutrální


def dp_recommendation(card: dict) -> tuple:
    """(štítek, tón) výsledného doporučení – ekonomický štítek plánovače, jinak jeho souhrnný stav."""
    label = card["econ_badge"] or card["badge"]
    tone = DP_REC_TONE.get(label, "bad" if card["op_badge"] == pv.OP_NO else "info")
    return label, tone


def dp_rec_chip(label: str, tone: str) -> str:
    """Štítek doporučení: u kladného/záporného výsledku symbol ✓/✕ ze štítku, jinak původní text plánovače."""
    return ui.dash_chip(label.removeprefix("★ ") if tone != "info" else label, tone)


def dp_op_chip(card: dict) -> str:
    tone = {pv.OP_FITS: "good", pv.OP_NO: "bad"}.get(card["op_badge"], "warn")
    return ui.dash_chip(card["op_badge"].split(" ", 1)[1], tone)


def rent_kpis(card: dict, related: list) -> list:
    """Čtyři KPI karty výsledku vyhodnocení (Pronájem strojů) – jen hodnoty z výsledku plánovače."""
    start, end = card["period"]
    op = card["op_badge"]
    op_word = {pv.OP_FITS: "Vejde se", pv.OP_NO: "Nevejde se"}.get(op, "Nelze určit")
    blocked_days = sorted({b["Den"] for b in card["blocking"]})
    econ_value, econ_why = dp_economy_value(card)
    rec_label, rec_tone = dp_recommendation(card)
    if op == pv.OP_UNKNOWN:
        conf = dict(value="Nelze určit", chip=ui.dash_chip("chybí údaje", "warn"),
                    sub=html.escape(card["operational"].removeprefix("Provozně: ")))
    elif blocked_days:
        conf = dict(value=f"{len(blocked_days)}<small>{cv.plural(len(blocked_days), 'den', 'dny', 'dní').split(' ', 1)[1]}</small>",
                    chip=ui.dash_chip("kapacitu drží závazky", "bad"),
                    sub="dny " + html.escape(cv.days_text(blocked_days)) + " – kusů je méně, než požadavek potřebuje")
    else:
        conf = dict(value="Bez konfliktu", cls="word", chip=ui.dash_chip("kapacita stačí", "good"),
                    sub=f"dny {start}–{end}: dost volných kusů po potvrzených závazcích a servisu")
    if related:
        conf["sub"] += (" · " + cv.plural(len(related), "konflikt", "konflikty", "konfliktů")
                        + " potvrzených závazků na dotčených strojích (k ručnímu řešení)")
        if not blocked_days and op != pv.OP_UNKNOWN:
            conf["chip"] = ui.dash_chip("konflikt závazků k řešení", "warn")
    return [
        dict(icon="gauge", label="Dostupná kapacita", value=op_word, cls="word",
             tone={pv.OP_FITS: "good", pv.OP_NO: "bad"}.get(op, "neutral"),
             sub=html.escape(card["operational"].removeprefix("Provozně: "))),
        dict(icon="coins", label="Ekonomický přínos", tone="accent", value=econ_value,
             cls="word" if econ_value == "Nelze určit" or econ_value.startswith("nejvýše") else "",
             chip=ui.dash_chip("ztráta", "bad") if card["economy"] and card["economy"]["status"] == pv.EXACT
             and card["economy"]["value"] < 0 else "", sub=econ_why),
        dict(icon="alert", label="Kapacitní konflikt", tone="bad" if blocked_days else "good" if op == pv.OP_FITS
             else "neutral", **conf),
        dict(icon="spark", label="Výsledné doporučení", tone=rec_tone if rec_tone != "info" else "neutral",
             value=html.escape(rec_label), cls="word", sub=html.escape(card["econ_why"])),
    ]


def variant_items(cards: list, names: dict) -> list:
    """Karty porovnávaných variant (všechny požadavky včetně předběžných rezervací z dat)."""
    items = []
    for c in cards:
        econ_value, econ_why = dp_economy_value(c)
        rec_label, rec_tone = dp_recommendation(c)
        items.append(dict(
            title=c["ref"], tag="váš požadavek" if c["own"] else "předběžná rezervace z dat",
            sub=", ".join(f"{names.get(t, t)} × {q}" for t, q in c["needs"]) + f" · dny {c['period'][0]}–{c['period'][1]}",
            rows=[("Provoz", dp_op_chip(c)), ("Ekonomický přínos", econ_value),
                  ("Doporučení", dp_rec_chip(rec_label, rec_tone))],
            why=html.escape(c["econ_why"]) if c["econ_badge"] else econ_why,
        ))
    return items
