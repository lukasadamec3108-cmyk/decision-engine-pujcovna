"""Zobrazovací vrstva stránky „Chybějící údaje“ (etapa E3.2). Bez Streamlitu.

Převádí data_gaps.Report na souhrn, skupiny podle priority, doporučené pořadí doplnění a tabulku všech
položek. Nic nepočítá ani neodhaduje – jen pojmenuje a seskupí. Výstup jsou obyčejné hodnoty.
"""

import data_gaps as dg

TITLE = "Chybějící údaje"
INTRO = ("Které údaje v datech chybějí nebo jsou neplatné, u kterých strojů, zakázek a nastavení, proč jsou "
         "důležité a které výsledky aplikace kvůli nim nejsou přesné. Stránka je jen pro čtení – nic nedoplňuje "
         "ani neodhaduje.")
FICTITIOUS = "Ukázková (fiktivní) data – nejde o skutečná firemní čísla."
NO_FORM = ("Aplikace zatím nemá formulář pro úpravu dat – údaje se doplňují v datových souborech "
           "(složka data/model).")
PRIORITY_TITLE = {dg.HIGH: "Vysoká priorita", dg.MEDIUM: "Střední priorita", dg.LOW: "Nízká priorita"}
PRIORITY_TONE = {dg.HIGH: "bad", dg.MEDIUM: "warn", dg.LOW: "info"}
STATUS_TONE = {dg.MISSING: "warn", dg.INVALID: "bad", dg.NOT_TRACKED: "info"}
EXPLAIN = [
    ("Skutečná nula", "V datech je „0“ – aplikace s ní počítá jako s nulou (např. režie 0 Kč) a výsledek může "
                      "být přesný. Na této stránce se neobjeví."),
    ("Chybějící údaj", "Buňka je prázdná nebo řádek či sloupec v datech není. Aplikace ho nikdy nenahradí "
                       "nulou ani odhadem – výsledek je pak jen horní mez, nebo ho nelze spočítat."),
    ("Neplatná hodnota", "Text místo čísla, záporná částka, pravděpodobnost mimo 0–1, nepovolená volba nebo "
                         "duplicitní záznam. Výpočet s takovým údajem nelze provést."),
    ("Zatím není evidováno", "Datový model pro údaj nemá pole, nebo ho žádný výpočet zatím nepoužívá. Do počtu "
                             "chybějících údajů se nezapočítává."),
]


def _plural(n, one, few, many) -> str:
    return one if n == 1 else few if 2 <= n <= 4 else many


def areas_text(areas) -> str:
    return "; ".join(f"{dg.AREAS[a][0]} ({dg.AREAS[a][1]})" for a in areas) or "dnešní výpočty nemění"


def _record(g) -> str:
    if g.record_kind == "nastavení":
        return "nastavení dat"
    name = g.record_name if g.record_name and g.record_name != g.record_id else ""
    return f"{g.record_id} – {name}" if name and g.record_kind != "stroj" else (name or g.record_id)


def _status_text(g) -> str:
    return f"{g.status} („{g.value}“)" if g.status == dg.INVALID and g.value else g.status


def groups(report) -> list:
    """Skupiny položek se stejným údajem a prioritou; pořadí = doporučené pořadí doplnění.

    V rámci priority dřív údaj, který omezuje víc oblastí výpočtů, pak víc záznamů, pak pořadí v datech.
    """
    out, index = [], {}
    for i, g in enumerate(report.gaps):
        key = (g.priority, g.label)
        if key not in index:
            index[key] = len(out)
            out.append(dict(priority=g.priority, label=g.label, rule=g.rule, rule_text=dg.PRIORITY_RULES[g.rule][1],
                            why=g.why, areas=list(g.areas), areas_text=areas_text(g.areas), where=g.where,
                            items=[], first=i))
        grp = out[index[key]]
        grp["items"].append(dict(record=_record(g), record_kind=g.record_kind, status=g.status,
                                 status_text=_status_text(g), tone=STATUS_TONE[g.status]))
        for a in g.areas:
            if a not in grp["areas"]:
                grp["areas"].append(a)
        grp["areas_text"] = areas_text(grp["areas"])
    rank = {p: i for i, p in enumerate(dg.PRIORITIES)}
    out.sort(key=lambda x: (rank[x["priority"]], -len(x["areas"]), -len(x["items"]), x["first"]))
    for grp in out:
        n = len(grp["items"])
        kinds = {it["record_kind"] for it in grp["items"]}
        kind = kinds.pop() if len(kinds) == 1 else "záznam"
        word = {"stroj": ("stroj", "stroje", "strojů"), "zakázka": ("zakázka", "zakázky", "zakázek"),
                "pronájem": ("pronájem", "pronájmy", "pronájmů"), "servis": ("servis", "servisy", "servisů")}.get(
            kind, ("záznam", "záznamy", "záznamů"))
        grp["count_text"] = ("nastavení dat" if kind == "nastavení" else f"{n} {_plural(n, *word)}")
        statuses = sorted({it["status"] for it in grp["items"]})
        grp["status_text"] = " / ".join(statuses)
        grp["tone"] = STATUS_TONE[dg.INVALID if dg.INVALID in statuses else dg.MISSING]
    return out


def build(report, names=None) -> dict:
    """Obsah stránky: souhrn, pořadí doplnění, skupiny podle priority, nesledované údaje, vysvětlení, tabulka."""
    names = names or {}
    s = report.summary
    grps = groups(report)
    n_types = len(names) if names else len(s["machines"])
    kpis = [
        dict(key="celkem", label="Chybějící nebo neplatné údaje", value=str(s["total"]),
             sub=(f"z toho neplatné {s['invalid']} · " if s["invalid"] else "")
             + " · ".join(f"{PRIORITY_TITLE[p].split()[0].lower()} {s['by_priority'][p]}" for p in dg.PRIORITIES),
             negative=s["total"] > 0),
        dict(key="stroje", label="Dotčené stroje", value=str(len(s["machines"])),
             sub=f"z {n_types} typů strojů v datech" if n_types else "v datech nejsou stroje"),
        dict(key="zakazky", label="Dotčené zakázky",
             value=f"{len(s['jobs'])} z {s['jobs_total']}" if s["jobs_total"] else "0",
             sub="s vlastním chybějícím údajem nebo s neúplným příspěvkem"),
        dict(key="oblasti", label="Omezené oblasti výpočtů", value=f"{len(s['areas'])} z {len(dg.AREAS)}",
             sub=", ".join(dg.AREAS[a][0] for a in s["areas"]) or "žádná"),
    ]
    first = [dict(priority=g["priority"], text=f"{g['label']} – {g['count_text']}", why=g["rule_text"])
             for g in grps[:3]]
    by_priority = []
    for p in dg.PRIORITIES:
        pg = [g for g in grps if g["priority"] == p]
        used = {g["rule"] for g in pg} | {r.rule for r in report.gaps if r.priority == p}
        rules = [text for code, (prio, text) in dg.PRIORITY_RULES.items() if prio == p and (code in used or not pg)]
        by_priority.append(dict(priority=p, title=PRIORITY_TITLE[p], tone=PRIORITY_TONE[p], rules=rules, groups=pg,
                                count=s["by_priority"][p]))
    not_tracked = [dict(label=t["label"], status=t["status"], detail=t["detail"],
                        field_text="pole v datech je, výpočty ho nepoužívají" if t["field_exists"]
                        else "datový model pole nemá",
                        areas_text=areas_text(t["would_affect"]) if t["would_affect"] else "")
                   for t in report.not_tracked]
    table = [{"Priorita": g.priority, "Údaj": g.label, "Záznam": _record(g), "Stav": _status_text(g),
              "Proč je potřebný": g.why, "Ovlivňuje": areas_text(g.areas), "Kde doplnit": g.where}
             for g in report.gaps]
    return dict(title=TITLE, intro=INTRO, fictitious=report.fictitious, fictitious_text=FICTITIOUS, no_form=NO_FORM,
                kpis=kpis, first=first, priorities=by_priority, not_tracked=not_tracked, explain=list(EXPLAIN),
                table=table, notes=list(report.notes), summary=s)
