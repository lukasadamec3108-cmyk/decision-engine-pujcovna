"""Zobrazovací vrstva stránky „Ekonomika flotily“ (etapa E2.2 – K3.4a). Bez Streamlitu.

Převádí hotový výsledek fleet_economics.period_result (K3.1–K3.3) na řádky, texty, stavy
a upozornění. NIC NEPOČÍTÁ jinak: částky, meze a stavy přebírá, jen je pojmenuje a naformátuje.
Výstup build() jsou jen obyčejné hodnoty (texty, čísla, seznamy, slovníky) – lze je uložit
do mezipaměti Streamlitu a testovat bez aplikace. Stránku zatím nic nepoužívá.

STAVY ČÁSTEK (jediné povolené označení)
---------------------------------------
  úplný        … všechny údaje známé – přesná částka (i skutečná nula „0 Kč“),
  horní mez    … chybí jen NÁKLAD (variabilní náklad, přímé náklady, penále, režie, fixní náklady) –
                 skutečná částka může být jen nižší: „nejvýše X Kč“,
  podmíněný    … platí jen za podmínky, kterou nelze ověřit (souhlas zákazníka, storno, smluvní
                 náhrada) – částka je nejvýše X Kč a podmínka se vypíše,
  chybí údaje  … nelze určit ani mez (chybí tržba / cena, neplatný údaj, nezadaný náklad),
  nezapočítáno … závazek je vidět, ale do výsledku se nepočítá (konflikt, předběžný, přesah
                 období, nejistá obsazenost, zrušený).
Neúplný výsledek se NIKDY neoznačuje jako zisk – texty používají „výsledek z jistých závazků“
a „příspěvek“ a u neúplného výsledku je vždy upozornění, že nejde o čistý zisk.
"""

import re
from datetime import date, timedelta

import fleet_economics as fe
import fleet_usage as fu
from formatting import fmt_czk, fmt_num

COMPLETE, UPPER, CONDITIONAL, MISSING, EXCLUDED = "úplný", "horní mez", "podmíněný", "chybí údaje", "nezapočítáno"
STATUSES = (COMPLETE, UPPER, CONDITIONAL, MISSING, EXCLUDED)
TONE = {COMPLETE: "good", UPPER: "warn", CONDITIONAL: "info", MISSING: "bad", EXCLUDED: "neutral"}
_RANK = {COMPLETE: 0, UPPER: 1, CONDITIONAL: 2, MISSING: 3}

NOT_NET_PROFIT = "Neúplný výsledek – nejde o čistý zisk."
FICTITIOUS = "Ukázková (fiktivní) data – nejde o skutečná firemní čísla."
CONTRIBUTION_NOTE = ("Příspěvek = tržba − variabilní náklady (u zakázky i přímé náklady a penále); "
                     "fixní náklady flotily a režie se odečítají jen jednou, ve výsledku.")

# Chybějící údaj, který NENÍ náklad: bez něj neplatí ani horní mez (štítky z economics.py).
REVENUE_MARKERS = (": tržba", "cena pronájmu/den", "cena rezervace")
SETTING_MARKERS = ("metoda přepočtu", "direct_cost_includes_machines")

STATE_TEXT = {
    fe.CONFLICT: "v konfliktu – rozhodne majitel",
    fe.TENTATIVE: "předběžný – jen potenciál",
    fe.CROSSES: "přesahuje období – chybí pravidlo časového rozlišení",
    fe.UNKNOWN: "obsazenost nelze určit",
    fe.OUTSIDE: "mimo období",
    fe.CANCELLED: "zrušený",
}
VARIANT_ORDER = ["zachovat ostatní závazky a posunout zakázku", "zachovat zakázku – zrušení celé rezervace",
                 "zachovat zakázku – zkrácení rezervace o konfliktní dny", "posun servisu",
                 "externí zapůjčení / náhradní stroj", "ponechat k ručnímu rozhodnutí"]
VARIANT_TITLE = {
    "zachovat ostatní závazky a posunout zakázku": "Posun zakázky",
    "zachovat zakázku – zrušení celé rezervace": "Zrušení pronájmu",
    "zachovat zakázku – zkrácení rezervace o konfliktní dny": "Zkrácení pronájmu",
    "posun servisu": "Posun servisu",
    "externí zapůjčení / náhradní stroj": "Externí zapůjčení / náhradní stroj",
    "ponechat k ručnímu rozhodnutí": "Ponechat k ručnímu rozhodnutí",
}


# ---------------------------------------------------------------------------
# Částky a stavy
# ---------------------------------------------------------------------------

def money(value) -> str:
    """Jednotné formátování částky („389 500 Kč“, „−20 000 Kč“, „0 Kč“); None → „–“."""
    return "–" if value is None else fmt_czk(value)


def _cell(status, text, explanation="", value=None) -> dict:
    return dict(status=status, tone=TONE[status], text=text, explanation=explanation, value=value)


def _missing_only_costs(missing) -> bool:
    return not any(m for m in missing if any(x in m for x in REVENUE_MARKERS))


def contribution_cell(result) -> dict:
    """Příspěvek (economics.Result): úplný / horní mez / chybí údaje."""
    if result is None or result.value is None or result.invalid:
        why = "; ".join(result.invalid) if result is not None and result.invalid else "neplatný údaj"
        return _cell(MISSING, "nelze určit", f"neplatný údaj: {why}")
    if result.complete:
        return _cell(COMPLETE, money(result.value), "", result.value)
    if _missing_only_costs(result.missing):
        return _cell(UPPER, f"nejvýše {money(result.value)}", "chybí: " + "; ".join(result.missing), result.value)
    return _cell(MISSING, "nelze určit", "chybí: " + "; ".join(result.missing))


def cost_cell(result) -> dict:
    """Náklad (fixní náklady, režie): úplný (i skutečná 0 Kč) / chybí údaje. Nikdy nezadaná nula."""
    if result is None or result.value is None or result.invalid:
        why = "; ".join(result.invalid) if result is not None and result.invalid else "neplatný údaj"
        return _cell(MISSING, "nelze určit", f"neplatný údaj: {why}")
    if result.complete:
        return _cell(COMPLETE, money(result.value), "", result.value)
    if result.value == 0:
        return _cell(MISSING, "nezadáno", "chybí: " + "; ".join(result.missing))
    return _cell(MISSING, f"nejméně {money(result.value)}",
                 "známá je jen část – chybí: " + "; ".join(result.missing), result.value)


def bounds_cell(lo, hi, conditional: bool = False, conditions=()) -> dict:
    """Interval (dolní, horní mez) z econ_formula.bounds → text a stav."""
    cond = "; ".join(conditions)
    if lo is not None and hi is not None and abs(lo - hi) < 0.5:
        if conditional:
            return _cell(CONDITIONAL, money(hi), "platí jen pokud: " + cond, hi)
        return _cell(COMPLETE, money(hi), "", hi)
    if hi is not None:
        if conditional:
            return _cell(CONDITIONAL, f"nejvýše {money(hi)}", "platí jen pokud: " + cond, hi)
        rng = f"rozpětí {money(lo)} až {money(hi)}" if lo is not None else "dolní mez nelze určit"
        return _cell(UPPER, f"nejvýše {money(hi)}", rng + " – chybí některé náklady", hi)
    if lo is not None:
        return _cell(MISSING, f"nejméně {money(lo)}", "horní mez nelze určit")
    return _cell(MISSING, "nelze určit", cond and ("chybí: " + cond))


def _worst(*statuses) -> str:
    known = [s for s in statuses if s in _RANK]
    return max(known, key=_RANK.get) if known else MISSING


def day_text(day, plan_start=None) -> str:
    """„den 15“, s datem dne 1 plánu „den 15 (13. 2. 2026)“."""
    if not isinstance(day, int) or isinstance(day, bool):
        return "–"
    if isinstance(plan_start, date):
        d = plan_start + timedelta(day - 1)
        return f"den {day} ({d.day}. {d.month}. {d.year})"
    return f"den {day}"


def days_text(days, plan_start=None) -> str:
    """„dny 15–18“ (souvislé), „dny 3, 5“ (nesouvislé), „den 11“."""
    days = sorted(d for d in days if isinstance(d, int))
    if not days:
        return "–"
    if len(days) == 1:
        return day_text(days[0], plan_start)
    if days == list(range(days[0], days[-1] + 1)):
        first, last = day_text(days[0], plan_start), day_text(days[-1], plan_start)
        return f"dny {first[4:]}–{last[4:]}"
    return "dny " + ", ".join(day_text(d, plan_start)[4:] for d in days)


# ---------------------------------------------------------------------------
# Celkový pohled
# ---------------------------------------------------------------------------

def build(pr, names=None, plan_start=None) -> dict:
    """Celý obsah stránky z výsledku period_result – jen obyčejné hodnoty.

    names … {type_id: název stroje} (jinak kódy), plan_start … datum dne 1 (jen pro popisky dnů).
    """
    names = dict(names or {})
    if pr is None or pr.status == fu.NOT_COMPUTABLE:
        reasons = list(pr.notes) if pr is not None else ["výsledek chybí"]
        return dict(status=MISSING, computable=False,
                    banners=[dict(tone="bad", text="Výsledek nelze spočítat: " + "; ".join(reasons))],
                    kpis=[], breakdown=[], missing=[], conflicts=[], scenarios=dict(), pending=[],
                    ledger=[], ledger_groups=[], usage=[], usage_summary=None, methodology=list(METHODOLOGY),
                    method=dict(text="", chosen=False), headline=[],
                    components=[],
                    period=_period(pr, plan_start) if pr is not None else "")
    led = pr.options.ledger
    rentals, jobs = contribution_cell(pr.rentals), contribution_cell(pr.jobs)
    fixed, overhead = cost_cell(pr.fixed_costs), cost_cell(pr.overhead)
    excluded = _excluded(led)
    result = _result_cell(pr, rentals, jobs, fixed, overhead, excluded)
    return dict(
        status=result["status"], computable=True,
        banners=_banners(pr, result),
        kpis=_kpis(pr, rentals, jobs, fixed, overhead, result),
        breakdown=_breakdown(pr, rentals, jobs, fixed, overhead, result),
        missing=_missing(pr, excluded),
        conflicts=[_conflict(case, names, plan_start) for case in pr.options.cases],
        scenarios=_scenarios(pr),
        pending=_pending(led, plan_start, names),
        ledger=_ledger(led, plan_start, names),
        ledger_groups=_ledger_groups(led, plan_start, names),
        usage_summary=_usage_summary(led.usage, names),
        methodology=list(METHODOLOGY),
        usage=_usage(led.usage, names),
        method=dict(text=basis_text(pr.fixed_cost_basis), chosen=pr.months is not None),
        period=_period(pr, plan_start),
        headline=_headline(pr, rentals, jobs, fixed, overhead, result, plan_start),
        components=_components(pr, rentals, jobs, fixed, overhead, result),
    )


# ---------------------------------------------------------------------------
# Hlavní ukazatele stránky a rozpad výsledku (K3.4c)
# ---------------------------------------------------------------------------
# Jen skládá hotové částky: příspěvky jednotlivých JISTÝCH závazků z economics.py (ledger.results)
# a jejich položky (tržba, variabilní náklady, přímé náklady, penále). Položka, kterou economics
# označí jako chybějící, se do součtu nezapočte a zůstane „nezadáno“ – nikdy jako 0.

PART_MARKERS = {
    "tržba z pronájmu": ("cena pronájmu/den", "cena rezervace"),
    "variabilní náklady pronájmu": ("variabilní náklad pronájmu/den",),
    "tržba ze zakázky": (": tržba",),
    "přímé náklady": (": přímé náklady",),
    "penále": ("penále/den",),
    "variabilní náklady strojů": ("variabilní náklad na zakázce/den", "direct_cost_includes_machines"),
}
MACHINES_IN_DIRECT = "obsaženy v přímých nákladech zakázek"


def _period(pr, plan_start=None) -> str:
    s, e = getattr(pr, "start", None), getattr(pr, "end", None)
    if not (isinstance(s, int) and isinstance(e, int)) or isinstance(s, bool) or e < s:
        return "neplatné období"
    n = e - s + 1
    return f"{days_text(list(range(s, e + 1)), plan_start)} ({n} {_plural(n, 'den', 'dny', 'dní')})"


def _certain_results(pr, kind) -> list:
    led = pr.options.ledger
    return [led.results[(kind, r.ref)] for r in led.entries.itertuples() if r.counted_as_certain and r.kind == kind]


def part_total(results, key, cost: bool = False):
    """Součet jedné položky (např. „tržba z pronájmu“) přes závazky jako economics.Result.

    Závazek, u něhož economics položku označí jako chybějící, se do součtu nezapočte a štítek
    chybějícího údaje se přenese – výsledek je pak neúplný (u nákladu „nejméně“, u tržby „nelze určit“).
    """
    value, missing, invalid = 0.0, [], []
    for r in results:
        if r.invalid or r.value is None:
            invalid += [x for x in (r.invalid or ("neplatný údaj",)) if x not in invalid]
            continue
        hit = [m for m in r.missing if any(x in m for x in PART_MARKERS[key])]
        if hit:
            missing += [m for m in hit if m not in missing]
            continue
        v = r.parts.get(key)
        if v is not None:
            value += abs(v) if cost else v
    return fe.ec.Result(None if invalid else value, tuple(missing), tuple(invalid))


def _negative(cell) -> dict:
    """Náklad jako odečítaná položka: „−440 000 Kč“ (jen u známé nenulové částky)."""
    if cell["status"] == COMPLETE and cell.get("value"):
        return dict(cell, text="−" + cell["text"])
    return cell


def _headline(pr, rentals, jobs, fixed, overhead, result, plan_start) -> list:
    period = _period(pr, plan_start)
    ovh = ("režie započtena" if overhead["status"] == COMPLETE
           else "režie není známá – nezapočtena (výsledek je nejvýše uvedená částka)")
    fixed_note = (basis_text(pr.fixed_cost_basis) if pr.months is not None
                  else "metoda přepočtu na období není zvolena – fixní náklady nejsou odečteny")
    return [
        dict(key="pronajem", label="Výsledek z pronájmu strojů", period=period, **_kpi(rentals),
             meaning="příspěvek jistých pronájmů: tržby − variabilní náklady pronájmu"),
        dict(key="zakazky", label="Výsledek z vlastních zakázek", period=period, **_kpi(jobs),
             meaning="příspěvek jistých zakázek: tržby − přímé náklady − penále − variabilní náklady strojů "
                     "(pokud nejsou v přímých nákladech)"),
        dict(key="fixni", label="Fixní náklady flotily", period=period, **_kpi(fixed),
             meaning="odečteny jednou za celou flotilu (vlastněné a leasingové kusy, včetně průměrného servisu); "
                     + fixed_note),
        dict(key="celkem", label="Celkový výsledek flotily", period=period, **_kpi(result),
             meaning="jisté příspěvky − fixní náklady − režie; " + ovh
                     + "; závazky v konfliktu a předběžné se nezapočítávají"),
    ]


def _components(pr, rentals, jobs, fixed, overhead, result) -> list:
    """Rozpad výsledku po skupinách: [dict(group, rows=[dict(item, text, status, tone, explanation, …)])]."""
    rent_r, job_r = _certain_results(pr, "rental"), _certain_results(pr, "job")

    def revenue(key, results):
        return contribution_cell(part_total(results, key)) if results else _cell(COMPLETE, money(0), "žádný jistý závazek", 0.0)

    def cost(key, results):
        if not results:
            return _cell(COMPLETE, money(0), "žádný jistý závazek", 0.0)
        return _negative(cost_cell(part_total(results, key, cost=True)))

    machine = cost("variabilní náklady strojů", job_r)
    if job_r and not any("variabilní náklady strojů" in r.parts for r in job_r) \
            and not any(any(x in m for x in PART_MARKERS["variabilní náklady strojů"]) for r in job_r for m in r.missing):
        machine = _cell(COMPLETE, MACHINES_IN_DIRECT, "podle nastavení direct_cost_includes_machines – neodečítají se znovu")
    n_rent, n_job = len(rent_r), len(job_r)
    return [
        dict(group="Pronájem strojů", note=f"{n_rent} {_plural(n_rent, 'jistý pronájem', 'jisté pronájmy', 'jistých pronájmů')}",
             rows=[dict(item="Tržby z pronájmu", **revenue("tržba z pronájmu", rent_r)),
                   dict(item="Variabilní náklady pronájmu", **cost("variabilní náklady pronájmu", rent_r)),
                   dict(item="Příspěvek pronájmu", total=True, **rentals)]),
        dict(group="Vlastní zakázky", note=f"{n_job} {_plural(n_job, 'jistá zakázka', 'jisté zakázky', 'jistých zakázek')}",
             rows=[dict(item="Tržby ze zakázek", **revenue("tržba ze zakázky", job_r)),
                   dict(item="Přímé náklady zakázek", **cost("přímé náklady", job_r)),
                   dict(item="Penále za zpoždění", **cost("penále", job_r)),
                   dict(item="Variabilní náklady strojů na zakázkách", **machine),
                   dict(item="Příspěvek vlastních zakázek", total=True, **jobs)]),
        dict(group="Flotila", note="náklady se odečítají jednou za celou flotilu",
             rows=[dict(item="Fixní náklady flotily", **_negative(fixed)),
                   dict(item="Režie", **_negative(overhead)),
                   dict(item="Celkový výsledek flotily", total=True, **result)]),
    ]


def _excluded(led) -> dict:
    """{stav: [ref]} závazků, které se do výsledku nezapočítávají (bez mimo období / zrušených)."""
    out = {}
    for r in led.entries.itertuples():
        if r.state in (fe.CONFLICT, fe.TENTATIVE, fe.CROSSES, fe.UNKNOWN):
            out.setdefault(r.state, []).append(r.ref)
    return out


def _result_cell(pr, rentals, jobs, fixed, overhead, excluded) -> dict:
    """Výsledek z jistých závazků: úplný jen tehdy, je-li úplný celý výpočet (period_result)."""
    r = pr.result
    if r is None or r.value is None or r.invalid:
        return _cell(MISSING, "nelze určit", "neplatný údaj: " + "; ".join(r.invalid if r else ()))
    not_counted = sorted(excluded.get(fe.CONFLICT, []) + excluded.get(fe.CROSSES, []) + excluded.get(fe.UNKNOWN, []))
    note = ("nezapočítáno: " + ", ".join(not_counted)) if not_counted else ""
    if pr.status == fu.COMPLETE:
        return _cell(COMPLETE, money(r.value), "", r.value)
    if MISSING in (rentals["status"], jobs["status"]):
        # chybí tržba / cena nebo je údaj neplatný → ani horní mez neplatí
        return _cell(MISSING, "nelze určit", "chybí tržba, cena nebo platný údaj některého závazku")
    if any(c["status"] != COMPLETE for c in (rentals, jobs, fixed, overhead)):
        # chybí jen náklady – ty výsledek mohou jen snížit → známá částka je horní mez
        return _cell(UPPER, f"nejvýše {money(r.value)}",
                     "; ".join(x for x in ("chybí některé náklady", note) if x), r.value)
    # náklady úplné, ale některé závazky nejsou započteny (chybí rozhodnutí) – nejde o úplný výsledek
    return _cell(MISSING, f"{money(r.value)} bez nezapočítaných závazků",
                 "chybí rozhodnutí; " + note, r.value)


def _banners(pr, result) -> list:
    out = []
    fictitious = any(getattr(x, "fictitious", False) for x in (pr.rentals, pr.jobs, pr.fixed_costs, pr.overhead))
    if fictitious:
        out.append(dict(tone="info", text=FICTITIOUS))
    if result["status"] != COMPLETE:
        n = len(_missing_flat(pr))
        out.append(dict(tone="warn", text=f"{NOT_NET_PROFIT} Chybí {n} {_plural(n, 'údaj', 'údaje', 'údajů')} "
                                          "nebo rozhodnutí (viz seznam)."))
    if pr.months is None:
        out.append(dict(tone="info", text="Fixní náklady a režie se neodečítají – metoda přepočtu na období "
                                          "není zvolena. Výsledek je proto jen horní mez."))
    if pr.options.cases:
        out.append(dict(tone="info", text="Konflikty se automaticky neřeší – nic se neruší ani nepřesouvá. "
                                          "Příspěvky závazků v konfliktu se do výsledku nezapočítávají."))
    return out


def _plural(n, one, few, many) -> str:
    return one if n == 1 else few if 2 <= n <= 4 else many


def _kpis(pr, rentals, jobs, fixed, overhead, result) -> list:
    both = pr.rentals + pr.jobs if pr.rentals is not None and pr.jobs is not None else None
    contrib = contribution_cell(both)
    costs_status = _worst(fixed["status"], overhead["status"])
    costs_text = (money(-(pr.fixed_costs.value + pr.overhead.value)) if costs_status == COMPLETE
                  else f"fixní {fixed['text']} · režie {overhead['text']}")
    n = len(pr.options.cases)
    return [
        dict(key="vysledek", label="Výsledek z jistých závazků", **_kpi(result)),
        dict(key="prispevky", label="Jisté příspěvky (pronájmy + zakázky)", **_kpi(contrib)),
        dict(key="naklady", label="Fixní náklady a režie za období", text=costs_text, status=costs_status,
             tone=TONE[costs_status], explanation=basis_text(pr.fixed_cost_basis)),
        dict(key="konflikty", label="Konflikty", text=f"{n} {_plural(n, 'případ', 'případy', 'případů')}",
             status=EXCLUDED if n else COMPLETE, tone=TONE[EXCLUDED if n else COMPLETE],
             explanation="příspěvky obou stran nezapočítány – rozhodne majitel" if n else "žádný konflikt"),
    ]


def _kpi(cell) -> dict:
    return dict(text=cell["text"], status=cell["status"], tone=cell["tone"], explanation=cell["explanation"])


def _breakdown(pr, rentals, jobs, fixed, overhead, result) -> list:
    neg = lambda c: dict(c, text=("−" + c["text"]) if c["status"] == COMPLETE and c["value"] else c["text"])
    rows = [
        dict(item="Jisté příspěvky pronájmů", **rentals),
        dict(item="Jisté příspěvky vlastních zakázek", **jobs),
        dict(item="Fixní náklady flotily za období", **neg(fixed)),
        dict(item="Režie za období", **neg(overhead)),
        dict(item="Výsledek z jistých závazků", total=True, **result),
    ]
    rows[2]["explanation"] = "; ".join(x for x in (basis_text(pr.fixed_cost_basis), fixed["explanation"],
                                                   "obsahují i průměrný servis – neodečítá se znovu") if x)
    return rows


def _missing_flat(pr) -> list:
    out = []
    for r in (pr.rentals, pr.jobs, pr.fixed_costs, pr.overhead):
        if r is not None:
            out += [m for m in r.missing if m not in out]
    out += [m for m in pr.missing if m not in out and _is_decision(m)]
    return out


def _is_decision(text) -> bool:
    return any(x in text for x in ("konflikt", "časové rozlišení", "obsazenost nelze"))


def _missing(pr, excluded) -> list:
    """Seznam chybějících údajů a rozhodnutí ve skupinách."""
    costs, settings = [], []
    for r in (pr.rentals, pr.jobs, pr.fixed_costs, pr.overhead):
        for m in (r.missing if r is not None else ()):
            target = settings if any(x in m for x in SETTING_MARKERS) else costs
            if m not in target:
                target.append(m)
    decisions = []
    if excluded.get(fe.CONFLICT):
        decisions.append("řešení konfliktů (" + ", ".join(sorted(excluded[fe.CONFLICT])) + ")")
    if excluded.get(fe.CROSSES):
        decisions.append("pravidlo časového rozlišení závazků přesahujících období ("
                         + ", ".join(sorted(excluded[fe.CROSSES])) + ")")
    if excluded.get(fe.UNKNOWN):
        decisions.append("závazky s nejistou obsazeností (" + ", ".join(sorted(excluded[fe.UNKNOWN])) + ")")
    groups = [("Náklady a ceny", costs), ("Nastavení", settings), ("Rozhodnutí majitele", decisions)]
    return [dict(group=g, items=items) for g, items in groups if items]


# ---------------------------------------------------------------------------
# Konflikty a varianty
# ---------------------------------------------------------------------------

def _conflict(case, names, plan_start) -> dict:
    machines = [names.get(t, t) for t in case["types"]]
    days = sorted({d for ds in case["days"].values() for d in ds})
    commitments = [dict(ref=c["ref"], kind="pronájem" if c["kind"] == "rental" else "vlastní zakázka",
                        period=days_text(list(range(c["start"], c["end"] + 1)), plan_start),
                        **{k: v for k, v in _contribution_from(c).items()}) for c in case["commitments"]]
    variants = sorted(case["variants"], key=lambda v: (VARIANT_ORDER.index(v["kind"])
                                                       if v["kind"] in VARIANT_ORDER else 99, v["key"]))
    return dict(
        id=case["id"],
        title=f"{', '.join(machines)} · {days_text(days, plan_start)} · "
              + " × ".join([c["ref"] for c in case["commitments"]] + (["servis"] if case["services"] else [])),
        machines=machines, types=list(case["types"]), days=days_text(days, plan_start),
        services=list(case["services"]), commitments=commitments,
        jobs=[c["ref"] for c in case["commitments"] if c["kind"] == "job"],
        variants=[_variant(v, plan_start, names) for v in variants],
        comparisons=[dict(a=c["a"], b=c["b"], text=c["text"],
                          kind=COMPARISON_KIND.get(c.get("status"), "nelze porovnat"))
                     for c in case["comparisons"]],
        comparison_note=_comparison_note(case),
        note="Nic se automaticky nevybírá, neruší ani nepřesouvá – rozhodne majitel.",
    )


# Porovnání variant (K3.4d) – jen informace z econ_formula.compare, nic se nevybírá ani nedoporučuje.
COMPARISON_KIND = {"jistý": "podložené porovnání", "podmíněný": "závisí na neznámých nákladech",
                   "nelze rozhodnout": "nelze porovnat"}
COMPARISON_DISCLAIMER = ("Jen informace: porovnání vychází z dostupných údajů. Nejde o doporučení ani rozhodnutí – "
                         "o řešení konfliktu rozhoduje majitel.")


def _comparison_note(case) -> str:
    """Proč porovnání chybí nebo je omezené (vždy výslovně – nic se nedoplňuje nulou)."""
    evaluable = [v for v in case["variants"] if v["feasibility"] in (fe.FEASIBLE, fe.CONDITIONAL)
                 and v["formula"] is not None and v["formula"].valid]
    if len(evaluable) < 2:
        return ("Varianty nelze porovnat – vyčíslit lze jen "
                + ("jednu variantu" if evaluable else "žádnou variantu")
                + "; ostatní závisí na údajích, které chybějí.")
    if all(c.get("status") == "nelze rozhodnout" for c in case["comparisons"]):
        return "Varianty nelze spolehlivě porovnat – výsledek závisí na chybějících údajích (viz u každé dvojice)."
    return ""


def _contribution_from(c) -> dict:
    """Příspěvek závazku v konfliktu – vždy „nezapočítáno“, hodnota jen pro informaci."""
    value, status, missing = c["contribution"], c["contribution_status"], c["missing"]
    if value is None or status == "nelze spočítat":
        amount = "nelze určit"
    elif status == "úplný":
        amount = money(value)
    elif _missing_only_costs(missing):
        amount = f"nejvýše {money(value)}"
    else:
        amount = "nelze určit"
    return dict(status=EXCLUDED, tone=TONE[EXCLUDED], text=amount,
                explanation="v konfliktu – do výsledku se nezapočítává" + (
                    "; chybí: " + "; ".join(missing) if missing else ""))


FEASIBILITY_TEXT = {fe.FEASIBLE: "proveditelná", fe.CONDITIONAL: "podmíněná", fe.PARTIAL: "jen částečně",
                    fe.NOT_EVALUABLE: "nelze vyhodnotit", fe.MANUAL: "ruční rozhodnutí"}


def _variant(v, plan_start, names) -> dict:
    title = VARIANT_TITLE.get(v["kind"], v["kind"])
    feas = v["feasibility"]
    if feas == fe.CONDITIONAL:
        conds = [m for m in v["missing"] if _is_condition(m)]
        rest = [m for m in v["missing"] if not _is_condition(m)]
        amount = bounds_cell(*v["bounds"], conditional=True, conditions=conds)
        if rest:
            amount["explanation"] += " · chybí: " + "; ".join(rest)
    elif feas == fe.FEASIBLE:
        amount = bounds_cell(*v["bounds"])
        if amount["status"] == UPPER:
            amount["explanation"] = "chybí: " + "; ".join(v["missing"]) if v["missing"] else amount["explanation"]
    elif feas == fe.MANUAL:
        amount = _cell(EXCLUDED, "nelze vyčíslit", "konflikt trvá – oba závazky nelze splnit současně")
    else:
        amount = _cell(MISSING, "nelze určit", "chybí: " + "; ".join(v["missing"]))
    return dict(key=v["key"], title=title, feasibility=FEASIBILITY_TEXT.get(feas, feas),
                resolves={fe.RESOLVED: "ano", fe.PARTLY: "jen částečně", fe.NOT_RESOLVED: "ne"}.get(v["resolves"], "–"),
                change=_change_text(v, plan_start), warnings=warnings_text(v.get("warnings", []), names, plan_start),
                blockers=[f"{b['source']} ({days_text(b['days'], plan_start)})" for b in v.get("blockers", [])],
                missing=list(v.get("missing", [])), amount=amount)


def _change_text(v, plan_start) -> str:
    ch = v.get("changes") or {}
    if not ch:
        return v.get("description", "")
    old = days_text(list(range(ch["start"], ch["end"] + 1)), plan_start)
    if "new_days_late" in ch:                                         # posun zakázky
        new = days_text(list(range(ch["new_start"], ch["new_end"] + 1)), plan_start)
        parts = [f"{ch['ref']}: {old} → {new}",
                 f"zpoždění {ch['days_late']} → {ch['new_days_late']} dní"]
        if ch.get("contribution_change") is not None:
            parts.append(f"změna příspěvku zakázky {money(ch['contribution_change'])}")
        if ch.get("penalty_free_starts") == []:
            parts.append("bez penále posunout nelze")
        return " · ".join(parts)
    if ch.get("new_start") is None:                                   # zrušení pronájmu
        return f"{ch['ref']}: zrušit ({old}) · uvolní {_md(ch['released_machine_days'])}"
    new = days_text(list(range(ch["new_start"], ch["new_end"] + 1)), plan_start)
    return f"{ch['ref']}: {old} → {new} · uvolní {_md(ch['released_machine_days'])}"


def _md(n) -> str:
    return f"{n} {_plural(n, 'strojoden', 'strojodny', 'strojodnů')}"


def _is_condition(text) -> bool:
    return any(x in text for x in ("souhlas zákazníka", "storno", "smluvní náhrad"))


_WARNING = re.compile(r"^(?P<type>\S+) den (?P<day>\d+): (?P<text>.+)$")


def warnings_text(warnings, names=None, plan_start=None) -> list:
    """Varování po dnech sloučená podle stroje a důvodu: „Smykový nakladač, dny 20–23: …“."""
    names = names or {}
    grouped, other = {}, []
    for w in warnings:
        m = _WARNING.match(w)
        if not m:
            other.append(w)
            continue
        grouped.setdefault((m["type"], m["text"]), []).append(int(m["day"]))
    return [f"{names.get(t, t)}, {days_text(d, plan_start)}: {text}" for (t, text), d in grouped.items()] + other


def basis_text(text) -> str:
    """Popis metody přepočtu s českou desetinnou čárkou (0.8333 → 0,8333)."""
    return re.sub(r"(\d)\.(\d)", r"\1,\2", text or "")


def _scenarios(pr) -> dict:
    lo, hi = pr.range
    if not pr.options.cases:
        cell = bounds_cell(lo, hi)
        return dict(applicable=False, cell=cell, counted=0, total=0, notes=[])
    counted = [s for s in pr.scenarios if s.get("valid") and s.get("in_period")]
    conditional = any(_has_condition(pr, s) for s in counted)
    cell = bounds_cell(lo, hi, conditional=conditional and hi is not None,
                       conditions=["souhlas zákazníka a výše storna / smluvní náhrady"] if conditional else ())
    if not counted:
        cell = _cell(MISSING, "nelze určit", "žádná kombinace variant nezůstane celá v období a bez konfliktu")
    notes = [n for n in pr.notes if "kombinac" in n or "rozpětí" in n]
    return dict(applicable=True, cell=cell, counted=len(counted), total=len(pr.scenarios), notes=notes,
                text="Rozpětí výsledku přes ověřené kombinace variant konfliktů – žádná varianta se nevybírá.")


def _has_condition(pr, scenario) -> bool:
    keys = set(scenario["variants"].values())
    return any(v["key"] in keys and v["feasibility"] == fe.CONDITIONAL
               for c in pr.options.cases for v in c["variants"])


# ---------------------------------------------------------------------------
# Předběžné a nejisté závazky, evidence, využití
# ---------------------------------------------------------------------------

# Předběžné, přesahující a nejisté závazky (K3.4e1): do výsledku se NEZAPOČÍTÁVAJÍ – částka jen pro informaci.
PENDING_STATES = (fe.TENTATIVE, fe.CROSSES, fe.UNKNOWN)
PENDING_NOTE = ("Nezapočítané částky – nejde o potvrzený výsledek. Částky se nesčítají s výsledkem ani mezi sebou; "
                "jde o příspěvky celých závazků (při chybějících nákladech jen horní meze).")
# Evidence závazků: skupiny v pevném pořadí (každý závazek z dat právě v jedné skupině)
LEDGER_GROUPS = [
    (fe.CERTAIN, "Jisté využití – započteno do výsledku"),
    (fe.CONFLICT, "V konfliktu – nezapočítáno, rozhodne majitel"),
    (fe.TENTATIVE, "Předběžné – nezapočítáno, jen potenciál"),
    (fe.CROSSES, "Přesah mimo období – nezapočítáno, chybí pravidlo časového rozlišení"),
    (fe.UNKNOWN, "Obsazenost nelze určit – nezapočítáno"),
    (fe.OUTSIDE, "Mimo zvolené období – nezapočítáno"),
    (fe.CANCELLED, "Zrušené – nezapočítáno"),
]


def _pending(led, plan_start, names=None) -> list:
    return [_entry_row(r, plan_start, names) for r in led.entries.itertuples() if r.state in PENDING_STATES]


def _ledger(led, plan_start, names=None) -> list:
    return [_entry_row(r, plan_start, names) for r in led.entries.itertuples()]


def _ledger_groups(led, plan_start, names=None) -> list:
    rows = _ledger(led, plan_start, names)
    known = {state for state, _ in LEDGER_GROUPS}
    groups = [dict(state=state, title=title, rows=[r for r in rows if r["state"] == state])
              for state, title in LEDGER_GROUPS]
    other = [r for r in rows if r["state"] not in known]                # pojistka – nic nezmizí
    if other:
        groups.append(dict(state="jiné", title="Jiný stav – nezapočítáno", rows=other))
    return [g for g in groups if g["rows"]]


def _machines_text(needs, names) -> str:
    """„MINI_BAGR×1; PASOVY_DUMPER×1“ → „Mini bagr, Pásový dumper“ (počet kusů jen když je víc než 1)."""
    out = []
    for part in str(needs or "").split("; "):
        if not part:
            continue
        type_id, _, qty = part.partition("×")
        name = (names or {}).get(type_id, type_id)
        out.append(name if qty in ("", "1") else f"{name} ×{qty}")
    return ", ".join(out)


def _entry_row(r, plan_start, names=None) -> dict:
    res = fe.ec.Result(r.contribution, tuple(r.missing), tuple(r.invalid))
    cell = contribution_cell(res)
    if r.state != fe.CERTAIN:
        cell = dict(cell, status=EXCLUDED, tone=TONE[EXCLUDED],
                    explanation="; ".join(x for x in (STATE_TEXT.get(r.state, r.state), cell["explanation"]) if x))
    period = (days_text(list(range(r.start, r.end + 1)), plan_start)
              if isinstance(r.start, int) and isinstance(r.end, int) else "bez termínu")
    return dict(ref=r.ref, kind="pronájem" if r.kind == "rental" else "vlastní zakázka", name=r.name or "",
                period=period, date_source=r.date_source, state=r.state,
                counted=bool(r.counted_as_certain), machines=r.needs, unit_assignment=r.unit_assignment,
                machines_text=_machines_text(r.needs, names),
                reason="započteno do výsledku" if r.state == fe.CERTAIN else STATE_TEXT.get(r.state, r.state),
                machine_days=r.machine_days, machine_days_in_period=r.md_in_period, **cell)


def _usage(usage, names) -> list:
    if usage is None or usage.by_type is None:
        return []
    rows = []
    for r in usage.by_type.itertuples():
        rows.append(dict(type_id=r.type_id, name=names.get(r.type_id, r.type_id),
                         capacity=int(r.capacity) if r.capacity_known else None,
                         rental=int(r.rental), job=int(r.job), maintenance=int(r.maintenance),
                         conflict=int(r.conflict), unknown=int(r.unknown), free=int(r.free),
                         capacity_known=bool(r.capacity_known), unit_level=r.unit_level))
    return rows



# ---------------------------------------------------------------------------
# K3.4e2 – stručný souhrn využití strojů a metodika „Jak se počítá“
# ---------------------------------------------------------------------------
# Souhrn je jen převzatý součet za flotilu z fleet_usage (Usage.fleet) – stejná čísla jako stránka
# Plánování kapacity → Přehled využití. Každý dostupný kus se v každém dni započte do PRÁVĚ JEDNÉ
# kategorie, takže se kategorie nepřekrývají a jejich součet je kapacita (fleet_usage to ověřuje: balanced).
# Nárok nad kapacitu a předběžné rezervace NEJSOU součástí kapacity – vykazují se zvlášť.

USAGE_ROWS = [
    ("rental", "Pronájem", "dny, kdy je kus obsazený potvrzeným pronájmem"),
    ("job", "Vlastní zakázky", "dny, kdy je kus na zakázce podle návrhu plánu"),
    ("maintenance", "Servis", "servisní odstávky kusu"),
    ("conflict", "Konflikt", "dny, kdy nároky převyšují kusy – všechny kusy typu toho dne, rozdělení nelze určit"),
    ("unknown", "Nelze určit", "chybějící nebo neplatné údaje – nikdy to neznamená volno"),
    ("free", "Volná kapacita", "kus je k dispozici a nic ho neblokuje"),
]


def _md_text(n) -> str:
    return f"{n} {_plural(n, 'strojoden', 'strojodny', 'strojodnů')}"


def _usage_summary(usage, names) -> dict:
    """Kompaktní souhrn strojodnů za období (nic se nepočítá znovu – Usage.fleet z fleet_usage)."""
    if usage is None or not getattr(usage, "fleet", None):
        return None
    f = usage.fleet
    rows = [dict(key=k, label=label, value=int(f[k]), text=_md_text(int(f[k])), meaning=meaning)
            for k, label, meaning in USAGE_ROWS if k != "unknown" or f.get(k)]
    notes = []
    if f.get("excess"):
        notes.append(f"Nárok nad kapacitu: {_md_text(int(f['excess']))} – potvrzené nároky, které se do kusů nevejdou; "
                     "nejsou součástí kapacity.")
    if f.get("tentative"):
        notes.append(f"Předběžné rezervace: {_md_text(int(f['tentative']))} – kapacitu neblokují; nejsou součástí "
                     "kapacity.")
    if f.get("excluded"):
        notes.append("Bez typů, u nichž nelze spolehlivě určit počet kusů: "
                     + ", ".join((names or {}).get(t, t) for t in f["excluded"]) + ".")
    balanced = bool(f.get("balanced"))
    return dict(capacity=int(f["capacity"]), capacity_text=_md_text(int(f["capacity"])), rows=rows, notes=notes,
                balanced=balanced,
                composition=("Kategorie se nepřekrývají – každý dostupný kus se v každém dni započte do právě jedné; "
                             "dohromady dávají kapacitu flotily." if balanced
                             else "Součet kategorií neodpovídá kapacitě – souhrn nelze spolehlivě sestavit."))


# Metodika „Jak se počítá“ – jen popis skutečné implementace (economics, fleet_usage, fleet_economics,
# fleet_view). Žádné závěry ani rady majiteli.
METHODOLOGY = [
    ("Pět stavů částek",
     "Úplný – všechny potřebné údaje jsou známé, částka je přesná (i skutečná nula „0 Kč“). "
     "Horní mez – chybí jen některý náklad (například variabilní náklad nebo režie); skutečná částka může být "
     "jen nižší, proto „nejvýše“. Podmíněný – platí jen za podmínky, kterou nelze ověřit (souhlas zákazníka, "
     "storno nebo smluvní náhrada). Chybí údaje – nejde určit ani mez (například chybí tržba, cena nebo "
     "zvolená metoda přepočtu) nebo nejsou rozhodnuté konflikty; nikdy se nenahrazuje nulou. Nezapočítáno – závazek je vidět, ale do výsledku "
     "se nepočítá."),
    ("Příspěvek a celkový výsledek",
     "Příspěvek pronájmu = tržby − variabilní náklady pronájmu. Příspěvek vlastní zakázky = tržby − přímé "
     "náklady − penále za zpoždění − variabilní náklady strojů (jen pokud nejsou už v přímých nákladech). "
     "Příspěvky neobsahují fixní náklady ani režii. Celkový výsledek flotily = jisté příspěvky pronájmů "
     "a zakázek − fixní náklady flotily − režie."),
    ("Fixní náklady jen jednou",
     "Fixní náklady se odečítají jednou za celou flotilu (vlastněné a leasingové kusy) a nerozpočítávají se do "
     "příspěvků pronájmů ani zakázek. Obsahují i průměrný servis, proto se servis znovu neodečítá."),
    ("Přepočet na zvolené období",
     "Fixní náklady a režie jsou v datech měsíční. Aplikace žádnou metodu přepočtu nepředpokládá: "
     "„Nezadávat“ – neodečtou se a výsledek je jen horní mez; „Zadat počet měsíců“ – měsíční částka × zadaný počet "
     "měsíců; „Podle kalendáře“ – za každý měsíc dny období v měsíci ÷ počet dní měsíce (vyžaduje datum "
     "prvního dne plánu). Režie se odečte jen tehdy, když je v datech zadaná."),
    ("Co se nezapočítává a proč",
     "Konflikt – stroj nemůže ve stejný den sloužit oběma závazkům a aplikace nerozhoduje, který ustoupí; "
     "nezapočítávají se obě strany. Předběžné rezervace – nejsou potvrzené a kapacitu neblokují. Přesah mimo "
     "období – chybí pravidlo, jakou část příspěvku přiřadit zvolenému období; příspěvek se poměrně nedělí. "
     "Nejistá obsazenost – neplatné nebo chybějící údaje, případně zakázka bez termínu v plánu."),
    ("Strojodny ve využití",
     "Strojoden = jeden kus stroje na jeden den. Každý dostupný kus se v každém dni započte do právě jedné "
     "kategorie (pronájem, vlastní zakázka, servis, konflikt, nelze určit, volno); dohromady dávají kapacitu. "
     "Termíny vlastních zakázek pocházejí z návrhu plánu. Nárok nad kapacitu a předběžné rezervace jsou "
     "uvedené zvlášť a do kapacity se nepočítají."),
    ("Proč zobrazený výsledek není přesný",
     "Pokud chybí náklady nebo režie, některé závazky nejsou započtené nebo jde o ukázková data, nejde o čistý "
     "zisk – je to výsledek z jistých závazků se známými položkami. Rozpětí scénářů jen ukazuje možné "
     "varianty řešení konfliktů; žádnou nevybírá."),
]
