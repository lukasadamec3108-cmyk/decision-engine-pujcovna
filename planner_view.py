"""Zobrazení rozhodovacího plánovače (etapa E2.5b) – záložka „Rozhodovací plánovač“.

Modul NIC nerozhoduje ani nepočítá: vše počítá planner.py (nad capacity.py a economics.py).
Tady se jen převádí formulář na požadavek a výsledek plánovače na obyčejné tabulky, štítky
a srozumitelná vysvětlení (obyčejné typy → lze uložit do mezipaměti Streamlitu).

Od E2.5c kroku 4B zobrazuje karta DVA oddělené štítky – Provoz (op_badge) a Ekonomika (econ_badge),
viz níže „Štítky karty“. Dřívější jednotný štítek (badge/why) zůstává ve výsledku jen kvůli kompatibilitě:

Štítek požadavku (schválená pravidla E2.5b):
  Doporučeno     … požadavek je ve variantě, kterou plánovač ekonomicky doporučil
                   (lepší i v nejhorším případě než každá jiná proveditelná varianta),
  Proveditelné   … vejde se do kapacity; samostatný požadavek bez soupeře nikdy není
                   „Doporučeno“ – ekonomické hodnocení je uvedené zvlášť,
  Neproveditelné … kapacitu drží potvrzené závazky nebo servis,
  Nelze rozhodnout … chybí údaje, nebo se ze soupeřících požadavků vejde jen část a
                   ekonomika chybí.
Nic se neukládá, nepotvrzuje ani nemění – požadavky existují jen v relaci uživatele.

E2.5c krok 4A: evaluate navíc předává strukturovanou ekonomiku (klíč „economy“ u karty i skupiny):
verdikt, doporučenou / shodné / nevyloučené varianty, vzorce příspěvků, hranice zlomu (Kč/den),
podmínky dvojic a chybějící údaje – jen obyčejné typy, bez nových výpočtů. Zobrazení se zatím nemění.
"""

import html
import re
from typing import Optional

import pandas as pd

import datamodel as dm
import econ_formula as ef
import economics as ec
import planner as pl

RECOMMENDED, FEASIBLE, NOT_FEASIBLE, UNDECIDED = "Doporučeno", "Proveditelné", "Neproveditelné", "Nelze rozhodnout"
BADGE_CLASS = {RECOMMENDED: "rec", FEASIBLE: "ok", NOT_FEASIBLE: "no", UNDECIDED: "und"}
MAX_REQUESTS = 5
REF_PREFIX = "NP-"


# ---------------------------------------------------------------------------
# Vstupy z formuláře
# ---------------------------------------------------------------------------

def md_escape(text) -> str:
    """Text (i od uživatele) bezpečně do st.markdown: bez HTML a bez markdownových odkazů či formátu."""
    out = html.escape(str(text), quote=False).replace("\\", "\\\\")
    for ch in "`*_{}[]()#+-.!|~":
        out = out.replace(ch, "\\" + ch)
    return out


def _days(n: int) -> str:
    return f"{n} " + ("den" if n == 1 else "dny" if 2 <= n <= 4 else "dní")


def _num(v) -> Optional[float]:
    s = "" if v is None else str(v).strip()
    if s in ("", "nan", "None"):
        return None
    try:
        return float(s.replace(",", "."))
    except ValueError:
        return None


def next_ref(model: dm.DataModel, existing) -> str:
    """Jedinečné označení nového požadavku (nekoliduje s daty ani s ostatními požadavky)."""
    taken = set(existing) | set(model.jobs["job_id"].astype(str)) | set(model.rental_bookings["booking_id"].astype(str))
    i = 1
    while f"{REF_PREFIX}{i}" in taken:
        i += 1
    return f"{REF_PREFIX}{i}"


def list_price(model: dm.DataModel, type_id: str) -> Optional[float]:
    """Ceník pronájmu typu stroje (None = v datech chybí)."""
    row = model.machine_types[model.machine_types["type_id"] == type_id]
    return None if row.empty else _num(row.iloc[0]["rental_price_day"])


def job_template(model: dm.DataModel, job_id: str) -> dict:
    """Údaje existující zakázky jako šablona NOVÉHO požadavku (původní zakázka se nemění).

    Chybějící hodnota zůstává None – nikdy se nenahrazuje nulou.
    """
    j = model.jobs[model.jobs["job_id"] == job_id].iloc[0]
    req = model.job_requirements
    req = req[(req["job_id"] == job_id) & (req["resource_kind"] == "machine")]
    dur = _num(j["duration_days"])
    return dict(source=job_id, name=f"{j['name']} (kopie údajů {job_id})",
                needs=[(str(r.resource_id), int(float(r.quantity))) for r in req.itertuples()],
                duration=int(dur) if dur else None,
                earliest=int(_num(j["earliest_start_day"])) if _num(j["earliest_start_day"]) else None,
                deadline=int(_num(j["deadline_day"])) if _num(j["deadline_day"]) else None,
                revenue=_num(j["revenue"]), direct_cost=_num(j["direct_cost"]),
                penalty_per_day=_num(j["penalty_per_day"]))


def rental_request(ref: str, type_id: str, qty: int, start: int, end: int, price_day=None) -> dict:
    """Požadavek na pronájem jako obyčejný slovník (do relace i do mezipaměti)."""
    return dict(ref=ref, kind="rental", needs=((type_id, int(qty)),), start=int(start), end=int(end),
                name="", earliest=None, deadline=None, price_day=_num(price_day),
                revenue=None, direct_cost=None, penalty_per_day=None)


def job_request(ref: str, name: str, needs, start: int, duration: int, earliest=None, deadline=None,
                revenue=None, direct_cost=None, penalty_per_day=None) -> dict:
    """Požadavek na vlastní zakázku. needs = [(type_id, kusů)] – prázdné řádky se vynechají."""
    merged = {}
    for t, q in needs:
        if t:
            merged[t] = merged.get(t, 0) + int(q)
    start, duration = int(start), int(duration)
    return dict(ref=ref, kind="job", needs=tuple(merged.items()), start=start, end=start + duration - 1,
                name=name or ref, earliest=None if earliest is None else int(earliest),
                deadline=None if deadline is None else int(deadline), price_day=None,
                revenue=_num(revenue), direct_cost=_num(direct_cost), penalty_per_day=_num(penalty_per_day))


def form_problems(req: dict) -> list:
    """Zjevné chyby formuláře (plánovač vstupy kontroluje i sám)."""
    out = []
    if not req["needs"]:
        out.append("Vyberte aspoň jeden stroj.")
    if req["end"] < req["start"]:
        out.append("Konec je před začátkem.")
    if req["kind"] == "job" and req["earliest"] is not None and req["deadline"] is not None \
            and req["deadline"] < req["earliest"]:
        out.append("Termín dokončení je před nejdřívějším začátkem.")
    if req["kind"] == "job" and req["earliest"] is not None and req["start"] < req["earliest"]:
        out.append("Začátek je před nejdřívějším začátkem okna zakázky.")
    return out


def to_request(d: dict) -> pl.Request:
    return pl.Request(d["ref"], d["kind"], pl.NEW, tuple(tuple(x) for x in d["needs"]), d["start"], d["end"],
                      name=d.get("name") or "", earliest=d.get("earliest"), deadline=d.get("deadline"),
                      price_day=d.get("price_day"), revenue=d.get("revenue"), direct_cost=d.get("direct_cost"),
                      penalty_per_day=d.get("penalty_per_day"))


def freeze(requests: list) -> tuple:
    """Seznam požadavků → n-tice (klíč mezipaměti)."""
    return tuple(tuple(sorted((k, tuple(map(tuple, v)) if k == "needs" else v) for k, v in r.items()))
                 for r in requests)


def unfreeze(frozen: tuple) -> list:
    return [dict(items) for items in frozen]


# ---------------------------------------------------------------------------
# Výsledek plánovače → obyčejné typy
# ---------------------------------------------------------------------------

def badge(entry: dict, group: Optional[dict]) -> tuple:
    """(štítek, proč) – štítek a srozumitelné zdůvodnění podle schválených pravidel."""
    op = entry["operational"]
    if op == pl.NOT_FEASIBLE:
        return NOT_FEASIBLE, ("Požadavek se v zadaném termínu nevejde – kapacitu drží potvrzené závazky "
                              "nebo servis (viz „Co blokuje“).")
    if op == pl.UNDETERMINED:
        return UNDECIDED, "Provozní proveditelnost nelze určit – chybí nebo jsou neplatné údaje (viz důvody)."
    if group is None or group["verdict"] == pl.ALL_FIT:
        return FEASIBLE, ("Požadavek se vejde do zbývající kapacity. Doporučení se nevydává – není s čím "
                          "porovnat; ekonomika je uvedená zvlášť.")
    if group["verdict"] == pl.RECOMMEND:
        if entry["ref"] in (group["choice"] or ()):
            return RECOMMENDED, ("Ze soupeřících požadavků se vejde jen část; varianta s tímto požadavkem je "
                                 "ekonomicky lepší i v nejhorším případě.")
        return FEASIBLE, ("Sám se vejde, ale soupeří o stejné stroje; ekonomicky lepší je jiná varianta "
                          f"({', '.join(group['choice'])}). Nic se neruší – rozhodne uživatel.")
    if group["verdict"] == pl.NONE_PAYS:
        return FEASIBLE, ("Sám se vejde, ale se soupeřícími požadavky se vejde jen část a každá taková varianta "
                          "je ztrátová – ekonomicky nejlépe vychází nepřijmout žádný z nich. Nic se neruší – "
                          "rozhodne uživatel.")
    if group["verdict"] == pl.CONDITIONAL:
        return FEASIBLE, ("Sám se vejde, ale se soupeřícími požadavky se vejde jen část a která varianta je "
                          "ekonomicky nejlepší, závisí na neznámých nákladech – hranice zlomu jsou uvedené u "
                          "soupeřících požadavků. Nic se neruší – rozhodne uživatel.")
    if group["verdict"] == pl.EQUAL:
        return FEASIBLE, ("Sám se vejde; se soupeřícími požadavky mají varianty shodný ekonomický přínos – "
                          "rozhodne uživatel.")
    return UNDECIDED, ("Sám se vejde, ale se soupeřícími požadavky se vejde jen část a ekonomicky nelze "
                       "rozhodnout – chybí údaje (viz Chybějící údaje).")


def _group_of(p: pl.Plan, ref: str) -> Optional[dict]:
    return next((g for g in p.groups if ref in g["requests"]), None)


def _operational_text(e: dict) -> str:
    if e["operational"] == pl.OK_ALONE:
        return "Provozně: vejde se do zbývající kapacity (po potvrzených závazcích a servisu)."
    if e["operational"] == pl.NOT_FEASIBLE:
        return "Provozně: nevejde se – " + e["decision"] + "."
    return "Provozně: nelze určit – " + "; ".join(e["reasons"] or [e["decision"]]) + "."


def _late_text(model: dm.DataModel, r) -> str:
    """U zakázky po termínu: kolik dní a jaké penále je v příspěvku (nebo že sazba chybí)."""
    if r.kind != "job" or r.deadline is None or r.end <= r.deadline:
        return ""
    late = r.end - r.deadline
    rate = r.penalty_per_day
    if r.in_data:
        row = model.jobs[model.jobs["job_id"] == r.ref]
        rate = _num(row.iloc[0]["penalty_per_day"]) if not row.empty else None
    if rate is None:
        return f" Zakázka končí {_days(late)} po termínu – sazba penále chybí."
    czk = f"{late * rate:,.0f}".replace(",", " ")
    return f" Zakázka končí {_days(late)} po termínu – penále {czk} Kč je v příspěvku započtené."


def _economic_text(e: dict, late: str = "") -> str:
    if e.get("economic_note"):
        return "Ekonomicky (jen informace, odděleně od provozu): " + e["economic_note"] + "." + late
    if e["operational"] == pl.OK_ALONE:
        return "Ekonomicky: nehodnoceno."
    return "Ekonomicky: nehodnotí se – požadavek není provozně proveditelný."


# ---------------------------------------------------------------------------
# Strukturovaná ekonomika (E2.5c krok 4A) – jen obyčejné typy, žádné nové výpočty
# ---------------------------------------------------------------------------
#
# Vše vychází z planner.plan a econ_formula (nad economics.py); tady se výsledek jen převádí
# na slovníky, seznamy, řetězce a čísla (bezpečné pro mezipaměť Streamlitu).
#
# Stav údaje:  ZADANO … skutečná hodnota (i 0), NEZADANO … chybí (nikdy se nebere jako 0),
#              NEPLATNE … nelze použít (text, záporné číslo…).
# Příspěvek:   EXACT … přesná částka (i přesně 0 Kč), WITH_UNKNOWN … vzorec s neznámými,
#              INVALID … nelze spočítat (neplatný údaj).
# Jednotky:    částky v Kč, sazby za den v Kč/den, koeficient sazby ve dnech (strojodny nebo dny
#              po termínu); u neznámé částky (tržba, přímé náklady) je koeficient bezrozměrný.

ZADANO, NEZADANO, NEPLATNE = "zadáno", "nezadáno", "neplatné"            # stav vstupního údaje
EXACT, WITH_UNKNOWN, INVALID = "přesný", "závisí na neznámých", "nelze spočítat"   # stav příspěvku
KC, KC_DEN = "Kč", "Kč/den"
VERDICT_KIND = {pl.RECOMMEND: "recommend", pl.NONE_PAYS: "none_pays", pl.EQUAL: "equal",
                pl.CONDITIONAL: "conditional", pl.NO_DECISION: "undecided", pl.ALL_FIT: "all_fit"}
MAX_OTHER_VARIANTS = 10                 # kolik prokazatelně horších variant se předá (zbytek jen počtem)
_RATE_FIELDS = {"variable_cost_rental_day", "variable_cost_job_day", "rental_price_day", "penalty_per_day"}


def _unknown_unit(key) -> str:
    return KC_DEN if key and key[0] in _RATE_FIELDS else KC


def _coef_unit(key) -> str:
    if not key:
        return ""
    if key[0] == "penalty_per_day":
        return "dní po termínu"
    return "strojodnů" if key[0] in _RATE_FIELDS else ""


def _plain_key(key) -> list:
    return [str(x) for x in key] if key is not None else None


def _unknown_info(key, unknowns: dict) -> dict:
    u = unknowns.get(key)
    return dict(key=_plain_key(key), label=u.label if u else " / ".join(map(str, key)),
                kind=u.kind if u else ef.DATA, unit=_unknown_unit(key))


def _formula_plain(f, unknowns: dict) -> dict:
    """Vzorec jako slovník: known (Kč) + Σ coef × neznámá (koeficient a jednotky u každého členu)."""
    return dict(known=float(f.known), unit=KC, exact=f.exact,
                terms=[dict(_unknown_info(k, unknowns), coef=float(c), coef_unit=_coef_unit(k)) for k, c in f.terms])


def _contribution_plain(c) -> dict:
    """Příspěvek (požadavku nebo varianty) po scénářích + stav a meze; neplatný údaj → status NEPLATNE."""
    if not c.valid:
        return dict(status=INVALID, invalid=list(c.invalid), scenarios=[], lo=None, hi=None, unit=KC,
                    value=None, text=ef.describe(c))
    scen = [dict(scenario=sc, scenario_label=ef.SCENARIO_LABEL.get(sc, sc), **_formula_plain(f, c.unknowns))
            for sc, f in c.scenarios.items()]
    lo, hi = ef.bounds(c)
    exact = all(x["exact"] for x in scen) and len({x["known"] for x in scen}) == 1
    return dict(status=EXACT if exact else WITH_UNKNOWN, invalid=[], scenarios=scen,
                lo=None if lo is None else float(lo), hi=None if hi is None else float(hi), unit=KC,
                value=float(scen[0]["known"]) if exact else None, text=ef.describe(c))


def _input(label: str, raw, unit: str, field: str, ident: str = "") -> dict:
    """Stav jednoho vstupního údaje stejnými pravidly jako economics.py (prázdné ≠ 0).
    key = [pole, id] – stejný klíč jako neznámá ve vzorci (podle něj se pozná, zda se údaj uplatní)."""
    r = ec._amount(raw, label, False)
    key = [field, ident] if ident else [field]
    if r.invalid:
        return dict(field=field, key=key, label=label, unit=unit, state=NEPLATNE, value=None, reason=r.invalid[0])
    if r.missing:
        return dict(field=field, key=key, label=label, unit=unit, state=NEZADANO, value=None, reason="")
    return dict(field=field, key=key, label=label, unit=unit, state=ZADANO, value=float(r.value), reason="")


def request_inputs(model: dm.DataModel, r: pl.Request) -> list:
    """Ekonomické vstupy požadavku: skutečná nula / nezadáno / neplatné, s jednotkami."""
    out = []
    mt = model.machine_types.set_index("type_id")
    for t, _ in r.needs:
        name = str(mt.loc[t, "name"]) if t in mt.index else t
        if r.kind == "rental":
            if r.price_day is not None:
                out.append(_input(f"{name}: cena pronájmu/den", r.price_day, KC_DEN, "price_day", t))
            else:
                out.append(_input(f"{name}: ceníková cena pronájmu/den", mt.loc[t, "rental_price_day"]
                                  if t in mt.index else "", KC_DEN, "rental_price_day", t))
            out.append(_input(f"{name}: variabilní náklad pronájmu/den", mt.loc[t, "variable_cost_rental_day"]
                              if t in mt.index else "", KC_DEN, "variable_cost_rental_day", t))
        else:
            out.append(_input(f"{name}: variabilní náklad na zakázce/den", mt.loc[t, "variable_cost_job_day"]
                              if t in mt.index else "", KC_DEN, "variable_cost_job_day", t))
    if r.kind == "job":
        row = None
        if r.in_data:
            rows = model.jobs[model.jobs["job_id"] == r.ref]
            row = rows.iloc[0] if len(rows) else None
        for col, text, unit in (("revenue", "tržba", KC), ("direct_cost", "přímé náklady", KC),
                                ("penalty_per_day", "penále/den", KC_DEN)):
            raw = row[col] if row is not None else getattr(r, col)
            out.append(_input(f"{r.ref}: {text}", "" if raw is None else raw, unit, col, r.ref))
        incl = ec.direct_cost_includes_machines(model)
        out.append(dict(field="direct_cost_includes_machines", key=["direct_cost_includes_machines"],
                        label="zda přímé náklady zakázek obsahují vlastní stroje", unit="",
                        state=NEZADANO if incl is None else ZADANO,
                        value=None if incl is None else ("ano" if incl else "ne"), reason=""))
    return out


def request_economy(model: dm.DataModel, r: pl.Request) -> dict:
    """Ekonomika jednoho požadavku pro kartu: vzorec po scénářích a stav vstupů."""
    try:
        c = ef.request_formula(model, r)
        out = _contribution_plain(c)
    except ArithmeticError as exc:
        out = dict(status=INVALID, invalid=[str(exc)], scenarios=[], lo=None, hi=None, unit=KC, value=None,
                   text=[])
    out["inputs"] = request_inputs(model, r)
    # E2.5c 4C: nezadaný údaj, který ve vzorci (v žádném scénáři) nehraje roli → „neuplatňuje se“
    used = {tuple(t["key"]) for sc in out["scenarios"] for t in sc["terms"]}
    forms = {(sc["known"], tuple((tuple(t["key"]), t["coef"]) for t in sc["terms"])) for sc in out["scenarios"]}
    for i in out["inputs"]:
        if i["state"] != NEZADANO:
            i["applies"] = None
        elif i["field"] == "direct_cost_includes_machines":
            i["applies"] = len(forms) > 1                 # nastavení hraje roli, jen když scénáře vycházejí jinak
        else:
            i["applies"] = tuple(i["key"]) in used
    return out


def _variant_plain(v: tuple, c, role: str) -> dict:
    return dict(refs=list(v), label=ef.variant_label(v), nothing=not v, role=role, contribution=_contribution_plain(c))


def group_economy(model: dm.DataModel, p: pl.Plan, g: dict) -> dict:
    """Strukturovaný ekonomický výsledek skupiny soupeřících požadavků (jen obyčejné typy).

    verdict / kind … verdikt plánovače (beze změny) a jeho strojový kód
    choice         … doporučená varianta (seznam označení; [] = „nepřijmout nic“; None = žádné doporučení)
    best           … prokazatelně nejlepší (shodné) varianty, contenders … varianty, které nelze vyloučit
    variants       … vzorce variant s rolí: doporučeno / nejlepší / kandidát / horší
    profiles       … hranice zlomu podle jedné neznámé (Kč/den), conditions … podmínky dvojic (víc neznámých)
    missing        … neznámé, na kterých rozhodnutí závisí; blocking … chybějící základní / neplatné údaje
    """
    kind = VERDICT_KIND.get(g["verdict"], "undecided")
    base = dict(verdict=g["verdict"], kind=kind, requests=list(g["requests"]), text=g["text"])
    if g["verdict"] == pl.ALL_FIT or not g.get("candidates"):
        return dict(base, choice=None, best=[], contenders=[], variants=[], others_total=0, profiles=[],
                    conditions=[], missing=[], blocking=[], scenarios=[], all_exact=False)
    forms = ef.group_formulas(model, p, g)
    unknowns = {k: u for c in forms.values() for k, u in c.unknowns.items()}
    choice = g.get("choice")
    best = [tuple(v) for v in g.get("best", [])]
    contenders = [tuple(v) for v in g.get("contenders", [])]

    def role(v):
        if choice is not None and v == tuple(choice):
            return "doporučeno"
        if v in best:
            return "nejlepší"
        if v in contenders:
            return "kandidát"
        return "horší"
    order = {"doporučeno": 0, "nejlepší": 1, "kandidát": 2, "horší": 3}

    def score(v):
        c = forms[v]
        return max(f.known for f in c.scenarios.values()) if c.valid else float("-inf")
    ranked = sorted(forms, key=lambda v: (order[role(v)], -score(v), len(v), v))
    main = [v for v in ranked if role(v) != "horší" or v == pl.NOTHING]
    others = [v for v in ranked if v not in main]
    variants = [_variant_plain(v, forms[v], role(v)) for v in main + others[:MAX_OTHER_VARIANTS]]

    profiles = []
    for item in g.get("profiles", []):
        prof = item["profile"]
        profiles.append(dict(
            scenario=item["scenario"], scenario_label=ef.SCENARIO_LABEL.get(item["scenario"], item["scenario"]),
            unknown=_unknown_info(prof["key"], unknowns) if prof["key"] is not None else None,
            unit=_unknown_unit(prof["key"]) if prof["key"] is not None else "",
            segments=[dict(start=float(lo), end=None if hi is None else float(hi), winners=list(w))
                      for lo, hi, w in prof["segments"]],
            # přesně v hranici zlomu vycházejí varianty sousedních úseků stejně (úseky jsou otevřené)
            boundaries=[dict(at=float(x), winners=list(w)) for x, w in prof.get("boundaries", [])],
            at_zero=list(prof["at_zero"])))
    conditions = []
    for c in g.get("pair_conditions", []):
        conditions.append(dict(
            scenario=c["scenario"], scenario_label=ef.SCENARIO_LABEL.get(c["scenario"], c["scenario"]),
            a=list(c["a"]), b=list(c["b"]), a_label=ef.variant_label(c["a"]), b_label=ef.variant_label(c["b"]),
            status=c["status"], winner=c["winner"], strict=bool(c["strict"]),
            known=None if c["known"] is None else float(c["known"]), unit=KC,
            terms=[dict(_unknown_info(k, unknowns), coef=float(x), coef_unit=_coef_unit(k)) for k, x in c["terms"]]))
    # chybějící údaje, na kterých rozhodnutí závisí: neznámé ve vzorcích nevyloučených variant
    missing, seen = [], set()
    for v in (contenders if kind in ("conditional", "undecided") else []):
        for f in forms[v].scenarios.values() if forms[v].valid else ():
            for k, _ in f.terms:
                if k not in seen:
                    seen.add(k)
                    missing.append(_unknown_info(k, unknowns))
    scen = sorted({sc for v in contenders for sc in (forms[v].scenarios if forms[v].valid else {})} - {ef.ONLY})
    if scen and kind in ("conditional", "undecided"):
        missing.append(dict(key=["direct_cost_includes_machines"], kind=ef.DATA, unit="",
                            label="zda přímé náklady zakázek obsahují vlastní stroje"))
    blocking = list(dict.fromkeys(r for v in forms.values() for r in v.invalid))
    if kind == "undecided":
        blocking += [x for x in dict.fromkeys(
            t["label"] for v in contenders if forms[v].valid
            for f in forms[v].scenarios.values() for t in [_unknown_info(k, unknowns) for k, _ in f.terms]
            if t["kind"] != ef.RATE) if x not in blocking]
    # rozdíly v Kč se smějí počítat jen tehdy, když mají VŠECHNY porovnávané varianty přesnou částku
    all_exact = all(c.valid and len({f.known for f in c.scenarios.values()}) == 1
                    and all(f.exact for f in c.scenarios.values()) for c in forms.values())
    return dict(base, choice=None if choice is None else list(choice), best=[list(v) for v in best],
                contenders=[list(v) for v in contenders], variants=variants, others_total=len(others),
                all_exact=all_exact,
                profiles=profiles, conditions=conditions, missing=missing, blocking=blocking, scenarios=scen)



# ---------------------------------------------------------------------------
# Štítky karty a ekonomické řádky (E2.5c krok 4B) – odvozeno z výsledku plánovače
# ---------------------------------------------------------------------------
#
# Karta má DVA oddělené štítky: Provoz (vejde se do kapacity?) a Ekonomika (vyplatí se?).
# Zelená je vyhrazená jen pro ekonomické „Doporučeno“; „Vejde se“ je neutrální.

OP_FITS, OP_NO, OP_UNKNOWN = "✓ Vejde se", "✕ Nevejde se", "? Provoz nelze určit"
EC_REC, EC_NO, EC_EQ, EC_COND, EC_UND, EC_NONE = ("★ Doporučeno", "Nevyplatí se", "= Rovnocenné",
                                                  "⇄ Podmíněné", "? Nelze rozhodnout", "Bez porovnání")
OP_CLASS = {OP_FITS: "op-ok", OP_NO: "op-no", OP_UNKNOWN: "op-und"}
EC_CLASS = {EC_REC: "ec-rec", EC_NO: "ec-no", EC_EQ: "ec-eq", EC_COND: "ec-cond", EC_UND: "ec-und",
            EC_NONE: "ec-none"}
NOT_A_REJECTION = ("Jde jen o ekonomické doporučení – požadavek se tím nezamítá a žádná rezervace se neruší; "
                   "rozhodnutí je na vás.")


def op_badge(entry: dict) -> str:
    """Provozní štítek přímo z provozního stavu plánovače."""
    return {pl.OK_ALONE: OP_FITS, pl.NOT_FEASIBLE: OP_NO}.get(entry["operational"], OP_UNKNOWN)


def _in(ref: str, variants) -> bool:
    return any(ref in v for v in variants)


def _labels(variants) -> str:
    return ", ".join(f"„{' + '.join(v) if v else 'nepřijmout nic'}“" for v in variants)


def _rate(x: float) -> str:
    return ef.rate_text(x) + "/den"


def _span(seg: dict) -> str:
    """Otevřený úsek (bez hranic – v hranici vycházejí sousední varianty stejně)."""
    if seg["end"] is None:
        return f"nad {_rate(seg['start'])}"
    if seg["start"] == 0:
        return f"pod {_rate(seg['end'])}"
    return f"mezi {ef.rate_text(seg['start']).removesuffix(' Kč')} a {_rate(seg['end'])}"


def _ties_at_boundaries(prof: dict) -> dict:
    return {b["at"]: b["winners"] for b in prof.get("boundaries", [])}


def _never_wins(ref: str, ge: dict) -> bool:
    """Hranice zlomu existují pro každý scénář a varianta s požadavkem v žádném úseku nevyhrává."""
    profiles = {p["scenario"]: p for p in ge["profiles"]}
    scenarios = ge["scenarios"] or [""]
    if any(sc not in profiles for sc in scenarios):
        return False                                   # víc neznámých → nelze takto vyloučit
    with_ref = {v["label"] for v in ge["variants"] if ref in v["refs"]}
    return not any(set(seg["winners"]) & with_ref for sc in scenarios for seg in profiles[sc]["segments"])


def _where_text(ref: str, ge: dict) -> str:
    """Kdy vychází nejlépe varianta s požadavkem – z hranic zlomu plánovače, po scénářích nastavení."""
    with_ref = {v["label"] for v in ge["variants"] if ref in v["refs"]}
    profiles = {p["scenario"]: p for p in ge["profiles"]}
    scenarios = ge["scenarios"] or [""]
    out = []
    for sc in scenarios:
        prof = profiles.get(sc)
        if prof is None:
            text = "rozhoduje víc nezadaných nákladů najednou"
        else:
            segs = [seg for seg in prof["segments"] if set(seg["winners"]) & with_ref]
            what = _plain_label(prof["unknown"]["label"]) if prof["unknown"] else ""
            if not segs:
                text = "nevychází nejlépe při žádné výši nezadaného nákladu"
            elif len(prof["segments"]) == 1 or not what:
                text = "vychází nejlépe při jakékoli výši nezadaného nákladu"
            else:
                text = f"vychází nejlépe, pokud je {what} " + " nebo ".join(_span(x) for x in segs)
                eq = [(x, [o for o in w if o not in with_ref]) for x, w in _ties_at_boundaries(prof).items()
                      if set(w) & with_ref]
                eq = [(x, o) for x, o in eq if o]
                if eq:
                    text += " (" + "; ".join(f"přesně při {_rate(x)} stejně jako {_and(o)}" for x, o in eq) + ")"
        if sc:
            label = SCENARIO_SHORT.get(sc, sc)
            out.append(f"Pokud {label}: {text}.")
        else:
            out.append(text[0].upper() + text[1:] + ".")
    return " ".join(out)


SCENARIO_SHORT = {"ano": "přímé náklady zakázek obsahují stroje", "ne": "přímé náklady zakázek neobsahují stroje"}


SEE_PANEL = "viz Ekonomické porovnání výše"


def econ_badge(ref: str, entry: dict, g: Optional[dict], ge: Optional[dict]) -> tuple:
    """(štítek, krátké vysvětlení) ekonomiky – výhradně z verdiktu plánovače a jeho variant.

    g  … skupina z planner.plan (verdikt, choice, best, contenders), ge … její strukturovaná ekonomika (4A).
    Podrobnosti (čísla, podmínky, co doplnit a upozornění, že se nic nezamítá) jsou jednou
    v panelu „Ekonomické porovnání“ (krok 4C); karta má jen krátkou větu s odkazem na panel.
    """
    if entry["operational"] != pl.OK_ALONE:
        return None, ("Ekonomicky se nehodnotí – požadavek se nevejde do kapacity." if entry["operational"]
                      == pl.NOT_FEASIBLE else "Ekonomicky se nehodnotí – provozní proveditelnost nelze určit.")
    if g is None or g["verdict"] == pl.ALL_FIT:
        return EC_NONE, ("Nemá soupeřící požadavek – vejde se spolu s ostatními, proto se doporučení nevydává. "
                         "Příspěvek je uvedený jen pro informaci.")
    v = g["verdict"]
    if v == pl.RECOMMEND:
        if ref in (g["choice"] or ()):
            return EC_REC, f"Je součástí doporučené varianty {_labels([g['choice']])} – {SEE_PANEL}."
        return EC_NO, f"Výhodnější je varianta {_labels([g['choice']])} – {SEE_PANEL}."
    if v == pl.NONE_PAYS:
        return EC_NO, f"Každá varianta, která se vejde, je prokazatelně ztrátová – {SEE_PANEL}."
    if v == pl.EQUAL:
        best = [tuple(b) for b in g.get("best", [])]
        if _in(ref, best):
            others = [b for b in best if ref not in b]
            return EC_EQ, (f"Stejně výhodná jako {_labels(others) or 'jiná varianta'}"
                           + (" (0 Kč)" if () in others else "") + f" – {SEE_PANEL}.")
        return EC_NO, f"Výhodnější jsou shodně varianty {_labels(best)} – {SEE_PANEL}."
    if v == pl.CONDITIONAL:
        contenders = [tuple(c) for c in g.get("contenders", [])]
        if not _in(ref, contenders):
            return EC_NO, f"Při jakékoli výši nezadaných nákladů je výhodnější jiná varianta – {SEE_PANEL}."
        if ge and _never_wins(ref, ge):
            with_ref = {v["label"] for v in ge["variants"] if ref in v["refs"]}
            at = sorted({x for pr in ge["profiles"] for x, w in _ties_at_boundaries(pr).items() if set(w) & with_ref})
            return EC_NO, ("Nevychází nejlépe při žádné výši nezadaného nákladu"
                           + (f" (stejně jako nejlepší jen přesně při {' a '.join(_rate(x) for x in at)})" if at else "")
                           + f" – {SEE_PANEL}.")
        where = _where_text(ref, ge) if ge else ""
        return EC_COND, (where + " " if where else "Záleží na nezadaných nákladech. ") + f"Podrobně {SEE_PANEL}."
    blocking = ", ".join((ge or {}).get("blocking", []))
    if blocking:
        return EC_UND, f"Nelze rozhodnout – chybí nebo je neplatné: {blocking} – {SEE_PANEL}."
    # bez chybějících údajů (pojistka plánovače, např. příliš mnoho variant) → důvod přímo z plánovače
    return EC_UND, g["text"].rstrip(".") + "."


def _czk(x: float) -> str:
    text = f"{abs(x):,.0f}".replace(",", " ")
    return ("−" if x < 0 else "") + text + " Kč"


def _plain_label(label: str) -> str:
    return label[:-4] if label.endswith("/den") else label


def _count(n: float, unit: str) -> str:
    forms = {"strojodnů": ("strojoden", "strojodny", "strojodnů"),
             "dní po termínu": ("den po termínu", "dny po termínu", "dní po termínu")}
    if unit not in forms:
        return f"{n:g}"
    if float(n).is_integer():
        k = int(n)
        word = forms[unit][0] if k == 1 else forms[unit][1] if 2 <= k <= 4 else forms[unit][2]
        return f"{k} {word}"
    return f"{n:g} {forms[unit][2]}"


def formula_words(sc: dict) -> str:
    """Vzorec slovy se správnými jednotkami; neznámá je vždy označená „nezadáno“, nikdy 0 Kč."""
    text = _czk(sc["known"])
    for t in sc["terms"]:
        sign = "−" if t["coef"] < 0 else "+"
        what = f"{_plain_label(t['label'])} ({t['unit']}, nezadáno)"
        if t["unit"] == KC and abs(abs(t["coef"]) - 1) < 1e-9:
            text += f" {sign} {what}"
        else:
            text += f" {sign} {_count(abs(t['coef']), t['coef_unit'])} × {what}"
    return text


def contribution_lines(economy: Optional[dict], late: str = "") -> list:
    """Řádky „Příspěvek…“ na kartě: přesná částka / vzorec s nezadanými údaji / proč nelze spočítat.

    Rozlišuje: zadáno (i skutečná 0 – výslovně označená), nezadáno (ve vzorci jako „nezadáno“,
    nikdy 0 Kč) a „neuplatňuje se“ (nezadaný údaj, který na tento výsledek nemá vliv).
    """
    if economy is None:
        return []
    lines = []
    st = economy["status"]
    if st == EXACT:
        v = economy["value"]
        lines.append("Příspěvek: přesně 0 Kč (vychází ze zadaných údajů)." if v == 0 else
                     f"Příspěvek: {_czk(v)} – ztráta." if v < 0 else f"Příspěvek: {_czk(v)}.")
    elif st == WITH_UNKNOWN:
        scen = economy["scenarios"]
        hi, lo = economy["hi"], economy["lo"]
        if any(x["terms"] for x in scen):
            upper = f" (nejvýše {_czk(hi)})" if hi is not None else ""
        else:                                         # neznámé je jen nastavení, žádný náklad
            upper = (f" (podle nastavení {_czk(lo).removesuffix(' Kč')} až {_czk(hi)})"
                     if lo is not None and hi is not None else "")
        if len(scen) == 1:
            lines.append(f"Příspěvek: {formula_words(scen[0])}{upper}.")
        else:
            lines.append("Příspěvek podle nezadaného nastavení, zda přímé náklady zakázek obsahují vlastní stroje"
                         + upper + ":")
            lines += [f"– {SCENARIO_WORD.get(s['scenario'], s['scenario'])}: {formula_words(s)}." for s in scen]
    else:
        reasons = "; ".join(economy["invalid"]) or "neplatný údaj"
        lines.append(f"Příspěvek nelze spočítat – neplatný údaj: {reasons}.")
    main_end = len(lines) - 1                         # poslední řádek samotného příspěvku
    inputs = economy.get("inputs", [])
    zeros = [i for i in inputs if i["state"] == ZADANO and i["value"] == 0
             and i["field"] != "direct_cost_includes_machines"]
    if zeros:
        lines.append("Zadaná nula (skutečná hodnota, ne chybějící údaj): "
                     + ", ".join(f"{_plain_label(i['label'])} 0 {i['unit']}" for i in zeros) + ".")
    unused = [i for i in inputs if i["state"] == NEZADANO and i.get("applies") is False]
    if unused and st != INVALID:
        lines.append("Nezadáno, ale neuplatňuje se (na výsledek nemá vliv): "
                     + ", ".join(_unused_label(i, economy) for i in unused) + ".")
    if late and lines:
        lines[main_end] += late
    return lines


SCENARIO_WORD = {"ano": "pokud obsahují stroje", "ne": "pokud neobsahují stroje"}


def _unused_label(i: dict, economy: dict) -> str:
    """Název nezadaného údaje bez vlivu s krátkým důvodem, je-li známý."""
    text = _plain_label(i["label"])
    if i["field"] == "penalty_per_day":
        return text + " (zakázka končí v termínu)"
    if i["field"] == "variable_cost_job_day":
        return text + " (náklady strojů jsou v přímých nákladech zakázky)"
    if i["field"] == "direct_cost_includes_machines":
        return text + " (oba případy vycházejí stejně)"
    return text


# ---------------------------------------------------------------------------
# Panel „Ekonomické porovnání“ (E2.5c krok 4C) – jen z výsledku plánovače, bez nových výpočtů
# ---------------------------------------------------------------------------
#
# Pořadí: výsledek jednou větou → Proč (částky a rozdíly jen při úplných datech) → porovnávané
# varianty včetně „nepřijmout nic“ → Kdy se co vyplatí (podmínky slovy) → Co doplnit → upozornění.

PANEL_BADGE = {"recommend": EC_REC, "none_pays": EC_NO, "equal": EC_EQ, "conditional": EC_COND,
               "undecided": EC_UND}
PANEL_NOTE = ("Jde jen o ekonomické doporučení – žádný požadavek se tím nezamítá a žádná rezervace se neruší; "
              "rozhodnutí je na vás.")
BREAK_EVEN = "Hranice zlomu je výše nákladu, při které vycházejí obě varianty stejně."
ROLE_TAG = {"doporučeno": "doporučeno", "nejlepší": "shodně nejlepší", "kandidát": "připadá v úvahu",
            "horší": "prokazatelně horší"}
MAX_PANEL_CONDITIONS = 4


def _vlabel(refs) -> str:
    return " + ".join(refs) if refs else "nepřijmout nic"


def _and(refs) -> str:
    refs = list(refs)
    return refs[0] if len(refs) == 1 else ", ".join(refs[:-1]) + " a " + refs[-1]


def _unknown_name(t: dict) -> str:
    """Neznámá s jednotkou, např. „Mini bagr: variabilní náklad pronájmu (Kč/den)“."""
    return f"{_plain_label(t['label'])} ({t['unit']})" if t.get("unit") else _plain_label(t["label"])


def _amount_text(c: dict) -> str:
    """Příspěvek varianty krátce: přesná částka / vzorec po scénářích / nelze spočítat."""
    if c["status"] == EXACT:
        return _czk(c["value"])
    if c["status"] == INVALID:
        return "nelze spočítat (neplatný údaj)"
    if len(c["scenarios"]) == 1:
        return formula_words(c["scenarios"][0])
    return "; ".join(f"{SCENARIO_WORD.get(x['scenario'], x['scenario'])}: {formula_words(x)}" for x in c["scenarios"])


def condition_sentence(c: dict) -> str:
    """Podmínka dvojice srozumitelně: obě varianty a jednotky. Rozdíl a − b = known + Σ coef × neznámá.

    Zjednodušení jen tam, kde je přesné: jedna neznámá → hranice zlomu; dvě neznámé se stejným
    počtem dní a opačným znaménkem → rozdíl sazeb. Jinak úplná nerovnice.
    """
    a, b = c["a_label"], c["b_label"]
    if c["status"] == ef.CERTAIN:
        if c["winner"] == ef.EQUAL:
            return f"{a} a {b} vycházejí vždy stejně."
        w, l = (a, b) if c["winner"] == ef.A else (b, a)
        return f"{w} je vždy lepší než {l}" + ("" if c["strict"] else " (při nulovém nákladu stejně)") + "."
    k, terms = c["known"], c["terms"]
    head = f"{a} je lepší než {b}, pokud "
    if len(terms) == 1:
        t = terms[0]
        limit = -k / t["coef"]
        return head + f"je {_unknown_name(t)} {'nižší' if t['coef'] < 0 else 'vyšší'} než {_rate(limit)}."
    if len(terms) == 2 and abs(terms[0]["coef"] + terms[1]["coef"]) < 1e-9:
        neg = next(t for t in terms if t["coef"] < 0)
        pos = next(t for t in terms if t["coef"] > 0)
        d = k / pos["coef"]                         # neg − pos < d
        x, y = _unknown_name(neg), _unknown_name(pos)
        if abs(d) < 1e-9:
            return head + f"je {x} nižší než {y}."
        if d > 0:
            return head + f"je {x} vyšší než {y} o méně než {_rate(d)} (nebo nižší)."
        return head + f"je {x} nižší než {y} o více než {_rate(-d)}."
    left = " + ".join(f"{_count(-t['coef'], t['coef_unit'])} × {_unknown_name(t)}" for t in terms if t["coef"] < 0)
    right = " + ".join(f"{_count(t['coef'], t['coef_unit'])} × {_unknown_name(t)}" for t in terms if t["coef"] > 0)
    return head + f"{left or '0'} < {_czk(k)}" + (f" + {right}" if right else "") + "."


def _profile_lines(prof: dict, contenders: list) -> list:
    """Pásma hranic zlomu jako seznam (bez grafiky – ta je v kroku 4D).

    Úseky jsou otevřené („pod“, „mezi“, „nad“); přesně v hranici zlomu se uvádí rovnost variant,
    nikoli vítězství jedné z nich.
    """
    out = []
    ties = _ties_at_boundaries(prof)
    for seg in prof["segments"]:
        if seg["start"] in ties:
            out.append(f"přesně {_rate(seg['start'])}: {_and(ties[seg['start']])} vycházejí stejně (rovnocenné)")
        who = _and(seg["winners"]) + (" (shodně)" if len(seg["winners"]) > 1 else "")
        if prof["unknown"] is None or (seg["start"] == 0 and seg["end"] is None):
            out.append(f"při jakékoli výši nákladu: {who}")
        else:
            out.append(f"{_span(seg)}: {who}")
    winners = {w for seg in prof["segments"] for w in seg["winners"]}
    for label in [_vlabel(c) for c in contenders if _vlabel(c) not in winners]:
        at = [x for x, w in ties.items() if label in w]
        out.append(f"{label}: nevychází nejlépe při žádné výši nákladu"
                   + (f" (stejně jako nejlepší jen přesně při {' a '.join(_rate(x) for x in at)})" if at else ""))
    first = set(prof["segments"][0]["winners"])
    extra = [w for w in prof["at_zero"] if w not in first]
    if extra and prof["unknown"] is not None:
        out.append(f"při nulovém nákladu vychází stejně i {_and(extra)}")
    return out


# ---------------------------------------------------------------------------
# Graf hranic zlomu (E2.5c krok 4D-A) – jen z již vypočtených hranic, žádný nový výpočet
# ---------------------------------------------------------------------------
#
# Vodorovný pruh: úseky v pořadí podle výše nezadaného nákladu, mezi nimi očíslované hranice.
# Úseky mají stejnou šířku (nejsou v měřítku) – poslední úsek je otevřený a hranice mohou ležet
# blízko sebe; rozhodují uvedené hodnoty. Graf se tvoří jen při JEDNÉ neznámé a aspoň jedné hranici.

NOTHING_LABEL = "nepřijmout nic"
CHART_SCALE_NOTE = "Úseky nejsou v měřítku – ukazují pořadí; rozhodují uvedené hodnoty v Kč/den."
CHART_MULTI_NOTE = ("Graf se nezobrazuje: rozhoduje víc nezadaných nákladů najednou a jednoduchý pruh by výsledek "
                    "zkreslil – platí podmínky uvedené níže.")
CHART_COLORS = 5                                  # počet barev variant (cyklicky); „nic“ má vlastní šrafování
MAX_CHART_NEVER = 3                               # kolik „nikdy nejvýhodnějších“ variant se vypíše, zbytek se sbalí
MAX_CHART_TIE = 4                                 # nad tolik shodných variant v hranici se jmenují jen sousední
MAX_CHART_ROW = 3                                 # nad tolik úseků se pruh na střední šířce (do 900 px) skládá svisle


def _variants(n: int) -> str:
    return f"další {n} varianta" if n == 1 else f"další {n} varianty" if 2 <= n <= 4 else f"dalších {n} variant"


def profile_chart(prof: dict, contenders: list) -> Optional[dict]:
    """Data grafu z profilu hranic zlomu; None, když hranici nelze vypočítat (nic se nevymýšlí)."""
    if prof["unknown"] is None or len(prof["segments"]) < 2 or not prof.get("boundaries"):
        return None
    colors = {}
    for seg in prof["segments"]:
        for w in seg["winners"]:
            if w != NOTHING_LABEL and w not in colors:
                colors[w] = len(colors) % CHART_COLORS
    segments = []
    for seg in prof["segments"]:
        w = list(seg["winners"])
        only_nothing = w == [NOTHING_LABEL]
        segments.append(dict(range=_span(seg), winners=w, tie=len(w) > 1, nothing=only_nothing,
                             color=None if only_nothing else colors[next(x for x in w if x != NOTHING_LABEL)],
                             label=_and(w) + (" (shodně)" if len(w) > 1 else "")
                             + (" (0 Kč)" if only_nothing else "")))
    boundaries = []
    for i, b in enumerate(prof["boundaries"]):
        # v grafu jmenovitě varianty sousedních úseků; další shodné jen počtem (úplný výčet se dá rozbalit)
        near = set(prof["segments"][i]["winners"]) | set(prof["segments"][i + 1]["winners"])
        many = len(b["winners"]) > MAX_CHART_TIE
        main = [w for w in b["winners"] if w in near or not many]
        others = [w for w in b["winners"] if w not in main]
        boundaries.append(dict(n=i + 1, value=_rate(b["at"]), winners=list(b["winners"]), others=others,
                               text=f"{_and(main)} vycházejí stejně (rovnocenné)"
                               + (f" – stejně vychází i {_variants(len(others))}" if others else "")))
    winners = {w for seg in prof["segments"] for w in seg["winners"]}
    ties = _ties_at_boundaries(prof)
    never = []
    for label in [_vlabel(c) for c in contenders if _vlabel(c) not in winners]:
        at = [x for x, w in ties.items() if label in w]
        never.append(f"{label} není nejvýhodnější při žádné výši nákladu"
                     + (f" (stejně jako nejlepší jen přesně při {' a '.join(_rate(x) for x in at)})" if at else "")
                     + ".")
    first = set(prof["segments"][0]["winners"])
    zero = [w for w in prof["at_zero"] if w not in first]
    summary = "; ".join(f"{s['range']}: {s['label']}" for s in segments)
    return dict(unknown=_unknown_name(prof["unknown"]), segments=segments, boundaries=boundaries,
                never=never[:MAX_CHART_NEVER], never_more=never[MAX_CHART_NEVER:],
                zero_note=(f"Při nulovém nákladu vychází stejně i {_and(zero)}." if zero else ""),
                summary=f"Nejvýhodnější varianta podle {_unknown_name(prof['unknown'])}: {summary}.")


# ---------------------------------------------------------------------------
# „Co kdyby“ (E2.5c krok 4D-B) – hypotetická hodnota JEDNOHO nezadaného nákladu
# ---------------------------------------------------------------------------
#
# Nic se nepočítá znovu ani nerozhoduje jinak: příspěvek varianty při hypotetické hodnotě x je dosazení do
# vzorce, který už plánovač spočítal z economics.py (known + coef × x). Simulace se nabízí jen tam, kde je
# výsledek „Podmíněné“, rozhoduje přesně jedna nezadaná sazba a nastavení „stroje v přímých nákladech“ je
# jednoznačné (jediný scénář). Nejlepší varianta je při každé hodnotě x mezi nevyloučenými variantami
# (ověřeno testy plánovače), proto stačí porovnat je a „nepřijmout nic“. Skutečné doporučení se nemění,
# nic se neukládá.

SIM_MULTI = ("Simulace „Co kdyby“ zatím není k dispozici: rozhoduje víc nezadaných údajů najednou a jedna "
             "hypotetická hodnota by výsledek neurčila.")
SIM_SETTING = ("Simulace „Co kdyby“ zatím není k dispozici: není zadáno, zda přímé náklady zakázek obsahují "
               "vlastní stroje – hypotetický výsledek by stál na nejasném předpokladu.")
SIM_TITLE = "Co kdyby – hypotetický výpočet (není doporučení)"
SIM_NOTE = ("Jde jen o hypotézu pro zadanou hodnotu: skutečné doporučení, data ani rezervace se nemění "
            "a nic se neukládá.")
SIM_EPS = 1e-6                                     # přesná rovnost podle výpočtu (ne podle zaokrouhlení)


def sim_spec(ge: dict) -> Optional[dict]:
    """Zda a pro jakou neznámou lze nabídnout „Co kdyby“; None = výsledek není podmíněný (není co simulovat)."""
    if ge["kind"] != "conditional":
        return None
    if ge["scenarios"]:
        return dict(available=False, reason=SIM_SETTING)
    rates = [m for m in ge["missing"] if m["kind"] == ef.RATE]
    if len(ge["missing"]) != 1 or len(rates) != 1 or len(ge["profiles"]) != 1 or ge["profiles"][0]["unknown"] is None:
        return dict(available=False, reason=SIM_MULTI)
    key = rates[0]["key"]
    prof = ge["profiles"][0]
    variants = []
    for v in ge["variants"]:
        if not (v["nothing"] or v["role"] in ("doporučeno", "nejlepší", "kandidát")):
            continue
        c = v["contribution"]
        if c["status"] == INVALID or len(c["scenarios"]) != 1:
            return dict(available=False, reason=SIM_MULTI)
        sc = c["scenarios"][0]
        if any(t["key"] != key for t in sc["terms"]):     # každý vzorec smí záviset jen na této neznámé
            return dict(available=False, reason=SIM_MULTI)
        variants.append(dict(label=v["label"], refs=list(v["refs"]), nothing=v["nothing"], known=float(sc["known"]),
                             coef=float(sum(t["coef"] for t in sc["terms"]))))
    return dict(available=True, reason="", key=list(key), unknown=_unknown_name(rates[0]),
                unknown_label=_plain_label(rates[0]["label"]), unit=KC_DEN, variants=variants,
                boundaries=[dict(n=i + 1, at=float(b["at"]), winners=list(b["winners"]))
                            for i, b in enumerate(prof["boundaries"])],
                segments=[dict(start=seg["start"], end=seg["end"], winners=list(seg["winners"]), range=_span(seg))
                          for seg in prof["segments"]])


SIM_KEY_PREFIX = "dp_sim_"


def sim_key(spec: dict, sig: str) -> str:
    """Klíč hypotetické hodnoty v relaci: stejný nákladový parametr (pole + typ stroje) = stejný klíč.
    Simulace je jen v jediném scénáři nastavení, takže stejný klíč znamená i stejný scénář.
    sig = otisk dat a požadavků – po jejich změně vznikne nový klíč (hodnota se vymaže)."""
    return SIM_KEY_PREFIX + "_".join(spec["key"]) + "_" + sig


def _czk_exact(x: float) -> str:
    text = f"{abs(x):,.2f}".replace(",", " ").replace(".", ",")
    return ("−" if x < -SIM_EPS else "") + text + " Kč"


def _gap(d: float) -> str:
    return "méně než 0,01 Kč/den" if d < 0.005 else _rate(d)


def simulate(spec: dict, x: float) -> dict:
    """Příspěvky nevyloučených variant při hypotetické hodnotě x (Kč/den) a poloha x vůči hranicím zlomu.

    x se zaokrouhlí na haléře – přesně na hodnotu, kterou pole ukazuje (formát na 2 desetinná místa); počítá se
    tedy vždy s tím, co uživatel vidí. Rovnost i poloha vůči hranici se určují z výpočtu, ne ze zaokrouhlených
    hodnot hranic."""
    if not spec or not spec.get("available") or x is None or x < 0:
        raise ValueError("Simulace není pro tuto skupinu dostupná nebo hodnota není platná.")
    x = round(float(x), 2)
    vals = [(v, v["known"] + v["coef"] * x) for v in spec["variants"]]
    top = max(val for _, val in vals)
    winners = [v["label"] for v, val in vals if val >= top - SIM_EPS]
    rest = [val for _, val in vals if val < top - SIM_EPS]
    gap = top - max(rest) if rest else None
    near = gap is not None and gap < 1.0                 # prakticky stejné: rozdíl pod 1 Kč
    rows = []
    for v, val in sorted(vals, key=lambda p: (-p[1], p[0]["label"])):
        win = v["label"] in winners
        rows.append(dict(label=v["label"], nothing=v["nothing"], value=val,
                         amount=_czk_exact(val) if near else _czk(round(val)),
                         tag=("vychází nejlépe" if len(winners) == 1 else "vychází stejně – nejlépe") if win else ""))
    # poloha vůči hranicím zlomu: „přesně na hranici“ jen při skutečné rovnosti z výpočtu; jinak úsek, ve kterém
    # vítěz opravdu vyhrává (u hodnoty těsně u hranice rozhoduje výpočet, ne zaokrouhlená hodnota hranice)
    bnds = spec["boundaries"]
    if len(winners) > 1 and bnds:
        on = min(bnds, key=lambda b: abs(b["at"] - x))
        position = dict(kind="boundary", n=on["n"], text=f"Hodnota leží přesně na hranici {on['n']} ({_rate(on['at'])}).")
    else:
        segment = sum(1 for b in bnds if b["at"] < x)
        for cand in (segment, segment - 1, segment + 1):         # oprava těsně u hranice podle skutečného vítěze
            if 0 <= cand < len(spec["segments"]) and set(winners) <= set(spec["segments"][cand]["winners"]):
                segment = cand
                break
        seg = spec["segments"][segment]
        text = f"Hodnota leží v úseku „{seg['range']}“"
        lower = bnds[segment - 1] if segment >= 1 else None            # hranice pod úsekem
        upper = bnds[segment] if segment < len(bnds) else None          # hranice nad úsekem
        opts = [b for b in (lower, upper) if b is not None]
        if opts:
            nb_ = min(opts, key=lambda b: abs(b["at"] - x))
            d = abs(nb_["at"] - x)
            text += (f" – k hranici {nb_['n']} ({_rate(nb_['at'])}) zbývá {_gap(d)}" if nb_ is upper
                     else f" – nad hranicí {nb_['n']} ({_rate(nb_['at'])}) o {_gap(d)}")
        position = dict(kind="segment", index=segment, text=text + ".")
    if len(winners) == 1 and near:
        rival = next(v["label"] for v, val in sorted(vals, key=lambda p: -p[1]) if v["label"] not in winners)
        diff = "méně než 0,01 Kč" if gap < 0.005 else _czk_exact(gap)
        verdict = f"Prakticky stejné: {winners[0]} vychází lépe než {rival} jen o {diff}."
    elif len(winners) == 1:
        verdict = f"Při tomto předpokladu vychází nejlépe {winners[0]}."
    else:
        verdict = f"Při tomto předpokladu vycházejí přesně stejně: {_and(winners)}."
    return dict(x=x, x_text=_rate(x) + (" (hypotetická nula)" if x == 0 else ""), winners=winners, near=near,
                gap=gap, rows=rows, verdict=verdict, position=position)


def sim_html(spec: dict, result: dict) -> str:
    """Rámeček „Co kdyby“ – vizuálně oddělený od skutečného doporučení (bez zelené a bez štítku Doporučeno)."""
    t = lambda x: _h(_nowrap_amounts(x))
    nb = lambda x: x.replace("-", "\u2011")              # názvy variant (FIKT-P07) se nelámou na spojovníku
    verdict = t(result["verdict"])
    for r in result["rows"]:
        verdict = verdict.replace(_h(r["label"]), nb(_h(r["label"])))
    rows = "".join(f'<div class="dp-sim-row{" win" if r["tag"] else ""}"><b>{nb(_h(r["label"]))}</b>'
                   f'<span class="amt">{t(r["amount"])}</span>'
                   + (f'<span class="dp-sim-win">◀ {_h(r["tag"])}</span>' if r["tag"] else "") + "</div>"
                   for r in result["rows"])
    return (f'<div class="dp-sim"><div class="dp-sim-t">{_h(SIM_TITLE)}</div>'
            f'<div class="dp-sec">Předpoklad: {t(spec["unknown_label"])} = <b>{t(result["x_text"])}</b></div>'
            f'<div class="dp-sim-v">{verdict}</div>{rows}'
            f'<div class="dp-note">{t(result["position"]["text"])}</div>'
            f'<div class="dp-note">{_h(SIM_NOTE)} Skutečné doporučení: {_h(EC_COND)}.</div></div>')


def group_panel(ge: dict) -> Optional[dict]:
    """Obsah panelu „Ekonomické porovnání“ pro skupinu soupeřících požadavků (jen obyčejné typy)."""
    kind = ge["kind"]
    if kind not in PANEL_BADGE:
        return None
    var = {tuple(v["refs"]): v for v in ge["variants"]}
    choice = tuple(ge["choice"]) if ge["choice"] is not None else None
    why, conditions, missing = [], [], []

    if kind == "recommend":
        w = var[choice]["contribution"]
        headline = f"Přijmout {_and(choice)}."
        if ge["all_exact"]:
            rivals = [v for v in ge["variants"] if tuple(v["refs"]) != choice and not v["nothing"]]
            rival = max(rivals, key=lambda v: v["contribution"]["value"]) if rivals else None
            if rival:
                diff = w["value"] - rival["contribution"]["value"]
                why.append(f"{_vlabel(choice)} přinese {_czk(w['value'])}, nejlepší jiná varianta "
                           f"({rival['label']}) {_czk(rival['contribution']['value'])} – o {_czk(diff)} méně.")
            else:
                why.append(f"{_vlabel(choice)} přinese {_czk(w['value'])}.")
            why.append("Je lepší než všechny ostatní varianty, které se vejdou, i než nepřijetí ničeho.")
        else:
            if w["status"] == EXACT:
                why.append(f"{_vlabel(choice)} přinese přesně {_czk(w['value'])}.")
            why.append("Je prokazatelně lepší než všechny ostatní varianty i než nepřijetí ničeho – při jakékoli "
                       "výši nezadaných nákladů. Rozdíly v Kč se bez úplných údajů neuvádějí.")
    elif kind == "none_pays":
        headline = "Nepřijmout žádný z těchto požadavků – každá varianta, která se vejde, je prokazatelně ztrátová."
        losers = sorted((v for v in ge["variants"] if not v["nothing"]),
                        key=lambda v: -(v["contribution"]["value"] if v["contribution"]["value"] is not None
                                        else (v["contribution"]["hi"] or 0)))[:3]
        for v in losers:
            c = v["contribution"]
            if c["status"] == EXACT:
                why.append(f"{v['label']}: ztráta {_czk(-c['value'])}.")
            elif c["hi"] is not None:
                why.append(f"{v['label']}: ztráta nejméně {_czk(-c['hi'])} (nezadané náklady ji mohou jen zvětšit).")
        why.append("Nepřijetím ničeho je výsledek 0 Kč, tedy lepší než kterákoli z variant.")
    elif kind == "equal":
        best = [tuple(b) for b in ge["best"]]
        headline = "Stejně výhodné jsou: " + " / ".join(_vlabel(b) for b in best) + " – vyberte podle sebe."
        c = var[best[0]]["contribution"]
        why.append(f"Každá z nich přinese přesně {_czk(c['value'])}." if c["status"] == EXACT
                   else "Mají prokazatelně shodný příspěvek.")
        why.append("Ostatní varianty jsou prokazatelně horší.")
    elif kind == "conditional":
        rates = [m for m in ge["missing"] if m["kind"] == ef.RATE]
        setting = any(m["key"] == ["direct_cost_includes_machines"] for m in ge["missing"])
        headline = (f"Záleží na nezadaném nákladu: {_unknown_name(rates[0])}." if len(rates) == 1 and not setting
                    else "Záleží na nezadaných údajích – žádná varianta není nejlepší při každé jejich výši.")
        why.append("Neznámý náklad se nepočítá jako 0 Kč; rozhoduje jeho skutečná výše – viz „Kdy se co vyplatí“.")
        contenders = [tuple(x) for x in ge["contenders"]]
        profiles = {p["scenario"]: p for p in ge["profiles"]}
        for sc in (ge["scenarios"] or [""]):
            title = f"Pokud {SCENARIO_SHORT[sc]}" if sc else ""
            if sc in profiles:
                prof = profiles[sc]
                lead = (f"Rozhoduje {_unknown_name(prof['unknown'])}:" if prof["unknown"] else "")
                conditions.append(dict(title=title, lead=lead, lines=_profile_lines(prof, contenders),
                                       more=[], break_even=bool(prof["unknown"]) and len(prof["segments"]) > 1,
                                       chart=profile_chart(prof, contenders), chart_note=""))
            else:
                sent = [condition_sentence(c) for c in ge["conditions"] if c["scenario"] == sc]
                conditions.append(dict(title=title, lead="Rozhoduje víc nezadaných nákladů najednou:",
                                       lines=sent[:MAX_PANEL_CONDITIONS], more=sent[MAX_PANEL_CONDITIONS:],
                                       break_even=False, chart=None, chart_note=CHART_MULTI_NOTE))
    else:
        blocking = ge["blocking"]
        headline = ("Nelze rozhodnout – chybí nebo je neplatné: " + ", ".join(blocking) + "." if blocking
                    else ge["text"].rstrip(".") + ".")
        for v in ge["variants"]:
            if not v["nothing"] and v["contribution"]["status"] == EXACT and v["role"] != "horší":
                why.append(f"{v['label']} přinese přesně {_czk(v['contribution']['value'])}.")
        why.append("Bez těchto údajů nelze varianty spolehlivě porovnat – nic se nedoporučuje.")

    if kind in ("conditional", "undecided"):
        missing = [_unknown_name(m) if m["key"] != ["direct_cost_includes_machines"]
                   else "zda přímé náklady zakázek obsahují vlastní stroje (ano/ne)" for m in ge["missing"]]
        if kind == "undecided":
            missing += [b for b in ge["blocking"] if b not in " ".join(missing)]
        missing = list(dict.fromkeys(missing))

    shown = [v for v in ge["variants"] if v["role"] != "horší" or v["nothing"]]
    if kind in ("recommend", "none_pays", "equal"):
        # nejsilnější vyřazené varianty zůstávají vidět (s čím se porovnává), zbytek je sbalený
        rivals = [v for v in ge["variants"] if v["role"] == "horší" and not v["nothing"]][:2]
        shown = [v for v in shown if not v["nothing"]] + rivals + [v for v in shown if v["nothing"]]
    hidden = [v for v in ge["variants"] if v not in shown]
    row = lambda v: dict(label=v["label"], amount=_amount_text(v["contribution"]), tag=ROLE_TAG[v["role"]],
                         role=v["role"], nothing=v["nothing"])
    intro = (f"Požadavky {_and(ge['requests'])} soupeří o stejné stroje a současně se nevejdou – porovnávají "
             "se jen varianty, které se vejdou do kapacity, a vždy i „nepřijmout nic“." if ge["variants"] else
             f"Požadavky {_and(ge['requests'])} soupeří o stejné stroje a současně se nevejdou – varianty se "
             "tentokrát ekonomicky neporovnávaly.")
    return dict(requests=list(ge["requests"]), kind=kind, badge=PANEL_BADGE[kind], headline=headline, why=why,
                intro=intro, sim=sim_spec(ge),
                variants=[row(v) for v in shown], hidden=[row(v) for v in hidden],
                hidden_total=ge["others_total"] - len([v for v in shown if v["role"] == "horší" and not v["nothing"]]),
                conditions=conditions, missing=missing,
                missing_note=("Nezadaný údaj se nepočítá jako 0 Kč. Po doplnění rozhodne aplikace jednoznačně."
                              if missing else ""),
                note=PANEL_NOTE)


def _relevant_missing(p: pl.Plan, cards: list, economies: dict) -> list:
    """Chybějící údaje, které opravdu ovlivňují zobrazené výsledky (E2.5c 4C).

    Ekonomické: jen neznámé, které jsou ve vzorci zobrazeného požadavku nebo ve verdiktu jeho skupiny
    (ne údaje nesouvisejících předběžných rezervací, ne údaje bez vlivu). Neekonomické nálezy plánovače
    (data flotily, neplatné požadavky) zůstávají. Bez technických klíčů a bez duplicit.
    """
    econ_labels = set()
    for e in p.requests.values():
        if e["economics"] is not None:
            econ_labels |= set(e["economics"].missing)
    for ge in economies.values():
        econ_labels |= {m["label"] for m in ge["missing"]} | set(ge["blocking"])
    econ_labels.add("zda přímé náklady zakázek obsahují vlastní stroje (direct_cost_includes_machines)")
    out = [x for x in p.missing_data if x not in econ_labels]
    for c in cards:
        if c["economy"] and c["economy"]["status"] == INVALID:   # neplatný údaj se nesmí ze seznamu ztratit
            out += list(c["economy"]["invalid"])
        for sc in (c["economy"] or {}).get("scenarios", []):
            out += [_unknown_name(t) for t in sc["terms"]]
        if c["economy"] and len({(x["known"], len(x["terms"])) for x in c["economy"]["scenarios"]}) > 1:
            out.append("zda přímé náklady zakázek obsahují vlastní stroje (ano/ne)")
    for ge in economies.values():
        if ge["kind"] in ("conditional", "undecided"):
            out += [b for b in ge["blocking"] if not any(x.startswith(b) for x in out)]
    return list(dict.fromkeys(out))


def evaluate(model: dm.DataModel, job_starts: dict, requests: list) -> dict:
    """Spustí planner.plan a vrátí výsledek jen z obyčejných typů (bez objektů plánovače)."""
    p = pl.plan(model, job_starts, [to_request(r) for r in requests])
    mine = [r["ref"] for r in requests]
    competing = {ref for g in p.groups if set(g["requests"]) & set(mine) for ref in g["requests"]}
    economies = {id(g): group_economy(model, p, g) for g in p.groups if set(g["requests"]) & set(mine)}
    cards = []
    for ref, e in p.requests.items():
        if ref not in mine and ref not in competing:
            continue                                   # předběžné bez vztahu k požadavkům jen v přehledu
        g = _group_of(p, ref)
        label, why = badge(dict(e, ref=ref), g)               # dřívější jednotný štítek (kompatibilita)
        r = e["request"]
        economy = request_economy(model, r) if e["operational"] == pl.OK_ALONE else None
        ec_label, ec_why = econ_badge(ref, e, g, economies.get(id(g)) if g is not None else None)
        cards.append(dict(
            ref=ref, title=e["label"], kind=r.kind, own=ref in mine, name=r.name,
            period=(r.start, r.end), needs=[tuple(x) for x in r.needs], badge=label, why=why,
            operational=_operational_text(e), economic=_economic_text(e, _late_text(model, r)),
            # záporné volno se nezobrazuje: den je už přetížený potvrzenými závazky
            blocking=[dict(Stroj=t, Den=int(d), **{"Volné kusy": max(0, int(free)), "Potřeba": int(need),
                                                   "Co stroj drží": ("; ".join(c) or "kus není k dispozici")
                                                   + (f" – už přetíženo o {-int(free)} ks" if free < 0 else "")})
                      for t, d, free, need, c in e["problems"]],
            alternatives=[a["text"] for a in e["alternatives"]],
            note=e["note"],
            # E2.5c 4A: strukturovaná ekonomika požadavku (jen proveditelné; jinak None)
            economy=economy,
            # E2.5c 4B: dva oddělené štítky a ekonomické řádky karty
            op_badge=op_badge(e), econ_badge=ec_label, econ_why=ec_why,
            economy_lines=contribution_lines(economy, _late_text(model, r))))
    cards.sort(key=lambda c: (not c["own"], c["ref"]))      # nejdřív požadavky uživatele
    # E2.5c 4C: karta ví, ke kterému panelu „Ekonomické porovnání“ patří (stejná skupina plánovače)
    for c in cards:
        g = _group_of(p, c["ref"])
        c["group_refs"] = list(g["requests"]) if g is not None and g["verdict"] != pl.ALL_FIT else None
    panels = [x for x in (group_panel(economies[id(g)]) for g in p.groups if id(g) in economies) if x]
    groups = [dict(requests=list(g["requests"]), verdict=g["verdict"], text=g["text"],
                   options=[list(o) for o in g["options"]], choice=list(g["choice"]) if g.get("choice") else None,
                   economy=economies[id(g)])
              for g in p.groups if set(g["requests"]) & set(mine)]
    types = {t for r in requests for t, _ in r["needs"]}
    lo = min((r["start"] for r in requests), default=None)
    hi = max((r["end"] for r in requests), default=None)
    manual = []
    unit_type = dict(zip(model.fleet_units["unit_id"].astype(str), model.fleet_units["type_id"].astype(str)))
    for m in p.manual:
        c = m["conflict"]
        t = c["id"] if c["scope"] == "typ" else unit_type.get(c["id"], c["id"])
        related = t in types and lo is not None and c["days"][0] <= hi and c["days"][-1] >= lo
        manual.append(dict(text=m["text"], alternatives=[a["text"] for a in m["alternatives"]], related=related))
    tentative = [dict(ref=r.ref, type_id=r.needs[0][0], start=r.start, end=r.end)
                 for r in pl.tentative_requests(model, job_starts)]
    return dict(cards=cards, groups=groups, panels=panels, missing_relevant=_relevant_missing(p, cards, economies),
                manual=manual, missing=list(p.missing_data), notes=list(p.notes),
                explanation=pl.explain(p), tentative=tentative)


# ---------------------------------------------------------------------------
# HTML karty
# ---------------------------------------------------------------------------

CSS = """
<style>
.dp-card{border:1px solid #E3E6EA;border-radius:10px;padding:.8rem 1rem;margin:.5rem 0 .3rem 0;background:#fff}
.dp-head{display:flex;flex-wrap:wrap;align-items:center;gap:.5rem;margin-bottom:.35rem}
.dp-title{font-weight:700;color:#1F2933}
.dp-sub{color:#4B5563;font-size:.85rem}
.dp-badge{display:inline-block;font-weight:700;font-size:.8rem;border-radius:999px;padding:.12rem .6rem}
.dp-badge.rec{background:#1E7B4A;color:#fff}
.dp-badge.ok{background:#E3F1E8;color:#14532D;border:1px solid #1E7B4A}
.dp-badge.no{background:#FDECEA;color:#8A1C12;border:1px solid #B42318}
.dp-badge.und{background:#FFF4DB;color:#7A4E00;border:1px solid #B7791F}
.dp-badges{display:flex;flex-wrap:wrap;gap:.35rem;margin-bottom:.3rem}
.dp-badge small{font-weight:600;opacity:.8;margin-right:.25rem}
.dp-badge.op-ok{background:#EEF2F7;color:#334155;border:1px solid #64748B}
.dp-badge.op-no{background:#FDECEA;color:#8A1C12;border:1px solid #B42318}
.dp-badge.op-und{background:#FFF4DB;color:#7A4E00;border:1px solid #B7791F}
.dp-badge.ec-rec{background:#1E7B4A;color:#fff;border:1px solid #1E7B4A}
.dp-badge.ec-no{background:#F3F4F6;color:#1F2933;border:1px solid #4B5563}
.dp-badge.ec-eq{background:#E8F0FE;color:#1E3A8A;border:1px solid #2563EB}
.dp-badge.ec-cond{background:#F3E8FF;color:#5B21B6;border:1px solid #7C3AED}
.dp-badge.ec-und{background:#FFF4DB;color:#7A4E00;border:1px solid #B7791F}
.dp-badge.ec-none{background:#fff;color:#4B5563;border:1px dashed #9CA3AF}
.dp-sec{color:#1F2933;margin:.25rem 0 .1rem 0;font-size:.92rem;overflow-wrap:anywhere}
.dp-panel{border:1px solid #CBD5E1;border-left:4px solid #475569;border-radius:10px;padding:.8rem 1rem;
  margin:.4rem 0 .8rem 0;background:#F8FAFC;overflow-wrap:anywhere}
.dp-panel.k-recommend{border-left-color:#1E7B4A}.dp-panel.k-none_pays{border-left-color:#4B5563}
.dp-panel.k-equal{border-left-color:#2563EB}.dp-panel.k-conditional{border-left-color:#7C3AED}
.dp-panel.k-undecided{border-left-color:#B7791F}
.dp-ptitle{font-size:.8rem;color:#475569;font-weight:700;text-transform:uppercase;letter-spacing:.02em}
.dp-answer{font-size:1.05rem;font-weight:700;color:#111827;margin:.3rem 0 .4rem 0}
.dp-h{font-weight:700;color:#1F2933;font-size:.9rem;margin:.5rem 0 .15rem 0}
.dp-panel ul{margin:.1rem 0 .2rem 1.1rem;padding:0}.dp-panel li{margin:.1rem 0;font-size:.9rem;color:#1F2933}
.dp-var{display:flex;flex-wrap:wrap;gap:.15rem .5rem;align-items:baseline;padding:.25rem 0;border-top:1px solid #E5E7EB}
.dp-var b{min-width:7rem}.dp-var .amt{flex:1 1 12rem;color:#1F2933;font-size:.9rem}
.dp-var .tag{font-size:.72rem;color:#4B5563;border:1px solid #D1D5DB;border-radius:4px;padding:0 .3rem}
.dp-var.rec .tag{color:#fff;background:#1E7B4A;border-color:#1E7B4A}
.dp-note{font-size:.85rem;color:#374151;margin-top:.5rem}
.dp-panel details{margin:.2rem 0}.dp-panel summary{cursor:pointer;font-size:.85rem;color:#374151}
.dp-chart{margin:.35rem 0 .45rem 0;padding:.55rem .6rem;background:#fff;border:1px solid #E5E7EB;border-radius:8px}
.dp-chart-t{font-size:.85rem;font-weight:700;color:#1F2933;margin-bottom:.4rem}
.dp-bar{display:flex;align-items:stretch;gap:0}
.dp-seg{flex:1 1 0;min-width:0;padding:.45rem .5rem;border:1px solid #94A3B8;border-radius:6px;
  overflow-wrap:break-word;word-break:normal;hyphens:manual}
.dp-seg-name{font-weight:700;font-size:.88rem;color:#111827;line-height:1.25}
.dp-seg-range{font-size:.8rem;color:#374151;margin-top:.15rem}
.dp-seg.c0{background:#DBEAFE;border-color:#3B82F6}.dp-seg.c1{background:#EDE9FE;border-color:#8B5CF6}
.dp-seg.c2{background:#FEF3C7;border-color:#D97706}.dp-seg.c3{background:#E0F2FE;border-color:#0284C7}
.dp-seg.c4{background:#FCE7F3;border-color:#DB2777}
.dp-seg.tie{border-style:dashed}
.dp-seg.nothing{background:repeating-linear-gradient(135deg,#F3F4F6 0 6px,#E5E7EB 6px 12px);border-color:#6B7280;
  border-style:dashed}
.dp-bnd{flex:0 0 auto;display:flex;flex-direction:column;align-items:center;justify-content:center;
  padding:0 .3rem;min-width:3.6rem}
.dp-bnd-n{display:inline-flex;align-items:center;justify-content:center;width:1.3rem;height:1.3rem;border-radius:50%;
  background:#1F2933;color:#fff;font-size:.72rem;font-weight:700}
.dp-bnd-v{font-size:.72rem;font-weight:700;color:#1F2933;white-space:nowrap;margin-top:.15rem}
.dp-bnd-list{margin:.45rem 0 .1rem 1.1rem;padding:0}.dp-bnd-list li{font-size:.85rem}
.dp-chart-n{font-size:.8rem;color:#4B5563;margin-top:.2rem}
.dp-sim-tag{display:inline-block;margin-top:.25rem;font-size:.72rem;font-weight:700;color:#0F172A;background:#fff;
  border:1px dashed #0F172A;border-radius:4px;padding:0 .3rem;white-space:nowrap}
.dp-seg.sim{box-shadow:0 0 0 2px #0F172A inset}
.dp-sim{border:2px dashed #64748B;border-radius:10px;padding:.7rem .9rem;margin:.2rem 0 .9rem 0;background:#F8FAFC;
  overflow-wrap:break-word}
.dp-sim-t{font-size:.8rem;font-weight:700;color:#334155;text-transform:uppercase;letter-spacing:.02em}
.dp-sim-v{font-weight:700;color:#0F172A;margin:.3rem 0 .2rem 0}
.dp-sim-row{display:flex;flex-wrap:wrap;gap:.1rem .5rem;align-items:baseline;padding:.2rem 0;border-top:1px solid #E5E7EB}
.dp-sim-row b{min-width:7rem}.dp-sim-row .amt{flex:1 1 9rem;font-size:.9rem;color:#1F2933}
.dp-sim-win{font-size:.78rem;font-weight:700;color:#0F172A}
@media (max-width: 640px){
  .dp-bar{flex-direction:column}
  .dp-bnd{flex-direction:row;justify-content:flex-start;gap:.4rem;padding:.25rem 0 .25rem .6rem;min-width:0}
  .dp-bnd-v{margin-top:0}
}
@media (max-width: 900px){
  .dp-chart.many .dp-bar{flex-direction:column}
  .dp-chart.many .dp-bnd{flex-direction:row;justify-content:flex-start;gap:.4rem;padding:.25rem 0 .25rem .6rem;
    min-width:0}
  .dp-chart.many .dp-bnd-v{margin-top:0}
}
.dp-title,.dp-sub{overflow-wrap:anywhere}
.dp-econ{color:#374151;font-size:.9rem;margin:.1rem 0 .1rem .1rem;overflow-wrap:anywhere}
.dp-why{color:#1F2933;margin:.2rem 0 .35rem 0}
.dp-line{color:#374151;font-size:.9rem;margin:.1rem 0}
.dp-tag{display:inline-block;font-size:.72rem;color:#4B5563;border:1px solid #D1D5DB;border-radius:4px;padding:0 .3rem}
</style>
"""


def _h(text) -> str:
    """Text do HTML karty: escapování HTML a závorek (ani Markdown z něj neudělá odkaz)."""
    return html.escape(str(text)).replace("[", "&#91;").replace("]", "&#93;").replace("(", "&#40;").replace(")", "&#41;")


def _nowrap_amounts(text: str) -> str:
    """Nezlomitelné mezery uvnitř částek („16 000 Kč“, „2 750–9 000 Kč/den“) – jen pro zobrazení."""
    text = re.sub(r"(?<=\d) (?=\d{3}\b)", "\u00a0", text)
    return re.sub(r"(?<=\d) (?=Kč)", "\u00a0", text)


def card_html(card: dict, names: dict) -> str:
    """Karta požadavku: dva oddělené štítky (Provoz, Ekonomika), provoz a ekonomika zvlášť."""
    machines = ", ".join(f"{names.get(t, t)} × {q}" for t, q in card["needs"])
    origin = "váš požadavek" if card["own"] else "předběžná rezervace z dat – neruší se"
    badges = (f'<span class="dp-badge {OP_CLASS[card["op_badge"]]}"><small>Provoz:</small>'
              f'{_h(card["op_badge"])}</span>')
    if card["econ_badge"]:
        badges += (f'<span class="dp-badge {EC_CLASS[card["econ_badge"]]}"><small>Ekonomika:</small>'
                   f'{_h(card["econ_badge"])}</span>')
    econ = "".join(f'<div class="dp-econ">{_h(_nowrap_amounts(x))}</div>' for x in card["economy_lines"])
    return (f'<div class="dp-card"><div class="dp-badges">{badges}</div>'
            f'<div class="dp-head"><span class="dp-title">{_h(card["title"])}</span>'
            f'<span class="dp-tag">{_h(origin)}</span></div>'
            f'<div class="dp-sub">{_h(machines)} · dny {card["period"][0]}–{card["period"][1]}'
            + (f' · {_h(card["name"])}' if card["name"] else "") + "</div>"
            f'<div class="dp-sec"><b>Provoz:</b> {_h(card["operational"].removeprefix("Provozně: "))}</div>'
            f'<div class="dp-sec"><b>Ekonomika:</b> {_h(_nowrap_amounts(card["econ_why"]))}</div>{econ}</div>')


def chart_html(chart: dict, marker: Optional[dict] = None) -> str:
    """Pruh hranic zlomu: úseky s nejvýhodnější variantou a očíslované hranice; text nese celou informaci,
    barva jen pomáhá. Na úzkém displeji se pruh skládá svisle (CSS).
    marker … poloha hypotetické hodnoty z „Co kdyby“ (úsek nebo hranice) – jen popisek, úseky se nemění."""
    t = lambda x: _h(_nowrap_amounts(x))
    tag = (f'<div class="dp-sim-tag">▼ předpoklad {t(marker["x_text"])}</div>' if marker else "")
    parts = []
    for i, seg in enumerate(chart["segments"]):
        if i > 0:
            b = chart["boundaries"][i - 1]
            here = marker and marker["position"]["kind"] == "boundary" and marker["position"]["n"] == b["n"]
            parts.append(f'<div class="dp-bnd{" sim" if here else ""}" title="{t(b["value"] + ": " + b["text"])}">'
                         f'<span class="dp-bnd-n">{b["n"]}</span><span class="dp-bnd-v">{t(b["value"])}</span>'
                         + (tag if here else "") + "</div>")
        cls = "dp-seg nothing" if seg["nothing"] else f'dp-seg c{seg["color"]}' + (" tie" if seg["tie"] else "")
        here = marker and marker["position"]["kind"] == "segment" and marker["position"]["index"] == i
        parts.append(f'<div class="{cls}{" sim" if here else ""}"><div class="dp-seg-name">{_h(seg["label"])}</div>'
                     f'<div class="dp-seg-range">{t(seg["range"])}</div>' + (tag if here else "") + "</div>")
    items = "".join(f'<li><b>Hranice {b["n"]} – přesně {t(b["value"])}:</b> {_h(b["text"])}.'
                    + (f'<details><summary>Zobrazit všechny shodné varianty ({len(b["winners"])})</summary>'
                       f'{_h(", ".join(b["winners"]))}</details>' if b.get("others") else "") + "</li>"
                    for b in chart["boundaries"])
    more = chart.get("never_more", [])
    notes = [x for x in chart["never"] if x]
    tail = [x for x in [chart["zero_note"], CHART_SCALE_NOTE] if x]
    many = " many" if len(chart["segments"]) > MAX_CHART_ROW else ""
    return (f'<div class="dp-chart{many}" role="img" aria-label="{_h(chart["summary"])}">'
            f'<div class="dp-chart-t">Nejvýhodnější varianta podle: {t(chart["unknown"])}</div>'
            f'<div class="dp-bar">{"".join(parts)}</div>'
            f'<ul class="dp-bnd-list">{items}</ul>'
            + "".join(f'<div class="dp-chart-n">{t(x)}</div>' for x in notes)
            + (f'<details class="dp-chart-n"><summary>Další varianty, které nejsou nejvýhodnější nikdy ({len(more)})'
               f'</summary>' + "".join(f"<div>{t(x)}</div>" for x in more) + "</details>" if more else "")
            + "".join(f'<div class="dp-chart-n">{t(x)}</div>' for x in tail) + "</div>")


def panel_html(panel: dict, sim_result: Optional[dict] = None) -> str:
    """Panel „Ekonomické porovnání“ – bloky pod sebou (bez tabulek), funguje i na úzkém displeji.
    sim_result … výsledek „Co kdyby“ (jen značka v grafu; skutečné doporučení se nemění)."""
    t = lambda x: _h(_nowrap_amounts(x))

    def var_rows(rows):
        return "".join(
            f'<div class="dp-var{" rec" if r["role"] == "doporučeno" else ""}"><b>{_h(r["label"])}</b>'
            f'<span class="amt">{t(r["amount"])}</span><span class="tag">{_h(r["tag"])}</span></div>' for r in rows)

    out = [f'<div class="dp-panel k-{panel["kind"]}">',
           f'<div class="dp-ptitle">Ekonomické porovnání · {_h(", ".join(panel["requests"]))}</div>',
           f'<div class="dp-badges"><span class="dp-badge {EC_CLASS[panel["badge"]]}"><small>Ekonomika:</small>'
           f'{_h(panel["badge"])}</span></div>',
           f'<div class="dp-answer">{t(panel["headline"])}</div>',
           f'<div class="dp-note">{_h(panel["intro"])}</div>']
    if panel["why"]:
        out.append('<div class="dp-h">Proč?</div><ul>' + "".join(f"<li>{t(x)}</li>" for x in panel["why"]) + "</ul>")
    if panel["variants"]:
        out.append('<div class="dp-h">Porovnávané varianty</div>' + var_rows(panel["variants"]))
    if panel["hidden_total"]:
        out.append(f'<details><summary>Další varianty ({panel["hidden_total"]}) – prokazatelně horší</summary>'
                   + var_rows(panel["hidden"])
                   + (f'<div class="dp-note">… a další ({panel["hidden_total"] - len(panel["hidden"])}).</div>'
                      if panel["hidden_total"] > len(panel["hidden"]) else "") + "</details>")
    if panel["conditions"]:
        out.append('<div class="dp-h">Kdy se co vyplatí</div>')
        for block in panel["conditions"]:
            if block["title"]:
                out.append(f'<div class="dp-sec"><b>{_h(block["title"])}:</b></div>')
            if block.get("chart"):
                out.append(chart_html(block["chart"], sim_result))
            elif block.get("chart_note"):
                out.append(f'<div class="dp-note">{_h(block["chart_note"])}</div>')
            if block["lead"]:
                out.append(f'<div class="dp-sec">{t(block["lead"])}</div>')
            out.append("<ul>" + "".join(f"<li>{t(x)}</li>" for x in block["lines"]) + "</ul>")
            if block["more"]:
                out.append(f'<details><summary>Další podmínky ({len(block["more"])})</summary><ul>'
                           + "".join(f"<li>{t(x)}</li>" for x in block["more"]) + "</ul></details>")
        if any(b["break_even"] for b in panel["conditions"]):
            out.append(f'<div class="dp-note">{_h(BREAK_EVEN)}</div>')
        sim = panel.get("sim")
        if sim and not sim["available"]:
            out.append(f'<div class="dp-note">{_h(sim["reason"])}</div>')
    if panel["missing"]:
        out.append('<div class="dp-h">Co doplnit</div><ul>' + "".join(f"<li>{t(x)}</li>" for x in panel["missing"])
                   + f'</ul><div class="dp-note">{_h(panel["missing_note"])}</div>')
    out.append(f'<div class="dp-note">{_h(panel["note"])}</div></div>')
    return "".join(out)


def blocking_table(card: dict, names: dict) -> pd.DataFrame:
    df = pd.DataFrame(card["blocking"], columns=["Stroj", "Den", "Volné kusy", "Potřeba", "Co stroj drží"])
    df["Stroj"] = df["Stroj"].map(lambda t: names.get(t, t))
    return df
