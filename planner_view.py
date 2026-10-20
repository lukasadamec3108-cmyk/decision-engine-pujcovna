"""Zobrazení rozhodovacího plánovače (etapa E2.5b) – záložka „Rozhodovací plánovač“.

Modul NIC nerozhoduje ani nepočítá: vše počítá planner.py (nad capacity.py a economics.py).
Tady se jen převádí formulář na požadavek a výsledek plánovače na obyčejné tabulky, štítky
a srozumitelná vysvětlení (obyčejné typy → lze uložit do mezipaměti Streamlitu).

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


def _input(label: str, raw, unit: str, field: str) -> dict:
    """Stav jednoho vstupního údaje stejnými pravidly jako economics.py (prázdné ≠ 0)."""
    r = ec._amount(raw, label, False)
    if r.invalid:
        return dict(field=field, label=label, unit=unit, state=NEPLATNE, value=None, reason=r.invalid[0])
    if r.missing:
        return dict(field=field, label=label, unit=unit, state=NEZADANO, value=None, reason="")
    return dict(field=field, label=label, unit=unit, state=ZADANO, value=float(r.value), reason="")


def request_inputs(model: dm.DataModel, r: pl.Request) -> list:
    """Ekonomické vstupy požadavku: skutečná nula / nezadáno / neplatné, s jednotkami."""
    out = []
    mt = model.machine_types.set_index("type_id")
    for t, _ in r.needs:
        name = str(mt.loc[t, "name"]) if t in mt.index else t
        if r.kind == "rental":
            if r.price_day is not None:
                out.append(_input(f"{name}: cena pronájmu/den (zadaná)", r.price_day, KC_DEN, "price_day"))
            else:
                out.append(_input(f"{name}: cena pronájmu/den (ceník)", mt.loc[t, "rental_price_day"]
                                  if t in mt.index else "", KC_DEN, "rental_price_day"))
            out.append(_input(f"{name}: variabilní náklad pronájmu/den", mt.loc[t, "variable_cost_rental_day"]
                              if t in mt.index else "", KC_DEN, "variable_cost_rental_day"))
        else:
            out.append(_input(f"{name}: variabilní náklad na zakázce/den", mt.loc[t, "variable_cost_job_day"]
                              if t in mt.index else "", KC_DEN, "variable_cost_job_day"))
    if r.kind == "job":
        row = None
        if r.in_data:
            rows = model.jobs[model.jobs["job_id"] == r.ref]
            row = rows.iloc[0] if len(rows) else None
        for col, text, unit in (("revenue", "tržba", KC), ("direct_cost", "přímé náklady", KC),
                                ("penalty_per_day", "penále/den", KC_DEN)):
            raw = row[col] if row is not None else getattr(r, col)
            out.append(_input(f"{r.ref}: {text}", "" if raw is None else raw, unit, col))
        incl = ec.direct_cost_includes_machines(model)
        out.append(dict(field="direct_cost_includes_machines",
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
                    conditions=[], missing=[], blocking=[], scenarios=[])
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
    return dict(base, choice=None if choice is None else list(choice), best=[list(v) for v in best],
                contenders=[list(v) for v in contenders], variants=variants, others_total=len(others),
                profiles=profiles, conditions=conditions, missing=missing, blocking=blocking, scenarios=scen)



def evaluate(model: dm.DataModel, job_starts: dict, requests: list) -> dict:
    """Spustí planner.plan a vrátí výsledek jen z obyčejných typů (bez objektů plánovače)."""
    p = pl.plan(model, job_starts, [to_request(r) for r in requests])
    mine = [r["ref"] for r in requests]
    competing = {ref for g in p.groups if set(g["requests"]) & set(mine) for ref in g["requests"]}
    cards = []
    for ref, e in p.requests.items():
        if ref not in mine and ref not in competing:
            continue                                   # předběžné bez vztahu k požadavkům jen v přehledu
        g = _group_of(p, ref)
        label, why = badge(dict(e, ref=ref), g)
        r = e["request"]
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
            economy=request_economy(model, r) if e["operational"] == pl.OK_ALONE else None))
    cards.sort(key=lambda c: (not c["own"], c["ref"]))      # nejdřív požadavky uživatele
    groups = [dict(requests=list(g["requests"]), verdict=g["verdict"], text=g["text"],
                   options=[list(o) for o in g["options"]], choice=list(g["choice"]) if g.get("choice") else None,
                   economy=group_economy(model, p, g))
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
    return dict(cards=cards, groups=groups, manual=manual, missing=list(p.missing_data), notes=list(p.notes),
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
.dp-why{color:#1F2933;margin:.2rem 0 .35rem 0}
.dp-line{color:#374151;font-size:.9rem;margin:.1rem 0}
.dp-tag{display:inline-block;font-size:.72rem;color:#4B5563;border:1px solid #D1D5DB;border-radius:4px;padding:0 .3rem}
</style>
"""


def _h(text) -> str:
    """Text do HTML karty: escapování HTML a závorek (ani Markdown z něj neudělá odkaz)."""
    return html.escape(str(text)).replace("[", "&#91;").replace("]", "&#93;").replace("(", "&#40;").replace(")", "&#41;")


def card_html(card: dict, names: dict) -> str:
    machines = ", ".join(f"{names.get(t, t)} × {q}" for t, q in card["needs"])
    origin = "váš požadavek" if card["own"] else "předběžná rezervace z dat – neruší se"
    lines = "".join(f'<div class="dp-line">{_h(x)}</div>' for x in (card["operational"], card["economic"]))
    return (f'<div class="dp-card"><div class="dp-head"><span class="dp-badge {BADGE_CLASS[card["badge"]]}">'
            f'{_h(card["badge"])}</span><span class="dp-title">{_h(card["title"])}</span>'
            f'<span class="dp-tag">{_h(origin)}</span></div>'
            f'<div class="dp-sub">{_h(machines)} · dny {card["period"][0]}–{card["period"][1]}'
            + (f' · {_h(card["name"])}' if card["name"] else "") + "</div>"
            f'<div class="dp-why"><b>Proč:</b> {_h(card["why"])}</div>{lines}</div>')


def blocking_table(card: dict, names: dict) -> pd.DataFrame:
    df = pd.DataFrame(card["blocking"], columns=["Stroj", "Den", "Volné kusy", "Potřeba", "Co stroj drží"])
    df["Stroj"] = df["Stroj"].map(lambda t: names.get(t, t))
    return df
