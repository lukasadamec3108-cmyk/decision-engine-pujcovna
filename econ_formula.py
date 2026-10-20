"""Ekonomický příspěvek jako vzorec s neznámými údaji (etapa E2.5c, krok 2).

Výpočetní vrstva bez Streamlitu; aplikace ji zatím nezobrazuje (navazuje krok 3).

PROČ
Když chybí variabilní náklad stroje, economics.py umí dát jen horní mez příspěvku. Dvě varianty,
které používají stejnou neznámou sazbu, se pak nedají porovnat, přestože se neznámá část
v rozdílu vyruší. Tento modul proto drží příspěvek jako vzorec

    příspěvek = známá částka + Σ koeficient × neznámá hodnota        (neznámé ≥ 0)

U nákladové sazby je koeficient záporný (−počet strojodnů), takže jde o tvar
„známá částka − Σ (jednotky využití × neznámá sazba)“.

JAK (žádná druhá ekonomická logika)
Příspěvek z economics.py je v každém chybějícím údaji lineární. Koeficienty se proto zjistí
přímo z economics.py (přes planner.contribution): jednou s chybějícími údaji = 0 (známá
částka) a pro každou neznámou zvlášť s hodnotou 1 (koeficient). Kontrola s více neznámými
najednou ověří linearitu. Doplnění 0/1 je jen v kopii v paměti – data se nemění a výsledek
nikdy nepředstírá, že chybějící údaj je nula: neznámá zůstává ve vzorci jako proměnná.

- skutečná nula v datech („0“) není neznámá → ve vzorci se neobjeví,
- prázdná buňka → neznámá proměnná (koeficient 0 = na výsledku nezáleží, vypadne),
- neplatný údaj (text, záporné číslo) → vzorec nelze sestavit.

SCÉNÁŘE direct_cost_includes_machines
Chybí-li nastavení, počítají se dva scénáře: „ano“ (náklady strojů jsou v přímých nákladech
zakázky – sazba stroje na zakázce se neodečítá) a „ne“ (odečítá se). Jde o jedno společné
nastavení firmy, takže se varianty porovnávají vždy ve stejném scénáři. V žádném scénáři se
náklad stroje neodečte dvakrát – rozhoduje o tom economics.job_contribution.

POROVNÁNÍ DVOU VARIANT
Rozdíl A − B je zase vzorec; společné neznámé části se v něm vyruší (koeficient 0).
- JISTÝ výsledek: znaménko rozdílu je stejné pro všechny přípustné hodnoty neznámých,
- PODMÍNĚNÝ výsledek: záleží jen na neznámých variabilních nákladových sazbách (nebo na
  scénáři nastavení) – podmínka se vypíše (u jedné sazby jako hranice zlomu),
- NELZE ROZHODNOUT: výsledek závisí na chybějícím základním údaji požadavku (tržba, přímé
  náklady, cena, sazba penále), nebo je údaj neplatný.
Modul nic nedoporučuje – jen porovnává. Doporučení postaví krok 3.
"""

import dataclasses
import math
from dataclasses import dataclass, field
from typing import Optional

import pandas as pd

import datamodel as dm
import economics as ec
import planner as pl

CERTAIN, CONDITIONAL, UNDECIDABLE = "jistý", "podmíněný", "nelze rozhodnout"
A, B, EQUAL = "A", "B", "shoda"
RATE, DATA = "sazba", "údaj"             # druh neznámé: variabilní nákladová sazba / základní údaj
ONLY = ""                                # jediný scénář (nastavení je zadané nebo nehraje roli)
SCENARIO_LABEL = {ONLY: "", "ano": "přímé náklady zakázek obsahují stroje",
                  "ne": "přímé náklady zakázek neobsahují stroje"}
EPS = 1e-6


# ---------------------------------------------------------------------------
# Vzorec
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class Unknown:
    key: tuple          # (pole, id) – např. ("variable_cost_rental_day", "MINI_BAGR")
    label: str          # srozumitelný název
    kind: str           # RATE | DATA
    unit: str = ""      # jednotka koeficientu (např. „strojodnů“)


@dataclass(frozen=True)
class Formula:
    """known + Σ coef × neznámá; terms = ((klíč, coef), …) seřazené, bez nulových koeficientů."""
    known: float
    terms: tuple = ()

    @staticmethod
    def of(known: float, coefs: dict) -> "Formula":
        terms = tuple(sorted((k, _clean(c)) for k, c in coefs.items() if abs(c) > EPS))
        return Formula(_clean(known), terms)

    def coefs(self) -> dict:
        return dict(self.terms)

    def __add__(self, other: "Formula") -> "Formula":
        c = self.coefs()
        for k, v in other.terms:
            c[k] = c.get(k, 0.0) + v
        return Formula.of(self.known + other.known, c)

    def __neg__(self) -> "Formula":
        return Formula.of(-self.known, {k: -v for k, v in self.terms})

    def __sub__(self, other: "Formula") -> "Formula":
        return self + (-other)

    def value(self, values: dict) -> float:
        """Hodnota pro konkrétní hodnoty neznámých (jen pro testy a příklady)."""
        return self.known + sum(c * values[k] for k, c in self.terms)

    @property
    def exact(self) -> bool:
        return not self.terms


ZERO = Formula(0.0)


def _clean(x: float) -> float:
    r = round(x, 6)
    return 0.0 if r == 0 else r


@dataclass(frozen=True)
class Contribution:
    """Příspěvek požadavku nebo varianty: vzorec pro každý scénář nastavení.

    scenarios … {ONLY: Formula} nebo {"ano": Formula, "ne": Formula}; None = nelze sestavit
    unknowns  … {klíč: Unknown} – popisy neznámých, které se ve vzorcích vyskytují
    invalid   … důvody, proč vzorec nelze sestavit (neplatné údaje)
    """
    scenarios: Optional[dict]
    unknowns: dict = field(default_factory=dict)
    invalid: tuple = ()

    @property
    def valid(self) -> bool:
        return self.scenarios is not None


NOTHING = Contribution({ONLY: ZERO})     # varianta „nepřijmout nic“: přírůstkový příspěvek přesně 0 Kč


def _expand(a: dict, b: dict) -> tuple:
    """Sjednotí scénáře dvou vzorců (jediný scénář platí pro „ano“ i „ne“)."""
    if set(a) == set(b):
        return a, b
    keys = ("ano", "ne")
    return ({k: a.get(k, a.get(ONLY)) for k in keys}, {k: b.get(k, b.get(ONLY)) for k in keys})


def combine(parts) -> Contribution:
    """Příspěvek varianty = součet příspěvků jejích požadavků (každý požadavek jednou)."""
    out = NOTHING
    for p in parts:
        if not out.valid or not p.valid:
            out = Contribution(None, {**out.unknowns, **p.unknowns}, out.invalid + p.invalid)
            continue
        a, b = _expand(out.scenarios, p.scenarios)
        out = Contribution({k: a[k] + b[k] for k in a}, {**out.unknowns, **p.unknowns})
    return out


# ---------------------------------------------------------------------------
# Sestavení vzorce z economics.py
# ---------------------------------------------------------------------------

def _empty(v) -> bool:
    return v is None or str(v).strip() == "" or str(v).strip().lower() == "nan"


def _type_name(model: dm.DataModel, type_id: str) -> str:
    rows = model.machine_types[model.machine_types["type_id"] == type_id]
    return str(rows.iloc[0]["name"]) if len(rows) else type_id


def _type_cell(model, type_id, column):
    rows = model.machine_types[model.machine_types["type_id"] == type_id]
    return rows.iloc[0][column] if len(rows) else ""


def _unknowns(model: dm.DataModel, req: pl.Request) -> list:
    """Chybějící údaje, na kterých příspěvek požadavku může záviset."""
    out = []
    for t, _ in req.needs:
        name = _type_name(model, t)
        if req.kind == "rental":
            if _empty(_type_cell(model, t, "variable_cost_rental_day")):
                out.append(Unknown(("variable_cost_rental_day", t), f"{name}: variabilní náklad pronájmu/den",
                                   RATE, "strojodnů"))
            if req.price_day is None and _empty(_type_cell(model, t, "rental_price_day")):
                out.append(Unknown(("rental_price_day", t), f"{name}: cena pronájmu/den", DATA, "strojodnů"))
        elif _empty(_type_cell(model, t, "variable_cost_job_day")):
            out.append(Unknown(("variable_cost_job_day", t), f"{name}: variabilní náklad na zakázce/den",
                               RATE, "strojodnů"))
    if req.kind == "job":
        label = req.name or req.ref
        row = None
        if req.in_data:
            rows = model.jobs[model.jobs["job_id"] == req.ref]
            row = rows.iloc[0] if len(rows) else None
        for col, text in (("revenue", "tržba"), ("direct_cost", "přímé náklady"), ("penalty_per_day", "penále/den")):
            missing = _empty(row[col]) if row is not None else getattr(req, col) is None
            if missing:
                out.append(Unknown((col, req.ref), f"{req.ref} ({label}): {text}", DATA,
                                   "dní po termínu" if col == "penalty_per_day" else "Kč"))
    return list({u.key: u for u in out}.values())


def _patched(model: dm.DataModel, req: pl.Request, values: dict, scenario: str):
    """Kopie modelu a požadavku s doplněnými hodnotami neznámých (jen v paměti)."""
    mt = model.machine_types.copy()
    jobs = model.jobs.copy()
    settings = model.settings.copy()
    changes = {}
    for (col, ident), v in values.items():
        if col in ("variable_cost_rental_day", "variable_cost_job_day", "rental_price_day"):
            mt.loc[mt["type_id"] == ident, col] = str(v)
        elif req.in_data:
            jobs.loc[jobs["job_id"] == ident, col] = str(v)
        else:
            changes[col] = float(v)
    if scenario != ONLY:
        settings = pd.concat([settings[settings["key"] != "direct_cost_includes_machines"],
                              pd.DataFrame([dict(key="direct_cost_includes_machines", value=scenario,
                                                 unit="", note="")])], ignore_index=True)
    tables = {n: getattr(model, n) for n in dm.SCHEMA}
    tables.update(machine_types=mt, jobs=jobs, settings=settings)
    return dm.DataModel(**tables), dataclasses.replace(req, **changes) if changes else req


def _evaluate(model, req, values, scenario):
    """(částka, neplatné údaje) z economics.py přes planner.contribution."""
    m, r = _patched(model, req, values, scenario)
    v = pl.contribution(m, r)
    res = v.result
    if res is None or res.value is None or res.invalid:
        return None, tuple(res.invalid) if res is not None else ("neplatný údaj",)
    return float(res.value), ()


def request_formula(model: dm.DataModel, req: pl.Request) -> Contribution:
    """Příspěvek požadavku jako vzorec (pro každý scénář nastavení zvlášť)."""
    unknowns = _unknowns(model, req)
    scenarios = [ONLY]
    if req.kind == "job" and ec.direct_cost_includes_machines(model) is None:
        scenarios = ["ano", "ne"]
    out = {}
    for sc in scenarios:
        base = {u.key: 0.0 for u in unknowns}
        known, invalid = _evaluate(model, req, base, sc)
        if invalid:
            return Contribution(None, {u.key: u for u in unknowns}, invalid)
        coefs = {}
        for u in unknowns:
            val, invalid = _evaluate(model, req, {**base, u.key: 1.0}, sc)
            if invalid:
                return Contribution(None, {u.key: u for u in unknowns}, invalid)
            coefs[u.key] = val - known
        f = Formula.of(known, coefs)
        if unknowns:                                  # kontrola linearity: všechny neznámé najednou
            probe = {u.key: 7.0 + i for i, u in enumerate(unknowns)}
            val, _ = _evaluate(model, req, probe, sc)
            if val is None or not math.isclose(val, f.value(probe), abs_tol=1e-3):
                raise ArithmeticError(f"Příspěvek {req.ref} není lineární v chybějících údajích.")
        out[sc] = f
    used = {k for f in out.values() for k, _ in f.terms}
    return Contribution(out, {u.key: u for u in unknowns if u.key in used})


def variant_formula(model: dm.DataModel, requests) -> Contribution:
    """Příspěvek varianty (kombinace požadavků); prázdná varianta = „nepřijmout nic“ (0 Kč)."""
    return combine(request_formula(model, r) for r in requests)


def group_formulas(model: dm.DataModel, plan: pl.Plan, group: dict) -> dict:
    """Vzorce všech porovnávaných variant soupeřící skupiny z planner.plan (včetně „nic“)."""
    reqs = {ref: plan.requests[ref]["request"] for ref in group["requests"]}
    cache = {ref: request_formula(model, r) for ref, r in reqs.items()}
    return {v: combine(cache[ref] for ref in v) for v in group.get("candidates", [])}


# ---------------------------------------------------------------------------
# Porovnání dvou variant
# ---------------------------------------------------------------------------

def _czk(x: float) -> str:
    return f"{x:,.0f} Kč".replace(",", " ")


def _num(x: float) -> str:
    return f"{x:,.2f}".replace(",", " ").rstrip("0").rstrip(".")


def _compare_one(diff: Formula, unknowns: dict) -> dict:
    """Znaménko rozdílu A − B pro všechny neznámé ≥ 0."""
    k, terms = diff.known, diff.terms
    if not terms:
        winner = A if k > EPS else B if k < -EPS else EQUAL
        return dict(status=CERTAIN, winner=winner, strict=True, diff=diff, condition="")
    pos = all(c > 0 for _, c in terms)
    neg = all(c < 0 for _, c in terms)
    if pos and k >= -EPS:                     # min rozdílu = k (při nulových neznámých)
        return dict(status=CERTAIN, winner=A, strict=k > EPS, diff=diff, condition="")
    if neg and k <= EPS:
        return dict(status=CERTAIN, winner=B, strict=k < -EPS, diff=diff, condition="")
    data = [key for key, _ in terms if unknowns.get(key) is None or unknowns[key].kind != RATE]
    if data:
        return dict(status=UNDECIDABLE, winner=None, strict=False, diff=diff, condition="",
                    blocking=[unknowns[key].label if key in unknowns else str(key) for key in data])
    return dict(status=CONDITIONAL, winner=None, strict=False, diff=diff, condition=condition(diff, unknowns))


def condition(diff: Formula, unknowns: dict) -> str:
    """Kdy je A lepší než B (rozdíl > 0) – u jedné sazby jako hranice zlomu."""
    label = lambda key: f"[{unknowns[key].label}]" if key in unknowns else str(key)
    if len(diff.terms) == 1:
        key, c = diff.terms[0]
        t = -diff.known / c
        return f"A je lepší, pokud {label(key)} {'<' if c < 0 else '>'} {_czk(t)} (hranice zlomu)"
    left = " + ".join(f"{_num(-c)} × {label(key)}" for key, c in diff.terms if c < 0)
    right = " + ".join(f"{_num(c)} × {label(key)}" for key, c in diff.terms if c > 0)
    rhs = _czk(diff.known) + (f" + {right}" if right else "")
    return f"A je lepší, pokud {left or '0'} < {rhs}"


def compare(a: Contribution, b: Contribution) -> dict:
    """Porovná varianty A a B. Vrací status (jistý/podmíněný/nelze rozhodnout), vítěze
    (A/B/shoda, jen u jistého výsledku), rozdíl A − B po scénářích a srozumitelný text."""
    unknowns = {**a.unknowns, **b.unknowns}
    if not a.valid or not b.valid:
        reasons = tuple(dict.fromkeys(a.invalid + b.invalid))
        return dict(status=UNDECIDABLE, winner=None, scenarios={}, reasons=reasons,
                    text="Nelze rozhodnout – neplatné údaje: " + "; ".join(reasons))
    sa, sb = _expand(a.scenarios, b.scenarios)
    per = {sc: _compare_one(sa[sc] - sb[sc], unknowns) for sc in sa}
    statuses = {r["status"] for r in per.values()}
    winners = {r["winner"] for r in per.values()}
    if UNDECIDABLE in statuses:
        blocking = list(dict.fromkeys(x for r in per.values() for x in r.get("blocking", [])))
        return dict(status=UNDECIDABLE, winner=None, scenarios=per, reasons=tuple(blocking),
                    text="Nelze rozhodnout – výsledek závisí na chybějících údajích: " + "; ".join(blocking))
    if statuses == {CERTAIN} and len(winners) == 1:
        w = winners.pop()
        strict = all(r["strict"] for r in per.values())
        if w == EQUAL:
            text = "Jistě shodné – rozdíl je přesně 0 Kč."
        else:
            text = (f"Jistě lepší je varianta {w}" + ("" if strict else
                    " (nikdy horší; shoda jen při nulových neznámých sazbách)") + ".")
        return dict(status=CERTAIN, winner=w, strict=strict, scenarios=per, reasons=(), text=text)
    parts = []
    for sc, r in per.items():
        prefix = f"pokud {SCENARIO_LABEL[sc]}: " if sc != ONLY else ""
        if r["status"] == CERTAIN:
            parts.append(prefix + ("shoda" if r["winner"] == EQUAL else f"lepší je {r['winner']}"))
        else:
            parts.append(prefix + r["condition"])
    return dict(status=CONDITIONAL, winner=None, scenarios=per, reasons=(),
                text="Výsledek závisí na neznámých nákladech – " + "; ".join(parts) + ".")


# ---------------------------------------------------------------------------
# Popis a meze (pro kontrolu se stávajícím výpočtem a pro krok 3)
# ---------------------------------------------------------------------------

def describe(c: Contribution) -> list:
    """Vzorec slovy, např. „67 500 Kč − 10 strojodnů × [Mini bagr: variabilní náklad pronájmu/den]“."""
    if not c.valid:
        return ["nelze sestavit – neplatné údaje: " + "; ".join(c.invalid)]
    out = []
    for sc, f in c.scenarios.items():
        text = _czk(f.known)
        for key, coef in f.terms:
            u = c.unknowns.get(key)
            unit = f" {u.unit}" if u and u.unit and u.unit != "Kč" else ""
            amount = "" if (u and u.unit == "Kč" and abs(abs(coef) - 1) < EPS) else f"{_num(abs(coef))}{unit} × "
            text += f" {'−' if coef < 0 else '+'} {amount}[{u.label if u else key}]"
        out.append((f"{SCENARIO_LABEL[sc]}: " if sc != ONLY else "") + text)
    return out


def bounds(c: Contribution) -> tuple:
    """(dolní, horní) mez příspěvku přes všechny scénáře a neznámé ≥ 0 (None = neomezená/neznámá)."""
    if not c.valid:
        return None, None
    los, his = [], []
    for f in c.scenarios.values():
        los.append(f.known if all(v > 0 for _, v in f.terms) else None)
        his.append(f.known if all(v < 0 for _, v in f.terms) else None)
    lo = None if any(x is None for x in los) else min(los)
    hi = None if any(x is None for x in his) else max(his)
    return lo, hi
