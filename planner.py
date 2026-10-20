"""Společný plánovací a rozhodovací model (etapa E2.5a) – strategie C: vyvážené rozhodování.

Samostatná logika bez Streamlitu; aplikace ji zatím nepoužívá. Nic nevymýšlí a nic nemění:
vstupní data, rezervace ani plán zakázek se nepřepisují – výsledkem je jen DOPORUČENÍ.

Využívá existující moduly (žádná druhá kapacitní ani ekonomická logika):
- capacity.py  … dostupnost, závazky, konflikty, ověření požadavku,
- economics.py … příspěvek z pronájmu a ze zakázky (bez fixních nákladů – ty jsou jen
                 jednou za flotilu, takže se nezapočítají dvakrát).

ROZHODOVACÍ PRAVIDLA
1. Servis a technická nedostupnost (kus mimo flotilu) blokují stroj – řeší capacity.py.
2. Potvrzené závazky (potvrzené pronájmy, servis, potvrzené zakázky v aktuálním plánu)
   mají přednost: blokují kapacitu, nové a předběžné požadavky se posuzují jen na zbytku.
3. Doporučení nikdy nepřekročí kapacitu – každá doporučená kombinace se ověří kapacitním modelem.
4. Konflikt potvrzených závazků se nerozhoduje – jde k ručnímu řešení. U zakázky lze nabídnout
   posun v okně bez penále, nikdy se ale automaticky nepřesouvá.
5. Předběžné a nové požadavky jsou rovnocenné; o zbývající kapacitě rozhoduje ekonomika.
   Předběžná rezervace zůstává vždy viditelná a nikdy se automaticky neruší.
6. Provozní proveditelnost a ekonomická výhodnost jsou oddělené. Ekonomicky se porovnávají jen
   kapacitně proveditelné varianty a vždy i „nepřijmout nic“ (přírůstkový příspěvek 0 Kč).
   Příspěvky jsou vzorce z econ_formula.py (nad economics.py): chybějící sazba zůstává neznámou,
   společné neznámé se v porovnání vyruší, oba scénáře direct_cost_includes_machines se respektují.
   Verdikt skupiny: doporučení (prokazatelně ostře lepší než vše ostatní), „žádný se nevyplatí“
   (všechny varianty prokazatelně ztrátové), rovnocenné (prokazatelně shodné), podmíněně lepší
   (záleží na neznámé sazbě – hranice zlomu), nebo nelze rozhodnout (chybí základní údaje).
7. Alternativy jen podložené daty: posun zakázky uvnitř jejího okna bez penále. Jiný stroj ani
   externí zapůjčení se nenabízí – chybí údaje o zaměnitelnosti strojů a o externích cenách.
"""

from dataclasses import dataclass, field
from itertools import combinations
from typing import Optional

import pandas as pd

import capacity as cp
import datamodel as dm
import economics as ec

TENTATIVE, NEW, CONFIRMED = "tentative", "new", "confirmed"
STATUS_LABEL = {TENTATIVE: "předběžný", NEW: "nový", CONFIRMED: "potvrzený"}

# Provozní stav požadavku
OK_ALONE, NOT_FEASIBLE, UNDETERMINED = "proveditelný", "neproveditelný", "nelze určit"
# Ekonomický verdikt skupiny
RECOMMEND, EQUAL, NO_DECISION, ALL_FIT = ("doporučení", "rovnocenné", "nelze rozhodnout",
                                          "vše proveditelné současně")
NONE_PAYS = "žádný se nevyplatí"       # nejlepší je nepřijmout žádný ze soupeřících požadavků
CONDITIONAL = "podmíněně lepší"         # vítěz závisí na neznámé nákladové sazbě (hranice zlomu)
NOTHING = ()                           # varianta „nepřijmout nic“ – přírůstkový příspěvek přesně 0 Kč
MAX_GROUP = 12                         # nejvíc požadavků v jedné skupině (2^12 kombinací)
MAX_FRONT = 40                         # nejvíc variant, které nelze navzájem vyloučit (jinak nelze rozhodnout)
MAX_CONDITIONS_TEXT = 6                # kolik podmínek se vypíše v souhrnném textu (všechny jsou v „conditions“)


# ---------------------------------------------------------------------------
# Vstupy
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class Request:
    """Požadavek, o kterém se rozhoduje (předběžný z dat nebo nový zadaný uživatelem).

    needs   … ((type_id, kusů), …) – zakázka potřebuje všechny stroje současně
    start, end … požadovaný termín (dny plánu, oba včetně)
    earliest, deadline … okno zakázky – jen pro nabídku posunu bez penále
    price_day … cena pronájmu/den (jinak ceník typu stroje)
    revenue, direct_cost, penalty_per_day … ekonomika nové zakázky (u zakázky z dat se berou z dat)
    """
    ref: str
    kind: str                         # rental | job
    status: str                       # tentative | new
    needs: tuple
    start: Optional[int]
    end: Optional[int]
    name: str = ""
    earliest: Optional[int] = None
    deadline: Optional[int] = None
    price_day: Optional[float] = None
    revenue: Optional[float] = None
    direct_cost: Optional[float] = None
    penalty_per_day: Optional[float] = None
    in_data: bool = False             # zakázka / rezervace existuje v datovém modelu

    @property
    def label(self) -> str:
        what = "pronájem" if self.kind == "rental" else "zakázka"
        return f"{what} {self.ref} ({STATUS_LABEL.get(self.status, self.status)})"


def tentative_requests(model: dm.DataModel, job_starts: Optional[dict] = None) -> list:
    """Předběžné požadavky z dat: předběžné rezervace a nepotvrzené zakázky s termínem v plánu."""
    out = []
    for r in model.rental_bookings.itertuples():
        if str(r.status) == TENTATIVE:
            out.append(Request(str(r.booking_id), "rental", TENTATIVE, ((str(r.type_id), _int(r.units)),),
                               _int(r.start_day), _int(r.end_day), price_day=_float(r.price_day), in_data=True))
    job_starts = job_starts or {}
    req = model.job_requirements[model.job_requirements["resource_kind"] == "machine"]
    for j in model.jobs.itertuples():
        if str(j.status) in ("tentative", "opportunity") and j.job_id in job_starts:
            s, dur = _int(job_starts[j.job_id]), _int(j.duration_days)
            needs = tuple((str(r.resource_id), _int(r.quantity)) for r in req[req["job_id"] == j.job_id].itertuples())
            out.append(Request(str(j.job_id), "job", TENTATIVE, needs, s,
                               s + dur - 1 if s is not None and dur else None, name=str(j.name),
                               earliest=_int(j.earliest_start_day), deadline=_int(j.deadline_day), in_data=True))
    return out


def _int(v):
    try:
        f = float(str(v).strip())
        return int(f) if f.is_integer() else None
    except (TypeError, ValueError):
        return None


def _float(v):
    try:
        return float(str(v).strip())
    except (TypeError, ValueError):
        return None


# ---------------------------------------------------------------------------
# Provozní proveditelnost (capacity.py)
# ---------------------------------------------------------------------------

def _validate(req: Request, cal: cp.Calendar) -> list:
    problems = []
    if req.kind not in ("rental", "job"):
        problems.append(f"neznámý druh požadavku „{req.kind}“")
    if req.status not in (TENTATIVE, NEW):
        problems.append(f"stav „{req.status}“ – rozhoduje se jen o předběžných a nových požadavcích")
    if not isinstance(req.start, int) or not isinstance(req.end, int) or req.end < req.start:
        problems.append("chybí nebo je neplatný termín")
    if not req.needs:
        problems.append("není uvedeno, jaké stroje požadavek potřebuje")
    for type_id, qty in req.needs:
        if not isinstance(qty, int) or isinstance(qty, bool) or qty < 1:
            problems.append(f"{type_id}: chybí nebo je neplatný počet strojů")
        if type_id in cal.unknown_capacity:
            problems.append(f"{type_id}: počet kusů ve flotile nelze spolehlivě určit")
        elif type_id not in cal.type_ids():
            problems.append(f"{type_id}: stroj není ve flotile")
    return problems


def _check(cal: cp.Calendar, req: Request, start: int, end: int) -> tuple:
    """(stav, problémové dny, důvody) pro všechny stroje požadavku na zbytku kapacity."""
    problems, reasons, statuses = [], [], []
    for type_id, qty in req.needs:
        r = cp.evaluate_request(cal, type_id, start, end, qty)
        statuses.append(r["status"])
        problems += [(type_id, d, free, need, causes) for d, free, need, causes in r["problem_days"]]
        reasons += r["reasons"]
    if cp.INFEASIBLE in statuses:
        return NOT_FEASIBLE, problems, reasons
    if cp.UNKNOWN in statuses:
        return UNDETERMINED, problems, reasons
    return OK_ALONE, problems, reasons


def _shift_options(cal: cp.Calendar, req: Request, limit: int = 3) -> list:
    """Posun zakázky uvnitř okna bez penále (začátek ≥ earliest, konec ≤ deadline)."""
    if req.kind != "job" or req.earliest is None or req.deadline is None or req.start is None:
        return []
    dur = req.end - req.start + 1
    out = []
    for s in range(req.earliest, req.deadline - dur + 2):
        if s != req.start and _check(cal, req, s, s + dur - 1)[0] == OK_ALONE:
            out.append((s, s + dur - 1))
            if len(out) == limit:
                break
    return out


def _residual(cal: cp.Calendar, requests: list) -> dict:
    """Volné kusy po závazcích pro dny a stroje požadavků: {(type, den): volno}."""
    out = {}
    for t in {t for r in requests for t, _ in r.needs}:
        lo = min(r.start for r in requests if any(x == t for x, _ in r.needs))
        hi = max(r.end for r in requests if any(x == t for x, _ in r.needs))
        df = cp.daily_capacity(cal, lo, hi, [t])
        for row in df.itertuples():
            out[(t, int(row.day))] = int(row.free) if row.status in (cp.FREE, cp.FULL) else None
    return out


def _fits(residual: dict, subset) -> bool:
    use = {}
    for r in subset:
        for t, q in r.needs:
            for d in range(r.start, r.end + 1):
                use[(t, d)] = use.get((t, d), 0) + q
    return all(residual.get(k) is not None and residual[k] - v >= 0 for k, v in use.items())


def _verify_with_capacity(cal: cp.Calendar, subset) -> bool:
    """Kontrola kapacitním modelem: kombinace přidaná jako potvrzená nesmí vytvořit přetížení."""
    added = [cp.Assignment(r.kind, r.ref, t, r.start, r.end, q) for r in subset for t, q in r.needs]
    test = cp.build_calendar(cal.units, cal.assignments + added, cal.open_jobs)
    for r in subset:
        for t, _ in r.needs:
            df = cp.daily_capacity(test, r.start, r.end, [t])
            if (df["free"] < 0).any() or (df["status"] == cp.UNKNOWN).any():
                return False
    return True


def _groups(requests: list) -> list:
    """Skupiny požadavků, které spolu soupeří o stejný stroj ve stejné dny (souvislé komponenty)."""
    parent = list(range(len(requests)))

    def find(i):
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i
    for i, a in enumerate(requests):
        for j in range(i + 1, len(requests)):
            b = requests[j]
            if ({t for t, _ in a.needs} & {t for t, _ in b.needs}) and cp.overlap(a.start, a.end, b.start, b.end):
                parent[find(i)] = find(j)
    comps = {}
    for i in range(len(requests)):
        comps.setdefault(find(i), []).append(requests[i])
    return list(comps.values())


# ---------------------------------------------------------------------------
# Ekonomika (economics.py)
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class Value:
    """Ekonomický přínos jako interval: lo/hi = jistá dolní / horní mez (None = neznámá)."""
    lo: Optional[float]
    hi: Optional[float]
    missing: tuple = ()
    result: Optional[ec.Result] = None

    @property
    def exact(self) -> bool:
        return self.lo is not None and self.lo == self.hi


def _with_new_job(model: dm.DataModel, req: Request) -> dm.DataModel:
    """Kopie modelu v paměti s novou zakázkou, aby šlo použít economics.job_contribution."""
    jobs = model.jobs.copy()
    row = {c: "" for c in jobs.columns}
    row.update(job_id=req.ref, name=req.name or req.ref, status="tentative", priority="Medium", risk_level="Medium",
               revenue="" if req.revenue is None else req.revenue,
               direct_cost="" if req.direct_cost is None else req.direct_cost,
               duration_days=req.end - req.start + 1, earliest_start_day=req.start,
               deadline_day=req.deadline if req.deadline is not None else req.end,
               # sazba penále se násobí dny po termínu: u zakázky dokončené v termínu je penále 0
               # bez ohledu na sazbu – nevyplněná sazba je chybějící údaj jen u zakázky po termínu
               penalty_per_day=(req.penalty_per_day if req.penalty_per_day is not None
                                else ("" if req.deadline is not None and req.end > req.deadline else "0")))
    jobs = pd.concat([jobs, pd.DataFrame([row])], ignore_index=True).astype(str)
    reqs = pd.concat([model.job_requirements, pd.DataFrame(
        [dict(job_id=req.ref, resource_kind="machine", resource_id=t, quantity=str(q), position=str(i + 1))
         for i, (t, q) in enumerate(req.needs)])], ignore_index=True)
    tables = {n: getattr(model, n) for n in dm.SCHEMA}
    tables.update(jobs=jobs, job_requirements=reqs)
    return dm.DataModel(**tables)


def contribution(model: dm.DataModel, req: Request) -> Value:
    """Příspěvek požadavku z ekonomického modulu (fixní náklady a ušlý zisk se nepřičítají)."""
    days = req.end - req.start + 1
    if req.kind == "rental":
        parts = []
        revenue_known = True
        for t, q in req.needs:
            row = model.machine_types[model.machine_types["type_id"] == t].iloc[0]
            if req.price_day is None and _float(row["rental_price_day"]) is None:
                revenue_known = False
            parts.append(ec.rental_contribution(model, t, days, req.price_day).scaled(q))
        result = ec.total(parts)
    else:
        # dny po termínu dokončení → penále (economics.job_contribution, days_late)
        late = max(0, req.end - req.deadline) if req.deadline is not None else 0
        if req.in_data:
            result = ec.job_contribution(model, req.ref, days_late=late)
            revenue_known = _float(model.jobs.set_index("job_id").loc[req.ref, "revenue"]) is not None
        else:
            result = ec.job_contribution(_with_new_job(model, req), req.ref, days_late=late)
            revenue_known = req.revenue is not None
    if result.value is None:                                   # neplatný údaj
        return Value(None, None, result.invalid + result.missing, result)
    lo = result.value if result.complete else None             # chybějí-li náklady, jen horní mez
    hi = result.value if (result.complete or revenue_known) else None
    return Value(lo, hi, result.missing, result)


def _sum(values) -> Value:
    values = list(values)
    lo = None if any(v.lo is None for v in values) else sum(v.lo for v in values)
    hi = None if any(v.hi is None for v in values) else sum(v.hi for v in values)
    missing = tuple(dict.fromkeys(m for v in values for m in v.missing))
    return Value(lo, hi, missing)


def economic_note(v: Value) -> str:
    """Ekonomický stav jednoho požadavku slovy (odděleně od provozní proveditelnosti)."""
    czk = lambda x: f"{x:,.0f} Kč".replace(",", " ")
    if v.hi is not None and v.hi < 0:
        return f"ztrátový – příspěvek nejvýše {czk(v.hi)}"
    if v.exact:
        return f"příspěvek {czk(v.lo)}"
    if v.hi is not None:
        return f"příspěvek nejvýše {czk(v.hi)} – chybí: " + "; ".join(v.missing)
    return "příspěvek nelze určit – chybí: " + ("; ".join(v.missing) or "neznámé")


def _decide(contrib: dict) -> dict:
    """Robustní ekonomický verdikt nad vzorci variant z econ_formula (krok 3 E2.5c).

    contrib … {varianta (tuple označení): econ_formula.Contribution}, vždy včetně NOTHING (0 Kč).
    - DOPORUČENO: jediná varianta je prokazatelně (ostře) lepší než všechny ostatní při jakýchkoli
      hodnotách neznámých sazeb a v obou scénářích direct_cost_includes_machines; je-li jí NOTHING,
      verdikt je „žádný se nevyplatí“.
    - ROVNOCENNÉ: nejlepší varianty mají prokazatelně shodný výsledek (stejný vzorec).
    - PODMÍNĚNĚ LEPŠÍ: vítěz závisí jen na neznámých variabilních sazbách (nebo na scénáři) –
      vrací hranice zlomu. Sem patří i „nikdy horší, ale při nulových sazbách shodné“.
    - NELZE ROZHODNOUT: rozhoduje chybějící základní údaj požadavku nebo neplatný údaj.
    Varianty, které jsou prokazatelně horší než jiná varianta, se z dalšího výkladu vyřadí.
    """
    import econ_formula as ef               # pozdní import: econ_formula sám používá planner
    variants = list(contrib)
    cache = {}

    def cmp(a, b):
        if (a, b) not in cache:
            cache[(a, b)] = ef.compare(contrib[a], contrib[b], (ef.variant_label(a), ef.variant_label(b)))
        return cache[(a, b)]

    def beats(a, b):                         # a je ostře lepší než b pro všechny hodnoty neznámých
        r = cmp(a, b)
        return r["status"] == ef.CERTAIN and r["winner"] == ef.A and r["strict"]

    def score(v):
        c = contrib[v]
        return max(f.known for f in c.scenarios.values()) if c.valid else float("-inf")

    front = []                               # varianty, které žádná jiná prokazatelně nepřekoná
    for v in sorted(variants, key=lambda v: (-score(v), len(v), v)):
        if any(beats(u, v) for u in front):
            continue
        front = [u for u in front if not beats(v, u)] + [v]
        if len(front) > MAX_FRONT:
            return dict(verdict=NO_DECISION, choice=None, contenders=front, best=[], conditions=[],
                        missing=(), text=f"Ekonomicky nelze rozhodnout – příliš mnoho variant, které nelze "
                                          f"navzájem vyloučit (víc než {MAX_FRONT}).")
    unknowns = {k: u for c in contrib.values() for k, u in c.unknowns.items()}
    labels = [ef.variant_label(v) for v in front]

    if len(front) == 1:
        w = front[0]
        if all(beats(w, v) for v in variants if v != w):
            if w == NOTHING:
                return dict(verdict=NONE_PAYS, choice=NOTHING, contenders=front, best=[w], conditions=[],
                            missing=(), text="Ekonomicky se nevyplatí přijmout žádný ze soupeřících požadavků – "
                                             "každá varianta, která se vejde, je prokazatelně ztrátová. Nic se "
                                             "neruší – rozhodne uživatel.")
            return dict(verdict=RECOMMEND, choice=w, contenders=front, best=[w], conditions=[], missing=(),
                        text=f"Varianta {labels[0]} je prokazatelně ekonomicky lepší než všechny ostatní "
                             "varianty včetně nepřijetí ničeho – při jakýchkoli hodnotách chybějících nákladů.")

    pairs = [(a, b) for i, a in enumerate(front) for b in front[i + 1:]]
    results = [cmp(a, b) for a, b in pairs]
    if results and all(r["status"] == ef.CERTAIN and r["winner"] == ef.EQUAL for r in results):
        return dict(verdict=EQUAL, choice=None, contenders=front, best=list(front), conditions=[], missing=(),
                    text="Varianty " + ", ".join(labels) + " mají prokazatelně shodný ekonomický výsledek "
                         "a ostatní jsou horší – rozhodne uživatel.")
    undecided = [r for r in results if r["status"] == ef.UNDECIDABLE]
    if undecided or not results:
        reasons = tuple(dict.fromkeys(x for r in undecided for x in r["reasons"]))
        return dict(verdict=NO_DECISION, choice=None, contenders=front, best=[], conditions=[], missing=reasons,
                    text="Ekonomicky nelze rozhodnout – výsledek závisí na chybějících nebo neplatných údajích: "
                         + ("; ".join(reasons) or "neznámé"))

    # Podmíněně lepší: výklad po scénářích – hranice zlomu (jedna neznámá) nebo podmínky dvojic
    scenarios = list(dict.fromkeys(sc for v in front for sc in contrib[v].scenarios))
    if len(scenarios) > 1:
        scenarios = [sc for sc in scenarios if sc != ef.ONLY]
    conditions = []
    for sc in scenarios:
        prefix = f"Pokud {ef.SCENARIO_LABEL[sc]}: " if sc != ef.ONLY else ""
        lines = {ef.variant_label(v): ef.scenario_formula(contrib[v], sc) for v in front}
        prof = ef.profile(lines)
        if prof is not None:
            conditions.append(dict(scenario=sc, profile=prof, text=prefix + ef.profile_text(prof, unknowns)))
            continue
        leader = max(front, key=lambda v: ef.scenario_formula(contrib[v], sc).known)
        for v in front:
            if v != leader:
                r = ef.compare(ef.restrict(contrib[leader], sc), ef.restrict(contrib[v], sc),
                               (ef.variant_label(leader), ef.variant_label(v)))
                one = r["scenarios"][ef.ONLY] if r["scenarios"] else None
                # jen podmínka bez úvodní věty a koncové tečky (věty se skládají níže)
                cond = r["text"].removeprefix("Výsledek závisí na neznámých nákladech – ").rstrip(".")
                conditions.append(dict(scenario=sc, profile=None, text=prefix + cond, pair=dict(
                    a=leader, b=v, status=r["status"], winner=r.get("winner"), strict=r.get("strict", False),
                    known=one["diff"].known if one else None, terms=one["diff"].terms if one else ())))
    missing = tuple(dict.fromkeys(unknowns[k].label for v in front for f in contrib[v].scenarios.values()
                                  for k, _ in f.terms if k in unknowns))
    shown = " ".join(c["text"] + "." for c in conditions[:MAX_CONDITIONS_TEXT])
    if len(conditions) > MAX_CONDITIONS_TEXT:
        shown += f" … a dalších {len(conditions) - MAX_CONDITIONS_TEXT} podmínek."
    return dict(verdict=CONDITIONAL, choice=None, contenders=front, best=[], conditions=conditions,
                missing=missing, text="Podmíněně – nejlepší varianta závisí na neznámých nákladech: " + shown)


def _economic_verdict(model: dm.DataModel, candidates: list) -> dict:
    """Vzorce všech porovnávaných variant (z econ_formula, nad economics.py) a robustní verdikt."""
    import econ_formula as ef
    reqs = {r.ref: r for o in candidates for r in o}
    try:
        forms = {ref: ef.request_formula(model, r) for ref, r in reqs.items()}
    except ArithmeticError as exc:            # pojistka: příspěvek není lineární – nic se nedoporučí
        return dict(verdict=NO_DECISION, choice=None, contenders=[], best=[], conditions=[], missing=(),
                    formulas={}, text=f"Ekonomicky nelze rozhodnout – {exc}")
    variants = [NOTHING] + [tuple(r.ref for r in o) for o in candidates if o]
    contrib = {v: ef.combine(forms[ref] for ref in v) if v else ef.NOTHING for v in variants}
    out = _decide(contrib)
    out["formulas"] = {v: ef.describe(contrib[v]) for v in set(out["contenders"]) | {NOTHING}}
    return out


# ---------------------------------------------------------------------------
# Hlavní vyhodnocení
# ---------------------------------------------------------------------------

@dataclass
class Plan:
    manual: list = field(default_factory=list)       # konflikty potvrzených závazků k ručnímu řešení
    requests: dict = field(default_factory=dict)     # ref → provozní a ekonomické vyhodnocení
    groups: list = field(default_factory=list)       # soupeřící požadavky a verdikt
    missing_data: list = field(default_factory=list) # chybějící údaje napříč výsledkem
    availability: Optional[pd.DataFrame] = None      # dostupnost po závazcích pro zadané období
    notes: list = field(default_factory=list)


def _job_shift_for_conflict(model, cal, job_ref: str, limit: int = 3) -> list:
    """U potvrzené zakázky v konfliktu: termíny v okně bez penále, kde by se vešla (jen nabídka)."""
    a_job = [a for a in cal.assignments if a.kind == "job" and a.ref == job_ref and a.status == cp.BLOCKING]
    row = model.jobs[model.jobs["job_id"] == job_ref]
    if not a_job or row.empty:
        return []
    others = [a for a in cal.assignments if not (a.kind == "job" and a.ref == job_ref)]
    rest = cp.build_calendar(cal.units, others, cal.open_jobs)
    req = Request(job_ref, "job", NEW, tuple((a.type_id, a.quantity) for a in a_job), a_job[0].start, a_job[0].end,
                  earliest=_int(row.iloc[0]["earliest_start_day"]), deadline=_int(row.iloc[0]["deadline_day"]))
    return _shift_options(rest, req, limit)


def plan(model: dm.DataModel, job_starts: Optional[dict] = None, new_requests=(), period=None,
         include_data_tentative: bool = True) -> Plan:
    """Vyhodnotí dostupnost, konflikty a požadavky. Nic nemění – vrací jen doporučení.

    job_starts … aktuální provozní plán zakázek {job_id: první den} (např. ze záložky 6) –
                 blokuje kapacitu, ale není smluvně neměnný (konflikty jdou k ručnímu řešení).
    new_requests … nové požadavky (Request se stavem „new“).
    period … (od, do) pro přehled dostupnosti.
    """
    cal = cp.from_model(model, job_starts=job_starts)
    out = Plan()
    if cal.issues:
        out.missing_data += [text for _, text in cal.issues]

    # 1) Konflikty potvrzených závazků → ruční řešení (nic se nerozhoduje ani nepřesouvá)
    for c in cp.find_conflicts(cal):
        item = dict(conflict=c, decision="k ručnímu řešení",
                    text=f"{c['kind']}: {', '.join(c['sources'])} (dny {c['days'][0]}–{c['days'][-1]})",
                    alternatives=[])
        for src in c["sources"]:
            if src.startswith(cp.KIND_LABEL["job"] + " "):
                ref = src.split(" ")[2]
                shifts = _job_shift_for_conflict(model, cal, ref)
                item["alternatives"].append(
                    dict(job=ref, shifts=shifts,
                         text=(f"zakázku {ref} lze bez penále posunout na dny "
                               + ", ".join(f"{s}–{e}" for s, e in shifts)) if shifts
                         else f"zakázku {ref} nelze v jejím okně bez penále posunout"))
        out.manual.append(item)
    for o in cp.overlapping_maintenance(cal):
        out.notes.append(f"Překrývající se servisní záznamy kusu {o['unit_id']} (dny {o['days'][0]}–{o['days'][-1]}): "
                         "kus se v servisu počítá jen jednou; zkontrolujte, zda nejde o duplicitní záznam.")
    if cal.open_jobs:
        out.notes.append("Potvrzené zakázky bez termínu v plánu: obsazení nelze určit – "
                         + ", ".join(sorted({j.ref for j in cal.open_jobs})))

    # 2) Požadavky: předběžné z dat + nové; každý zůstává ve výsledku viditelný
    requests = (tentative_requests(model, job_starts) if include_data_tentative else []) + list(new_requests)
    refs = [r.ref for r in requests]
    dups = sorted({x for x in refs if refs.count(x) > 1})
    if dups:
        raise ValueError("Označení požadavků musí být jedinečná (i vůči předběžným rezervacím v datech): "
                         + ", ".join(dups))
    # Nový požadavek nesmí nést označení existující zakázky ani rezervace – ekonomika by jinak
    # převzala údaje původního záznamu (economics hledá zakázku podle job_id).
    data_ids = set(model.jobs["job_id"].astype(str)) | set(model.rental_bookings["booking_id"].astype(str))
    clash = sorted(r.ref for r in new_requests if not r.in_data and r.ref in data_ids)
    if clash:
        raise ValueError("Označení nového požadavku koliduje s existující zakázkou nebo rezervací: "
                         + ", ".join(clash))
    economics_needed = []
    for r in requests:
        entry = dict(request=r, label=r.label, operational=None, problems=[], reasons=[], alternatives=[],
                     economics=None, economic_note="", decision="", note="")
        invalid = _validate(r, cal)
        if invalid:
            entry.update(operational=UNDETERMINED, reasons=invalid,
                         decision="nelze rozhodnout – chybí nebo jsou neplatné údaje požadavku")
            out.missing_data += [f"{r.ref}: {p}" for p in invalid]
        else:
            status, problems, reasons = _check(cal, r, r.start, r.end)
            entry.update(operational=status, problems=problems, reasons=reasons)
            if status == NOT_FEASIBLE:
                shifts = _shift_options(cal, r)
                entry["alternatives"] = [dict(shift=s, text=f"posun bez penále na dny {s[0]}–{s[1]}") for s in shifts]
                entry["decision"] = "nelze přijmout v požadovaném termínu – kapacitu drží závazky"
            elif status == UNDETERMINED:
                entry["decision"] = "nelze rozhodnout – obsazení nelze spolehlivě určit"
                out.missing_data += [f"{r.ref}: {x}" for x in reasons]
            else:
                economics_needed.append(r)
        if r.status == TENTATIVE:
            entry["note"] = "předběžná rezervace zůstává – nic se automaticky neruší ani nepotvrzuje"
        out.requests[r.ref] = entry

    values = {r.ref: contribution(model, r) for r in economics_needed}
    for r in economics_needed:
        out.requests[r.ref]["economics"] = values[r.ref]
        out.requests[r.ref]["economic_note"] = economic_note(values[r.ref])

    # 3) Soupeřící proveditelné požadavky → kombinace, které se vejdou, a ekonomický verdikt
    residual = _residual(cal, economics_needed) if economics_needed else {}
    for group in _groups(economics_needed):
        refs = [r.ref for r in group]
        if _fits(residual, group) and _verify_with_capacity(cal, group):
            for ref in refs:
                out.requests[ref]["decision"] = "lze přijmout (vejde se spolu s ostatními požadavky)"
            if len(group) > 1:
                out.groups.append(dict(requests=refs, verdict=ALL_FIT, options=[tuple(refs)], choice=tuple(refs),
                                       text="Všechny požadavky skupiny se vejdou současně."))
            continue
        if len(group) > MAX_GROUP:
            for ref in refs:
                out.requests[ref]["decision"] = "nelze rozhodnout – příliš mnoho soupeřících požadavků"
            out.groups.append(dict(requests=refs, verdict=NO_DECISION, options=[], choice=None,
                                   text=f"Skupina má {len(group)} požadavků (limit {MAX_GROUP})."))
            continue
        feasible = [s for k in range(len(group), 0, -1) for s in combinations(group, k) if _fits(residual, s)]
        maximal = [s for s in feasible if not any(set(s) < set(t) for t in feasible)]
        verified = [s for s in maximal if _verify_with_capacity(cal, s)]
        # ekonomicky se porovnávají všechny proveditelné kombinace (podmnožiny ověřených variant)
        candidates = [s for s in feasible if any(set(s) <= set(v) for v in verified)]
        maximal = verified
        verdict = _economic_verdict(model, candidates)
        sums = {o: _sum(values[r.ref] for r in o) for o in [NOTHING] + list(candidates)}
        for ref in refs:
            out.requests[ref]["decision"] = ("v konfliktu s jinými požadavky – vejde se jen část; "
                                             + verdict["verdict"])
        if verdict["missing"]:
            out.missing_data += list(verdict["missing"])
        out.groups.append(dict(requests=refs, verdict=verdict["verdict"],
                               options=[tuple(r.ref for r in o) for o in maximal],
                               # všechny ekonomicky porovnávané varianty včetně „nepřijmout nic“
                               candidates=[NOTHING] + [tuple(r.ref for r in o) for o in candidates],
                               choice=verdict["choice"],
                               # robustní výklad (E2.5c krok 3): varianty, které nelze vyloučit,
                               # prokazatelně nejlepší, hranice zlomu a vzorce příspěvků
                               contenders=verdict["contenders"], best=verdict["best"],
                               conditions=[c["text"] for c in verdict["conditions"]],
                               profiles=[dict(scenario=c["scenario"], profile=c["profile"])
                                         for c in verdict["conditions"] if c["profile"] is not None],
                               formulas=verdict["formulas"],
                               # podmínky dvojic strukturovaně (víc neznámých): rozdíl a − b = known + Σ coef × neznámá
                               pair_conditions=[dict(scenario=c["scenario"], **c["pair"])
                                                for c in verdict["conditions"] if c.get("pair")],
                               # intervaly (dolní/horní mez) jen pro informaci, verdikt z nich nevzniká
                               values={tuple(r.ref for r in o): v for o, v in sums.items()
                                       if o in maximal or o == NOTHING
                                       or tuple(r.ref for r in o) == verdict["choice"]},
                               text=verdict["text"]))

    for ref, v in values.items():
        if v.missing:
            out.missing_data += list(v.missing)
    out.missing_data = list(dict.fromkeys(out.missing_data))
    if period:
        out.availability = cp.daily_capacity(cal, period[0], period[1])
    return out


def explain(p: Plan) -> list:
    """Srozumitelné vysvětlení výsledku (řádky textu)."""
    lines = []
    for m in p.manual:
        lines.append(f"K ručnímu řešení – {m['text']}.")
        lines += [f"  Posun zakázky: {a['text']}." for a in m["alternatives"]]
    for e in p.requests.values():
        lines.append(f"{e['label']}: provozně {e['operational']} – {e['decision']}.")
        lines += [f"  Důvod: {x}" for x in e["reasons"]]
        lines += [f"  Blokuje: {t} den {d} ({', '.join(c) or 'kus není k dispozici'})"
                  for t, d, _, _, c in e["problems"][:5]]
        lines += [f"  Alternativa: {a['text']}." for a in e["alternatives"]]
        if e["note"]:
            lines.append(f"  Pozn.: {e['note']}.")
    for g in p.groups:
        lines.append(f"Skupina {', '.join(g['requests'])}: {g['verdict']} – {g['text']}")
        if g["verdict"] != ALL_FIT and g["options"]:
            lines.append("  Provozně proveditelné varianty (vejdou se do kapacity): "
                         + " | ".join(" + ".join(o) for o in g["options"]))
        if g.get("choice") and g["verdict"] == RECOMMEND:
            lines.append(f"  Doporučeno (jen doporučení, nic se nemění): {', '.join(g['choice'])}")
        if g["verdict"] == CONDITIONAL:
            lines += [f"  Podmínka: {c}" for c in g.get("conditions", [])]
        if g["verdict"] == NONE_PAYS:
            lines.append("  Ekonomicky nejlépe vychází nepřijmout žádný z nich (příspěvek 0 Kč); "
                         "předběžné rezervace se tím neruší.")
    lines += [f"Poznámka: {n}" for n in p.notes]
    return lines
