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
6. Provozní proveditelnost a ekonomická výhodnost jsou oddělené. Ekonomické doporučení vznikne
   jen tehdy, když je jedna kombinace lepší i v nejhorším případě (při chybějících nákladech je
   známá částka jen horní mez). Jinak „nelze rozhodnout“ s výčtem chybějících údajů.
   Mezi variantami je vždy i „nepřijmout nic“ (přírůstkový příspěvek 0 Kč): jsou-li všechny varianty
   jistě ztrátové, výsledek je „žádný se nevyplatí“, nikdy doporučení nejmenší ztráty.
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
NOTHING = ()                           # varianta „nepřijmout nic“ – přírůstkový příspěvek přesně 0 Kč
MAX_GROUP = 12                         # nejvíc požadavků v jedné skupině (2^12 kombinací)


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


def _economic_verdict(options: list, values: dict) -> dict:
    """options = všechny proveditelné kombinace (i menší – ztrátový požadavek je lepší vynechat).
    Vždy se přidá i varianta „nepřijmout nic“ s přírůstkovým příspěvkem přesně 0 Kč – když jsou všechny
    ostatní varianty ztrátové, nesmí se doporučit „nejmenší ztráta“.
    Doporučí jen kombinaci, která je lepší i v nejhorším případě než každá jiná (včetně „nic“);
    „nic“ vyhraje jen tehdy, když je každá jiná varianta ztrátová i v nejlepším případě (horní mez < 0)."""
    options = [NOTHING] + [o for o in options if o != NOTHING]
    sums = {o: _sum(values[r.ref] for r in o) for o in options}
    missing = tuple(dict.fromkeys(m for v in sums.values() for m in v.missing))
    known = [o for o in options if sums[o].lo is not None]
    if known:
        best = max(known, key=lambda o: sums[o].lo)
        others = [o for o in options if o is not best]
        if all(sums[o].hi is not None and sums[best].lo > sums[o].hi for o in others):
            if best == NOTHING:
                return dict(verdict=NONE_PAYS, choice=NOTHING, values=sums, missing=missing,
                            text="Ekonomicky se nevyplatí přijmout žádný ze soupeřících požadavků – každá "
                                 "varianta, která se vejde, je ztrátová i v nejlepším případě. Nic se neruší "
                                 "– rozhodne uživatel.")
            return dict(verdict=RECOMMEND, choice=best, values=sums, missing=missing,
                        text="Kombinace je ekonomicky lepší i v nejhorším případě.")
    if all(v.exact for v in sums.values()) and len({v.lo for v in sums.values()}) == 1:
        return dict(verdict=EQUAL, choice=None, values=sums, missing=missing,
                    text="Varianty mají shodný ekonomický přínos – rozhodne uživatel.")
    if all(v.exact for v in sums.values()):
        best = max(v.lo for v in sums.values())
        tied = [o for o in options if sums[o].lo == best]
        nothing = " (včetně varianty nepřijmout nic)" if NOTHING in tied else ""
        return dict(verdict=EQUAL, choice=None, values=sums, missing=missing,
                    text=f"Nejvyšší přínos má shodně {len(tied)} variant{nothing} – rozhodne uživatel.")
    return dict(verdict=NO_DECISION, choice=None, values=sums, missing=missing,
                text="Ekonomicky nelze rozhodnout – chybí údaje: " + ("; ".join(missing) or "neznámé"))


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
        verdict = _economic_verdict(candidates, values)
        for ref in refs:
            out.requests[ref]["decision"] = ("v konfliktu s jinými požadavky – vejde se jen část; "
                                             + verdict["verdict"])
        if verdict["missing"]:
            out.missing_data += list(verdict["missing"])
        out.groups.append(dict(requests=refs, verdict=verdict["verdict"],
                               options=[tuple(r.ref for r in o) for o in maximal],
                               # všechny ekonomicky porovnávané varianty včetně „nepřijmout nic“
                               candidates=[NOTHING] + [tuple(r.ref for r in o) for o in candidates],
                               choice=(tuple(r.ref for r in verdict["choice"])
                                       if verdict["choice"] is not None else None),
                               values={tuple(r.ref for r in o): v for o, v in verdict["values"].items()
                                       if o in maximal or o == verdict["choice"] or o == NOTHING},
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
        if g["verdict"] == NONE_PAYS:
            lines.append("  Ekonomicky nejlépe vychází nepřijmout žádný z nich (příspěvek 0 Kč); "
                         "předběžné rezervace se tím neruší.")
    lines += [f"Poznámka: {n}" for n in p.notes]
    return lines
