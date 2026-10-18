"""Ekonomický modul Decision Engine (etapa E2.1).

Společný výpočtový základ pro pronájem strojů i vlastní zakázky nad datovou vrstvou
`datamodel.py`. Modul zatím žádná záložka aplikace nepoužívá – v E2.1 se jen připravuje
a testuje; dosavadní výpočty v app.py se nemění.

EKONOMICKÁ LOGIKA (srozumitelně)
--------------------------------
Každý den může stroj sloužit jen jednomu účelu: pronájem, vlastní zakázka, údržba, nebo stojí.

1. Náklady flotily (fixní): platí se každý měsíc za každý vlastněný kus, i když stojí.
   - vlastnické: splátka, odpis, pojištění  (ownership_cost_month)
   - provozní fixní: garáž, pevný servis    (fixed_operating_cost_month, service_cost_month_mean)
   - nerozdělené fixní: dočasně, dokud majitel nedodá rozdělení (fixed_cost_month_unsplit)
   Počítají se JEN JEDNOU za celou flotilu a nerozpočítávají se na pronájmy ani zakázky.
2. Příspěvek z využití (variabilní):
   - pronájem: tržba z pronájmu − variabilní náklad pronájmu za dny pronájmu
   - zakázka:  tržba − přímé náklady − variabilní náklady strojů za dny na zakázce − penále
               (náklady strojů se odečtou jen tehdy, když nejsou v přímých nákladech zakázky)
3. Celkový výsledek = Σ příspěvků pronájmů + Σ příspěvků zakázek − fixní náklady flotily − režie.
   Variabilní náklady jsou jen uvnitř příspěvků, fixní jen v nákladech flotily → nic se
   nezapočte dvakrát.
4. Ušlý zisk: když zakázka zabere stroj ve dny, kdy by se pronajal, přijde firma o příspěvek
   z pronájmu za tyto dny. Modul dává podklad (příspěvek jednoho dne pronájmu); kolik dní by se
   skutečně pronajalo, určí až plánovač (E2.5).

STAVY ÚDAJŮ
-----------
- skutečná nula: v datech je "0" → počítá se jako 0 (výsledek je přesný),
- chybějící údaj: prázdná buňka → výsledek je NEÚPLNÝ a uvádí, co chybí,
- neplatný údaj: text, záporné číslo… → výsledek NELZE SPOČÍTAT,
- fiktivní data: celý datový soubor je označený jako fiktivní (settings: data_origin) →
  každý výsledek nese příznak „fiktivní“.
Neúplný výsledek obsahuje součet známých položek, ale nikdy se netváří jako přesný.
"""

from dataclasses import dataclass, field
from typing import Optional

import datamodel as dm

FIXED_STATUSES = ("owned", "leased")   # kusy, za které se platí fixní náklady flotily


# ---------------------------------------------------------------------------
# Výsledek výpočtu se stavem
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class Result:
    """Částka v Kč se stavem úplnosti.

    value     – součet známých položek (None, pokud je některý údaj neplatný)
    missing   – popisy chybějících údajů (prázdné = výsledek je úplný)
    invalid   – popisy neplatných údajů (neprázdné = výsledek nelze spočítat)
    fictitious – výsledek vychází z fiktivních dat
    parts     – rozpad na položky (název → částka nebo None)
    """
    value: Optional[float]
    missing: tuple = ()
    invalid: tuple = ()
    fictitious: bool = False
    parts: dict = field(default_factory=dict, compare=False)

    @property
    def complete(self) -> bool:
        return not self.missing and not self.invalid

    @property
    def status(self) -> str:
        if self.invalid:
            return "nelze spočítat"
        return "úplný" if not self.missing else "neúplný"

    def describe(self) -> str:
        """Srozumitelný popis stavu pro budoucí zobrazení v aplikaci."""
        if self.invalid:
            text = "Nelze spočítat – neplatné údaje: " + "; ".join(self.invalid)
        elif self.missing:
            text = "Neúplný výpočet – nezadáno: " + "; ".join(self.missing)
        else:
            text = "Úplný výpočet"
        return text + (" (fiktivní data)" if self.fictitious else "")

    def __add__(self, other: "Result") -> "Result":
        return _combine(self, other, 1)

    def __sub__(self, other: "Result") -> "Result":
        return _combine(self, other, -1)

    def scaled(self, factor: float, label: str = "") -> "Result":
        value = None if self.value is None else self.value * factor
        return Result(value, self.missing, self.invalid, self.fictitious,
                      {label: value} if label else dict(self.parts))


def _unique(*seqs) -> tuple:
    out = []
    for s in seqs:
        for x in s:
            if x not in out:
                out.append(x)
    return tuple(out)


def _combine(a: Result, b: Result, sign: int) -> Result:
    value = None if a.value is None or b.value is None else a.value + sign * b.value
    parts = dict(a.parts)
    for k, v in b.parts.items():
        parts[k] = None if v is None else sign * v
    return Result(value, _unique(a.missing, b.missing), _unique(a.invalid, b.invalid),
                  a.fictitious or b.fictitious, parts)


def total(results, label: str = "") -> Result:
    """Součet více výsledků (prázdný seznam = přesná nula)."""
    out = Result(0.0)
    for r in results:
        out = out + r
    return out if not label else Result(out.value, out.missing, out.invalid, out.fictitious,
                                         {label: out.value})


# ---------------------------------------------------------------------------
# Čtení jednotlivých údajů
# ---------------------------------------------------------------------------

def is_fictitious(model: dm.DataModel) -> bool:
    s = model.settings.set_index("key")["value"] if len(model.settings) else {}
    return "fiktivní" in str(s.get("data_origin", "")).lower() or "fiktivni" in str(s.get("data_origin", "")).lower()


def _amount(raw, label: str, fictitious: bool) -> Result:
    """Převede buňku na Result: prázdná = nezadáno, nečíselná/záporná = neplatná."""
    text = "" if raw is None else str(raw).strip()
    if text == "" or text.lower() == "nan":
        return Result(0.0, missing=(label,), fictitious=fictitious, parts={label: None})
    try:
        value = float(text.replace(",", "."))
    except ValueError:
        return Result(None, invalid=(f"{label} = „{text}“",), fictitious=fictitious, parts={label: None})
    if value < 0:
        return Result(None, invalid=(f"{label} je záporné ({text})",), fictitious=fictitious,
                      parts={label: None})
    return Result(value, fictitious=fictitious, parts={label: value})


def _type_row(model: dm.DataModel, type_id: str):
    rows = model.machine_types[model.machine_types["type_id"] == type_id]
    if rows.empty:
        raise KeyError(f"Neznámý kód stroje: {type_id}")
    return rows.iloc[0]


def _job_row(model: dm.DataModel, job_id: str):
    rows = model.jobs[model.jobs["job_id"] == job_id]
    if rows.empty:
        raise KeyError(f"Neznámá zakázka: {job_id}")
    return rows.iloc[0]


def setting(model: dm.DataModel, key: str) -> str:
    s = model.settings
    hit = s[s["key"] == key]
    return "" if hit.empty else str(hit.iloc[0]["value"]).strip()


# ---------------------------------------------------------------------------
# 1. Náklady flotily
# ---------------------------------------------------------------------------

def unit_fixed_costs(model: dm.DataModel, type_id: str) -> dict:
    """Měsíční fixní náklady JEDNOHO kusu podle skupin.

    Vrací {"vlastnické": Result, "provozní fixní": Result, "nerozdělené fixní": Result}.
    Je-li zadán nerozdělený fixní náklad, vlastnické a provozní fixní (mimo servis) se nečtou –
    částka je celá v „nerozdělené“ (jinak by se započetla dvakrát). Současné zadání obou
    variant je neplatné.
    """
    fic = is_fictitious(model)
    row = _type_row(model, type_id)
    name = row["name"]
    unsplit_raw = str(row["fixed_cost_month_unsplit"]).strip()
    split_given = any(str(row[c]).strip() for c in ("ownership_cost_month", "fixed_operating_cost_month"))
    service = _amount(row["service_cost_month_mean"], f"{name}: servis/měs.", fic)
    if unsplit_raw and split_given:
        bad = Result(None, invalid=(f"{name}: fixní náklady zadané nerozdělené i rozdělené zároveň "
                                    "(dvojí započtení)",), fictitious=fic)
        return {"vlastnické": bad, "provozní fixní": bad, "nerozdělené fixní": bad}
    if unsplit_raw:
        # Celková částka je známá, jen není rozdělená: vlastnické a provozní fixní (mimo servis)
        # jsou obsažené v „nerozdělených“ – v rozpadu je proto označíme jako nezadané.
        split_note = f"{name}: rozdělení fixních nákladů na vlastnické a provozní"
        return {
            "vlastnické": Result(0.0, missing=(split_note,), fictitious=fic,
                                 parts={f"{name}: vlastnické (součást nerozdělených)": None}),
            "provozní fixní": service + Result(0.0, missing=(split_note,), fictitious=fic),
            "nerozdělené fixní": _amount(unsplit_raw, f"{name}: nerozdělené fixní/měs.", fic),
        }
    return {
        "vlastnické": _amount(row["ownership_cost_month"], f"{name}: vlastnické náklady/měs.", fic),
        "provozní fixní": _amount(row["fixed_operating_cost_month"], f"{name}: provozní fixní/měs.", fic)
        + service,
        "nerozdělené fixní": Result(0.0, fictitious=fic),
    }


def unit_fixed_total(model: dm.DataModel, type_id: str) -> Result:
    """Měsíční fixní náklad JEDNOHO kusu celkem (všechny skupiny, každá položka jednou).

    Při nerozdělených nákladech je celek známý a přesný, i když rozpad do skupin není.
    """
    fic = is_fictitious(model)
    row = _type_row(model, type_id)
    name = row["name"]
    groups = unit_fixed_costs(model, type_id)
    if any(g.invalid for g in groups.values()):
        return groups["vlastnické"]
    service = _amount(row["service_cost_month_mean"], f"{name}: servis/měs.", fic)
    if str(row["fixed_cost_month_unsplit"]).strip():
        return _amount(row["fixed_cost_month_unsplit"], f"{name}: nerozdělené fixní/měs.", fic) + service
    return total(groups.values())


def units_in_fleet(model: dm.DataModel, type_id: str, statuses=FIXED_STATUSES) -> int:
    f = model.fleet_units
    return int(((f["type_id"] == type_id) & f["status"].isin(statuses)).sum())


def fleet_fixed_costs(model: dm.DataModel, months: float = 1.0, statuses=FIXED_STATUSES) -> dict:
    """Fixní náklady celé flotily za `months` měsíců podle skupin a celkem.

    Vrací {"vlastnické", "provozní fixní", "nerozdělené fixní", "celkem", "podle typu"}.
    „celkem“ se počítá z celkového nákladu každého kusu (každá položka jednou), takže je přesné
    i tehdy, když rozpad do skupin kvůli nerozděleným nákladům přesný není.
    Počítá jen kusy ve stavu `statuses` (výchozí: vlastněné a pronajaté na leasing).
    """
    groups = {"vlastnické": [], "provozní fixní": [], "nerozdělené fixní": []}
    by_type = {}
    for type_id in model.machine_types["type_id"]:
        n = units_in_fleet(model, type_id, statuses)
        if n == 0:
            continue
        costs = unit_fixed_costs(model, type_id)
        for g in groups:
            groups[g].append(costs[g].scaled(n * months))
        by_type[type_id] = unit_fixed_total(model, type_id).scaled(n * months)
    out = {g: total(v) for g, v in groups.items()}
    out["celkem"] = total(by_type.values())
    out["podle typu"] = by_type
    return out


def variable_costs(model: dm.DataModel, usage: dict) -> dict:
    """Variabilní náklady podle využití: usage = {type_id: {"rental_days": x, "job_days": y}}.

    Vrací {"pronájem": Result, "zakázky": Result, "celkem": Result}. Nula dní znamená nulový
    náklad i při nezadané sazbě (nic se neprovedlo, nic nechybí).
    """
    rent, jobs = [], []
    for type_id, u in usage.items():
        rent.append(rental_variable_cost(model, type_id, u.get("rental_days", 0)))
        jobs.append(job_machine_variable_cost(model, type_id, u.get("job_days", 0)))
    out = {"pronájem": total(rent), "zakázky": total(jobs)}
    out["celkem"] = out["pronájem"] + out["zakázky"]
    return out


def rental_variable_cost(model: dm.DataModel, type_id: str, days: float) -> Result:
    row = _type_row(model, type_id)
    if days == 0:
        return Result(0.0, fictitious=is_fictitious(model))
    rate = _amount(row["variable_cost_rental_day"], f"{row['name']}: variabilní náklad pronájmu/den",
                   is_fictitious(model))
    return rate.scaled(days, f"{row['name']}: variabilní náklad pronájmu")


def job_machine_variable_cost(model: dm.DataModel, type_id: str, days: float) -> Result:
    row = _type_row(model, type_id)
    if days == 0:
        return Result(0.0, fictitious=is_fictitious(model))
    rate = _amount(row["variable_cost_job_day"], f"{row['name']}: variabilní náklad na zakázce/den",
                   is_fictitious(model))
    return rate.scaled(days, f"{row['name']}: variabilní náklad na zakázce")


# ---------------------------------------------------------------------------
# 2. Příspěvek z pronájmu
# ---------------------------------------------------------------------------

def rental_contribution(model: dm.DataModel, type_id: str, days: float,
                        price_day: Optional[float] = None) -> Result:
    """Příspěvek z pronájmu = tržba (cena/den × dny) − variabilní náklad pronájmu × dny.

    Fixní náklady stroje se zde NEODEČÍTAJÍ – patří do nákladů flotily.
    `price_day` umožní jinou cenu (např. z rezervace); jinak ceník typu stroje.
    """
    if days < 0:
        raise ValueError("Počet dní pronájmu nemůže být záporný.")
    fic = is_fictitious(model)
    row = _type_row(model, type_id)
    price = (Result(float(price_day), fictitious=fic, parts={f"{row['name']}: cena/den": price_day})
             if price_day is not None
             else _amount(row["rental_price_day"], f"{row['name']}: cena pronájmu/den", fic))
    revenue = price.scaled(days, "tržba z pronájmu")
    return revenue - rental_variable_cost(model, type_id, days).scaled(1, "variabilní náklady pronájmu")


def rental_day_contribution(model: dm.DataModel, type_id: str) -> Result:
    """Příspěvek jednoho dne pronájmu (podklad pro ušlý zisk)."""
    return rental_contribution(model, type_id, 1)


# ---------------------------------------------------------------------------
# 3. Příspěvek z vlastní zakázky
# ---------------------------------------------------------------------------

def direct_cost_includes_machines(model: dm.DataModel) -> Optional[bool]:
    """Nastavení `direct_cost_includes_machines` (ano/ne); chybí-li, vrací None (nezadáno)."""
    v = setting(model, "direct_cost_includes_machines").lower()
    if v in ("ano", "true", "1", "yes"):
        return True
    if v in ("ne", "false", "0", "no"):
        return False
    return None


def job_machine_days(model: dm.DataModel, job_id: str) -> dict:
    """Strojodny zakázky podle typu stroje: délka zakázky × požadovaný počet kusů."""
    job = _job_row(model, job_id)
    dur = float(job["duration_days"])
    req = model.job_requirements
    req = req[(req["job_id"] == job_id) & (req["resource_kind"] == "machine")]
    out = {}
    for r in req.itertuples():
        out[r.resource_id] = out.get(r.resource_id, 0.0) + dur * float(r.quantity)
    return out


def job_contribution(model: dm.DataModel, job_id: str, days_late: float = 0.0,
                     includes_machines: Optional[bool] = None) -> Result:
    """Příspěvek zakázky = tržba − přímé náklady − variabilní náklady strojů − penále.

    Variabilní náklady strojů na zakázce se odečtou jen tehdy, když NEJSOU obsažené
    v přímých nákladech (`includes_machines` nebo nastavení direct_cost_includes_machines).
    Není-li známo, zda obsažené jsou, výsledek je neúplný. Fixní náklady strojů se na
    zakázku nerozpočítávají (patří do nákladů flotily).
    """
    if days_late < 0:
        raise ValueError("Zpoždění nemůže být záporné.")
    fic = is_fictitious(model)
    job = _job_row(model, job_id)
    name = job["name"]
    revenue = _amount(job["revenue"], f"{name}: tržba", fic).scaled(1, "tržba ze zakázky")
    direct = _amount(job["direct_cost"], f"{name}: přímé náklady", fic).scaled(1, "přímé náklady")
    penalty = _amount(job["penalty_per_day"], f"{name}: penále/den", fic).scaled(days_late, "penále")
    result = revenue - direct - penalty

    incl = direct_cost_includes_machines(model) if includes_machines is None else includes_machines
    if incl is True:
        return result
    machine = total(job_machine_variable_cost(model, t, d) for t, d in job_machine_days(model, job_id).items())
    machine = machine.scaled(1, "variabilní náklady strojů")
    if incl is None:
        note = "zda přímé náklady zakázek obsahují vlastní stroje (direct_cost_includes_machines)"
        return Result(result.value, _unique(result.missing, (note,)), result.invalid, fic,
                      dict(result.parts, **{"variabilní náklady strojů": None}))
    return result - machine


# ---------------------------------------------------------------------------
# 4. Podklady pro ušlý zisk
# ---------------------------------------------------------------------------

def opportunity_cost_basis(model: dm.DataModel) -> dict:
    """Pro každý typ stroje: příspěvek dne pronájmu a očekávaná poptávka po pronájmu.

    Ušlý zisk vznikne jen tehdy, když by stroj ve dny zakázky jinak opravdu byl pronajatý
    (poptávka převyšuje volnou kapacitu). Kolik takových dní je, určí plánovač (E2.5).
    """
    demand = model.rental_demand
    out = {}
    for r in model.machine_types.itertuples():
        d = demand[demand["type_id"] == r.type_id]
        out[r.type_id] = {
            "příspěvek dne pronájmu": rental_day_contribution(model, r.type_id),
            "očekávaná poptávka (dny/období)": _amount(d.iloc[0]["expected_days"] if len(d) else "",
                                                      f"{r.name}: poptávka po pronájmu",
                                                      is_fictitious(model)),
            "období poptávky": d.iloc[0]["period"] if len(d) else "",
        }
    return out


def opportunity_cost(model: dm.DataModel, displaced_rental_days: dict) -> Result:
    """Ušlý zisk = Σ (dny pronájmu, o které firma přišla) × příspěvek dne pronájmu."""
    return total(rental_day_contribution(model, t).scaled(days, f"ušlý zisk {t}")
                 for t, days in displaced_rental_days.items() if days)


# ---------------------------------------------------------------------------
# 5. Celkový ekonomický výsledek
# ---------------------------------------------------------------------------

def overhead(model: dm.DataModel, months: float = 1.0) -> Result:
    return _amount(setting(model, "overhead_cost_month"), "společná režie/měs.",
                   is_fictitious(model)).scaled(months, "režie")


def total_result(model: dm.DataModel, rentals=(), jobs=(), months: float = 1.0) -> dict:
    """Celkový výsledek za `months` měsíců.

    rentals = [(type_id, dny[, cena/den])], jobs = [(job_id, dny zpoždění)].
    Výsledek = Σ příspěvků pronájmů + Σ příspěvků zakázek − fixní náklady flotily − režie.
    Variabilní náklady jsou jen v příspěvcích, fixní jen v nákladech flotily.
    """
    rental = total((rental_contribution(model, *r) for r in rentals), "příspěvky z pronájmů")
    job = total((job_contribution(model, j, late) for j, late in jobs), "příspěvky ze zakázek")
    fleet = fleet_fixed_costs(model, months)["celkem"].scaled(1, "fixní náklady flotily")
    ovh = overhead(model, months)
    result = rental + job - fleet - ovh
    return {
        "příspěvky z pronájmů": rental,
        "příspěvky ze zakázek": job,
        "fixní náklady flotily": fleet,
        "režie": ovh,
        "výsledek": result,
    }
