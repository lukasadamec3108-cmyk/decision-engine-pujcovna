"""Zobrazení kapacitního kalendáře strojů (etapa E2.3B) – záložka „Kapacita strojů“.

Modul NIC nepočítá sám: všechny počty kusů, obsazenost, konflikty i vyhodnocení požadavku
přebírá z capacity.py. Tady se jen převádí na tabulky a HTML mřížku pro Streamlit.

Dny jsou čísla dní plánu (den 1 = první den plánu, stejně jako v záložce 6). Je-li zadané
datum prvního dne, zobrazí se u dní i kalendářní datum – výpočet se tím nemění.
"""

import html
from datetime import date, timedelta
from typing import Optional

import pandas as pd

import capacity as cp

# Barvy: pronájem a zakázka navazují na barvy grafů aplikace (modrá / oranžová).
COLORS = {
    "free": "#2F8F5B",        # volný kus
    "rental": "#1F5FA8",      # pronájem
    "job": "#D97706",         # vlastní zakázka
    "maintenance": "#6B7280",  # servis / údržba
    "over": "#B42318",        # nad kapacitu / konflikt
}
LABELS = {
    "free": "volný kus",
    "rental": "pronájem",
    "job": "vlastní zakázka",
    "maintenance": "servis / údržba",
    "over": "nad kapacitu (konflikt)",
    "unknown": "nelze určit",
    "tentative": "předběžná rezervace",
}
STATUS_TONE = {cp.FEASIBLE: "good", cp.INFEASIBLE: "bad", cp.UNKNOWN: "warn"}
WEEKDAYS = ["Po", "Út", "St", "Čt", "Pá", "So", "Ne"]


# ---------------------------------------------------------------------------
# Vstupy
# ---------------------------------------------------------------------------

def job_starts_from_plan(jobs: pd.DataFrame, plan: pd.DataFrame) -> dict:
    """Začátky zakázek z plánu záložky 6 → {job_id: první den}.

    `plan` je výstup scheduling.evaluate (index = index zakázky v `jobs`). Nenaplánované zakázky
    se nevracejí – capacity.py je pak bere jako potvrzené zakázky bez termínu („nelze určit“).
    Zpožděné zakázky se berou celé, i dny po termínu.
    """
    return {str(jobs.loc[i, "job_id"]): int(r["start_day"])
            for i, r in plan.iterrows() if bool(r["scheduled"])}


def job_info(model) -> dict:
    """{job_id: (název, termín dokončení)} pro popisky."""
    return {str(r.job_id): (str(r.name), int(float(r.deadline_day))) for r in model.jobs.itertuples()}


def type_names(model) -> dict:
    return {str(r.type_id): str(r.name) for r in model.machine_types.itertuples()}


def build(model, job_starts: dict) -> cp.Calendar:
    """Kalendář z datového modelu a plánu zakázek – jen přes capacity.from_model."""
    return cp.from_model(model, job_starts=job_starts)


DEFAULT_PERIOD_DAYS = 42      # výchozí zobrazené období nejvýš 6 týdnů (delší lze zvolit posuvníkem)


def period_bounds(cal: cp.Calendar, minimum: int = 14, extra: int = 5) -> tuple:
    """(výchozí konec období, nejzazší volitelný den).

    Posuvník pokryje všechna známá využití; výchozí zobrazení je kvůli rychlosti a čitelnosti
    nejvýš DEFAULT_PERIOD_DAYS dní.
    """
    ends = [a.end for a in cal.assignments] + [j.latest for j in cal.open_jobs if j.latest is not None]
    last = max(ends + [minimum])
    return min(last, DEFAULT_PERIOD_DAYS), max(last + extra, 30)


def plural(n: int, one: str, few: str, many: str) -> str:
    """České tvary: 1 pronájem, 2–4 pronájmy, 0 a 5+ pronájmů."""
    return f"{n} " + (one if n == 1 else few if 2 <= n <= 4 else many)


def data_info(model, data_version: str = "") -> str:
    """Nenápadný řádek o načtených datech – pro ověření, že aplikace používá aktuální data."""
    b = model.rental_bookings
    confirmed = int((b["status"] == cp.BLOCKING).sum())
    text = (f"Načtená data: {plural(len(b), 'pronájem', 'pronájmy', 'pronájmů')} "
            f"({plural(confirmed, 'potvrzený', 'potvrzené', 'potvrzených')}, "
            f"{plural(len(b) - confirmed, 'předběžný', 'předběžné', 'předběžných')}) · "
            f"{plural(len(model.maintenance), 'servisní odstávka', 'servisní odstávky', 'servisních odstávek')}")
    return text + (f" · verze dat {data_version[:8]}" if data_version else "")


# ---------------------------------------------------------------------------
# Popisky dnů a zdrojů
# ---------------------------------------------------------------------------

def to_date(day: int, plan_start: Optional[date]) -> Optional[date]:
    return plan_start + timedelta(int(day) - 1) if plan_start else None


def day_label(day: int, plan_start: Optional[date] = None, short: bool = False) -> str:
    """„den 5“, nebo s datem „den 5 · Čt 5. 11. 2026“ (krátce: „Čt 5. 11.“)."""
    d = to_date(day, plan_start)
    if d is None:
        return f"den {day}"
    text = f"{WEEKDAYS[d.weekday()]} {d.day}. {d.month}."
    return text if short else f"den {day} · {text} {d.year}"


def days_text(day_list, plan_start: Optional[date] = None, limit: int = 8) -> str:
    """Výčet dní pro souhrn: „11, 15, 16“, nebo s datem „St 30. 12., Ne 3. 1.“."""
    shown = [day_label(d, plan_start, short=True) if plan_start else str(d) for d in day_list[:limit]]
    return ", ".join(shown) + (" …" if len(day_list) > limit else "")


def job_title(ref: str, jobs: dict) -> str:
    return f"{ref} – {jobs[ref][0]}" if ref in jobs else ref


def describe(label: str, jobs: dict) -> str:
    """Popisek z capacity.py doplněný o název zakázky (např. „vlastní zakázka J005 – Výkop …“)."""
    prefix = cp.KIND_LABEL["job"] + " "
    if label.startswith(prefix):
        ref = label[len(prefix):].split(" ")[0]
        if ref in jobs:
            return f"{label} – {jobs[ref][0]}"
    return label


# ---------------------------------------------------------------------------
# Mřížka
# ---------------------------------------------------------------------------

def grid(cal: cp.Calendar, start: int, end: int, type_ids) -> pd.DataFrame:
    """Denní kapacita z capacity.daily_capacity + počet obsazených kusů (used)."""
    df = cp.daily_capacity(cal, start, end, list(type_ids))
    df["used"] = df["maintenance"] + df["rental"] + df["job"]
    return df


def segments(row) -> list:
    """Jednotlivé kusy v buňce (pořadí: servis, pronájem, zakázka, volné, nad kapacitu).

    Jen převod počtů z capacity.py na pruh; při stavu „nelze určit“ se pruh nekreslí.
    """
    out = (["maintenance"] * int(row["maintenance"]) + ["rental"] * int(row["rental"])
           + ["job"] * int(row["job"]))
    units = int(row["units"])
    if len(out) > units:
        return out[:units] + ["over"] * (len(out) - units)
    return out + ["free"] * (units - len(out))


def cell_title(row, name: str, plan_start) -> str:
    parts = [f"{name}, {day_label(int(row['day']), plan_start)}", f"kusů k dispozici: {int(row['units'])}"]
    for key in ("rental", "job", "maintenance"):
        if row[key]:
            parts.append(f"{LABELS[key]}: {int(row[key])}")
    if row["tentative"]:
        parts.append(f"předběžně: {int(row['tentative'])}")
    if row["open_job_risk"]:
        parts.append(f"může obsadit zakázka bez termínu: {int(row['open_job_risk'])}")
    parts.append(f"stav: {row['status']}")
    return " · ".join(parts)


def cell_html(row, name: str, plan_start) -> str:
    status = row["status"]
    title = html.escape(cell_title(row, name, plan_start), quote=True)
    mark = '<div class="cc-tent"></div>' if row["tentative"] else ""
    if status == cp.UNKNOWN:
        return (f'<td class="cc unknown" title="{title}"><div class="cc-num">?</div>'
                f'<div class="cc-bar"><span class="seg unknown"></span></div>{mark}</td>')
    free = int(row["free"])
    if status == cp.OVER:
        num = f'<div class="cc-num over">+{-free}</div>'
    else:
        num = f'<div class="cc-num">{free}<small>/{int(row["units"])}</small></div>'
    segs = "".join(f'<span class="seg {s}"></span>' for s in segments(row)) or '<span class="seg none"></span>'
    cls = {cp.OVER: "over", cp.FULL: "full", cp.FREE: "free"}.get(status, "")
    return f'<td class="cc {cls}" title="{title}">{num}<div class="cc-bar">{segs}</div>{mark}</td>'


CSS = """
<style>
.cc-wrap{overflow-x:auto;-webkit-overflow-scrolling:touch;border:1px solid #E3E6EA;border-radius:10px;margin:.4rem 0 .6rem 0;background:#fff}
.cc-table{border-collapse:separate;border-spacing:0;font-size:.78rem;min-width:100%}
.cc-table th,.cc-table td{border-bottom:1px solid #EEF0F3;border-right:1px solid #EEF0F3;padding:4px 2px;text-align:center;vertical-align:middle}
.cc-table thead th{background:#F6F7F9;color:#4B5563;font-weight:600;position:sticky;top:0;white-space:nowrap;min-width:34px}
.cc-table thead th small{display:block;font-weight:400;color:#4B5563;font-size:.68rem}
.cc-table .cc-type{position:sticky;left:0;z-index:2;background:#fff;text-align:left;min-width:120px;max-width:150px;padding:4px 8px;font-weight:600;color:#1F2933;white-space:normal}
.cc-table thead .cc-type{z-index:3;background:#F6F7F9}
.cc-type small{display:block;font-weight:400;color:#4B5563}
.cc{position:relative;min-width:34px}
.cc-num{font-weight:700;color:#1F2933;font-variant-numeric:tabular-nums;line-height:1.1}
.cc-num small{font-weight:400;color:#4B5563}
.cc-num.over{color:#B42318}
.cc-bar{display:flex;gap:2px;height:8px;margin-top:3px}
.seg{flex:1;border-radius:2px;min-width:6px}
.seg.free{background:#2F8F5B}.seg.rental{background:#1F5FA8}.seg.job{background:#D97706}
.seg.maintenance{background:#6B7280}.seg.over{background:#B42318}.seg.none{background:#E3E6EA}
.seg.unknown,.cc-legend .unknown{background:repeating-linear-gradient(45deg,#9CA3AF 0 3px,#F3F4F6 3px 6px)}
.cc.over{background:#FDECEA;box-shadow:inset 0 0 0 2px #B42318}
.cc.full{background:#F7F8FA}
.cc.unknown{background:#F3F4F6}
.cc.unknown .cc-num{color:#4B5563}
.cc-tent{height:0;border-top:2px dashed #1F5FA8;margin-top:3px}
.cc-legend{display:flex;flex-wrap:wrap;gap:.35rem 1rem;font-size:.78rem;color:#4B5563;margin:.2rem 0 .3rem 0}
.cc-legend span{display:inline-flex;align-items:center;gap:.35rem}
.cc-legend i{display:inline-block;width:18px;height:10px;border-radius:2px}
.cc-legend i.overbox{background:#FDECEA;box-shadow:inset 0 0 0 2px #B42318}
.cc-legend i.tent{height:0;border-top:2px dashed #1F5FA8;border-radius:0}
@media (max-width:640px){.cc-table .cc-type{min-width:88px;max-width:96px;font-size:.72rem}.cc-table{font-size:.72rem}}
</style>
"""


# Čitelnost ovládacích prvků záložky (jen widgety s klíčem cap_…, ostatní záložky se nemění).
TAB_CSS = """
<style>
.st-key-cap_period [data-testid="stSliderTickBar"], .st-key-cap_period [data-testid="stSliderTickBar"] p{color:#4B5563 !important}
.st-key-cap_plan_start .react-aria-DateField span{color:#4B5563 !important}
</style>
"""


def legend_html() -> str:
    items = [f'<span><i style="background:{COLORS[k]}"></i>{LABELS[k]}</span>'
             for k in ("free", "rental", "job", "maintenance")]
    items += ['<span><i class="overbox"></i>konflikt / nad kapacitu (+ počet kusů navíc)</span>',
              '<span><i class="unknown"></i>? nelze určit – chybí údaje, nejde o volno</span>',
              '<span><i class="tent"></i>předběžná rezervace (kapacitu neblokuje)</span>']
    return f'<div class="cc-legend">{"".join(items)}</div>'


def grid_html(df: pd.DataFrame, names: dict, units: dict, plan_start=None) -> str:
    """HTML mřížka: řádky = typy strojů (v pořadí `names`), sloupce = dny.
    Číslo v buňce = volné kusy / kusy k dispozici."""
    day_list = sorted(df["day"].unique())
    head = '<th class="cc-type">Stroj / den plánu</th>' + "".join(
        f'<th>{d}' + (f'<small>{html.escape(day_label(d, plan_start, short=True))}</small>' if plan_start else "")
        + "</th>" for d in day_list)
    body = ""
    present = set(df["type_id"])
    order = [t for t in names if t in present] + sorted(present - set(names))
    for t in order:
        rows = df[df["type_id"] == t]
        name = names.get(t, t)
        cells = "".join(cell_html(r, name, plan_start) for _, r in rows.sort_values("day").iterrows())
        n = units.get(t, 0)
        body += (f'<tr><td class="cc-type">{html.escape(name)}<small>{n} ks ve flotile</small></td>'
                 f"{cells}</tr>")
    return (CSS + legend_html() + f'<div class="cc-wrap"><table class="cc-table"><thead><tr>{head}'
            f"</tr></thead><tbody>{body}</tbody></table></div>")


def fleet_units(cal: cp.Calendar) -> dict:
    out = {}
    for u in cal.units:
        out[u.type_id] = out.get(u.type_id, 0) + 1
    return out


# ---------------------------------------------------------------------------
# Souhrn, detail dne, konflikty
# ---------------------------------------------------------------------------

def summary(df: pd.DataFrame, names: dict) -> dict:
    """Souhrnná čísla za zobrazené období (z hodnot capacity.daily_capacity)."""
    conflict_days = sorted(df.loc[df["status"] == cp.OVER, "day"].unique())
    unknown_cells = int((df["status"] == cp.UNKNOWN).sum())
    known = df[df["status"] != cp.UNKNOWN]
    # Obsazenost nejvýš 100 % kusů – přetížení (konflikt) se vykazuje zvlášť v conflict_days.
    known = known.assign(used=known[["used", "units"]].min(axis=1))
    by_type = known.groupby("type_id")[["used", "units"]].sum()
    by_type = by_type[by_type["units"] > 0]
    busiest, share = None, None
    if len(by_type):
        ratio = (by_type["used"] / by_type["units"]).sort_values(ascending=False)
        busiest, share = names.get(ratio.index[0], ratio.index[0]), float(ratio.iloc[0]) * 100
    return dict(conflict_days=conflict_days, unknown_cells=unknown_cells, busiest=busiest,
                busiest_share=share, tentative_cells=int((df["tentative"] > 0).sum()))


def day_status_text(row) -> str:
    """Stav jednoho dne a stroje slovy – nikdy záporný počet volných kusů."""
    units, free = int(row["units"]), int(row["free"])
    if row["status"] == cp.UNKNOWN:
        text = "počet volných kusů nelze určit"
    elif free < 0:
        text = (f"volné 0 z {units} ks – obsazení je o {-free} ks víc, než je k dispozici "
                "(konflikt)")
    else:
        text = f"volné {free} z {units} ks"
    return f"{text} · stav: {row['status']}"


def day_detail(cal: cp.Calendar, type_id: str, day: int, jobs: dict, plan_start=None) -> pd.DataFrame:
    """Co v daný den drží stroje daného typu (potvrzené, předběžné i zakázky bez termínu)."""
    rows = []
    for a in cal.assignments:
        if a.type_id == type_id and a.start <= day <= a.end:
            note = ""
            if a.kind == "job" and a.ref in jobs:
                deadline = jobs[a.ref][1]
                note = (f"po termínu (termín {day_label(deadline, plan_start)})" if day > deadline
                        else f"termín {day_label(deadline, plan_start)}")
            rows.append({
                "Druh": cp.KIND_LABEL[a.kind],
                "Označení": job_title(a.ref, jobs) if a.kind == "job" else a.ref,
                "Od": day_label(a.start, plan_start), "Do": day_label(a.end, plan_start),
                "Kusů": int(a.quantity),
                "Kus": a.unit_id or "kterýkoli kus typu",
                "Stav": "potvrzeno" if a.status == cp.BLOCKING else "předběžně (neblokuje)",
                "Poznámka": note,
            })
    for j in cal.open_jobs:
        if j.type_id == type_id and j.may_occupy(day):
            rows.append({"Druh": cp.KIND_LABEL["job"], "Označení": job_title(j.ref, jobs),
                         "Od": day_label(j.earliest, plan_start),
                         "Do": "neznámo", "Kusů": int(j.quantity), "Kus": "kterýkoli kus typu",
                         "Stav": "potvrzeno, bez termínu → nelze určit", "Poznámka": "není v plánu"})
    return pd.DataFrame(rows, columns=["Druh", "Označení", "Od", "Do", "Kusů", "Kus", "Stav", "Poznámka"])


def unit_states(cal: cp.Calendar, type_id: str, day: int) -> pd.DataFrame:
    occ = cp.unit_occupancy(cal, day, day)
    occ = occ[occ["type_id"] == type_id]
    states = [s if s != "bez přiřazení" else "bez přiřazení ke konkrétnímu kusu (rozhodují počty kusů)"
              for s in occ["state"]]
    return pd.DataFrame({"Kus": occ["unit_id"].values, "Stav kusu": states})


def conflicts_table(cal: cp.Calendar, start: int, end: int, type_ids, names: dict, jobs: dict,
                    plan_start=None) -> pd.DataFrame:
    rows = []
    for c in cp.find_conflicts(cal, start, end):
        type_id = c["id"] if c["scope"] == "typ" else next(
            (u.type_id for u in cal.units if u.unit_id == c["id"]), c["id"])
        if type_id not in type_ids:
            continue
        first, last = c["days"][0], c["days"][-1]
        rows.append({
            "Stroj": names.get(type_id, type_id) + ("" if c["scope"] == "typ" else f" ({c['id']})"),
            "Konflikt": c["kind"].replace("překročení kapacity typu – ", ""),
            "Dny": day_label(first, plan_start) if first == last
            else f"{day_label(first, plan_start)} – {day_label(last, plan_start)}",
            "Chybí kusů": c.get("max_over", 1),
            "Příčina": "; ".join(describe(s, jobs) for s in c["sources"]),
            "Jistota": "jisté" if c.get("certain", True) else "nejisté (neúplná data o flotile)",
        })
    return pd.DataFrame(rows, columns=["Stroj", "Konflikt", "Dny", "Chybí kusů", "Příčina", "Jistota"])


def late_jobs(cal: cp.Calendar, jobs: dict, type_ids=None, start=None, end=None) -> list:
    """[(job_id, název, konec, termín)] zakázek, které v plánu končí po termínu.

    Volitelně jen ty, které drží vybrané typy strojů ve dnech po termínu v zobrazeném období.
    """
    out = {}
    for a in cal.assignments:
        if a.kind != "job" or a.ref not in jobs or a.end <= jobs[a.ref][1]:
            continue
        if type_ids is not None and a.type_id not in type_ids:
            continue
        late_from = max(a.start, jobs[a.ref][1] + 1)            # první den, kdy stroj drží po termínu
        if (start is not None and a.end < start) or (end is not None and late_from > end):
            continue
        out[a.ref] = (a.ref, jobs[a.ref][0], a.end, jobs[a.ref][1])
    return sorted(out.values())


# ---------------------------------------------------------------------------
# Ověření požadavku
# ---------------------------------------------------------------------------

def check_request(cal: cp.Calendar, type_id: str, start: int, end: int, quantity: int,
                  include_tentative: bool, jobs: dict, plan_start=None) -> dict:
    """Výsledek capacity.evaluate_request doplněný o čitelné texty a tabulku problémových dnů."""
    r = cp.evaluate_request(cal, type_id, start, end, quantity, include_tentative)
    problems = pd.DataFrame(
        [{"Den": day_label(d, plan_start), "Volné kusy": free, "Požadováno": need,
          "Co stroj drží": "; ".join(describe(s, jobs) for s in causes) or "kus není ve flotile"}
         for d, free, need, causes in r["problem_days"]],
        columns=["Den", "Volné kusy", "Požadováno", "Co stroj drží"])
    if r["status"] == cp.FEASIBLE:
        headline = (f"Požadavek lze přijmout – v každém dni je volných aspoň {r['min_free']} ks "
                    f"(potřeba {quantity} ks).")
    elif r["status"] == cp.INFEASIBLE and not r["problem_days"]:
        headline = "Požadavek nelze přijmout – " + "; ".join(r["reasons"]) + "."
    elif r["status"] == cp.INFEASIBLE:
        headline = (f"Požadavek nelze přijmout – volná kapacita chybí "
                    f"(problémové dny: {len(r['problem_days'])} z {len(r['free_by_day'])}).")
    elif r["min_free"] is None:                       # neplatný požadavek (např. konec před začátkem)
        headline = "Nelze vyhodnotit – " + "; ".join(r["reasons"]) + "."
    else:
        headline = "Nelze určit – chybí údaje, které by mohly výsledek změnit. Nejde o potvrzení volna."
    return dict(status=r["status"], tone=STATUS_TONE[r["status"]], headline=headline, problems=problems,
                reasons=[describe(x, jobs) for x in r["reasons"]],
                warnings=[(day_label(d, plan_start), describe(t, jobs)) for d, t in r["warnings"]],
                raw=r)
