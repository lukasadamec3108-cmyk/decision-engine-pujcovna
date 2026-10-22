"""Regresní snímek aplikace Decision Engine – pro porovnání dvou verzí (např. main × pracovní větev).

Použití (z kořene repozitáře, se stejným Pythonem, jakým běží testy):

    python tools/regression/snapshot.py <adresář_aplikace> <výstup.json>
    python tools/regression/compare.py <před.json> <po.json>

Co snímek obsahuje:
  calc     … výpočty záložek 1–6 ve 4 scénářích postranního panelu (simulace, plán, nákupní sestava)
  docx     … otisky všech částí exportu DOCX (pro každý scénář)
  ui       … výchozí stav všech 8 záložek: popisky záložek, viditelná čísla, widgety, otisk textu
  tab7     … souhrnné ukazatele a tabulky záložky 7 pro několik kombinací filtrů
  planner  … otisky verdiktů rozhodovacího plánovače pro vzorové požadavky

Přihlašuje se náhodným jednorázovým heslem; žádné heslo se neukládá ani nevypisuje.
Nástroj aplikaci ani data nemění.
"""

import datetime
import hashlib
import json
import logging
import os
import secrets
import sys
import tempfile
from pathlib import Path

logging.disable(logging.WARNING)   # hlášky Streamlitu v režimu AppTest (např. „missing ScriptRunContext“)
HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import extract  # noqa: E402

# Scénáře postranního panelu (stejné jako při dosavadních regresních kontrolách)
SCENARIOS = {
    "S1_hlavni": dict(m1="Mini bagr", m2="Pásový dumper", scen="Realistický", fixed=100000, days=18, n=5000,
                      budget=3000000),
    "S2_pesimisticky": dict(m1="Mini bagr", m2="Pásový dumper", scen="Pesimistický", fixed=120000, days=18,
                            n=1000, budget=3000000),
    "S3_optimisticky": dict(m1="Kolový bagr", m2="Silniční válec", scen="Optimistický", fixed=60000, days=25,
                            n=2000, budget=1500000),
    "S4_jine": dict(m1="Smykový nakladač", m2="Nákladní auto", scen="Realistický", fixed=120000, days=12, n=1000,
                    budget=1000000),
}

# Filtry záložky 7 (typy strojů, datum prvního dne, období, zobrazení)
TAB7_FILTERS = [
    dict(), dict(period=(5, 12)), dict(types=["SILNICNI_VALEC", "KOLOVY_BAGR"]),
    dict(period=(11, 11), types=["KOLOVY_BAGR"]), dict(start=datetime.date(2026, 11, 2), period=(14, 20)),
    dict(view="Přehled využití"), dict(view="Přehled využití", period=(15, 18), types=["SILNICNI_VALEC"]),
]


def _find(widgets, text):
    for w in widgets:
        if text.lower() in w.label.lower():
            return w
    raise KeyError(text)


def calc_snapshot(root: Path) -> dict:
    """Výpočty a DOCX pro každý scénář – aplikace běží v AppTest přes obal _app_wrapper.py."""
    from streamlit.testing.v1 import AppTest

    out = {}
    os.chdir(root)
    with tempfile.TemporaryDirectory() as tmp:
        for name, p in SCENARIOS.items():
            os.environ["DE_REG_APP"] = str(root / "app.py")
            os.environ["DE_REG_OUT"] = os.path.join(tmp, name + ".json")
            password = secrets.token_urlsafe(12)
            at = AppTest.from_file(str(HERE / "_app_wrapper.py"), default_timeout=300)
            at.secrets["password"] = password
            at.run()
            at.text_input[0].set_value(password).run()
            sb = at.sidebar
            _find(sb.selectbox, "Stroj 1").set_value(p["m1"])
            _find(sb.selectbox, "Stroj 2").set_value(p["m2"])
            _find(sb.selectbox, "Scénář").set_value(p["scen"])
            _find(sb.number_input, "Fixní").set_value(p["fixed"])
            _find(sb.number_input, "rozpočet").set_value(p["budget"])
            _find(sb.slider, "poptávk").set_value(p["days"])
            _find(sb.slider, "simulac").set_value(p["n"])
            at.run()
            if at.exception:
                raise RuntimeError(f"{name}: " + "; ".join(e.value[:200] for e in at.exception))
            out[name] = json.load(open(os.environ["DE_REG_OUT"], encoding="utf-8"))
    return out


def tab7_snapshot(at) -> list:
    """Souhrn a tabulky záložky 7 pro každou kombinaci filtrů (klíče widgetů jsou stabilní)."""
    rows = []
    for f in TAB7_FILTERS:
        at.multiselect(key="cap_types").set_value(f.get("types", [])).run()
        at.date_input(key="cap_plan_start").set_value(f.get("start")).run()
        at.slider(key="cap_period").set_value(f.get("period", (1, 25))).run()
        at.radio(key="cap_view").set_value(f.get("view", "Kalendář dostupnosti")).run()
        t = at.tabs[6]
        kpis = [m.value for m in t.markdown if "de-kpis" in m.value and "<style>" not in m.value]
        tables = [d.value.to_csv() for d in t.dataframe]
        rows.append(dict(filters={k: str(v) for k, v in f.items()}, kpis=kpis, tables=tables,
                         numbers=extract.visible_numbers(t),
                         exceptions=[e.value[:200] for e in at.exception]))
    return rows


def planner_snapshot(root: Path) -> list:
    """Otisky verdiktů plánovače – stejné případy jako dřívější kontroly (data z modelu, nic se neukládá)."""
    sys.path.insert(0, str(root))
    sys.path.insert(0, str(root / "tests"))
    import datamodel as dm
    import planner as pl
    import planner_view as pv
    import scheduling as sch
    from test_planner import model

    m = dm.load_model()
    jobs = dm.legacy_jobs(m)
    plan = sch.evaluate(jobs, sch.optimize_schedule(jobs, dm.fleet_capacity(m))["starts"])
    starts = {jobs.loc[i, "job_id"]: int(r.start_day) for i, r in plan.iterrows() if r.scheduled}
    cases = [
        (m, starts, [pv.rental_request("NP-1", "SMYKOVY_NAKLADAC", 1, 18, 24)]),
        (m, starts, [pv.rental_request("NP-1", "PASOVY_DUMPER", 1, 20, 26, 9000),
                     pv.rental_request("NP-2", "MINI_BAGR", 1, 13, 15)]),
        (m, starts, [pv.rental_request("NP-1", "PASOVY_DUMPER", 1, 22, 23)]),
        (m, starts, [pv.rental_request("NP-1", "MINI_BAGR", 1, 40, 44, 6100),
                     pv.rental_request("NP-2", "MINI_BAGR", 1, 42, 43, 9000)]),
        (model(), {}, [pv.job_request("NP-1", "J", [("MINI_BAGR", 1)], 1, 10, None, None, 150000, 60000, None),
                       pv.rental_request("NP-2", "MINI_BAGR", 1, 1, 5, 6500),
                       pv.rental_request("NP-3", "MINI_BAGR", 1, 6, 10, 7000)]),
        (model(var_rental=7000), {}, [pv.rental_request("NP-1", "MINI_BAGR", 1, 1, 3, 6000),
                                      pv.rental_request("NP-2", "MINI_BAGR", 1, 2, 4, 6500)]),
    ]
    out = []
    for mm, st, reqs in cases:
        p = pl.plan(mm, st, [pv.to_request(x) for x in reqs])
        key = [(g["verdict"], g["choice"], sorted(g.get("contenders", [])), sorted(g.get("best", [])),
                sorted((k, v.lo, v.hi) for k, v in g.get("values", {}).items())) for g in p.groups]
        key += [(r, e["operational"], None if e["economics"] is None else (e["economics"].lo, e["economics"].hi))
                for r, e in sorted(p.requests.items())]
        out.append(hashlib.sha256(json.dumps(key, sort_keys=True, default=str).encode()).hexdigest()[:16])
    return out


def main(argv):
    if len(argv) != 3:
        print(__doc__)
        return 2
    root, target = Path(argv[1]).resolve(), Path(argv[2]).resolve()
    if not (root / "app.py").exists():
        print(f"V adresáři {root} není app.py.")
        return 2
    snap = dict(app=str(root))
    snap["calc"] = calc_snapshot(root)
    at = extract.logged_in_app(root)
    snap["ui"] = extract.ui_state(at)
    snap["tab7"] = tab7_snapshot(at)
    snap["planner"] = planner_snapshot(root)
    target.write_text(json.dumps(snap, ensure_ascii=False, indent=1, sort_keys=True), encoding="utf-8")
    print(f"Snímek uložen: {target}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
