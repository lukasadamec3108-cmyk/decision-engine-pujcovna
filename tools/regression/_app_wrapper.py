"""Obal pro snapshot.py: spustí aplikaci (cesta v DE_REG_APP) v AppTest a uloží vypočtené hodnoty.

Nespouštět ručně. Aplikace se vykoná beze změny; obal jen po doběhnutí přečte její proměnné
(výsledky simulace, plán zakázek, nákupní sestavu) a vytvoří DOCX stejnou funkcí jako tlačítko.
"""

import hashlib
import io
import json
import os
import sys
import zipfile

import numpy as np

APP = os.environ["DE_REG_APP"]
OUT = os.environ["DE_REG_OUT"]
sys.path.insert(0, os.path.dirname(APP))
g = {"__name__": "__main__", "__file__": APP}
exec(compile(open(APP, encoding="utf-8").read(), APP, "exec"), g)   # při přihlašovací obrazovce skončí st.stop()

KEYS = ["machine_1", "machine_2", "fixed_cost", "scenario", "min_days", "max_days", "expected_days", "n", "budget",
        "expected_profit_mini", "expected_profit_dumper", "prob_loss_mini", "prob_loss_dumper",
        "p10_mini", "p50_mini", "p90_mini", "p10_dumper", "p50_dumper", "p90_dumper", "var95_mini", "var95_dumper",
        "prob_dumper_better", "break_even_mini", "break_even_dumper", "roi_mini", "roi_dumper",
        "expected_profit_mini_real", "expected_profit_dumper_real", "required_price_1", "required_price_2",
        "scorecard_mini", "scorecard_dumper", "better_machine", "labels", "values",
        "profit_mini", "profit_dumper", "profit_matrix", "profit_curve_mini", "profit_curve_dumper"]


def conv(v):
    if isinstance(v, np.ndarray):
        a = np.round(v.astype(float), 6)
        return {"sha": hashlib.sha256(a.tobytes()).hexdigest()[:16], "sum": round(float(a.sum()), 4), "n": int(a.size)}
    if isinstance(v, (np.floating, float)):
        return round(float(v), 6)
    if isinstance(v, (np.integer, int)) and not isinstance(v, bool):
        return int(v)
    if isinstance(v, dict):
        return {str(k): conv(x) for k, x in v.items()}
    if isinstance(v, (list, tuple)):
        return [conv(x) for x in v]
    return v if isinstance(v, (str, type(None), bool)) else str(v)


res = {k: conv(g[k]) for k in KEYS if k in g}
res["schedule"] = [[r.job_name, conv(r.start_day), conv(r.end_day), r.schedule_status]
                   for r in g["scheduled_jobs"].itertuples()]
pb = g.get("purchase_best") or {}
res["purchase_best"] = conv({k: pb.get(k) for k in ("counts", "units", "investment", "revenue", "costs", "profit")})
docx = g["build_docx_report"]()
z = zipfile.ZipFile(io.BytesIO(docx))
res["docx_parts"] = {n: hashlib.sha256(z.read(n)).hexdigest()[:16] for n in sorted(z.namelist())}
json.dump(res, open(OUT, "w"), ensure_ascii=False, indent=1)
