"""Obnovení dat po změně datových souborů bez restartu aplikace.

Streamlit Cloud při novém nasazení neukončí běžící proces – mezipaměť (st.cache_data) by bez
otisku dat vracela tabulky načtené před nasazením. Test spustí kopii aplikace v samostatném
procesu, změní CSV soubor a ověří, že se změna projeví v záložce Kapacita strojů i v datech
pro ostatní záložky. Heslo je náhodné, jen pro test.

Spuštění: python -m pytest tests   nebo   python tests/test_data_refresh.py
"""

import json
import secrets
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import datamodel as dm  # noqa: E402

# Skript běží v novém procesu, aby moduly i mezipaměť patřily jen kopii aplikace.
SCRIPT = r'''
import json, os, sys, logging, warnings
logging.disable(logging.CRITICAL); warnings.filterwarnings("ignore")
app_dir, password = sys.argv[1], sys.argv[2]
sys.path.insert(0, app_dir); os.chdir(app_dir)
from streamlit.testing.v1 import AppTest
at = AppTest.from_file(os.path.join(app_dir, "app.py"), default_timeout=300)
at.secrets["password"] = password
at.run(); at.text_input[0].set_value(password).run()

def snapshot():
    md = [m.value for m in at.tabs[6].markdown] if len(at.tabs) == 7 else []
    return dict(kpi=[m for m in md if "Dny s konfliktem" in m],
                info=[m for m in md if "Načtená data" in m],
                plan=[d.value["Konec (den)"].tolist() for d in at.tabs[5].dataframe
                      if "Konec (den)" in d.value.columns],
                exceptions=[e.value for e in at.exception])

def edit(name, keep):
    path = os.path.join(app_dir, "data", "model", name)
    lines = open(path, encoding="utf-8").read().splitlines()
    open(path, "w", encoding="utf-8").write("\n".join(l for l in lines if keep(l)) + "\n")

out = {"start": snapshot()}
edit("rental_bookings.csv", lambda l: not l.startswith("FIKT-P05"))      # zruší konflikt válce
at.run(); out["bez_FIKT_P05"] = snapshot()
edit("maintenance.csv", lambda l: not l.startswith("KOLOVY_BAGR-01"))    # zruší konflikt bagru
at.run(); out["bez_servisu"] = snapshot()
path = os.path.join(app_dir, "data", "model", "jobs.csv")                 # data záložek 1–6
text = open(path, encoding="utf-8").read()
open(path, "w", encoding="utf-8").write(text.replace("J010,", "J010,", 1).replace(",2,1,5,", ",3,1,5,", 1))
at.run(); out["delsi_J010"] = snapshot()
print("VYSLEDEK" + json.dumps(out, ensure_ascii=False))
'''


def run_copy(app_root: Path = ROOT) -> dict:
    tmp = Path(tempfile.mkdtemp())
    try:
        shutil.copytree(app_root, tmp / "app", ignore=shutil.ignore_patterns(
            ".git", "tests", "__pycache__", "secrets.toml", "*.xlsx"))
        (tmp / "run.py").write_text(SCRIPT, encoding="utf-8")
        res = subprocess.run([sys.executable, str(tmp / "run.py"), str(tmp / "app"), secrets.token_urlsafe(12)],
                             capture_output=True, text=True, timeout=600, cwd=tmp)
        line = [x for x in res.stdout.splitlines() if x.startswith("VYSLEDEK")]
        assert line, res.stdout[-2000:] + res.stderr[-2000:]
        return json.loads(line[0][len("VYSLEDEK"):])
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def conflict_days(snap) -> str:
    kpi = snap["kpi"][0]
    return kpi.split('class="val ')[1].split(">")[1].split("<")[0]


def test_changed_data_files_are_reloaded_without_restart():
    out = run_copy()
    for step in out.values():
        assert step["exceptions"] == []
    assert conflict_days(out["start"]) == "5"                 # válec 15–18 + kolový bagr 11
    assert conflict_days(out["bez_FIKT_P05"]) == "1"           # po odebrání pronájmu jen bagr
    assert conflict_days(out["bez_servisu"]) == "0"            # po odebrání servisu žádný
    assert "7 pronájmů" in out["start"]["info"][0] and "3 servisní odstávky" in out["start"]["info"][0]
    assert "6 pronájmů" in out["bez_FIKT_P05"]["info"][0]
    assert "2 servisní odstávky" in out["bez_servisu"]["info"][0]
    # plán zakázek (záložka 6) pronájmy ani servis nepoužívá – zůstává stejný
    assert out["start"]["plan"] == out["bez_FIKT_P05"]["plan"] == out["bez_servisu"]["plan"]
    # změna dat zakázky (delší J010) se projeví i v plánu záložky 6 – bez restartu
    assert out["delsi_J010"]["plan"] != out["bez_servisu"]["plan"]


def test_data_fingerprint_changes_only_with_csv_content():
    tmp = Path(tempfile.mkdtemp())
    try:
        shutil.copytree(dm.DEFAULT_DIR, tmp, dirs_exist_ok=True)
        first = dm.data_fingerprint(tmp)
        assert first == dm.data_fingerprint(tmp) == dm.data_fingerprint(dm.DEFAULT_DIR)
        (tmp / "README.md").write_text("jiný popis", encoding="utf-8")         # popis není data
        assert dm.data_fingerprint(tmp) == first
        path = tmp / "rental_bookings.csv"
        original = path.read_text(encoding="utf-8")
        path.write_text(original.replace("FIKT-P01,MINI_BAGR,13,16", "FIKT-P01,MINI_BAGR,13,17"), encoding="utf-8")
        assert dm.data_fingerprint(tmp) != first                                 # změna jednoho dne
        path.write_text(original, encoding="utf-8")
        assert dm.data_fingerprint(tmp) == first                                 # zpět stejný obsah
        (tmp / "novy.csv").write_text("a\n1\n", encoding="utf-8")
        assert dm.data_fingerprint(tmp) != first                                 # nový soubor
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def test_data_info_line():
    import capacity_view as cv
    text = cv.data_info(dm.load_model(), "abcdef1234567890")
    assert text == ("Načtená data: 7 pronájmů (5 potvrzených, 2 předběžné) · 3 servisní odstávky"
                    " · verze dat abcdef12")
    assert cv.plural(1, "pronájem", "pronájmy", "pronájmů") == "1 pronájem"
    assert cv.plural(0, "pronájem", "pronájmy", "pronájmů") == "0 pronájmů"


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_"):
            fn()
            print("OK", name)
