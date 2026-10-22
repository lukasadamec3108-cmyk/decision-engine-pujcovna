"""Vizuální regrese: pixelové porovnání záložek dvou běžících verzí aplikace (např. main × větev).

Jen vývojový nástroj – potřebuje Playwright (Chromium) a Pillow, které nejsou v requirements.txt
a aplikace je nepotřebuje.

Postup:
  1. Spusťte obě verze lokálně na různých portech, každou s vlastním lokálním .streamlit/secrets.toml
     se STEJNÝM testovacím heslem (soubor je v .gitignore – nikdy ho necommitujte):
        streamlit run app.py --server.port 8501      (verze „před“)
        streamlit run app.py --server.port 8502      (verze „po“)
  2. Heslo předejte proměnnou prostředí (nikdy ne parametrem ani v kódu):
        DE_TEST_PASSWORD=… python tools/regression/visual.py http://localhost:8501 http://localhost:8502 vystup/
     Volitelně DE_CHROMIUM=/cesta/k/chrome a DE_WIDTHS=1440,900,768,390.

Porovnává se obsah každé záložky (celá výška) na zvolených šířkách, navíc záložka 7 v pohledu
„Přehled využití“ a záložka 8 se vzorovým požadavkem. Před snímkem se myš přesune do rohu, aby se
nezachytily efekty najetí myší. Odchylky barvy nejvýš DE_NOISE (výchozí 2 z 255) jsou šum vyhlazování
hran v prohlížeči – vzniká i při porovnání verze se sebou samou; vypíší se, ale nejsou chybou.
Výsledek: seznam shod a rozdílů, u rozdílu obrázek „…_rozdil.png“. Návratový kód 1 při rozdílu.
"""

import os
import sys
from pathlib import Path


def _idle(p):
    p.wait_for_timeout(1500)
    p.wait_for_function("!document.querySelector('[data-testid=stStatusWidget]')", timeout=180000)
    p.wait_for_timeout(1500)


def _views(p):
    """Pořadí pohledů: (název, akce před snímkem). Akce jen kliká – nic se neukládá."""
    def tab(i):
        return lambda: p.locator("[role=tab]").nth(i).click()

    def usage():
        p.locator("[role=tab]").nth(6).click()
        _idle(p)
        p.locator("[role=tabpanel]:visible").get_by_text("Přehled využití").click()

    def calendar_back():
        p.locator("[role=tabpanel]:visible").get_by_text("Kalendář dostupnosti").click()

    def request():
        p.locator("[role=tab]").nth(7).click()
        _idle(p)
        p.locator(".st-key-dp_r_type input").click()
        p.locator(".st-key-dp_r_type input").fill("Pásový dumper")
        p.keyboard.press("Enter")
        _idle(p)
        p.get_by_role("button", name="Vyhodnotit samostatně").click()

    views = [(f"t{i + 1}", tab(i)) for i in range(8)]
    views += [("t7_vyuziti", usage), ("t7_kalendar_zpet", calendar_back), ("t8_pozadavek", request)]
    return views


def shoot(url: str, width: int, out: Path, tag: str, password: str, chromium: str) -> list:
    from playwright.sync_api import sync_playwright

    files = []
    with sync_playwright() as pw:
        b = pw.chromium.launch(executable_path=chromium) if chromium else pw.chromium.launch()
        mobile = width < 500
        p = b.new_page(viewport={"width": width, "height": 900}, is_mobile=mobile, has_touch=mobile)
        p.goto(url)
        p.wait_for_selector("input[type=password]", timeout=90000)
        p.fill("input[type=password]", password)
        p.keyboard.press("Enter")
        p.wait_for_selector(".de-header", timeout=180000)
        _idle(p)
        for name, action in _views(p):
            action()
            p.mouse.move(0, 0)            # bez najetí myší: jinak se např. u posuvníku ukazuje stupnice
            _idle(p)
            f = out / f"{tag}_{width}_{name}.png"
            p.locator("[role=tabpanel]:visible").screenshot(path=str(f), animations="disabled")
            files.append((name, f))
        b.close()
    return files


def main(argv):
    if len(argv) != 4:
        print(__doc__)
        return 2
    password = os.environ.get("DE_TEST_PASSWORD")
    if not password:
        print("Chybí proměnná prostředí DE_TEST_PASSWORD (testovací heslo obou lokálních instancí).")
        return 2
    from PIL import Image, ImageChops

    before, after, out = argv[1], argv[2], Path(argv[3])
    out.mkdir(parents=True, exist_ok=True)
    widths = [int(w) for w in os.environ.get("DE_WIDTHS", "1440,900,768,390").split(",")]
    chromium = os.environ.get("DE_CHROMIUM", "")
    noise = int(os.environ.get("DE_NOISE", "2"))      # odchylka barvy (0–255), která se bere jako šum vykreslení
    differences = 0
    for w in widths:
        a = dict(shoot(before, w, out, "pred", password, chromium))
        b = dict(shoot(after, w, out, "po", password, chromium))
        for name in a:
            ia, ib = Image.open(a[name]).convert("RGB"), Image.open(b[name]).convert("RGB")
            if ia.size != ib.size:
                differences += 1
                print(f"{w} px {name}: ROZDÍL – jiná velikost {ia.size} → {ib.size}")
                continue
            diff = ImageChops.difference(ia, ib)
            box = diff.getbbox()
            if box is None:
                print(f"{w} px {name}: shodné pixelově")
                continue
            hist = diff.convert("L").histogram()
            changed = sum(hist[1:])
            worst = max(i for i, v in enumerate(hist) if v)
            if worst <= noise:
                # vyhlazování hran v prohlížeči: jednotky pixelů o 1–2 odstíny, okem neviditelné
                # (stejný šum vzniká i při porovnání verze se sebou samou)
                print(f"{w} px {name}: shodné (šum vykreslení: {changed} px, odchylka nejvýš {worst}/255)")
                continue
            differences += 1
            diff.save(out / f"{w}_{name}_rozdil.png")
            print(f"{w} px {name}: ROZDÍL v oblasti {box} – {changed} px, odchylka až {worst}/255")
    print("VÝSLEDEK:", "bez viditelných rozdílů" if not differences else f"{differences} rozdílů")
    return 1 if differences else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
