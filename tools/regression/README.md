# Regresní nástroje (UX/UI 2, etapa 0)

Nástroje ověřují, že úprava vzhledu nezměnila výpočty, data ani chování aplikace. Aplikaci ani data
nemění. Přihlašují se náhodným jednorázovým heslem (AppTest) nebo heslem z proměnné prostředí
(vizuální porovnání). **Žádné heslo nepatří do kódu, parametrů příkazové řádky ani do repozitáře.**

## 1. Automatický test v běžné sadě

`tests/test_ui_regression.py` – spouští se s ostatními testy (`python -m pytest tests`).
Porovná výchozí stav všech 8 záložek a postranního panelu se základem `tests/fixtures/ui_baseline.json`:

- popisky záložek,
- všechna viditelná čísla (nic nesmí zmizet, změnit se ani přibýt),
- widgety (typ, klíč, popisek, výchozí hodnota, volby; obsah textových polí se neukládá),
- počet grafů, tabulek a tlačítek exportu DOCX.

Viditelný text a rozbalovací bloky se nekontrolují – při úpravách vzhledu se smějí měnit.
Po **záměrné a schválené** změně čísel nebo widgetů základ obnovte a změnu popište v PR:

```
python tests/test_ui_regression.py --update
```

## 2. Snímek a porovnání dvou verzí

```
python tools/regression/snapshot.py <adresář_verze_před> pred.json
python tools/regression/snapshot.py <adresář_verze_po> po.json
python tools/regression/compare.py pred.json po.json [--allow-added-numbers 7] [--strict-text]
```

Verzi „před“ lze připravit např. příkazem `git archive origin/main | tar -x -C /tmp/main`.
Snímek obsahuje:

| Část | Co zachycuje |
|---|---|
| `calc` | výpočty záložek 1–6 ve 4 scénářích postranního panelu (simulace, plán zakázek, nákupní sestava) a otisky všech částí exportu DOCX |
| `ui` | výchozí stav všech záložek (jako test výše) včetně otisku textu |
| `tab7` | souhrnné ukazatele a tabulky záložky 7 pro 7 kombinací filtrů |
| `planner` | otisky verdiktů rozhodovacího plánovače pro 6 vzorových situací |

`compare.py` vrací 1 při přísném rozdílu (výpočty, DOCX, widgety, změněná či chybějící čísla, souhrn
záložky 7, plánovač). Změna textu a rozbalovacích bloků je jen informace.

## 3. Vizuální porovnání (pixelové)

Vývojový nástroj – potřebuje Playwright s prohlížečem Chromium a Pillow (nejsou v `requirements.txt`).

1. Spusťte obě verze na různých portech, každou s lokálním `.streamlit/secrets.toml` se stejným
   testovacím heslem (soubor je v `.gitignore`).
2. `DE_TEST_PASSWORD=… python tools/regression/visual.py http://localhost:8501 http://localhost:8502 vystup/`
   (volitelně `DE_WIDTHS=1440,900,768,390`, `DE_CHROMIUM=/cesta/k/chrome`).

Porovná obsah všech 8 záložek na zvolených šířkách, záložku 7 v obou pohledech a záložku 8 se vzorovým
požadavkem. U rozdílu uloží obrázek `…_rozdil.png`. Odchylky barvy nejvýš 2 z 255 (`DE_NOISE`) jsou šum
vyhlazování hran v prohlížeči – stejný šum vzniká i při porovnání verze se sebou samou – a chybou nejsou.

## Doporučený postup u každé etapy UX/UI

1. `python -m pytest tests` – celá sada.
2. Snímek `main` × větev a `compare.py` – bez přísných rozdílů.
3. `visual.py` – ostatní záložky pixelově shodné; u upravené záložky snímky před a po.
4. Kontrola diffu: jen zamýšlené soubory, žádná hesla ani tokeny.
