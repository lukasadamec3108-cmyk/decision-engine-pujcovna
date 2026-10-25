# Decision Engine – rozhodovací nástroj pro půjčovnu stavebních strojů

**Interaktivní webová aplikace pro hodnocení investic do stavebních strojů, analýzu rizika a plánování zakázek.**
Vytvořeno v Pythonu a Streamlitu.

> ⚠️ **Stav projektu: funkční pilotní prototyp připravovaný k odbornému a uživatelskému ověření.**
> Nejde o hotový produkční systém.
>
> Aplikace pracuje **výhradně s fiktivními demonstračními hodnotami**. Stroje, ceny, náklady,
> zakázky i všechny zobrazené výsledky slouží jen k ukázce funkčnosti a **nepředstavují
> skutečné ekonomické výsledky žádné společnosti**.

---

## Představení projektu

Decision Engine je samostatný analytický a vývojový projekt. Ukazuje, jak lze typická
manažerská rozhodnutí provozovatele stavebních strojů podpořit kvantitativním modelem
a přehlednou webovou aplikací.

Projekt vzniká jako pilotní řešení pro společnost působící v oblasti stavební mechanizace
a je připravován za odborných konzultací. Cílem současné fáze je:

- vytvořit funkční prototyp, na kterém lze přístup srozumitelně předvést,
- získat odbornou a uživatelskou zpětnou vazbu,
- připravit základ pro případné ověření na reálných datech.

## Podnikatelský problém

Provozovatel půjčovny stavebních strojů pravidelně řeší otázky, které přímo ovlivňují ziskovost:

- **Do kterého stroje investovat?** Nákup stroje je vysoká jednorázová investice, jejíž návratnost
  závisí na nejisté poptávce, ceně pronájmu, servisních nákladech a poruchovosti.
- **Jak velké je riziko ztráty?** Průměrný odhad nestačí. Je potřeba vědět, jak pravděpodobný
  je ztrátový měsíc a jak nepříznivé mohou být horší scénáře.
- **Při jaké ceně a vytížení se stroj vyplatí?** Kde leží bod zvratu a jaká je minimální denní sazba.
- **Jak rozdělit investiční rozpočet** mezi více typů strojů.
- **Jak seřadit a naplánovat zakázky**, když více zakázek potřebuje stejné stroje.

Decision Engine tyto otázky spojuje do jednoho interaktivního nástroje. Uživatel mění vstupní
předpoklady a okamžitě vidí dopad na očekávaný zisk, riziko a výsledné doporučení.

## Funkce aplikace

Aplikace má **levé menu** se sekcemi (etapa M2):

| Sekce | Obsah |
|---|---|
| **Dashboard** | Přehled existujících výsledků (kapacita, plán zakázek, rozpracované požadavky, doporučený stroj a nákupní sestava) s odkazy do sekcí. Nic nového nepočítá. |
| **Pronájem strojů** | Formulář nového požadavku na pronájem a jeho vyhodnocení (stejný plánovač jako Rozhodovací centrum, sdílený seznam požadavků). |
| **Vlastní zakázky** | Dřívější záložka 🗓️ Plán zakázek. |
| **Strojový park** | Počty kusů ve flotile a ekonomika jednoho kusu podle dat (stejná tabulka jako v Nákupu strojů). |
| **Plánování kapacity** | Dřívější záložka 📅 Kapacita strojů. |
| **Rozhodovací centrum** | Dřívější záložka 🧭 Rozhodovací plánovač (ekonomické porovnání, hranice zlomu, Co kdyby, konflikty). |
| **Investice** | Dřívější záložky 1–5 (jako záložky uvnitř sekce), investiční vstupy v panelu „Nastavení investičního scénáře“ nad výsledky a export DOCX. |
| **Nastavení** | Přihlášený uživatel, odhlášení, správa účtů (jen administrátor) a informace o datech. |

Investiční vstupy se zobrazují jen v sekci Investice; jejich hodnoty se při přechodu mezi sekcemi
zachovají (do odhlášení) a Dashboard počítá ze stejných hodnot.

Vstupy se nastavují na stránce Investice ve sbalovacím panelu **Nastavení investičního scénáře** nad výsledky
(pod panelem je stručné shrnutí zvolených hodnot; postranní panel obsahuje jen menu a přihlášeného uživatele):
- výběr dvou porovnávaných strojů,
- fixní měsíční náklad,
- scénář trhu (pesimistický / realistický / optimistický),
- investiční rozpočet,
- v části „Pokročilé nastavení“: odhad poptávky pro rychlý výpočet (dny v měsíci) a počet simulací,
- kapacitní scénář (nákup druhého pásového dumperu) – zatím vypnutý, zapojení je plánováno.

Výchozí hodnoty tvoří **ukázkový scénář** s fiktivními čísly (realistický trh, fixní náklad 100 000 Kč).

Výsledky byly původně rozděleny do osmi záložek; jejich obsah je beze změny v sekcích uvedených výše:

| Záložka | Obsah |
|---|---|
| 🏆 **Doporučení** | Doporučený stroj s krátkým zdůvodněním, tři hlavní ukazatele (měsíční zisk, riziko ztráty, bod zvratu), srovnání obou strojů, další kroky a **export investičního reportu do formátu DOCX**. Podrobnosti jsou v rozbalovacích sekcích. |
| 📈 **Možné výsledky** | Simulace měsíčního zisku (Monte Carlo): slabý / typický / dobrý měsíc (P10 / medián / P90), histogram; v detailu pro analytiky kumulativní pravděpodobnost (CDF) a interval nejistoty. |
| ⚠️ **Riziko** | Semafor rizika, šance ztrátového měsíce, nejhorší měsíc z 20 (VaR 5 %), pravděpodobnost poruchy a graf riziko vs. výnos. |
| 🎯 **Co ovlivňuje zisk** | Tornádo graf citlivosti na změnu vstupů o ±10 %, minimální denní cena pronájmu, zisk podle počtu dní (bod zvratu) a mapa ziskovosti (cena × počet dní). |
| 🛒 **Nákup strojů** | Optimální nákup strojů v rámci rozpočtu, křivka rozpočet vs. maximální zisk; v detailu kontrola lineární (LP) a diskrétní (DP) optimalizací. |
| 🗓️ **Plán zakázek** | Harmonogram (Ganttův diagram) se smluvními termíny, zakázky po termínu a penále za zpoždění, detail zakázek, vytížení strojů, úzké hrdlo a prioritní zakázka. |
| 📅 **Kapacita strojů** | Kalendář dostupnosti strojů po dnech: volné kusy, pronájmy, zakázky z plánu (i po termínu), servis, konflikty a stav „nelze určit“; filtr typů a období, detail dne a ověření nového požadavku. Přepínač **Přehled využití**: kapacita flotily v kusodnech, pronájmy, vlastní zakázky, servis, volno, využití a konfliktní dny – celkem i podle typu stroje. |
| 🧭 **Rozhodovací plánovač** | Zadání nového pronájmu nebo zakázky (i podle šablony existující zakázky) a vyhodnocení: provozní proveditelnost na zbývající kapacitě, co požadavek blokuje, posun zakázky v okně bez penále, soupeření s předběžnými rezervacemi a odděleně ekonomika. Na kartě dva oddělené štítky – **Provoz** (Vejde se / Nevejde se / Provoz nelze určit) a **Ekonomika** (Doporučeno / Nevyplatí se / Rovnocenné / Podmíněné / Nelze rozhodnout / Bez porovnání) – se zdůvodněním a příspěvkem; zelená jen pro „Doporučeno“. Nad kartami panel **Ekonomické porovnání** pro každou skupinu soupeřících požadavků: výsledek jednou větou, Proč (částky jen z úplných dat), porovnávané varianty včetně „nepřijmout nic“, podmínky slovy (Kč/den) s grafem hranic zlomu (pruh nejvýhodnějších variant), simulace **Co kdyby** (hypotetická hodnota jednoho nezadaného nákladu, odděleně od doporučení) a co doplnit. Nic se neukládá ani nemění. |

Čísla jsou zobrazena v českém formátu, grafy jsou interaktivní (Plotly) a vzhled je přizpůsoben i menším obrazovkám.

Přístup do aplikace je chráněn heslem.

## Analytické metody a výpočty

**Monte Carlo simulace měsíčního výsledku** (100–10 000 simulovaných měsíců) pro každý stroj zvlášť,
s pevným náhodným základem (`SIMULATION_SEED`), takže stejné vstupy dávají stejné výsledky:
- poptávka ve dnech: trojúhelníkové rozdělení v rozmezí daném scénářem trhu,
- servisní náklady: normální rozdělení (oříznuté na nezáporné hodnoty),
- poruchy: náhodný výskyt podle pravděpodobnosti poruchy, délka odstávky z Poissonova rozdělení,
- denní cena: normální rozdělení kolem ceníkové ceny (směrodatná odchylka 8 %, omezeno na ±20 %),
- měsíční zisk = skutečně pronajaté dny × denní cena − fixní náklad − servisní náklad.

**Rizikové metriky:** pravděpodobnost ztráty, percentily P10 / P50 / P90, Value at Risk (5% kvantil)
a pravděpodobnost, že jeden stroj vydělá více než druhý.

**Ekonomické ukazatele:** bod zvratu (počet dní pronájmu potřebný k pokrytí fixních nákladů),
minimální denní cena pronájmu a doba návratnosti investice (využitá v bodovém hodnocení).

**Citlivostní analýza:** jednorozměrná citlivost na cenu, vytížení, fixní a servisní náklady (±10 %)
a dvourozměrná mapa ziskovosti cena × dny.

**Optimalizace:**
- výběr nákupní sestavy v rámci investičního rozpočtu (`portfolio.py`): úplné prohledání
  kombinací 0–2 kusů u všech typů strojů včetně varianty „nic nekupovat“; poptávka po typu
  stroje se mezi kusy dělí (nejvýše 30 dní na kus), náklady jsou individuální podle dat.

**Rozhodovací pravidla:** doporučený stroj je ten s vyšším očekávaným měsíčním ziskem.
Doplňkové bodové hodnocení (očekávaný zisk, riziko ztráty, VaR, návratnost; při shodě bez bodu)
a textové zdůvodnění se generují **pevně danými pravidly** ze stejných vypočtených metrik.
Aplikace nevyužívá generativní umělou inteligenci.

**Plánování zakázek** (`scheduling.py`): celočíselná optimalizace (SciPy `milp`, řešič HiGHS)
hledá rozvrh s nejnižším celkovým penále za zpoždění, tedy s nejvyšším ziskem po penále,
při respektování počtu kusů každého stroje v každém dni. Pokud řešič optimum neprokáže,
aplikace to uvede. Pro srovnání zobrazuje i původní plán podle rozhodovacího skóre.

## Ukázkový scénář rozhodování

*Všechny hodnoty v tomto scénáři jsou fiktivní a pocházejí z demonstračních dat aplikace.*

Provozovatel zvažuje, zda dokoupit **mini bagr**, nebo **pásový dumper**.

1. **Nastavení vstupů:** v sekci Investice v panelu „Nastavení investičního scénáře“ vybere oba stroje, zadá fixní měsíční náklad,
   zvolí scénář trhu (např. realistický) a odhad poptávky v dnech za měsíc.
2. **Shrnutí:** aplikace porovná očekávaný měsíční zisk, pravděpodobnost ztráty a bod zvratu
   a označí relativně výhodnější variantu. Pokud jsou při daném nastavení obě varianty ztrátové,
   výslovně na to upozorní.
3. **Monte Carlo a Riziko:** uživatel vidí rozptyl možných výsledků. Nejen průměr, ale i nepříznivý
   scénář (P10), hodnotu v riziku a pravděpodobnost ztrátového měsíce.
4. **Citlivost:** zjistí, které vstupy mají na výsledek největší vliv, při jakém vytížení se stroj
   dostává do zisku a jaká je minimální denní cena pronájmu.
5. **Nákup strojů:** aplikace projde všechny kombinace strojů v rámci rozpočtu a ukáže sestavu
   s nejvyšším očekávaným provozním ziskem, její návratnost a zisk podle výše rozpočtu.
   Pokud se nákup nevyplatí, doporučí nic nekupovat.
6. **Plán zakázek:** uvidí rozvrh zakázek s nejnižším penále, srovnání s původním plánem
   a zakázku s největším čistým přínosem pro plán.
7. **Kapacita strojů:** v kalendáři zjistí, kolik kusů je který den volných, kde se pronájem
   nebo servis střetává se zakázkou a zda lze přijmout nový požadavek na stroj.
8. **Report:** stáhne investiční report ve formátu DOCX jako podklad pro další diskusi.

Výstupem je **strukturovaný podklad pro rozhodnutí**, ne automatické investiční doporučení.

## Technologie

| Oblast | Nástroje |
|---|---|
| Jazyk | Python |
| Webová aplikace | Streamlit, nasazení na Streamlit Community Cloud |
| Práce s daty | pandas, NumPy |
| Simulace a optimalizace | NumPy (náhodná rozdělení s pevným základem), SciPy (`milp`) |
| Vizualizace | Plotly |
| Reporting | python-docx |
| Správa kódu a zabezpečení | Git, GitHub (větve, pull requesty), Streamlit secrets |

## Technické informace

### Struktura projektu

```
├── app.py                           # Streamlit aplikace: přihlášení, menu, vstupy a obsah sekcí; výpočty volá z modulů
├── navigace.py                      # registr sekcí levého menu (pořadí, názvy, ikony, skupiny, adresy)
├── stranky/                         # značkové soubory sekcí pro st.navigation (obsah vykresluje app.py)
├── investice.py                     # investiční výpočty záložek 1–4 (Monte Carlo, metriky, LP/DP, citlivost)
├── report_docx.py                   # investiční report DOCX (export z první záložky)
├── formatting.py                    # české formátování čísel (bez závislosti na Streamlitu)
├── ui_theme.py                      # vzhled: barvy, české formátování čísel, karty, šablona grafů
├── auth.py                          # přihlašování: jedno heslo nebo uživatelské účty (hash hesel)
├── datamodel.py                     # společná datová vrstva: načtení, kontrola, převod pro výpočty
├── economics.py                     # ekonomický modul (E2.1): náklady flotily, příspěvky, celkový výsledek
├── capacity.py                      # kapacitní plánování strojů (E2.3A): dostupnost, konflikty, požadavky
├── planner.py                       # plánovací a rozhodovací model (E2.5a)
├── planner_view.py                  # záložka Rozhodovací plánovač (E2.5b): formulář → požadavek, výsledek → karty
├── econ_formula.py                 # příspěvek jako vzorec s neznámými náklady a porovnání variant (E2.5c); základ verdiktu plánovače
├── capacity_view.py                 # kapacitní kalendář (E2.3B): převod výsledků capacity.py na mřížku
├── portfolio.py                     # výběr nákupní sestavy strojů v rámci rozpočtu
├── scheduling.py                    # plán zakázek (celočíselná optimalizace)
├── tools/hash_password.py           # vytvoření hashe hesla pro secrets (heslo se zadává skrytě)
├── tools/regression/                # regresní nástroje pro úpravy vzhledu (snímek, porovnání, pixelová kontrola)
├── docs/SLOVNIK.md                  # slovník pojmů aplikace (druhy zisku, kapacita, rezervace, stavy údajů)
├── tests/                           # testy (python -m pytest tests)
│   └── fixtures/legacy/             # původní tabulky pro kontrolu zpětné kompatibility
├── assets/decision_engine_mark.svg  # vlastní neutrální symbol aplikace
├── data/
│   ├── model/                       # společný datový model (fiktivní ukázka), popis v model/README.md
│   └── *.xlsx                       # rozšířený fiktivní datový model (aplikace ho zatím nenačítá)
├── .streamlit/config.toml           # barevné téma aplikace
├── .streamlit/secrets.toml.example  # šablona pro nastavení přístupového hesla
├── requirements.txt
└── README.md
```

### Demonstrační data

Složka `data/` obsahuje **výhradně fiktivní hodnoty** vytvořené pro testování a ukázku:
typy strojů, denní sazby, pořizovací ceny, náklady, parametry poptávky a poruchovosti a vzorové zakázky.
Neobsahuje žádná skutečná data žádné společnosti ani osobní údaje.

Data jsou uložená ve **společném datovém modelu** (`data/model/`) pro pronájem strojů i vlastní
zakázky: typy strojů, jednotlivé kusy flotily, údržba, rezervace pronájmů, poptávka, zakázky,
jejich požadavky na stroje, obsluha a společná nastavení. Náklady jsou rozdělené na vlastnické,
provozní fixní a variabilní, aby se později nezapočítaly dvakrát. Při načtení se data
zkontrolují (typy, rozsahy, vazby mezi tabulkami); při chybě aplikace výpočty nespustí.
Podrobný popis tabulek a jejich budoucího využití je v [`data/model/README.md`](data/model/README.md).

**Ekonomický modul** (`economics.py`, etapa E2.1) je připravený výpočtový základ pro společnou
ekonomiku pronájmu i vlastních zakázek: fixní náklady flotily (vlastnické, provozní, nerozdělené),
variabilní náklady podle využití, příspěvek z pronájmu, příspěvek ze zakázky, podklady pro ušlý
zisk a celkový výsledek. Každý výsledek nese stav **úplný / neúplný (co chybí) / nelze spočítat**
a příznak fiktivních dat. Aplikace modul zatím nepoužívá – zobrazené výsledky se nemění.

**Kapacitní modul** (`capacity.py`, etapa E2.3A) počítá, kolik kusů každého typu stroje je
v jednotlivých dnech k dispozici a čím jsou obsazené (pronájem, vlastní zakázka, servis), hledá
konflikty (překryv rezervací, pronájem × zakázka, konflikt se servisem, překročení kapacity typu)
a vyhodnotí nový požadavek na typ stroje v termínu: **realizovatelné / nerealizovatelné / nelze
určit**, s problémovými dny a příčinami. Nic nepřesouvá ani neurčuje přednost. Chybějící údaje se
nikdy nepovažují za volno. Zobrazuje ho záložka **Kapacita strojů** (E2.3B, `capacity_view.py`),
která přebírá termíny zakázek z plánu záložky 6 a sama nic nepočítá. Pronájmy (`FIKT-…`)
a servisní odstávky v ukázkových datech jsou fiktivní a obsahují dva záměrné konflikty;
ostatní záložky tyto tabulky nepoužívají.

### Spuštění lokálně

```bash
git clone <URL tohoto repozitáře>
cd decision-engine-pujcovna
pip install -r requirements.txt

cp .streamlit/secrets.toml.example .streamlit/secrets.toml
# v souboru .streamlit/secrets.toml nastavte vlastní heslo

streamlit run app.py
```

### Nastavení přístupového hesla

Heslo se načítá výhradně ze Streamlit secrets (klíč `password`) a není součástí zdrojového kódu.

- **Lokálně:** soubor `.streamlit/secrets.toml`. Je uvedený v `.gitignore`, takže se do repozitáře nedostane.
- **Streamlit Community Cloud:** nastavení aplikace → *Settings* → *Secrets* → řádek `password = "..."`.

### Uživatelské účty (volitelné)

Bez sekce `[users]` v secrets funguje přihlášení jedním heslem jako dosud. Po přidání sekce
`[users]` se přihlašuje **uživatelským jménem a heslem** a každý účet má vlastní údaje:

```toml
password = "…"                 # stávající heslo – zůstává

[users.lukas]
name = "Lukáš"
role = "admin"                 # admin | tester | owner
active = true
legacy_password = true         # přihlašuje se stávajícím heslem `password`

[users.honza]
name = "Honza Podroužek"
role = "tester"
active = true
password_hash = "pbkdf2_sha256$600000$…"
```

- V secrets se ukládá jen **hash hesla** (PBKDF2-HMAC-SHA256 se solí), ne heslo samotné.
  Hash vytvoří administrátor v aplikaci (*Správa účtů*) nebo příkazem `python tools/hash_password.py`.
- Účet bez `active = true` se nepřihlásí; deaktivace nebo změna hesla jednoho účtu ostatní neovlivní
  a přihlášeného uživatele odhlásí při jeho další akci.
- Nástroj *Správa účtů* vidí jen role `admin`. Veřejná registrace neexistuje.
- V postranním panelu je jméno přihlášeného uživatele a tlačítko **Odhlásit se**.

### Vizuální identita

Aplikace používá vlastní neutrální symbol **Decision Engine** (jednoduchý sloupcový graf se symbolem
rozhodnutí) uložený v `assets/decision_engine_mark.svg`. Nepoužívá logo žádné společnosti.

### Online verze

Pilotní verze běží na adrese <https://decision-engine-pujcovna.streamlit.app>.
Aplikace je **chráněna heslem**. Přístup je poskytován pouze pro účely ověření prototypu.

## Omezení současného prototypu

- Model pracuje s **fiktivními demonstračními daty** a nebyl ověřen na skutečných provozních datech.
- Model používá zjednodušené předpoklady. Například zadaný fixní náklad je stejný pro oba porovnávané stroje
  a model nepočítá se sezónností poptávky.
- Simulace nepoužívá pevný seed, takže se výsledky při každém přepočtu mírně liší.
- Celá aplikace je zatím v jednom souboru `app.py`. Výpočty nejsou pokryty automatickými testy.
- Výstupy slouží jako **podklad pro odbornou diskusi**, ne jako investiční doporučení.

## Plánované další kroky

1. **Odborná validace** předpokladů, vstupních parametrů a metodiky.
2. **Uživatelské ověření**: předvedení prototypu budoucím uživatelům a sběr zpětné vazby.
3. **Případné ověření na reálných datech**, pouze se souhlasem dotčené společnosti
   a s odpovídajícím zabezpečením. Reálná data nebudou ukládána do veřejného repozitáře.
4. **Technický rozvoj:**
   - rozdělení kódu na moduly (výpočty / uživatelské rozhraní / exporty),
   - automatické testy výpočtů a opakovatelné simulace,
   - možnost nahrát vlastní data a ukládat projekty.

## Přínos z hlediska datové a business analytiky

Projekt propojuje programování, datovou analýzu a business analytiku na konkrétním podnikatelském problému:

- **převod obchodní otázky na kvantitativní model:** investice, riziko, kapacity,
- **stochastické modelování nejistoty** pomocí Monte Carlo simulace a práce s rizikovými metrikami (VaR, percentily),
- **optimalizační úlohy:** lineární a diskrétní optimalizace, výběr portfolia v rámci rozpočtu,
- **citlivostní analýza** jako nástroj pro identifikaci klíčových faktorů ziskovosti,
- **interaktivní aplikace pro netechnické uživatele**, kde se výsledky okamžitě přepočítávají podle předpokladů,
- **automatizovaný manažerský reporting** ve formátu DOCX,
- **profesionální vývojový postup:** Git, pull requesty a bezpečná správa přístupových údajů.
