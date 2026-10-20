# Společný datový model (etapa E1)

> **Všechna data v této složce jsou fiktivní ukázka.** Nejde o skutečná firemní data žádné
> společnosti. Prázdné buňky znamenají „zatím nezadáno“ – hodnoty doplní majitel půjčovny,
> model si je nevymýšlí.

Jedna flotila strojů slouží dvěma způsobům využití: **pronájmu zákazníkům** a **vlastním
zakázkám**. Tabulky proto popisují stroje, jejich kusy, kapacitu, náklady, pronájmy
i zakázky na jednom místě. Načtení a kontrolu zajišťuje `datamodel.py`.

## Tabulky

| Soubor | Klíč | Obsah | Vazby |
|---|---|---|---|
| `machine_types.csv` | `type_id` | typ stroje: název, cena pronájmu/den, pořizovací cena, náklady, poruchovost | – |
| `fleet_units.csv` | `unit_id` | jednotlivé kusy flotily, stav (`owned` / `leased` / `ordered`), dostupnost od–do (den) | `type_id` → typ stroje |
| `maintenance.csv` | – | plánované odstávky kusu (začátek, počet dní) – 3 fiktivní ukázkové odstávky | `unit_id` → kus flotily |
| `rental_bookings.csv` | `booking_id` | konkrétní rezervace pronájmu (od–do, počet kusů, cena, s obsluhou, `confirmed` / `tentative`) – 7 fiktivních ukázkových rezervací `FIKT-…` | `type_id` → typ stroje |
| `rental_demand.csv` | `type_id` + `period` | očekávaná poptávka po pronájmu (dny za období, rozptyl, zdroj) – **odhad, ne závazek** | `type_id` → typ stroje |
| `jobs.csv` | `job_id` | vlastní zakázky: stav (`confirmed` / `tentative` / `opportunity`), pravděpodobnost získání, tržba, přímé náklady, délka, nejdřívější začátek, termín, penále | – |
| `job_requirements.csv` | `job_id` + druh + zdroj | co zakázka potřebuje: typ stroje nebo roli obsluhy, počet, pořadí | `job_id` → zakázka, `resource_id` → typ stroje / role |
| `operator_roles.csv` | `role_id` | role obsluhy, počet lidí, denní náklad, které stroje obsluhuje – zatím prázdné | `machine_types` → typy strojů |
| `settings.csv` | `key` | společná nastavení (původ dat, měna, dny v měsíci, horizont plánu, režie) | – |

Dny (`start_day`, `deadline_day` …) se počítají od 1. dne plánu, stejně jako v plánu zakázek.

## Jednotné identifikátory

- Typ stroje má technický kód `type_id` (velká písmena bez diakritiky, např. `MINI_BAGR`)
  a jedinečný český název `name` (např. „Mini bagr“). Vazby mezi tabulkami používají **jen kód**.
- Kus flotily: `TYP-01`, `TYP-02` … Zakázka: `J001` …

## Náklady – oddělené skupiny (žádné dvojí započtení)

| Skupina | Sloupce | Kdy vzniká | Kam patří |
|---|---|---|---|
| Vlastnické | `ownership_cost_month` | každý měsíc, i když stroj stojí (splátka, odpis, pojištění) | flotila |
| Provozní fixní | `fixed_operating_cost_month`, `service_cost_month_mean` / `_std` | každý měsíc za kus | flotila |
| Nerozdělené fixní | `fixed_cost_month_unsplit` | dočasně, dokud není známé rozdělení | flotila |
| Variabilní – pronájem | `variable_cost_rental_day` | jen za dny pronájmu | pronájem |
| Variabilní – zakázka | `variable_cost_job_day` | jen za dny na vlastní zakázce | zakázka |
| Přímé náklady zakázky | `jobs.direct_cost` | materiál, práce apod. **bez vlastních strojů** | zakázka |
| Režie půjčovny | `settings: overhead_cost_month` | společné náklady firmy | samostatně, nerozpočítává se |

Kontrola dat odmítne typ stroje, který má zadané nerozdělené i rozdělené fixní náklady
zároveň (dvojí započtení), nebo nemá zadané žádné.

## Kontroly (`datamodel.validate`)

- povinné sloupce a hodnoty, číselné typy, nezáporné hodnoty, pravděpodobnosti 0–1, povolené stavy,
- jedinečné klíče a názvy strojů,
- vazby mezi tabulkami (neexistující typ, kus, zakázka, role),
- každá zakázka potřebuje aspoň jeden stroj,
- potvrzené rezervace nepřekročí počet kusů v žádném dni,
- upozornění: zakázka nesplnitelná v termínu, nepotvrzená zakázka bez pravděpodobnosti,
  chybějící rozdělení fixních nákladů.

Při chybě aplikace výpočty nespustí a vypíše srozumitelné hlášení.

## Zpětná kompatibilita

Aplikace zatím počítá stejně jako dříve. `legacy_machines`, `legacy_jobs` a `fleet_capacity`
převádějí nový model do dřívějšího tvaru (`machines_demo.csv`, `jobs_demo.csv`, počty kusů).
Původní soubory jsou uložené v `tests/fixtures/legacy/` a testy hlídají, že převod dává
přesně stejné tabulky.

## Převod z dřívějších dat

| Dříve | Nově |
|---|---|
| `machines_demo.daily_price` | `machine_types.rental_price_day` |
| `machines_demo.monthly_fixed_cost` | `machine_types.fixed_cost_month_unsplit` (rozdělení zatím neznámé) |
| `machines_demo.service_cost_mean` / `_std` | `machine_types.service_cost_month_mean` / `_std` |
| `machines_demo.demand_days_mean` / `_std` | `rental_demand.expected_days` / `std_days` |
| `jobs_demo.cost_base` | `jobs.direct_cost` |
| `jobs_demo.required_machines` („A;B“) | řádky v `job_requirements` (kód typu + pořadí) |
| počty kusů natvrdo v `app.py` | řádky v `fleet_units` |
| — | `jobs.status = confirmed` (všechny ukázkové zakázky se plánují jako dosud) |

## Ekonomický modul (`economics.py`, E2.1)

Modul počítá nad těmito tabulkami, ale aplikace ho zatím nepoužívá.

- **Náklady flotily (fixní)** – za každý vlastněný nebo leasingový kus a měsíc, jen jednou za flotilu.
  Celek je přesný i s nerozdělenými náklady; rozpad na vlastnické / provozní je do doplnění neúplný.
- **Příspěvek z pronájmu** = cena/den × dny − variabilní náklad pronájmu × dny.
- **Příspěvek ze zakázky** = tržba − přímé náklady − variabilní náklady strojů × strojodny − penále;
  náklady strojů se odečtou jen tehdy, když nejsou v přímých nákladech.
- **Celkový výsledek** = příspěvky z pronájmů a zakázek − fixní náklady flotily − režie.
- **Ušlý zisk** = dny pronájmu, o které firma přijde kvůli zakázce, × příspěvek dne pronájmu.

Stavy údajů: „0“ = skutečná nula; prázdná buňka = nezadáno → výsledek je **neúplný** a vypíše,
co chybí; text nebo záporné číslo = neplatné → **nelze spočítat**. Výsledky z dat označených
jako fiktivní nesou příznak „fiktivní data“.

Volitelné nastavení `direct_cost_includes_machines` (`ano` / `ne`) v `settings.csv` určuje, zda přímé
náklady zakázek už obsahují vlastní stroje. Dokud není zadané, je příspěvek zakázky neúplný.

## Kapacitní modul (`capacity.py`, E2.3A)

Modul počítá nad `fleet_units`, `maintenance`, `rental_bookings`, `jobs` a `job_requirements`.
V aplikaci ho zobrazuje záložka **Kapacita strojů** (E2.3B) – termíny zakázek bere z plánu záložky 6.

**Fiktivní ukázková data (E2.3B):** pronájmy `FIKT-P01` až `FIKT-P07` (5 potvrzených, 2 předběžné)
a tři servisní odstávky. Dva konflikty jsou záměrné, aby kalendář ukázal, jak je hlásí:
`FIKT-P05` (silniční válec, dny 15–19) proti zakázce J006 a servis kolového bagru (dny 11–12)
proti zakázce J005. Záložky 1–6 tabulky pronájmů a servisu nepoužívají, jejich výsledky se nemění.

- **Dny:** celé dny, začátek i konec **včetně** (pronájem `start_day`–`end_day`, servis
  `start_day` + `days` − 1, zakázka začátek + `duration_days` − 1). Navazující termín (den po
  skončení) je v pořádku, předání ve stejný den je konflikt. Volitelné datum prvního dne plánu
  převede čísla dní na kalendářní data (správně přes konec měsíce i roku).
- **Kapacita typu v daný den** = kusy ve `fleet_units` k dispozici ten den (`owned`, `leased`;
  `ordered` až od `available_from_day`) − kusy v servisu (každý kus nejvýše jednou za den, i při
  překrývajících se servisních záznamech). Počítá se v kusech, ne ve strojodnech.
- **Stavy:** `confirmed` blokuje kapacitu; `tentative` a `opportunity` se vykazují zvlášť jako
  riziko; `cancelled` (zatím není v datech) se ignoruje.
- **Rezervace pronájmu** jsou jen na úrovni typu – konkrétní kus se jim nepřiřazuje.
- **Potvrzené zakázky** nemají v datech pevný termín: bez plánu se berou jako možné obsazení
  kdykoli od nejdřívějšího začátku – i po termínu dokončení, protože plán v záložce 6 dovoluje
  zpoždění s penále. Výsledek tam je „nelze určit“, pokud by na tom záleželo. Omezení do termínu
  jen na výslovné vyžádání (`assume_deadlines_met=True`, předpoklad). Spolehlivé termíny dodá
  plán zakázek (`job_starts`).
- **Neplatné nebo neúplné záznamy** (i s neznámým stavem, např. překlepem) se vypíšou a dotčený
  typ je „nelze určit“, nikdy „volno“. Neplatný předběžný záznam kapacitu neblokuje.
- **Pronájem nebo zakázka konkrétního kusu** v den, kdy kus není ve flotile (např. po vrácení
  leasingu), se nezahodí: ubírá kapacitu typu a vypíše se jako konflikt.

Pro spolehlivý kalendář chybí: datum prvního dne plánu, skutečné rezervace pronájmů, plánovaný
servis, plán zakázek, který bere ohled na pronájmy a servis, stav `cancelled`, dny na převoz / předání mezi nasazeními
a případně přiřazení rezervací ke konkrétním kusům.

## Přehled využití (E2.4)

Záložka **Kapacita strojů → Přehled využití** sčítá denní kapacitu z kapacitního modulu (stejná
data, filtry a plán zakázek jako kalendář):

- **Kapacita** = součet kusů k dispozici v jednotlivých dnech (kusodny).
- Každý den a typ stroje patří do jedné skupiny: běžný den → pronájem, vlastní zakázky, servis
  a volno (při neúplných údajích „nelze určit“ místo volna); konfliktní den (víc nároků než kusů)
  → celá kapacita jako **konflikt** – bez rozhodnutí o přednosti nelze určit, kdo stroj má.
  Je-li počet kusů nejistý, přetížení není jistý konflikt – den je „nelze určit“ (stejně jako v kalendáři).
  Nároky nad počet kusů = **poptávka nad kapacitu**, do kapacity se nepočítají.
- **Využití** = (pronájmy + vlastní zakázky) ÷ kapacita; servis se nepočítá. V konfliktním dni je
  jistě produktivních nejméně kusy − servis, nejvýše min(kusy, pronájmy + zakázky) → případné
  rozpětí. Nikdy nepřesáhne 100 %.
- **Předběžné rezervace** kapacitu neblokují a do využití se nepočítají.
- Hodnocení (orientační, nepodložené firemními daty): vysoké ≥ 80 %, nízké < 40 %; zasahuje-li
  rozpětí do dvou kategorií, „nelze určit“.
- **Obsazenost** v kalendáři zahrnuje servis; **využití** v přehledu servis nezahrnuje.

## Plánovací a rozhodovací model (`planner.py`, E2.5a)

Logika nad `capacity.py` a `economics.py`; v aplikaci ji zobrazuje záložka **Rozhodovací plánovač**
(E2.5b, `planner_view.py`). Štítek požadavku: **Doporučeno** jen ve variantě, kterou plánovač
ekonomicky doporučil při soupeření; samostatný požadavek bez soupeře je **Proveditelné** (ekonomika
zvlášť); **Neproveditelné** – kapacitu drží závazky nebo servis; **Nelze rozhodnout** – chybí údaje,
nebo se ze soupeřících požadavků vejde jen část a ekonomika chybí. Existující zakázka slouží jen jako
šablona nového požadavku. Nové požadavky nesmějí nést označení zakázek ani rezervací z dat.

- **Potvrzené závazky** (potvrzené pronájmy, servis, potvrzené zakázky v aktuálním plánu záložky 6)
  blokují kapacitu. Jejich vzájemný konflikt jde **k ručnímu řešení**; u zakázky se jen ověří,
  zda ji okno dovoluje posunout bez penále.
- **Předběžné rezervace a nové požadavky** jsou rovnocenné a posuzují se na zbylé kapacitě.
  Předběžná rezervace zůstává vždy viditelná a nikdy se automaticky neruší.
- **Provozně**: proveditelný / neproveditelný (co ho blokuje) / nelze určit (chybí údaje).
  Soupeřící požadavky → kombinace, které se vejdou (ověřené kapacitním modelem).
- **Ekonomicky** (E2.5c krok 3): porovnávají se jen kapacitně proveditelné kombinace soupeřících
  požadavků a vždy i **„nepřijmout nic“** (přírůstkový příspěvek přesně 0 Kč). Příspěvky jsou vzorce
  z `econ_formula.py` nad `economics.py` (bez fixních nákladů a bez ušlého zisku navíc); chybějící
  sazba zůstává neznámou, společné neznámé se v porovnání vyruší, oba scénáře
  `direct_cost_includes_machines` se respektují. Verdikt skupiny:
  - **doporučení** – jediná varianta je prokazatelně (ostře) lepší než všechny ostatní při jakýchkoli
    hodnotách neznámých sazeb a v obou scénářích; je-li jí „nic“, výsledek je **„žádný se nevyplatí“**
    (všechny varianty jsou prokazatelně ztrátové) – nikdy doporučení nejmenší ztráty,
  - **rovnocenné** – nejlepší varianty mají prokazatelně shodný výsledek (varianta s přesně 0 Kč je
    rovnocenná s „nic“, ne lepší),
  - **podmíněně lepší** – vítěz závisí na neznámé sazbě; vypíší se **hranice zlomu** (např. „do 2 750 Kč: A;
    2 750–9 000 Kč: B; nad 9 000 Kč: nepřijmout nic“). Patří sem i varianta, která je při nulové sazbě
    shodná s jinou a při kladné lepší – bez jednoznačného doporučení,
  - **nelze rozhodnout** – rozhoduje chybějící základní údaj (tržba, přímé náklady, cena, sazba penále)
    nebo neplatný údaj.
  Varianty prokazatelně horší než jiná varianta se z výkladu vyřadí. Nic se neruší – ani předběžné rezervace.
- Označení požadavků musí být jedinečná (i vůči předběžným rezervacím v datech).
- Překrývající se servisní záznamy téhož kusu: kus je v servisu nejvýše jednou za den (dva různé
  kusy zvlášť); překryv není konflikt využití, plánovač na něj upozorní jako na možnou duplicitu dat.
- **Alternativy**: jen posun zakázky v okně bez penále. Jiný stroj ani externí zapůjčení se
  nenabízí – chybí údaje o zaměnitelnosti strojů a o externích cenách a dostupnosti.

Pro ekonomická doporučení zatím chybí: variabilní náklady za den pronájmu a zakázky, nastavení
`direct_cost_includes_machines`, případně ceny a pružnost termínů nových požadavků.

## Příspěvek jako vzorec s neznámými náklady (`econ_formula.py`, E2.5c krok 2)

Výpočetní vrstva pro porovnání variant; od kroku 3 z ní vychází verdikt plánovače. Příspěvek varianty je
vzorec `známá částka − Σ (jednotky využití × neznámá sazba)`; chybějící údaj zůstává proměnnou,
nikdy se nenahradí nulou. Koeficienty se zjišťují přímo z `economics.py` (žádný druhý výpočet).

- skutečná nula („0“) není neznámá; prázdná buňka je neznámá; neplatný údaj → vzorec nelze sestavit,
- chybí-li `direct_cost_includes_machines`, počítají se scénáře „ano“ a „ne“ a varianty se porovnávají
  vždy ve stejném scénáři; náklad stroje na zakázce se odečte nejvýše jednou,
- v rozdílu dvou variant se společné neznámé vyruší → **jistý** výsledek; jinak **podmíněný**
  (podmínka, u jedné sazby hranice zlomu), nebo **nelze rozhodnout**, když rozhoduje chybějící základní
  údaj požadavku (tržba, přímé náklady, cena, sazba penále) nebo neplatný údaj,
- varianta „nepřijmout nic“ má vzorec přesně 0 Kč.

## Obnovení dat v běžící aplikaci

Aplikace drží načtená data v mezipaměti. Klíčem je otisk obsahu všech CSV souborů v této složce
(`datamodel.data_fingerprint`, SHA-256, < 1 ms), takže se po jakékoli změně dat – i po novém
nasazení na Streamlit Cloud bez restartu – načtou znovu. Stejně tak klíč obsahuje otisk
zdrojového kódu modulů, které uložené výpočty volají (`datamodel`, `scheduling`, `capacity`,
`capacity_view`), aby se po nasazení nového kódu nepoužily staré výsledky. Záložka **Kapacita strojů** ukazuje
pod kalendářem počet načtených pronájmů a servisních odstávek a zkrácenou verzi dat.

## Budoucí využití

- **E2 – jednotná ekonomika:** jeden zdroj nákladů a cen pro všechny záložky.
- **E3 – kapacita:** plán zakázek bude respektovat kusy flotily, údržbu a potvrzené rezervace.
- **E4–E6:** poptávka po pronájmu, obsluha a nepotvrzené zakázky v plánovači.
- **E7:** simulace nákupu stroje = nový řádek ve `fleet_units` (stav `ordered`).

## Údaje, které musí doplnit majitel

Rozdělení fixních nákladů (vlastnické / provozní), variabilní náklady za den pronájmu a zakázky,
zda přímé náklady zakázek obsahují stroje, skutečné počty kusů a odstávky, rezervace pronájmů,
historie poptávky, obsluha (počty, kvalifikace, náklady), stav a pravděpodobnost zakázek,
horizont plánu, pracovní dny a společná režie.
