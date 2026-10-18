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
| `maintenance.csv` | – | plánované odstávky kusu (začátek, počet dní) – zatím prázdné | `unit_id` → kus flotily |
| `rental_bookings.csv` | `booking_id` | konkrétní rezervace pronájmu (od–do, počet kusů, cena, s obsluhou, `confirmed` / `tentative`) – zatím prázdné | `type_id` → typ stroje |
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

Modul počítá nad `fleet_units`, `maintenance`, `rental_bookings`, `jobs` a `job_requirements`;
aplikace ho zatím nepoužívá.

- **Dny:** celé dny, začátek i konec **včetně** (pronájem `start_day`–`end_day`, servis
  `start_day` + `days` − 1, zakázka začátek + `duration_days` − 1). Navazující termín (den po
  skončení) je v pořádku, předání ve stejný den je konflikt. Volitelné datum prvního dne plánu
  převede čísla dní na kalendářní data (správně přes konec měsíce i roku).
- **Kapacita typu v daný den** = kusy ve `fleet_units` k dispozici ten den (`owned`, `leased`;
  `ordered` až od `available_from_day`) − kusy v servisu. Počítá se v kusech, ne ve strojodnech.
- **Stavy:** `confirmed` blokuje kapacitu; `tentative` a `opportunity` se vykazují zvlášť jako
  riziko; `cancelled` (zatím není v datech) se ignoruje.
- **Rezervace pronájmu** jsou jen na úrovni typu – konkrétní kus se jim nepřiřazuje.
- **Potvrzené zakázky** nemají v datech pevný termín: bez plánu se berou jako možné obsazení
  v okně nejdřívější začátek – termín, a výsledek tam je „nelze určit“, pokud by na tom záleželo.
- **Neplatné nebo neúplné záznamy** se vypíšou a dotčený typ je „nelze určit“, nikdy „volno“.

Pro spolehlivý kalendář chybí: datum prvního dne plánu, skutečné rezervace pronájmů, plánovaný
servis, pevné termíny zakázek, stav `cancelled`, dny na převoz / předání mezi nasazeními
a případně přiřazení rezervací ke konkrétním kusům.

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
