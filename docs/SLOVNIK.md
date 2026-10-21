# Slovník pojmů Decision Engine

Interní slovník pro texty aplikace. Při každé úpravě rozhraní používejte tyto pojmy ve stejném významu.
Definice popisují **současné výpočty** – pokud se výpočet změní, upravte i slovník.

Všechna data v aplikaci jsou ukázková (fiktivní).

## Ekonomické výsledky

Aplikace počítá několik různých „zisků“. **Nejsou zaměnitelné** a nesmějí se vydávat jeden za druhý.

| Pojem v aplikaci | Kde | Jak se počítá | Co nezahrnuje |
|---|---|---|---|
| **Očekávaný měsíční zisk stroje** (simulace) | záložky 1–3 | Průměr simulovaných měsíců: tržby − fixní náklad (z postranního panelu) − servis. Simulace mění počet dní pronájmu podle scénáře trhu, poruchy, cenu (±) a servisní náklady. | Společné náklady půjčovny (prostory, mzdy, administrativa), daně. |
| **Zisk při N dnech** (rychlý odhad) | záložka 1 (rozbalovací blok), záložka 4 (vliv ±10 %, minimální cena) | Ceníková cena × N dní − fixní náklad − průměrný servis. Bez simulace. | Poruchy, kolísání ceny a servisu, rozpětí dní podle scénáře – proto se liší od simulace. |
| **Tržby − fixní náklad** | záložka 4 (graf „Kdy stroj pokryje fixní náklad“, mapa ziskovosti) | Denní cena × dny − fixní náklad. | Servis a poruchy. Není to zisk. |
| **Očekávaný měsíční provozní zisk** (sestavy) | záložka 5 | Tržby z průměrné poptávky po odečtení poruch − fixní náklady jednotlivých strojů **z dat** − průměrný servis. | Fixní náklad z postranního panelu (nepoužívá se), společné náklady, úroky, daně, sezónnost. |
| **Zisk zakázky**, **zisk zakázek po penále** | záložka 6 | Tržba zakázky − přímé náklady zakázky (z dat); po penále navíc − penále za zpoždění. | Fixní náklady strojů a společné náklady. Data neuvádějí, zda přímé náklady obsahují náklady vlastních strojů. |
| **Příspěvek** | záložka 8 | Pronájem: tržba − variabilní náklad pronájmu × strojodny. Zakázka: tržba − přímé náklady − variabilní náklady strojů (jen pokud nejsou v přímých nákladech) − penále. | Fixní náklady flotily a režie. **Není to čistý zisk.** |
| **Čistý zisk firmy** | aplikace ho nepočítá | – | Chybí režie, daně a úplné náklady. Tento pojem v aplikaci nepoužívejte. |

## Náklady

- **Fixní náklad** – platí se každý měsíc bez ohledu na počet dní pronájmu (např. splátka, pojištění).
  Záložky 1–4 berou jednu ukázkovou hodnotu z postranního panelu pro oba stroje. Záložka 5 bere
  individuální hodnoty strojů z ukázkových dat.
- **Servisní náklad** – průměrný měsíční servis stroje z ukázkových dat (v simulaci kolísá).
- **Variabilní náklad** – náklad za den, kdy stroj pracuje (pronájem nebo zakázka). V ukázkových
  datech chybí, proto je ekonomika v záložce 8 často „Záleží na nákladech“.
- **Přímé náklady zakázky** – náklady konkrétní zakázky z dat.
- **Ukázkové náklady** – všechny náklady v aplikaci jsou zatím ukázkové, nejde o ověřené náklady firmy.

## Riziko a návratnost

- **Šance ztrátového měsíce** – podíl simulovaných měsíců, kdy stroj skončí ve ztrátě.
- **Velmi nepříznivý měsíc (VaR 5 %)** – hranice 5 % nejhorších simulovaných měsíců.
- **Bod zvratu** – počet dní pronájmu měsíčně, při kterém tržby pokryjí fixní náklad (bez servisu a poruch).
- **Minimální denní cena** – cena, při které se při zadaném počtu dní pokryje fixní náklad i průměrný servis.
- **Prostá návratnost** – pořizovací cena ÷ měsíční zisk; bez úroků a společných nákladů.

## Kapacita a vytížení

- **Strojoden** – 1 kus stroje na 1 den. Jednotka kapacity a využití (dříve také „kusoden“).
- **Kapacita dne** – kusy ve flotile k dispozici ten den minus kusy v servisu.
- **Vytížení zakázkami** (záložka 6) – dny strojů na vlastních zakázkách ÷ (kusy × délka plánu).
  Pronájmy a servis nezahrnuje.
- **Obsazenost** (záložka 7, kalendář) – pronájmy, zakázky i servis.
- **Využití** (záložka 7, přehled) – pronájmy a vlastní zakázky; servis se nepočítá.
- **Konflikt** – potvrzené nároky (pronájmy, zakázky z plánu, servis) převyšují počet kusů.
- **Kolize mezi zakázkami plánu** (záložka 6) – překrytí vlastních zakázek navzájem. Plán zakázek
  pronájmy a servis nezná, proto může mít 0 kolizí, i když kalendář ukazuje konflikty.
- **Nelze určit** – kapacitu nelze spočítat (chybí nebo je neplatný údaj); nikdy to neznamená volno.

## Rezervace a závazky

- **Potvrzená rezervace / závazek** – blokuje kapacitu.
- **Předběžná rezervace** – kapacitu zatím neblokuje.
  - Kalendář (záložka 7) ji vyznačí zvlášť jako riziko; při ověření požadavku ji lze volbou započítat.
  - Rozhodovací plánovač (záložka 8) ji bere jako požadavek, který soutěží o volnou kapacitu stejně
    jako nové požadavky.
  - Nikdy se automaticky neruší.

## Kvalita údajů a typy výsledků

- **Skutečná / zadaná data** – údaj, který zadal uživatel nebo je v datech vyplněný.
- **Ukázková data** – fiktivní hodnoty pro předvedení aplikace.
- **Chybějící údaj (nezadáno)** – prázdná hodnota. **Nikdy se nepočítá jako 0 Kč.**
- **Skutečně zadaná nula** – výslovně zadaná 0; počítá se jako 0 a aplikace ji tak označí.
- **Neuplatňuje se** – nezadaný údaj, který na daný výsledek nemá vliv.
- **Hypotetický výpočet „Co kdyby“** – dosazení zkusmé hodnoty nezadaného nákladu. Nemění
  doporučení, data ani rezervace a nic se neukládá.
- **Podmíněné doporučení** (štítek „⇄ Záleží na nákladech“) – nejlepší varianta závisí na výši
  nezadaného nákladu; aplikace ukáže hranice zlomu.
- **Hranice zlomu** – výše nákladu, při které vycházejí dvě varianty stejně.
- **Nelze rozhodnout / nelze určit** – chybí údaje, výsledek nelze spolehlivě určit.

## Štítky v Rozhodovacím plánovači

| Štítek | Význam |
|---|---|
| Provoz: ✓ Vejde se / ✕ Nevejde se / ? Provoz nelze určit | Kapacita po potvrzených závazcích a servisu |
| ★ Doporučeno | Varianta je prokazatelně nejlepší i v nejhorším případě |
| Méně výhodné | Lepší je jiná varianta |
| Nevyplatí se | Každá varianta, která se vejde, je ztrátová |
| = Rovnocenné | Varianty vycházejí stejně |
| ⇄ Záleží na nákladech | Podmíněné doporučení – rozhoduje nezadaný náklad |
| ? Nelze rozhodnout | Chybí údaje |
| Bez porovnání | Požadavek nemá soupeře – doporučení se nevydává |
