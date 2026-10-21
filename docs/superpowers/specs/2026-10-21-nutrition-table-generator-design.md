# Technická špecifikácia dizajnu: Generátor a exportér tabuliek výživových údajov

**Dátum:** 21. október 2026  
**Cieľ:** Samostatný, ľahko použiteľný offline nástroj na tvorbu, detailné formátovanie a export výživových tabuliek (pre etikety potravín a výživových doplnkov v programe Avery).

---

## 1. Prehľad a cieľ projektu

Používateľ (otec) potrebuje jednoduchý a rýchly nástroj, ktorý nahradí chýbajúce exportné funkcie Excelu / LibreOffice:
- Umožní tvoriť a upravovať nutričné tabuľky pre obaly produktov (napr. sušené ovocie, doplnky výživy).
- Poskytne plnú slobodu orámovania (rôzne hrúbky čiar, linky na úrovni jednotlivých riadkov, odsadenia podpoložiek).
- Umožní priamy import existujúcich `.xlsx` zošitov a automatické rozdelenie hárkov do záložiek produktov.
- Umožní okamžité kopírovanie do schránky (`Ctrl+V` v Avery) a export do čistého vektorového **SVG** (neobmedzene zväčšiteľné na plátno/obal bez straty ostrosti) a vysokorozlíšeného **PNG / JPG** (300 až 600 DPI pre tlač).
- Funguje offline priamo v prehliadači bez nutnosti akejkoľvek inštalácie.

---

## 2. Architektúra a technologický zásobník

1. **Používateľské rozhranie (Frontend):**
   - Samostatná webová aplikácia (HTML5, Vanilla CSS, ES2022 JavaScript).
   - Žiadna inštalácia runtime (Node.js/Python), žiadne servery: otvára sa priamo dvojklikom na `index.html` v Chrome, Edge alebo Firefoxe.
2. **Knižnica pre Excel:**
   - Lokálna samostatná knižnica SheetJS (`xlsx.full.min.js`), uložená v projekte pre 100 % offline fungovanie.
   - Spracovanie `.xlsx` / `.xls` / `.csv` súborov priamo v prehliadači cez File API a Drag & Drop.
3. **Grafický engine (Vektorové SVG & Canvas):**
   - **Vektorový renderer:** Tabuľka v náhľade sa generuje priamo ako čisté `<svg>`, čo zaručuje 100 % zhodu medzi obrazovkou a stiahnutým SVG súborom.
   - **Rasterizér (Canvas):** Pre export do PNG/JPG a schránky sa SVG prenesie do `<canvas>` elementu s násobným rozlíšením (1x, 2x, 4x = 300 DPI, 6x = 600 DPI) a vyhladenou typografiou.
4. **Perzistencia dát:**
   - Lokálne úložisko prehliadača (`localStorage`) s automatickým ukladaním pri každej zmene.
   - Možnosť exportu/importu celej databázy produktov do jedného záložného JSON súboru.

---

## 3. Funkčná špecifikácia

### 3.1 Správa produktov a záložiek (Sheets / Tabs)
- Horný panel so záložkami reprezentujúcimi jednotlivé produkty (napr. *Soft-Dried-Papaya*, *Dried-Mango*, *Soft-Dried-PassionFruit*).
- Možnosti záložiek:
  - Pridať nový produkt (`+`).
  - Premenovať záložku dvojklikom alebo ikonou.
  - Duplikovať produkt (ideálne pri podobných receptúrach).
  - Zmazať produkt s potvrdením.
- **Import z Excelu:**
  - Tlačidlo „Nahrať Excel súbor“ alebo potiahnutie súboru myšou na plochu.
  - Aplikácia načíta všetky hárky zo zošita a pre každý vytvorí samostatnú záložku.
  - Automaticky rozpozná hlavičku a riadky nutričných hodnôt.

### 3.2 Dátový model tabuľky a riadkov
Každý produkt obsahuje:
- **Hlavička tabuľky:**
  - Názov stĺpca hodnôt (predvolené: `Výživové údaje na`).
  - Merná jednotka / vzťažná hodnota (predvolené: `100g`, voliteľne `100ml`, `1 porcia`).
- **Položky (riadky):**
  - Štandardná štruktúra podľa legislatívy EÚ:
    1. *Energetická hodnota* (viacriadková hodnota, napr. `1372kJ / \n 328kcal`).
    2. *Tuky* (napr. `0,12g`).
    3. *– z toho nasýtené mastné kyseliny* (odsadené, napr. `0,08g`).
    4. *Sacharidy* (napr. `80,9g`).
    5. *– z toho cukry* (odsadené, napr. `70,4g`).
    6. *Bielkoviny* (napr. `1,54g`).
    7. *Soľ* (napr. `0,23g`).
  - Možnosť pridať ďalšie riadky (Vláknina, Vitamíny, Minerály, Sodík, vlastný text).
  - Možnosť preusporiadania riadkov (hore / dole).

### 3.3 Formátovanie a orámovanie (úroveň Excelu)
- **Globálne nastavenie tabuľky:**
  - Šírka tabuľky (nastaviteľná v mm / px alebo posuvníkom).
  - Pomer šírky stĺpca s popisom voči stĺpcu s hodnotami.
  - Výber písma (Calibri, Arial, Inter, Roboto, Segoe UI).
  - Základná veľkosť písma (pt / px).
  - Výška riadkov / vnútorný padding buniek.
- **Nastavenie orámovania (Borders):**
  - **Vonkajší rám (Outer Border):** vypnutý / tenký (1px) / stredný (2px) / hrubý (3px).
  - **Zvislá oddeľovacia čiara:** zapnutá / vypnutá / hrúbka.
  - **Orámovanie konkrétneho riadku:**
    - Spodná čiara riadku: *Žiadna* / *Tenká* / *Hrubá* / *Dvojitá*.
    - Špecifické odsadenie textu vľavo (Indent: 0 až 30 px).
    - Zvýraznenie: *Tučné písmo (Bold)* pre názov alebo hodnotu.
    - Špecifická výška riadku (napr. dvojitá výška pre Energetickú hodnotu).

### 3.4 Exportný systém a schránka
- **Kopírovanie do schránky:**
  - Tlačidlo „Kopírovať obrázok do schránky“ – vygeneruje rastrový obrázok vo vysokej kvalite a vloží ho do systémovej schránky cez `navigator.clipboard.write`.
  - V softvéri Avery stačí stlačiť `Ctrl + V`.
- **Export do SVG (Vektor):**
  - Stiahnutie `.svg` súboru s presnými krivkami a textami.
  - Avery a akýkoľvek grafický softvér zachováva 100 % vektorovú kvalitu pri ľubovoľnom zväčšení na etikete.
- **Export do PNG / JPG:**
  - Voľba rozlíšenia:
    - 1x (náhľad, 72 DPI)
    - 2x (retina, 150 DPI)
    - 4x (tlačová kvalita, 300 DPI – odporúčaná pre Avery)
    - 6x (ultra ostrá tlač, 600 DPI)
  - Voľba pozadia: Priehľadné (Transparent PNG) alebo Biele (Solid White JPG/PNG).
  - Možnosť nastaviť bezpečnostný okraj (margin okolo tabuľky 0 až 20 px).

---

## 4. Používateľské rozhranie a dizajn (UI/UX)

- **Vizuálny štýl:** Moderný, profesionálny, čistý vizuál inšpirovaný pravidlami `frontend-design` a `ui-ux-pro-max`. Žiadne zbytočné dekorácie, vysoký kontrast, veľké a zreteľné ovládacie prvky vhodné pre pohodlné používanie bez namáhania zraku.
- **Rozloženie obrazovky:**
  - **Hlavička:** Logo/názov, tlačidlá na import Excelu, zálohu a rýchly export/kopírovanie.
  - **Záložkový pás:** Zoznam produktov s jednoduchým prepínaním.
  - **Ľavá sekcia (Editor):** Políčka pre editáciu hodnôt, pridávanie/mazanie riadkov a záložky pre formátovanie orámovania a písma.
  - **Pravá sekcia (Živý náhľad):** Vektorový náhľad tabuľky v reálnom čase s možnosťou prepínania podkladu (biela vs. šachovnica) a posuvníkom lupy (zoom).

---

## 5. Štruktúra súborov projektu

```
Oco - Vyživové hodnoty/
├── index.html                   # Hlavná aplikácia (štruktúra, editor, náhľad)
├── styles.css                   # Dizajnový systém, layout, tlačidlá a ovládanie
├── app.js                       # Logika, SVG engine, export do PNG/JPG, schránka, Excel parser
├── libs/
│   └── xlsx.full.min.js         # Lokálna SheetJS knižnica pre offline import Excelu
├── docs/
│   └── superpowers/specs/       # Technická dokumentácia a špecifikácie
└── README.md                    # Používateľská príručka v slovenčine pre otca
```

---

## 6. Verifikácia a testovanie
- Overenie importu viac-hárkového `.xlsx` súboru.
- Overenie presnosti orámovania v SVG a exportovanom PNG.
- Overenie funkcie kopírovania do schránky a vloženia do externých aplikácií.
- Overenie správnosti diakritiky a špeciálnych znakov v slovenčine (kJ, kcal, %, g, ť, ž, š, č).
