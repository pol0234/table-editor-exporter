# Generátor výživových tabuliek pre obaly a tlač (Avery)

Jednoduchá, rýchla a **100 % offline webová aplikácia** vytvorená na mieru pre tvorbu, detailnú úpravu a export výživových tabuliek na etikety a obaly potravín (napr. sušené ovocie, doplnky výživy).

---

## 🚀 Ako aplikáciu spustiť (jediným kliknutím)

1. Dvakrát kliknite na súbor **`Spustiť-Aplikáciu.bat`** (alebo priamo otvorte **`index.html`**).
2. Aplikácia sa okamžite otvorí v prehliadači (Chrome, Edge, Firefox).
3. **Nevyžaduje žiadny internet, žiadnu inštaláciu ani žiadne administrátorské práva.**

---

## ✨ Hlavné funkcie a možnosti

### 1. Práca so záložkami produktov (ako v Exceli)
- Na hornom paneli vidíte záložky pre jednotlivé produkty (*Soft-Dried-Papaya*, *Dried-Mango*, *Soft-Dried-PassionFruit*, *Mixed-Dried-TropicalFruits*).
- Kliknutím na záložku prepnete produkt.
- Tlačidlom `+ Nový produkt` pridáte ďalší produkt do zoznamu.
- Ikonou ✏️ môžete produkt premenovať, ikonou 📄 ho duplikovať a krížikom ✕ zmazať.
- **Všetky zmeny sa automaticky ukladajú do pamäte prehliadača**, takže pri zatvorení okna o nič neprídete.

### 2. Import existujúceho Excelu (.xlsx / .xls)
- Kliknite na tlačidlo **„Nahrať Excel (.xlsx)“** vpravo hore, alebo **jednoducho potiahnite súbor z plochy myšou priamo do okna aplikácie**.
- Aplikácia automaticky prečíta všetky hárky zo zošita a pre každý vytvorí vlastnú záložku s presnými hodnotami a odsadeniami!

### 3. Voľnosť orámovania a riadkovania ako v Exceli
- **Pre celú tabuľku:**
  - Hrúbka vonkajšieho rámu (žiadny, tenký, stredný 1.5px, hrubý 2.5px, extra hrubý).
  - Zvislá deliaca čiara medzi názvom a hodnotami (zapnutá / skrytá).
  - Spodná čiara pod hlavičkou (žiadna, tenká, hrubá, dvojitá).
  - Nastaviteľná šírka tabuľky, pomer stĺpcov a výška riadkov.
  - Výber písma (Calibri, Arial, Segoe UI, Times New Roman) a plynulá zmena veľkosti textu.
- **Pre každý riadok samostatne:**
  - **Spodná linka riadku:** voľba *Bez čiary*, *Tenká čiara*, *Hrubá čiara*, *Dvojitá čiara*.
  - **Odsadenie (Indent):** tlačidlo *Odsadiť* pre podpoložky ako *„z toho cukry“* alebo *„z toho nasýtené mastné kyseliny“*.
  - **Tučné písmo (Bold):** tlačidlo *B*.
  - **Dvojriadkové texty:** v políčku hodnoty (napr. pri energetickej hodnote) stačí stlačiť `Enter` pre zalomenie riadku:
    ```
    1372kJ /
    328kcal
    ```

### 4. Export a vkladanie do Avery (1-klikový prenos)
- **Zelené tlačidlo „Kopírovať do schránky (Avery)“:**
  - Jedným klikom vygeneruje ostrý obrázok priamo do schránky Windows.
  - Otvorte softvér Avery a stlačte klávesovú skratku **`Ctrl + V`**. Hotovo!
- **Tlačidlo „Stiahnuť SVG (Vektor)“:**
  - Stiahne 100 % čisté vektorové SVG. V Avery ho môžete roztiahnuť na akúkoľvek veľkosť (aj veľké plátno) a písmo zostane dokonale ostré bez rozmazania.
- **Tlačidlá „Stiahnuť PNG“ a „Stiahnuť JPG“:**
  - Uloží obrázok v tlačovej kvalite **300 DPI** (alebo až 600 DPI podľa nastavenia).
  - Možnosť voľby priehľadného pozadia (Transparent PNG) alebo bieleho pozadia.

---

## 🧪 Automatizované testy

Všetky výpočty, SVG generovanie, parsovanie Excelu a exporty sú pokryté testami v `tests/`:
```bash
node --test tests/*.test.js
```
Všetkých 11 testov prechádza na 100 %.
