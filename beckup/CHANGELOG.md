# CHANGELOG - BiroA Word Add-in

## [1.0.0.29] - 2025-02-07 - V29

### ✨ NEW FEATURE - DRAG & DROP REORDERING
- **Drag & Drop**: Prevuci redove da promeniš njihov redosled
  - ⋮⋮ handle pored svakog reda
  - HTML5 Native Drag & Drop API (bez dependency-ja)
  - Vizuelni feedback: plava linija pokazuje gde će se dropovati
  - Dragged red: 50% opacity + plava isprekidana ivica
  - Smooth animacije i transitions

### 🎯 UX IMPROVEMENTS
- **Smart selection tracking**: Selektovani red ostaje selektovan i nakon drag-a
- **Focus retention**: Cursor ostaje u input-u ako si kucao
- **Status feedback**: "Polje 'naziv' premešteno."
- **Auto-save**: Novi redosled se automatski čuva u XML

### 💻 TECHNICAL
- Grid layout ažuriran: 32px za drag handle kolonu
- ~150 linija JS koda (drag handlers)
- ~70 linija CSS (drag stilovi)
- Touch support included (HTML5 API)

### 📝 UI UPDATES
- Help text updated: objašnjava drag & drop
- Table header: dodata prazna kolona za drag handle
- CSS klase: `.drag-handle`, `.dragging`, `.drag-over`

### 🧪 TESTING
- Testiraj: Dodaj 3 reda, prevuci srednji na vrh
- Testiraj: Drag sa selektovanim redom - ostaje selektovan
- Testiraj: Drag sa fokusom u input-u - zadržava focus

---

## [1.0.0.28] - 2025-02-07 - V28

### 🔧 FIXED - KONAČNO REŠENJE ZA DELETE
- **KRITIČNO**: DELETE dugme konačno radi! Novi pristup "Insert Outside, Then Delete"
  - Umesto `cc.delete(true)` koji ne radi pouzdano
  - Novi algoritam: umetni tekst VAN content control-a, pa obriši CC
  - `cc.getRange(Word.RangeLocation.after).insertText(finalText)`
  - Zatim `cc.delete(false)` - briše samo CC, tekst je već izvučen
  - **100% pouzdan** u Desktop i Online Word-u

### 🎯 KAKO RADI
1. Učita se tekst iz CC-a (ili iz tabele ako je popunjeno)
2. Tekst se umetne NAKON CC-a (van njega)
3. Sinhronizacija
4. CC se briše bez sadržaja (tekst je već izvučen van)

### 🧪 TESTIRANJE
- ✅ Popunjena polja: tekst ostaje, CC nestaju
- ✅ Prazna polja: {PLACEHOLDER} ostaje, CC nestaju
- ✅ Mixed scenario: sve radi kako treba

### 📝 MEMO
- Dodata dokumentacija za sledeći feature: SharePoint Template Picker
- Spremno za implementaciju u V29

---

## [1.0.0.27] - 2025-02-07 - V27

### 🔧 FIXED
- **KRITIČNO**: Klik na input polje sada selektuje red za ubacivanje
  - Dodati click event listeners na oba input polja (POLJE i ODGOVOR)
  - Dodati focus event listeners za Tab navigaciju
  - Implementirana focus retention - cursor ostaje u input-u nakon re-render
  - Koristi se `e.stopPropagation()` da spreči dupli event
  - Uslovni re-render samo kada je potrebno (`if (selectedRowIndex !== idx)`)

### 🎯 USER EXPERIENCE
- Korisnik sada može da klikne BILO GDE u redu i taj red će biti selektovan
- Tab navigacija kroz input polja automatski selektuje red
- Kucanje u input ne gubi focus (cursor ostaje u polju)
- Vizuelni feedback - selektovani red dobija plavu pozadinu

### 🧪 TESTING
- Testiraj: Dodaj 3 reda, klikni na ODGOVOR u 2. redu, klikni UBACI POLJE → ubacuje se polje iz 2. reda
- Testiraj: Tab navigacija kroz input polja → red se automatski selektuje
- Testiraj: Kucanje u input polje zadržava focus → cursor ne skače

---

## [1.0.0.26] - 2025-02-07 - V26

### 🔧 FIXED
- **KRITIČNO**: Dugme OBRIŠI sada pravilno briše content controls iz dokumenta
  - Implementiran dva-prolaza pristup: prvo umetni tekst, pa obriši CC
  - Promenjen parametar `cc.delete(false)` → `cc.delete(true)` da zadrži tekst
  - Dodato bolje sinhronizovanje između operacija
  - Dodato error handling sa console logging za debugging
  - XML state se pravilno briše iz dokumenta

### 🧪 TESTIRANJE
- Potrebno testirati u Word Desktop i Word Online
- Testirati scenario: dodaj polja → popuni → obriši
- Testirati scenario: dodaj polja → obriši (bez popunjavanja)

---

## [1.0.0.25] - 2025-02-07 - V25

### ✨ FEATURES
- Modal dijalog za podešavanje tipa i formata polja
- Dugme za edit (⚙) u svakom redu tabele
- Tri tipa polja: tekst, datum, broj
- Napredni formati za svaki tip:
  - Tekst: VELIKA/mala slova, Naslov
  - Datum: dd.mm.yyyy, yyyy-mm-dd, MMMM.yyyy, dd.MMMM.yyyy, danas
  - Broj: ceo broj, 2 decimale, RSD, €, $

### 🎨 UI/UX
- Moderna kartica-based tabela
- Radio buttons za tip polja u modalu
- Dropdown za format sa hint tekstom
- Status bar za feedback korisnicima

---

## [1.0.0.23] - 2025-02-07 - V23

### ✨ FEATURES
- Osnovna funkcionalnost Word add-in-a
- UBACI POLJE: ubacuje content control sa placeholder-om
- POPUNI: popunjava sva polja iz tabele
- OČISTI: vraća {POLJE} placeholder (čuva vrednosti u tabeli)
- OBRIŠI: briše content controls (sa confirm dijalogom) - **BUG: nije radilo**

### 📦 DATA
- CSV Export/Import funkcionalnost
- XML state sačuvan u Custom XML Parts dokumenta
- Automatsko čuvanje pri izmeni tabele

### 🔧 TECHNICAL
- Content Controls sa tag sistemom: `BA_FIELD|key=...|type=...|format=...`
- Parsiranje i formatiranje vrednosti prema tipu
- Serbian locale support (dd.mm.yyyy format, meseci na srpskom)

---

## POZNATI BUGOVI (TODO za V27)

### 🐛 BUG #1: Datum formatiranje
- `date:dd.mm.yyyy` i `date:yyyy-mm-dd` ne formatiraju unetu vrednost
- Trenutno samo vraćaju originalni string
- **FIX**: Dodati parsiranje i konverziju datuma

### 🐛 BUG #2: CSV Import gubi tip/format
- Pri importu CSV-a se sve postavlja na `type: "text"`, `format: "text:auto"`
- **FIX**: Proširiti CSV format sa dodatnim kolonama za tip i format

### 🐛 BUG #3: Auto-save performance
- `saveStateToDocument()` poziva se pri svakom keystroke-u
- Može biti sporo na većim dokumentima
- **FIX**: Dodati debounce (npr. 500ms nakon poslednje izmene)

---

## ROADMAP - Sledeće verzije

### V27 - Datum fix
- [ ] Implementirati parsiranje za `date:dd.mm.yyyy`
- [ ] Implementirati parsiranje za `date:yyyy-mm-dd`
- [ ] Testirati sa različitim input formatima

### V28 - CSV poboljšanja
- [ ] Export: dodati kolonu za tip i format
- [ ] Import: čitati tip i format iz CSV-a
- [ ] Backward compatibility sa starim CSV formatom

### V29 - Performance
- [ ] Debounce za auto-save (500ms)
- [ ] Show loading indicator za duge operacije
- [ ] Optimizacija Word API calls

### V30 - UX poboljšanja
- [ ] Preview formatiranja u modalu (live preview)
- [ ] Validacija datuma/brojeva pre formatiranja
- [ ] Search/filter u tabeli polja
- [ ] Bulk operacije (copy/paste između redova)
- [ ] Drag-and-drop reorder redova u tabeli

---

## TEHNIČKI INFO

### Tehnologije
- Office.js API
- Word JavaScript API
- Vanilla JavaScript (bez framework-a)
- Custom XML Parts za storage
- Content Controls za polja

### Browser support
- Microsoft Edge (Chromium)
- Chrome, Firefox (za razvoj)
- Word Desktop (Windows/Mac)
- Word Online

### Deployment
- GitHub Pages hosting
- Manifest sideloading za razvoj
- AppSource submission (budućnost)
