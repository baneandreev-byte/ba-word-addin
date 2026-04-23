# CHANGELOG - BiroA Word Add-in

## 2026-04-23 — FIX: XML parser, redosled polja, učitavanje iz skrivenog XML-a

### Problem

* BiroA Polja plugin nije učitavao sva polja iz skrivenog XML-a u dokumentu
* Redosled polja se menjao nakon ponovnog otvaranja dokumenta
* Template funkcije (load/save) nikada nisu nalazile sačuvane template

### Uzroci i popravke

#### ba-word-addin (BiroA Polja) — taskpane.js

| #   | Funkcija | Uzrok | Popravka |
| --- | --- | --- | --- |
| 1   | `loadStateFromDocument()` | Regex `[^/>]` u XML parseru isključuje `/` karakter — polja sa vrednostima poput "02/1", "P+Pk" ili putanjama se gube | Zamenjen sa **DOMParser** + regex fallback sa `[\s\S]` |
| 2   | `loadStateFromDocument()` | Iteracija svih customXmlParts nesigurna | Dodat **getByNamespace()** kao primarni metod sa fallback-om |
| 3   | `buildStateXml()` | Nema `id` ni `order` atributa u XML-u | Dodat `id` i `order` atribut za svaki `<item>` |
| 4   | `saveStateToDocument()` | Bez error handling-a | Dodat getByNamespace + try/catch sa fallback-om |
| 5   | `deleteSavedStateFromDocument()` | Isti problem | Isti fix |
| 6   | `loadTemplatesFromDocument()` | `namespaceUri` se NE UČITAVA pre `filter()` — template se nikad ne nalaze | Dodat `p.load("namespaceUri")` + `await context.sync()` pre filtera |
| 7   | `saveTemplatesToDocument()` | Isti bug | Isti fix |
| 8   | —   | —   | Dodate helper funkcije: `parseStateXml()`, `parseStateXmlRegexFallback()` |

#### ba-word-addin-admin (BiroA Admin) — admin.js

| #   | Funkcija | Uzrok | Popravka |
| --- | --- | --- | --- |
| 1   | `upsertClientStateXmlPart()` | Ne piše `id`/`order` u XML za klijenta | Dodati `id` i `order` atributi |
| 2   | `scanDocument()` | Isti regex bug `[^/>]` kao u klijentskom pluginu | Zamenjen sa DOMParser + regex fallback |

### Kako oba plugina sada rade sa XML-om

**XML format (namespace: `http://biroa.rs/word-addin/state`):**

    <state xmlns="http://biroa.rs/word-addin/state">
      <item id="uuid" order="0" field="Investitor" value="Tanja" type="text" format="text:auto"/>
      <item id="uuid" order="1" field="Objekat" value="Stambeno-poslovni" type="text" format="text:auto"/>
    </state>

**Admin** piše XML → **Polja** čita XML — isti format, isti parser (DOMParser), isti atributi.

### Fajlovi promenjeni

* `ba-word-addin/taskpane.js` — 220 novih linija, 62 uklonjene
* `ba-word-addin-admin/admin.js` — 34 nove linije, 17 uklonjenih
