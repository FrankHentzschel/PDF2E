# PDF2E 1.0.1 – PDF-Rechnungen als E-Rechnung (ZUGFeRD)

Kundenneutrales Werkzeug, Stand 26.10.2026. Erster Kunde ist die Krug-Bräu. Kundenspezifisch sind nur die Stammdaten-Blätter und das Rechnungslayout (Einstellung „Layout“, Einlese-Vorlage `modParserBSI` mit Kundenvariante als Parameter). Bis die Smart-Tools-Bibliothek eintrifft, läuft der **Stub-Modus**. In diesem Modus entsteht keine gültige E-Rechnung.

## Inhalt

| Datei | Zweck |
| --- | --- |
| `PDF2E.xlsx` | Arbeitsmappe ohne Makros; erzeugt mit `python vorlage/build_xlsx.py PDF2E.xlsx` oder als Anhang eines Releases |
| `src/modPDF2E.bas` | Name, Version, Splash beim Start, Info-Fenster, Schaltflächen anlegen (`PDF2E_Einrichten`) |
| `src/modHilfen.bas` | Einstellungen, Feldzugriff, Zahlen und Datumswerte, RegExp, Protokoll |
| `src/modPdfText.bas` | Ruft `pdftotext -layout` auf und liest den Text ein |
| `src/modParserBSI.bas` | Einlese-Vorlage für BSI-Rechnungen, Variante je Kunde über die Einstellung „Layout-Variante“ |
| `src/modEinlesen.bas` | Makro **RechnungEinlesen** |
| `src/modPruefen.bas` | Prüfung: Pflichtangaben, Menge × Preis, Summenabgleich mit dem PDF |
| `src/modZugferd.bas` | Makro **ZugferdErzeugen**, ruft den Adapter auf |
| `src/modAdapterStub.bas` | Stub: schreibt alle Übergabedaten nach `<Ziel>.daten.xml` |
| `src/modAdapterSmartTools.bas` | Vorbereitet für Smart Tools, mit vollständiger Feldzuordnung |
| `src/modEntwicklung.bas` | Nur Entwicklung: `PDF2E_Exportieren` schreibt alle Module für Git nach `src/` |
| `src/frmInfo_Code.txt` | Code für das Splash- und Info-Fenster (UserForm) |
| `src/ThisWorkbook_Code.txt` | Code für „DieseArbeitsmappe“ (Splash beim Öffnen) |

## Einrichtung (ca. 15 Minuten)

1. **pdftotext:** Lade auf https://www.xpdfreader.com/download.html die *Xpdf command line tools* für Windows und lege `pdftotext.exe` (64-Bit) z. B. unter `C:\Tools\xpdf\` ab.
2. **Mappe vorbereiten:** `PDF2E.xlsx` öffnen und als **Excel-Arbeitsmappe mit Makros (.xlsm)** speichern.
3. **Module importieren:** VBA-Editor mit Alt+F11 öffnen. Über *Datei > Datei importieren* alle `.bas`-Dateien aus `src/` einlesen.
4. **Info-Fenster anlegen:** *Einfügen > UserForm*. Im Eigenschaftenfenster bei (Name) `frmInfo` eintragen. Dann mit Rechtsklick auf `frmInfo` *Code anzeigen* wählen und den gesamten Inhalt von `frmInfo_Code.txt` einfügen. Auf der Form selbst nichts platzieren, die Elemente entstehen zur Laufzeit.
5. **Start-Code:** Doppelklick auf *DieseArbeitsmappe* und den Inhalt von `ThisWorkbook_Code.txt` einfügen.
6. **Kompilieren:** *Debuggen > Kompilieren VBAProject*. Es dürfen keine Fehler erscheinen.
7. **Einstellungen** im gleichnamigen Blatt ausfüllen:
   - Pfad zu pdftotext, Eingangs- und Ausgabeordner
   - **Firma**, Kontakt, Anwendungslogo (oben) und Firmenlogo (unten), jeweils BMP, JPG oder GIF, kein PNG
   - den Adapter auf `STUB` lassen
8. **Schaltflächen:** Mit Alt+F8 einmalig `PDF2E_Einrichten` ausführen. Das Makro legt auf „Start“ und „Rechnung“ die Schaltflächen *PDF einlesen*, *Prüfen*, *E-Rechnung erzeugen* und *Info* an.
9. Speichern, schließen und neu öffnen. Der Splash erscheint 3 Sekunden lang, ein Klick schließt ihn sofort.

## Versionsnummern

PDF2E verwendet das übliche Schema HAUPT.NEBEN.KORREKTUR (Semantic Versioning):

| Änderung | Beispiel |
| --- | --- |
| Fehlerbehebung | 1.0.0 → 1.0.1 |
| neue Funktion | 1.0.1 → 1.1.0 |
| inkompatible Änderung, z. B. neues Stammdatenformat | 1.4.2 → 2.0.0 |

Die Version steht als Konstante `TOOL_VERSION` in `modPDF2E`. Jede Änderung kommt zusätzlich ins Blatt „Versionshistorie“.

## Test mit den Mustern

Die Auswertungslogik ist an allen vier Mustern gegengeprüft. Warenwert, Pfand, Nettosumme, MwSt und Rechnungsbetrag stimmen jeweils mit dem PDF überein:

| Rechnung | Empfänger | Positionen | Pfand | Brutto |
| --- | --- | --- | --- | --- |
| 2603976 | Markgrafen, Lieferort Stadtsteinach | 13 + 2 Aktionen | 201,50 + 15,50 | 1.029,53 |
| 2603894 | Markgrafen, Lieferort Herzogenaurach | 3 | −18,60 (Rückgabe) | 141,19 |
| 2603816 | REWE Markt GmbH (RE-GLN 4388096000002) | 3 | 24,80 | 489,74 |
| 2603587 | REWE Markt GmbH, Lieferort Forchheim | 4 | −49,60 (Rückgabe) | 176,30 |

Im Ordner `tests/pdftotext` liegen die pdftotext-Ausgaben dieser Muster. Dateien mit `_windows` stammen aus pdftotext unter Windows (Xpdf), die übrigen aus Linux (Poppler). Die Windows-Version verteilt manche Leergutzeilen auf zwei Zeilen; der Parser fügt sie spaltengenau zusammen.

Im Stub-Modus legt **E-Rechnung erzeugen** zwei Dateien an: `<RechNr>_ZUGFeRD_STUB.pdf`, eine unveränderte Kopie, und daneben `.daten.xml` mit allen Übergabewerten.

## Noch nicht getestet

- Der VBA-Code ist noch nicht in Excel gelaufen. Gegengeprüft habe ich nur die Auswertungslogik, und zwar mit Poppler-pdftotext unter Linux.
- `VBScript.RegExp`: Microsoft kündigt die schrittweise Abschaltung von VBScript an. Ob und wann das RegExp in VBA betrifft, ist offen.

## Bekannte Grenzen

- Bei mehreren Lieferscheinen pro Rechnung wird nur der erste übernommen. Das Protokoll gibt dann eine Warnung aus.
- Bei mehreren MwSt-Sätzen muss der Satz je Position von Hand eingetragen werden.
- Als Einheit ist C62 (Stück) für Kasten und Fass eingestellt. Das lässt sich in den Einstellungen ändern.
