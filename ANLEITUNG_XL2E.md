# XL2E 1.0.0 – Excel-Rechnungen als E-Rechnung (ZUGFeRD)

Stand 04.11.2026. XL2E ist das Gegenstück zu PDF2E für Brauereien, deren BSI-Software die Rechnung als Excel-Datei ausgibt (Blatt `PA400_00`). Erster Kunde ist die Privatbrauerei Fiedler. Der Ablauf ist derselbe wie bei PDF2E: Rechnung einlesen, prüfen, E-Rechnung erzeugen, E-Mail-Entwurf.

Die Excel-Rechnung hat keinen Briefkopf, weil sie für vorgedrucktes Papier gedacht ist. XL2E erzeugt deshalb selbst ein PDF aus einer Kopie des Rechnungsblatts. Absender, Logo, Registerangaben, Steuernummer und Bankverbindung kommen aus den Stammdaten in Kopf- und Fußzeile. In dieses PDF bettet Smart Tools die E-Rechnung ein. Die Excel-Originaldatei bleibt unverändert.

## Inhalt

| Datei | Zweck |
| --- | --- |
| `XL2E.xlsx` | Arbeitsmappe ohne Makros, mit Platzhaltern in den Stammdaten |
| `vba/modProdukt.bas` | **nur XL2E:** Name, Version, Dateiauswahl, PDF mit Briefkopf erzeugen |
| `vba/modParserXL.bas` | **nur XL2E:** Einlese-Vorlage für das Blatt `PA400_00` |
| übrige Dateien in `vba/` | gemeinsamer Kern mit PDF2E (Einlesen, Prüfen, Smart Tools, E-Mail, Info-Fenster) |
| `logo/` | Logo der Brauerei für den Briefkopf, Entwicklerlogo für das Info-Fenster |

## Einrichtung (ca. 15 Minuten)

1. **Mappe vorbereiten:** `XL2E.xlsx` öffnen und als **Excel-Arbeitsmappe mit Makros (.xlsm)** speichern.
2. **Smart-Tools-Verweis:** VBA-Editor (Alt+F11) > *Extras > Verweise* > **SmartTools E-Rechnung-Bibliothek** anhaken. Ohne diesen Verweis lässt sich das Projekt nicht kompilieren.
3. **Module importieren:** *Datei > Datei importieren*, alle `.bas`-Dateien aus `vba/`.
4. **Info-Fenster:** *Einfügen > UserForm*, (Name) = `frmInfo`. Rechtsklick > *Code anzeigen*, Inhalt von `frmInfo_Code.txt` einfügen.
5. **Start-Code:** Doppelklick auf *DieseArbeitsmappe*, Inhalt von `ThisWorkbook_Code.txt` einfügen.
6. *Debuggen > Kompilieren VBAProject*. Es dürfen keine Fehler erscheinen.
7. **Stammdaten** pflegen. Alle mit PLATZHALTER markierten Werte durch die echten ersetzen:
   - Blatt „Brauerei“: USt-IdNr., Steuernummer, Registernummer, GLN falls vorhanden
   - Blatt „Bankkonten“: echte IBAN und BIC, Standardkonto mit `x`
   - Blatt „Zahlungskonditionen“: Standard mit `x`
8. **Einstellungen:**
   - Eingangsordner (Excel-Rechnungen) und Ausgabeordner
   - **Briefkopf-Logo**: Pfad zum Logo, z. B. `C:\XL2E\fiedler_logo_hoch.png`
   - **Sammelkundennummern**: vorbelegt mit `12100`
   - E-Rechnung-Adapter: `SMARTTOOLS` (oder `STUB` für einen Testlauf ohne E-Rechnung)
   - Firma, Kontakt und Logos für das Info-Fenster wie bei PDF2E
9. Alt+F8 > `XL2E_Einrichten` einmal ausführen. Es legt die Schaltflächen *Excel einlesen*, *Prüfen*, *E-Rechnung erzeugen*, *E-Mail erstellen* und *Info* an.
10. Speichern, schließen, neu öffnen.

Solange Platzhalter in den Stammdaten stehen, bricht die Prüfung mit Smart Tools ab. Im Testmodus STUB erscheint nur eine Warnung.

## Was XL2E aus der Excel-Rechnung liest

| Angabe | Quelle im Blatt `PA400_00` |
| --- | --- |
| Kundennummer, Rechnungsnummer, Datum | Zeile unter „Kun-Nr.“ (Spalten I und K) |
| Anschrift | Spalte D oberhalb der Kopfzeile |
| Lieferschein und Datum | Zeile „Lieferschein: … vom …“ |
| Lieferort | Zeile „Lieferort: …“, falls vorhanden |
| Positionen | ab „Art-Nr“: Art-Nr., Bezeichnung, Ka/Fa, Liter, Preis, Betrag |
| Pfand | „LEERGUT-KONTO“: geliefert und zurück je Gebinde, mit Pfandsatz |
| Summen zum Abgleich | Warenwert, Zwischensumme inkl. Pfand, MwSt-Satz und -Betrag, Rechnungsbetrag |

Besonderheiten:

- **Fassbier** wird je Hektoliter berechnet. XL2E erkennt das am Betrag (Liter / 100 × Preis) und übergibt die Menge in hl (Einheit HLT), z. B. 0,3 hl bei einem 30-Liter-Fass.
- **Pfand** steht wie von REWE verlangt getrennt: Lieferung mit positiver, Rücknahme mit negativer Menge. Einstellung „Pfand aufteilen“ = Nein fasst sie zur Differenz zusammen.
- **Zeilen „Leergutrückgabe:“** im Positionsteil sind nur Hinweise. Das Pfand kommt ausschließlich aus dem Leergutkonto.
- **Nachlässe** (negativer Betrag ohne Artikelnummer) werden als Abschlag übernommen.

## Sammelkundennummern

Unter der Kundennummer 12100 rechnet die Brauerei viele verschiedene Kunden ab. Bei Nummern aus der Einstellung **Sammelkundennummern** legt XL2E im Blatt „Empfaenger“ je Kunde einen eigenen Eintrag an. Der Schlüssel lautet `Kundennr. / Name / PLZ`. E-Mail-Adresse, Bankkonto und Zahlungskondition lassen sich so je Kunde hinterlegen.

## E-Mail-Versand

Wie bei PDF2E (siehe dort): Outlook-Entwurf mit der E-Rechnung als Anhang, Texte im Blatt „Mailvorlage“. Der Absender ist im Feld **Absenderkonto** wählbar, auch ein gemeinsames Postfach wie `rechnung@firma.de` (Recht „Senden als“ nötig).

## Briefkopf im PDF

| Bereich | Inhalt |
| --- | --- |
| Kopfzeile links | Name (fett), Straße, PLZ Ort, Telefon, E-Mail, Internet |
| Kopfzeile rechts | Briefkopf-Logo, 2,2 cm hoch |
| Fußzeile links | Name, Geschäftsführer, Registernummer |
| Fußzeile Mitte | USt-IdNr., Steuernummer, Seitenzahl |
| Fußzeile rechts | Bankverbindung der Rechnung (Bezeichnung, IBAN, BIC) |

Einstellung „Briefkopf“ = Nein lässt Kopf- und Fußzeile so, wie sie in der Excel-Datei eingestellt sind. Die Ränder sind auf oben 3,6 cm und unten 2,4 cm gesetzt, die Breite wird auf eine Seite angepasst.

## Test mit den Mustern

Die Auswertungslogik ist an drei Musterrechnungen gegengeprüft (Python-Nachbau `tests/xl_parser_test.py`). Warenwert, Pfand, Zwischensumme, MwSt und Rechnungsbetrag stimmen jeweils überein:

| Rechnung | Kundennr. | Besonderheit |
| --- | --- | --- |
| 2602271 | 11104 | Lieferort Reichenbrand, 1,5 hl Fassbier |
| 2602256 | 12100 (Sammelkunde) | Kasten, Fass und Einzelflaschen gemischt |
| 2602274 | 12100 (Sammelkunde) | drei Fassgrößen (20, 30, 50 Liter) |

Vorschlag für den ersten Test in Excel: Adapter `STUB`, Muster 2602271 einlesen, prüfen, *E-Rechnung erzeugen*. Dann das PDF ansehen (Briefkopf, Seitenumbruch) und die Datei `.daten.xml` kontrollieren. Danach die Platzhalter ersetzen, Adapter `SMARTTOOLS`, erneut erzeugen.

## Noch nicht getestet

- Der VBA-Code von XL2E ist noch nicht in Excel gelaufen. Gegengeprüft ist nur die Einleselogik, mit dem Python-Nachbau.
- Das Aussehen des Briefkopfs ist ungeprüft: Länge der Kopf- und Fußzeilentexte (Excel erlaubt rund 255 Zeichen je Abschnitt), Logogröße, Abstand zum Rechnungsinhalt.
- Ein `&` in Name oder Bankbezeichnung wird in Kopf- und Fußzeile als Steuerzeichen gelesen. Falls nötig, in den Stammdaten `&&` schreiben.

## Bekannte Grenzen

- Ein MwSt-Satz je Rechnung. Bei mehreren Sätzen muss der Satz je Position von Hand eingetragen werden.
- Nur ein Lieferschein je Rechnung wird übernommen.
- Bestellnummer, Kostenstelle und Konditionsschlüssel für REWE stehen nicht in der Excel-Rechnung und werden noch nicht übergeben.
