# Versionshistorie

Schema: HAUPT.NEBEN.KORREKTUR ([Semantic Versioning](https://semver.org/lang/de/)).
Fehlerbehebung → dritte Stelle, neue Funktion → zweite Stelle, inkompatible Änderung → erste Stelle.

## PDF2E 1.11.6 – 25.11.2026

- **Fehlerbehebung:** RE 2604315 (Barverkauf, Kd. 10026): Die Position `59 Rabatt  1-  15,000  15,00-*` (negative Menge, `*` direkt am Betrag) wurde nicht übernommen; Warenwert 104,87 statt 90,00. Jetzt Position mit Menge −1.
- **Fehlerbehebung:** Bei Bruttobeträgen verteilte 1.11.4 auch eine große Differenz (hier −14,87) teilweise auf die Preise. Jetzt werden nur echte Rundungsdifferenzen verteilt (höchstens 1 Cent je Position); eine größere Differenz bleibt stehen, wird gewarnt und von der Prüfung als Fehler gemeldet.
- Testfall `tests/pdftotext/Rechnung_2604315.txt`. XL2E unverändert.

## PDF2E 1.11.5 – 25.11.2026

- **Fehlerbehebung:** RE 2604314 (Verona Birre, Italien) schreibt den Hinweis anders als RE 2604264: `** Steuerfreie Lieferung nach Italien ab Brauereihof **` und `UST - ID: IT04962160232`. Beides wurde nicht erkannt, daher „Kein MwSt-Satz“. Jetzt reichen „Steuerfreie Lieferung“ bzw. „Steuerfreie innergemeinschaftliche Lieferung“ und USt-ID mit beliebigen Leerzeichen/Bindestrichen. Kategorie K wird nur gesetzt, wenn das Land des Empfängers ein anderes EU-Land ist; sonst (z. B. Schweiz, Ausfuhr) eine klare Fehlermeldung.
- Testfall `tests/pdftotext/Rechnung_2604314.txt`. XL2E unverändert.

## PDF2E 1.11.4 – 25.11.2026

- **Änderung (Vorgabe Frank):** Bei Bruttobeträgen sollen die Beträge der E-Rechnung denen der Rechnung entsprechen. Die Rundungsdifferenz zur Nettosumme laut Rechnung (Rechnungsbetrag − enthaltene MwSt) wird jetzt centweise auf die Positionen mit dem größten Rundungsrest verteilt, deren Nettopreis passend neu berechnet. Nettosumme, MwSt und Zahlbetrag stimmen damit genau (RE 2603475: −649,16 / −123,34 / −772,50; RE 2603531: 205,38 / 39,02 / 244,40). Der Rundungsbetrag BT-114 bleibt nur als Rückfall, wenn sich ein Cent nicht verteilen lässt.
- XL2E unverändert.

## PDF2E 1.11.3 – 25.11.2026

- **Fehlerbehebung:** RE 2603531 (Barverkauf ab Rampe, Bruttobeträge, positiver Betrag): Die Summenzeile `RECHNUNGSBETRAG (incl. 43,40 Pfand) EUR 244,40` mit Klammer wurde nicht erkannt; Pfand, Nettosumme und Rechnungsbetrag fehlten.
- **Fehlerbehebung:** Eine Telefonnummer im Anschriftenblock (`0170-4383514`) wurde Teil des Empfängernamens. Sie wird jetzt übergangen und im Protokoll genannt.
- Testfall `tests/pdftotext/Rechnung_2603531.txt`. XL2E unverändert.

## PDF2E 1.11.2 – 25.11.2026

- **Fehlerbehebung:** Bei Bruttobeträgen (RE 2603475) stand der Nettopreis in Excel nur mit 2 Stellen im Blatt Positionen (11,76 statt 11,7647). Die Prüfung meldete deshalb „Menge x Preis“ bei 5 Zeilen. Der Preis wird jetzt als Double mit Zahlenformat `#,##0.0000` geschrieben. XL2E unverändert (nur `modParserBSI` und `modProdukt` geändert).

## PDF2E 1.11.1 und XL2E 1.2.2 – 25.11.2026

- **Fehlerbehebung:** Nach dem Import neuer Module in eine geöffnete ältere Mappe (ohne Neustart) brach „PDF einlesen“ mit „Element in dieser Auflistung nicht gefunden“ ab (DAO-Fehler 3265). Die Migration des Backends auf Schema-Version 2 (Spalte ZRNummer) läuft nur beim Öffnen der Mappe; beim Einlesen fehlte die Spalte. Jetzt werden Schema-Version und neue Einstellungen auch vor jedem Laden der Stammdaten geprüft und bei Bedarf nachgeholt.

## PDF2E 1.11.0 und XL2E 1.2.1 – 25.11.2026

- **Neue Funktion:** Minus-Rechnung mit Rückware und Bruttobeträgen (RE 2603475, J & J Festbedarf): Summenblock `MINUS-RECHNUNGSBETRAG incl. … Pfand) EUR …` und `ENTHALTENE MWSt. … (aus …)` werden gelesen. Bisher brach die Prüfung mit 26 Fehlern ab (kein MwSt-Satz, Menge × Preis bei jeder Rückware-Zeile, Pfand/Netto/Rechnungsbetrag fehlten).
  - Rückware (negativer Betrag) bekommt eine negative Menge.
  - Bruttobeträge: Preise auf 4 Stellen netto umgerechnet, Betrag = Menge × Nettopreis; Rundungsdifferenz als Rundungsbetrag BT-114 (neues Feld `f_Rundung`, wird in älteren Mappen angehängt), Zahlbetrag BT-115 = Rechnungsbetrag laut PDF.
  - Prüfung bei Bruttobeträgen: Warenwert und Pfand brutto, Netto und MwSt mit Toleranz 1 Cent je Position.
  - Zeilen unter „Leergutrückgabe:“ im Positionsteil erscheinen nicht mehr als „Nicht übernommen“.
- **Fehlerbehebung:** Der Rechnungsbetrag im Verarbeitungsprotokoll (Backend) wurde mit `Val()` gelesen und verlor bei deutschen Ländereinstellungen vermutlich die Nachkommastellen. Jetzt Zahlbetrag als Currency, ebenso `{Betrag}` in der E-Mail.
- Testfall `tests/pdftotext/Rechnung_2603475.txt`. XL2E 1.2.1: dieselben gemeinsamen Module.

## PDF2E 1.10.0 und XL2E 1.2.0 – 25.11.2026

- **Neue Funktion:** Vorgaben der G5-Zentralregulierung (SGH Service GmbH, Merkblatt 11/2026 und E-Rechnung-Guideline EN 16931 V1.1, gültig ab 01.12.2026). Gilt nur für Empfänger mit `G5` in der Spalte „Vorgaben“ (von Hand eintragen), REWE, EDEKA und alle anderen Empfänger bleiben unverändert.
  - ZR-Nummer des Händlers (7 Ziffern, beginnt mit 27) aus der neuen Spalte „ZR-Nummer“ im Blatt „Empfaenger“ als Käuferreferenz BT-10 (`ram:BuyerReference`).
  - G5-Lieferantennummer der Brauerei (7 Ziffern, 27…0) aus der neuen Einstellung „G5 Lieferantennummer“ als Verkäuferkennung BT-29 ohne Schema (`ram:SellerTradeParty/ram:ID`), zusätzlich zur GLN.
  - E-Mail-Entwurf an die zentrale Adresse aus der Einstellung „G5 E-Mail“ (Vorgabe `g5.e-invoice@documents.sgh-net.de`).
  - Prüfung: beide Nummern Pflicht und im richtigen Aufbau.
- **Backend:** Schema-Version 2 (Spalte `ZRNummer` in der Tabelle Empfaenger). Die Umstellung läuft beim ersten Start der neuen Version automatisch, mit Sicherung vorher; dabei dürfen keine anderen Plätze PDF2E geöffnet haben. Ältere Mappen bekommen die Spalte vor den Verwaltungsspalten eingefügt; die Übernahme aus einer älteren Mappe liest nur deren Datenspalten.
- XL2E 1.2.0: dieselben gemeinsamen Module.

## PDF2E 1.9.1 – 23.11.2026

- **Fehlerbehebung:** Positionen ohne Angabe in der Spalte Hl./Fl. und mit `*` am Zeilenende (RE 2604307: `136 EURO-Palette  29  12,000  348,00 *`) wurden nicht übernommen; Netto fehlte um 348,00. Der Ort „Hamburg (Deutschland)“ wird ohne Zusatz übernommen.
- Testfall `tests/pdftotext/Rechnung_2604307.txt`. XL2E unverändert (nur `modParserBSI` und `modProdukt` von PDF2E geändert).

## PDF2E 1.9.0 und XL2E 1.1.7 – 23.11.2026

- **Neue Funktion:** Steuerfreie innergemeinschaftliche Lieferung (RE 2604264, Empfänger in Italien): Hinweistext erkannt, USt-Kategorie K mit 0 %, Befreiungsgrund VATEX-EU-IC (BT-120/121), USt-IdNr. des Empfängers aus der Zeile „USt-ID:“, Lieferort/Lieferland aus der Empfängeranschrift, wenn kein Lieferort angegeben ist. Prüfung: USt-IdNr. Pflicht, Land ungleich DE, Lieferland vorhanden.
- **Neue Funktion:** Anschriften mit Länderkennzeichen vor der PLZ (`I- 39012 Meran`), weitere Ländercodes im Smart-Tools-Adapter.
- **Fehlerbehebung:** WARENWERT ohne Betrag (reine Leergutrückgabe) wird als 0 gelesen. Ohne MwSt-Zeile und ohne Steuerbefreiung gibt es eine klare Meldung statt „USt-Satz fehlt“ je Zeile.
- Testfall `tests/pdftotext/Rechnung_2604264.txt`.

## PDF2E 1.8.3 und XL2E 1.1.6 – 23.11.2026

- **Fehlerbehebung:** Kompilierfehler in `modHilfen`: Die Konstante `UST_PRAEFIXE` stand zwischen den Prozeduren statt im Deklarationsteil (von Frank gemeldet und lokal korrigiert).
- **Neu:** `tests/vba_struktur_test.py` prüft alle VBA-Dateien darauf, dass Const/Dim/Declare/Type/Enum auf Modulebene vor der ersten Prozedur stehen.

## PDF2E 1.8.2 und XL2E 1.1.5 – 23.11.2026

- **Fehlerbehebung:** Beim Kunden stand im Blatt „Empfaenger“ der Kundenname in der Spalte USt-IdNr. Beim Einlesen wird das jetzt erkannt, das Feld geleert und im Protokoll gemeldet; andere ungültige USt-IdNr. werden gemeldet. Neue Empfänger bekommen nur eine gültige USt-IdNr.
- **Neu:** Prüfung der Feldnamen im Blatt „Rechnung“ vor dem Einlesen (fehlende Namen, zwei Namen auf derselben Zelle) mit klarer Meldung. Die Ursache beim Kunden ist noch nicht bestätigt; ein falsch verknüpfter Name `f_KUstId` würde genau dieses Bild ergeben.

## PDF2E 1.8.1 und XL2E 1.1.4 – 23.11.2026

- **Fehlerbehebung:** USt-IdNr. ohne Länderkennzeichen fiel erst im Smart-Tools-Validierungsbericht auf (BR-CO-09, RE 2604306). Jetzt wird die USt-IdNr. von Brauerei und Rechnungsempfänger bereinigt (Leerzeichen, Punkte, Bindestriche, Schrägstriche, Großbuchstaben) und beim Prüfen auf Länderkennzeichen und Form geprüft, mit Hinweis auf Blatt und Kundennummer. Bei der Eingabe im Blatt „Empfaenger“ wird sie sofort bereinigt und geprüft.

## PDF2E 1.8.0 und XL2E 1.1.3 – 23.11.2026

- **Neue Funktion:** Skonto- bzw. Nachlasszeile zwischen WARENWERT und NETTOWERT (z. B. `Skonto  Gesamtes Vergütung  299,75 EU  3,00  8,99-`) wird als Nachlass übernommen und in der E-Rechnung als Abschlag auf Belegebene (BG-20) übergeben. Der Warenwert-Abgleich nutzt den NETTOWERT. Auslöser: RE 2604304 (Prüffehler Netto 333,95 statt 324,96).
- **Neue Funktion:** SEPA-Lastschrift aus dem Rechnungstext: Mandatsreferenz, Gläubiger-ID und Einzugsdatum werden gelesen (neue Felder im Blatt „Rechnung“, in älteren Mappen automatisch angelegt). E-Rechnung mit Zahlungsart 59, BT-89, BT-90, Einzugsdatum als BT-9 und Einzugshinweis als BT-20; ebenso im Stub und in der E-Mail.
- Testfall `tests/pdftotext/Rechnung_2604304.txt`.
- XL2E 1.1.3: nur die gemeinsamen Module aktualisiert, keine Änderung im Ablauf.

## PDF2E 1.7.2 und XL2E 1.1.2 – 19.11.2026

- **Fehlerbehebung:** Das Blatt „Versionshistorie“ wird beim Start aus dem Programm nachgeführt (Funktion `Versionen` in `modProdukt`). Bisher befüllte es nur die Vorlage, deshalb fehlten in bestehenden Mappen die Versionen ab 1.5.0. Das Datum steht jetzt als echtes Datum statt als Text.

## PDF2E 1.7.1 und XL2E 1.1.1 – 19.11.2026

- **Fehlerbehebung:** Änderungen an vorhandenen Zeilen der Datenblätter werden über das Recordset gespeichert. Vorher lehnte Access Änderungen im Blatt Einstellungen ab („Sie müssen einen Wert in das Feld 'Einstellungen.Arbeitsplatz' eingeben“).
- **Fehlerbehebung:** Die nicht mehr genutzten Einstellungen Entwickler, Firma und Kontakt werden beim Start aus dem Backend entfernt.

## PDF2E 1.7.0 und XL2E 1.1.0 – 19.11.2026

Mehrplatzbetrieb (gemeinsame Module, in XL2E noch nicht getestet). Anleitung: Abschnitt „Mehrplatz- und Einzelplatzbetrieb“.

- **Neue Funktion:** Backend `.accdb` (DAO, späte Bindung) für Einstellungen, Mailvorlage, Brauerei, Bankkonten, Zahlungskonditionen, Empfänger, Pfandartikel und das Verarbeitungsprotokoll. Wird beim ersten Start angelegt und aus den Blättern vorbelegt; Tabelle Systeminfo mit Schema-Version, Migration mit vorheriger Sicherung.
- **Neue Funktion:** Stammdatenblätter laden beim Aktivieren und vor jedem Einlesen frisch; Änderungen werden sofort pro Zeile gespeichert, Konfliktprüfung über Änderungszähler; Datensatz löschen über das Kontextmenü.
- **Neue Funktion:** Rückfrage beim Einlesen einer bereits erzeugten Rechnung und vor dem Überschreiben, jeweils mit wer und wann.
- **Neue Funktion:** Tägliche Sicherung des Backends (Einstellungen „Backup-Ordner“, „Backup behalten“); Makro `StammdatenAusMappeUebernehmen`.
- **Neue Funktion:** Starter-Mappe (`src/starter`): Versionsprüfung aus der geschlossenen Mappe (Start!Z1), Kopie neuer Versionen, platzbezogene Werte ins neue Blatt „Arbeitsplatz“.
- **Neue Funktion:** Einstellungen je Arbeitsplatz: Spalte „Arbeitsplatz“ (Computername) im Blatt Einstellungen, leer = alle Plätze. Der Computername steht im Info-Fenster.
- **Geändert:** Entwickler, Firma und Kontakt im Info-Fenster sind fest im Programm (Einstellungen „Entwickler“, „Firma“, „Kontakt“ entfallen). „Lizenziert für“ kommt aus den Stammdaten (Brauerei, Name) im Backend, die beim Start vor dem Splash geladen werden.
- Ohne Eintrag „Backend-Pfad“ verhält sich das Programm wie 1.6.0.

## XL2E 1.0.0 – 17.11.2026

Erstversion für Brauereien, deren BSI-Software Rechnungen als Excel-Datei ausgibt. Anleitung: [ANLEITUNG_XL2E.md](ANLEITUNG_XL2E.md).

- Liest das Blatt `PA400_00`: Kopf, Anschrift, Lieferschein, Lieferort, Positionen, Leergutkonto und Summen zum Abgleich.
- Fassbier wird je Hektoliter übergeben (Einheit HLT). Pfand getrennt nach Lieferung und Rücknahme.
- Erzeugt das PDF selbst aus einer Kopie des Rechnungsblatts, mit Briefkopf aus den Stammdaten (Kopfzeile: Absender und Logo; Fußzeile: Register, Steuer, Bank).
- ZUGFeRD EN16931 mit Smart Tools, E-Mail-Entwurf, Prüfung und Info-Fenster aus dem gemeinsamen Kern mit PDF2E.

## PDF2E 1.6.0 – 17.11.2026

- **Umbau:** PDF2E und XL2E nutzen dieselben gemeinsamen Module. Im Repository hat jede Mappe ihren vollständigen Ordner (`src/pdf2e`, `src/xl2e`), die gemeinsamen Module sind identische Kopien (Liste `src/gemeinsam.txt`, Abgleich mit `tools/abgleich.py`). Produktspezifisch ist nur `modProdukt` mit Name, Version und Datenquelle. `modPDF2E` heißt jetzt `modAnwendung`, `PDF2E_Start` ruft `AnwendungStart` auf. Update-Schritte in der Anleitung.
- **Neue Funktion:** Einstellung „Sammelkundennummern“. Bei diesen Kundennummern legt das Programm je Kunde einen eigenen Empfänger an (Schlüssel Kundennr. / Name / PLZ).
- **Neue Funktion:** Die Prüfung erkennt Platzhalter in den Stammdaten (z. B. `DE123456789`, BIC `TESTDEFFXXX`). Mit Smart Tools ist das ein Fehler, im Testmodus eine Warnung.
- **Geändert:** E-Mail an REWE mit neutralem Text ohne Rechnungsdaten (neue Zeile „Text REWE“ im Blatt Mailvorlage, sonst eingebauter Standard). Kein Hinweisfenster mehr.
- **Neue Funktion:** Absender der E-Mail wählbar (Feld „Absenderkonto“ im Blatt Mailvorlage, Auswahl aus den Outlook-Konten). Ein gemeinsames Postfach wie rechnung@firma.de geht auch ohne eigenes Konto (Versand im Namen des Postfachs, Recht „Senden als“ nötig).
- **Pfand mit Differenz 0:** Sind Lieferung und Rückgabe eines Leerguts gleich (z. B. 5 geliefert, 5 zurück), steht im PDF kein Pfandbetrag. Bisher fiel die Zeile weg. Jetzt entstehen zwei Positionen, +5 und −5 zum Pfandsatz, wie EDEKA und REWE es verlangen.
- **Neues Blatt „Pfandartikel“** (Leergut-Nr., Bezeichnung, GTIN): Pfandartikel werden beim Einlesen automatisch angelegt, der Anwender trägt nur die GTIN ein. Sie geht als BT-157 (Schema 0160) in die E-Rechnung; die Prüfung kontrolliert die Prüfziffer.
- **EDEKA-Vorgaben:** Spalte „Vorgaben“ = `EDEKA` (automatisch, wenn der Empfängername „EDEKA“ enthält). Dann sind getrenntes Pfand und eine GTIN für jeden Pfandartikel Pflicht.
- Beim Öffnen ist immer das Blatt „Start“ aktiv, nicht das zuletzt bearbeitete.
- Quellen im VbaSync-Format: UserForm als `frmInfo.frm` + `.frx`, Dokumentmodule als `.doccls`.
- Beschriftungen „laut PDF“ heißen in der Prüfung jetzt „laut Rechnung“.

## PDF2E 1.5.0 – 04.11.2026

- **Neue Funktion: echte E-Rechnung mit Smart Tools.** Mit der Einstellung „E-Rechnung-Adapter“ = `SMARTTOOLS` erzeugt PDF2E eine ZUGFeRD-Rechnung im Profil EN16931. Die Bibliothek bettet `factur-x.xml` in eine Kopie der BSI-PDF ein und wandelt sie in PDF/A-3. Vor dem Speichern prüft die Bibliothek die Rechnung (`Validate`). Bei Regelverletzungen entsteht keine Datei, die Meldungen stehen im Protokoll und im Prüfbericht.
- Übergeben werden unter anderem: GLN von Rechnungsempfänger und Lieferort mit Schema 0088, USt-IdNr. des Empfängers, Lieferschein, Lieferdatum, Bankverbindung, Zahlungsbedingung und Fälligkeit. Aktionen gehen als Abschlag auf Belegebene in die Rechnung. Pfandpositionen sind als Pfand gekennzeichnet (DEPOSIT_SYSTEM / RETURNABLE).
- **Neue Funktion:** Leergut wird in Lieferung (positive Menge) und Rücknahme (negative Menge) aufgeteilt, statt nur die Differenz zu übernehmen. REWE verlangt das. Einstellung „Pfand aufteilen“ (Standard Ja).
- **Neue Funktion:** REWE-Vorgaben je Empfänger (Blatt „Empfaenger“, neue Spalte „Vorgaben“). Ein REWE-Empfänger mit RE-GLN wird beim Einlesen erkannt und markiert. Dann gilt: Hinweis auf Entgeltminderungen (Code AAK) in der Rechnung; RE-GLN, Lieferort-GLN und USt-IdNr. sind Pflicht; beim E-Mail-Entwurf erscheint ein Hinweis, wenn der Text Betrag, Fälligkeit oder Zahlungsbedingung enthält.
- Die Prüfung kontrolliert GLN auf 13 Ziffern und Prüfziffer.
- Neue optionale Einträge: Brauerei „GLN“; Einstellungen „Pfand aufteilen“, „Pfandart“, „Prüfbericht anzeigen“, „REWE Hinweis AAK“. Fehlen sie, gelten die Standardwerte.

## 1.4.0 – 26.10.2026

- **Neue Funktion:** Schaltfläche „E-Mail erstellen“ (vor „Info“). Outlook öffnet einen Entwurf an die E-Mail-Adresse des Empfängers, mit der E-Rechnung als Anhang. Die Mail wird nicht automatisch gesendet.
- Betreff, Text, CC, BCC, Signatur und Absenderkonto stehen im neuen Blatt „Mailvorlage“ mit Platzhaltern wie `{Rechnungsnummer}` oder `{Betrag}`. Fehlt das Blatt, legt das Programm es mit einer Standardvorlage an.
- Im Testmodus (Adapter STUB) beginnt der Betreff mit „TEST - keine gültige E-Rechnung“.
- Voraussetzung: klassisches Outlook für Windows.

## 1.3.0 – 26.10.2026

- **Neue Funktion:** Das Firmenlogo steht im Info-Fenster standardmäßig links, bündig mit dem Anwendungslogo. Mit der Einstellung „Firmenlogo-Position“ = RECHTS steht es wie bisher rechts.
- Anwendungslogo `assets/pdf2e_logo.jpg` durch Franks Entwurf ersetzt.

## 1.2.0 – 26.10.2026

- **Neue Funktion:** Ein Rechnungsempfänger, der noch nicht im Blatt „Empfaenger“ steht, wird beim Einlesen automatisch angelegt (Kundennr., Name, E-Mail, USt-IdNr.). Konto und Kondition bleiben leer, dann gilt der Standard-Eintrag. Nur Abweichungen trägt der Anwender ein.
- **Fehlerbehebung:** Die Korrektur aus 1.0.1 beruhte auf einer Simulation. In der echten Ausgabe von pdftotext unter Windows steht der Teil mit dem Pfandbetrag *vor* der Artikelzeile (RE 2603587: Pfand fehlte erneut). Beide Teile werden jetzt spaltengenau übereinandergelegt; das funktioniert in beiden Reihenfolgen. Eine Teilzeile, die keinem Artikel zugeordnet werden kann, erscheint als Warnung im Protokoll.
- Test: echte Windows-Ausgabe `tests/pdftotext/Rechnung_2603587_windows.txt` ersetzt die Simulation.

## 1.1.0 – 26.10.2026

- **Neue Funktion:** Das Info-Fenster zeigt zwei Logos. Oben steht das Logo der Anwendung (Einstellung „Anwendungslogo“), unten rechts neben den Entwicklerangaben das Firmenlogo (Einstellung „Firmenlogo“).
- Der bisherige Schlüssel „Logo-Datei“ gilt weiter als Firmenlogo, wenn „Firmenlogo“ fehlt.
- Neu: Anwendungslogo `assets/pdf2e_logo.jpg` (Entwurf).

## 1.0.3 – 26.10.2026

- **Fehlerbehebung:** Kann das Info-Fenster die Logo-Datei nicht laden, steht der Grund jetzt im Protokoll. Bisher blieb die Fläche ohne Hinweis leer. Anlass: Eine WebP-Datei mit der Endung .gif wurde nicht angezeigt.

## 1.0.2 – 26.10.2026

- **Fehlerbehebung:** Die Schaltflächen wurden in neueren Excel-Versionen teils als „Das Bild kann nicht angezeigt werden“ dargestellt. Sie sind jetzt abgerundete Formen statt Formularsteuerelemente. `PDF2E_Einrichten` ersetzt die alten Schaltflächen.
- **Fehlerbehebung:** Im Blatt „Start“ stand fest „Version 1.0.0“. Die Zeile wird jetzt beim Öffnen und bei `PDF2E_Einrichten` aus der Programmversion nachgeführt.
- Vorlage: Firma („office & database development“) und Kontakt im Blatt „Einstellungen“ vorbelegt.

## 1.0.1 – 26.10.2026

- **Fehlerbehebung:** pdftotext unter Windows (Xpdf) bricht Leergutzeilen teilweise in zwei Zeilen um. Dadurch fehlte der Pfandbetrag, z. B. −49,60 bei RE 2603587. Der Parser fügt solche Zeilen jetzt zusammen.
- Neue Einstellung „Textdatei behalten“: Die pdftotext-Ausgabe bleibt zur Fehlersuche erhalten.
- Neu: `modEntwicklung` mit `PDF2E_Exportieren` für die Pflege des Quellcodes in Git.

## 1.0.0 – 26.10.2026

- **Erstversion**
- BSI-Rechnung (PDF) einlesen mit `modParserBSI`, Variante je Kunde als Parameter.
- Prüfung: Pflichtangaben, Menge × Preis, Summenabgleich mit dem PDF.
- Stammdaten mit Standard-Einträgen für Bankkonten und Zahlungskonditionen.
- Stub-Adapter statt Smart Tools, bis die Bibliothek vorliegt.
- Splash- und Info-Fenster, Schaltflächen per `PDF2E_Einrichten`.
