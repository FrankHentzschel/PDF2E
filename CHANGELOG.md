# Versionshistorie

Schema: HAUPT.NEBEN.KORREKTUR ([Semantic Versioning](https://semver.org/lang/de/)).
Fehlerbehebung → dritte Stelle, neue Funktion → zweite Stelle, inkompatible Änderung → erste Stelle.

## XL2E 1.0.0 – 04.11.2026

Erstversion für Brauereien, deren BSI-Software Rechnungen als Excel-Datei ausgibt. Anleitung: [ANLEITUNG_XL2E.md](ANLEITUNG_XL2E.md).

- Liest das Blatt `PA400_00`: Kopf, Anschrift, Lieferschein, Lieferort, Positionen, Leergutkonto und Summen zum Abgleich.
- Fassbier wird je Hektoliter übergeben (Einheit HLT). Pfand getrennt nach Lieferung und Rücknahme.
- Erzeugt das PDF selbst aus einer Kopie des Rechnungsblatts, mit Briefkopf aus den Stammdaten (Kopfzeile: Absender und Logo; Fußzeile: Register, Steuer, Bank).
- ZUGFeRD EN16931 mit Smart Tools, E-Mail-Entwurf, Prüfung und Info-Fenster aus dem gemeinsamen Kern mit PDF2E.

## PDF2E 1.6.0 – 04.11.2026

- **Umbau:** PDF2E und XL2E nutzen denselben Kern (`src/kern`). Produktspezifisch ist nur `modProdukt` mit Name, Version und Datenquelle. `modPDF2E` heißt jetzt `modAnwendung`, `PDF2E_Start` ruft `AnwendungStart` auf. Update-Schritte in der Anleitung.
- **Neue Funktion:** Einstellung „Sammelkundennummern“. Bei diesen Kundennummern legt das Programm je Kunde einen eigenen Empfänger an (Schlüssel Kundennr. / Name / PLZ).
- **Neue Funktion:** Die Prüfung erkennt Platzhalter in den Stammdaten (z. B. `DE123456789`, BIC `TESTDEFFXXX`). Mit Smart Tools ist das ein Fehler, im Testmodus eine Warnung.
- **Geändert:** E-Mail an REWE mit neutralem Text ohne Rechnungsdaten (neue Zeile „Text REWE“ im Blatt Mailvorlage, sonst eingebauter Standard). Kein Hinweisfenster mehr.
- **Neue Funktion:** Absender der E-Mail wählbar (Feld „Absenderkonto“ im Blatt Mailvorlage, Auswahl aus den Outlook-Konten). Ein gemeinsames Postfach wie rechnung@firma.de geht auch ohne eigenes Konto (Versand im Namen des Postfachs, Recht „Senden als“ nötig).
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
