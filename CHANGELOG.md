# Versionshistorie

Schema: HAUPT.NEBEN.KORREKTUR ([Semantic Versioning](https://semver.org/lang/de/)).
Fehlerbehebung → dritte Stelle, neue Funktion → zweite Stelle, inkompatible Änderung → erste Stelle.

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
