# Versionshistorie

Schema: HAUPT.NEBEN.KORREKTUR ([Semantic Versioning](https://semver.org/lang/de/)).
Fehlerbehebung → dritte Stelle, neue Funktion → zweite Stelle, inkompatible Änderung → erste Stelle.

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
