# Versionshistorie

Schema: HAUPT.NEBEN.KORREKTUR ([Semantic Versioning](https://semver.org/lang/de/)).
Fehlerbehebung → dritte Stelle, neue Funktion → zweite Stelle, inkompatible Änderung → erste Stelle.

## Unveröffentlicht

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
