# PDF2E und XL2E

Zwei Excel/VBA-Werkzeuge mit gemeinsamem Kern:

- **PDF2E** – PDF-Rechnungen als E-Rechnung (ZUGFeRD), Anleitung: [ANLEITUNG.md](ANLEITUNG.md)
- **XL2E** – Excel-Rechnungen als E-Rechnung (ZUGFeRD), Anleitung: [ANLEITUNG_XL2E.md](ANLEITUNG_XL2E.md)

**PDF2E** ist ein Excel/VBA-Werkzeug für Betriebe, die ihre Rechnungen weiter als PDF erzeugen, etwa mit der BSI-Software, und sie als ZUGFeRD-E-Rechnung (Profil EN 16931) versenden müssen.

Das Tool liest die PDF ein, zeigt die Daten zur Kontrolle an und ergänzt sie aus Stammdaten wie Bankkonto und Zahlungskondition. Danach prüft es die Summen gegen das PDF und bettet das XML über die SmartTools-E-Rechnungsbibliothek ins PDF ein.

**XL2E** liest die Excel-Rechnung der BSI-Software (Blatt `PA400_00`), erzeugt daraus ein PDF mit Briefkopf aus den Stammdaten und bettet die E-Rechnung ein.

Aktueller Stand: **PDF2E 1.6.0** und **XL2E 1.0.0**, beide mit Smart-Tools-Anbindung (ZUGFeRD EN16931). Siehe [CHANGELOG.md](CHANGELOG.md).

## Aufbau

| Pfad | Inhalt |
| --- | --- |
| `src/kern/` | Gemeinsamer VBA-Code (Einlesen, Prüfen, Smart Tools, E-Mail, Info-Fenster), **maßgebliche Quelle** |
| `src/pdf2e/` | Nur PDF2E: `modProdukt` (Name, Version, Datenquelle), pdftotext, Einlese-Vorlage BSI-PDF |
| `src/xl2e/` | Nur XL2E: `modProdukt` (Name, Version, PDF mit Briefkopf), Einlese-Vorlage BSI-Excel |
| `vorlage/build_xlsx.py` | Erzeugt `PDF2E.xlsx` oder `XL2E.xlsx` (Blätter, Feldnamen, Formeln, Auswahllisten) |
| `vorlage/pakete.py` | Erzeugt die Auslieferungs-ZIPs beider Produkte |
| `tests/pdftotext/`, `tests/parser_test.py` | Muster und Regressionstest der PDF-Einleselogik |
| `tests/xl_parser_test.py` | Regressionstest der Excel-Einleselogik; die Muster in `tests/xlsx/` liegen nicht im Repository |
| `ANLEITUNG.md`, `ANLEITUNG_XL2E.md` | Einrichtung, Test und bekannte Grenzen |

Der Kern ruft nur, was `modProdukt` festlegt: die Konstanten `TOOL_NAME`, `TOOL_VERSION`, `TOOL_STAND`, `TOOL_UNTERTITEL`, `QUELLE_SCHALTFLAECHE` und die Prozeduren `QuelleAuswaehlen`, `QuelleEinlesen`, `QuellPdfErstellen`. Ein weiteres Quellformat braucht nur ein neues Produktverzeichnis.

## Arbeitsweise mit Git

- **Quelle:** Die Dateien in `src/` sind maßgeblich. Arbeitsmappen (`.xlsx/.xlsm`) stehen nicht im Repository. Für Kunden werden sie als Anhang eines GitHub-Releases verteilt.
- **Änderungen in Excel:** Makro `PDF2E_Exportieren` bzw. `XL2E_Exportieren` ausführen. Es schreibt alle Module in einen Ordner; danach nach `src/kern` bzw. Produktordner einsortieren. Danach committen. Einmalig dafür im Trust Center den „Zugriff auf das VBA-Projektobjektmodell“ erlauben.
- **Kodierung:** Im Arbeitsverzeichnis Windows-1252 mit CRLF, wie es der VBA-Editor braucht. Im Repository liegt UTF-8. Das regelt `.gitattributes` automatisch.
- **Versionen:** `TOOL_VERSION` in `src/<produkt>/modProdukt.bas`, Eintrag in `CHANGELOG.md` und im Blatt „Versionshistorie“ (`vorlage/build_xlsx.py`), dazu ein Git-Tag `vX.Y.Z`.
- **Vor jedem Release:** `python tests/parser_test.py` und `python tests/xl_parser_test.py` müssen für alle Muster `OK` melden. Pakete: `python vorlage/pakete.py <Zielordner> [<Logo-Ordner>]`.

## Voraussetzungen

Windows mit Excel 2016 oder neuer, für PDF2E zusätzlich `pdftotext.exe` aus den [Xpdf command line tools](https://www.xpdfreader.com/download.html). Für echte E-Rechnungen zusätzlich die SmartTools E-Rechnung VBA-Bibliothek mit Lizenz je Arbeitsplatz.
