# PDF2E und XL2E

Zwei Excel/VBA-Werkzeuge mit gemeinsamem Kern:

- **PDF2E** – PDF-Rechnungen als E-Rechnung (ZUGFeRD), Anleitung: [ANLEITUNG.md](ANLEITUNG.md)
- **XL2E** – Excel-Rechnungen als E-Rechnung (ZUGFeRD), Anleitung: [ANLEITUNG_XL2E.md](ANLEITUNG_XL2E.md)

**PDF2E** ist ein Excel/VBA-Werkzeug für Betriebe, die ihre Rechnungen weiter als PDF erzeugen, etwa mit der BSI-Software, und sie als ZUGFeRD-E-Rechnung (Profil EN 16931) versenden müssen.

Das Tool liest die PDF ein, zeigt die Daten zur Kontrolle an und ergänzt sie aus Stammdaten wie Bankkonto und Zahlungskondition. Danach prüft es die Summen gegen das PDF und bettet das XML über die SmartTools-E-Rechnungsbibliothek ins PDF ein.

**XL2E** liest die Excel-Rechnung der BSI-Software (Blatt `PA400_00`), erzeugt daraus ein PDF mit Briefkopf aus den Stammdaten und bettet die E-Rechnung ein.

Aktueller Stand: **PDF2E 1.9.1** und **XL2E 1.1.7**, beide mit Smart-Tools-Anbindung (ZUGFeRD EN16931). Siehe [CHANGELOG.md](CHANGELOG.md).

## Aufbau

| Pfad | Inhalt |
| --- | --- |
| `src/pdf2e/` | Alle VBA-Module der Mappe PDF2E (VbaSync-Format: `.bas`, `frmInfo.frm` + `.frx`, `.doccls`), **maßgebliche Quelle** |
| `src/xl2e/` | Alle VBA-Module der Mappe XL2E, gleiches Format |
| `src/gemeinsam.txt` | Liste der Module, die in allen Produktordnern identisch sein müssen |
| `tools/abgleich.py` | Prüft die gemeinsamen Module; `--von pdf2e` kopiert sie in die anderen Ordner |
| `vorlage/build_xlsx.py` | Erzeugt `PDF2E.xlsx` oder `XL2E.xlsx` (Blätter, Feldnamen, Formeln, Auswahllisten) |
| `vorlage/pakete.py` | Erzeugt die Auslieferungs-ZIPs beider Produkte |
| `tests/pdftotext/`, `tests/parser_test.py` | Muster und Regressionstest der PDF-Einleselogik |
| `tests/xl_parser_test.py` | Regressionstest der Excel-Einleselogik; die Muster in `tests/xlsx/` liegen nicht im Repository |
| `ANLEITUNG.md`, `ANLEITUNG_XL2E.md` | Einrichtung, Test und bekannte Grenzen |

**Gemeinsamer Code als Kopien:** Jede Mappe hat ihren vollständigen Satz Module, damit sie für sich allein läuft und VbaSync wie gewohnt einen Ordner exportiert. Die Module aus `src/gemeinsam.txt` sind in allen Produktordnern identisch. Sie rufen nur, was `modProdukt` festlegt: die Konstanten `TOOL_NAME`, `TOOL_VERSION`, `TOOL_STAND`, `TOOL_UNTERTITEL`, `QUELLE_SCHALTFLAECHE` und die Prozeduren `QuelleAuswaehlen`, `QuelleEinlesen`, `QuellPdfErstellen`. Nach jeder Änderung an einem gemeinsamen Modul `python tools/abgleich.py --von <produkt>` ausführen und die geänderten Module in die anderen Mappen importieren.

## Arbeitsweise mit Git

- **Quelle:** Die Dateien in `src/` sind maßgeblich. Arbeitsmappen (`.xlsx/.xlsm`) stehen nicht im Repository. Für Kunden werden sie als Anhang eines GitHub-Releases verteilt.
- **Änderungen in Excel:** Mit [VbaSync](https://github.com/FrankHentzschel/VbaSync) „Alle exportieren“ in einen lokalen Ordner je Mappe (z. B. `C:\Work\PDF2E\src` und `C:\Work\XL2E\src`); im Repository landet der Inhalt in `src/pdf2e` bzw. `src/xl2e`. Danach `python tools/abgleich.py --von <produkt>` und committen. Einmalig dafür im Trust Center den „Zugriff auf das VBA-Projektobjektmodell“ erlauben.
- **Kodierung:** Im Arbeitsverzeichnis Windows-1252 mit CRLF, wie es der VBA-Editor braucht. Im Repository liegt UTF-8. Das regelt `.gitattributes` automatisch.
- **Versionen:** `TOOL_VERSION` in `src/<produkt>/modProdukt.bas`, Eintrag in `CHANGELOG.md` und im Blatt „Versionshistorie“ (`vorlage/build_xlsx.py`), dazu ein Git-Tag `vX.Y.Z`.
- **Vor jedem Release:** `python tests/vba_struktur_test.py` (Deklarationen nur im Kopf der Module), `python tests/parser_test.py` und `python tests/xl_parser_test.py` müssen für alle Muster `OK` melden. Pakete: `python vorlage/pakete.py <Zielordner> [<Logo-Ordner>]`.

## Voraussetzungen

Windows mit Excel 2016 oder neuer, für PDF2E zusätzlich `pdftotext.exe` aus den [Xpdf command line tools](https://www.xpdfreader.com/download.html). Für echte E-Rechnungen zusätzlich die SmartTools E-Rechnung VBA-Bibliothek mit Lizenz je Arbeitsplatz.
