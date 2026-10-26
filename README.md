# PDF2E

**PDF-Rechnungen als E-Rechnung (ZUGFeRD).** Excel/VBA-Werkzeug für Betriebe, die ihre Rechnungen weiter als PDF erzeugen, etwa mit der BSI-Software, und sie als ZUGFeRD-E-Rechnung (Profil EN 16931) versenden müssen.

Das Tool liest die PDF ein, zeigt die Daten zur Kontrolle an und ergänzt sie aus Stammdaten wie Bankkonto und Zahlungskondition. Danach prüft es die Summen gegen das PDF und bettet das XML über die SmartTools-E-Rechnungsbibliothek ins PDF ein.

Aktueller Stand: **1.0.1**, Smart-Tools-Anbindung noch als Stub. Siehe [CHANGELOG.md](CHANGELOG.md).

## Aufbau

| Pfad | Inhalt |
| --- | --- |
| `src/` | VBA-Quellcode (`.bas`), Code der UserForm und von „DieseArbeitsmappe“ als `_Code.txt`, **maßgebliche Quelle** |
| `vorlage/build_xlsx.py` | Erzeugt die Arbeitsmappe `PDF2E.xlsx` (Blätter, Feldnamen, Formeln, Auswahllisten) |
| `tests/pdftotext/` | pdftotext-Ausgaben der Muster-Rechnungen |
| `tests/parser_test.py` | Regressionstest der Einleselogik (Python-Nachbau der Suchmuster) |
| `ANLEITUNG.md` | Einrichtung, Test und bekannte Grenzen |

## Arbeitsweise mit Git

- **Quelle:** Die Dateien in `src/` sind maßgeblich. Arbeitsmappen (`.xlsx/.xlsm`) stehen nicht im Repository. Für Kunden werden sie als Anhang eines GitHub-Releases verteilt.
- **Änderungen in Excel:** Makro `PDF2E_Exportieren` ausführen. Es schreibt alle Module nach `src/`. Danach committen. Einmalig dafür im Trust Center den „Zugriff auf das VBA-Projektobjektmodell“ erlauben.
- **Kodierung:** Im Arbeitsverzeichnis Windows-1252 mit CRLF, wie es der VBA-Editor braucht. Im Repository liegt UTF-8. Das regelt `.gitattributes` automatisch.
- **Versionen:** `TOOL_VERSION` in `src/modPDF2E.bas`, Eintrag in `CHANGELOG.md` und im Blatt „Versionshistorie“ (`vorlage/build_xlsx.py`), dazu ein Git-Tag `vX.Y.Z`.
- **Vor jedem Release:** `python tests/parser_test.py` muss für alle Muster `OK` melden.

## Voraussetzungen

Windows mit Excel 2016 oder neuer und `pdftotext.exe` aus den [Xpdf command line tools](https://www.xpdfreader.com/download.html). Später kommt die SmartTools E-Rechnung VBA-Bibliothek dazu.
