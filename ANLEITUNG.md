# PDF2E 1.5.0 – PDF-Rechnungen als E-Rechnung (ZUGFeRD)

Kundenneutrales Werkzeug, Stand 04.11.2026. Erster Kunde ist die Krug-Bräu. Kundenspezifisch sind nur die Stammdaten-Blätter und das Rechnungslayout (Einstellung „Layout“, Einlese-Vorlage `modParserBSI` mit Kundenvariante als Parameter). Echte E-Rechnungen entstehen mit der Smart-Tools-Bibliothek (Adapter `SMARTTOOLS`). Im **Stub-Modus** (Adapter `STUB`) entsteht keine gültige E-Rechnung, er dient nur zum Testen ohne Bibliothek.

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
| `src/modAdapterSmartTools.bas` | Smart Tools: ZUGFeRD EN16931 erzeugen, prüfen, in die PDF einbetten |
| `src/modEntwicklung.bas` | Nur Entwicklung: `PDF2E_Exportieren` schreibt alle Module für Git nach `src/` |
| `src/frmInfo.frm` + `src/frmInfo.frx` | Splash- und Info-Fenster (UserForm, VBE-Export) |
| `src/DieseArbeitsmappe.doccls` | Code für „DieseArbeitsmappe“ (Splash beim Öffnen) |

## Einrichtung (ca. 15 Minuten)

1. **pdftotext:** Lade auf https://www.xpdfreader.com/download.html die *Xpdf command line tools* für Windows und lege `pdftotext.exe` (64-Bit) z. B. unter `C:\Tools\xpdf\` ab.
2. **Mappe vorbereiten:** `PDF2E.xlsx` öffnen und als **Excel-Arbeitsmappe mit Makros (.xlsm)** speichern.
3. **Module importieren:** VBA-Editor mit Alt+F11 öffnen. Über *Datei > Datei importieren* alle `.bas`-Dateien aus `src/` einlesen.
4. **Info-Fenster importieren:** Ebenso `frmInfo.frm` importieren (`frmInfo.frx` muss im selben Ordner liegen).
5. **Start-Code:** Doppelklick auf *DieseArbeitsmappe* und den Inhalt von `DieseArbeitsmappe.doccls` einfügen.

   Mit [VbaSync](https://github.com/FrankHentzschel/VbaSync) erledigt „Alle importieren“ die Schritte 3 bis 5 in einem Durchgang.
6. **Kompilieren:** *Debuggen > Kompilieren VBAProject*. Es dürfen keine Fehler erscheinen.
7. **Einstellungen** im gleichnamigen Blatt ausfüllen:
   - Pfad zu pdftotext, Eingangs- und Ausgabeordner
   - **Firma**, Kontakt, Anwendungslogo (oben) und Firmenlogo (unten), jeweils BMP, JPG oder GIF, kein PNG
   - den Adapter auf `STUB` lassen, bis Smart Tools eingerichtet ist (siehe unten)
8. **Schaltflächen:** Mit Alt+F8 einmalig `PDF2E_Einrichten` ausführen. Das Makro legt auf „Start“ und „Rechnung“ die Schaltflächen *PDF einlesen*, *Prüfen*, *E-Rechnung erzeugen* und *Info* an.
9. Speichern, schließen und neu öffnen. Der Splash erscheint 3 Sekunden lang, ein Klick schließt ihn sofort.

## Smart Tools einrichten

1. SmartTools E-Rechnung VBA-Bibliothek installieren. Die Lizenzdatei `Lizenz xl-e-rechnung-vba.txt` (oder `Lizenz <Bestellnr> xl-e-rechnung-vba.txt`) liegt in `%APPDATA%\SmartTools\Excel E-Rechnung VBA`. Die Lizenz gilt je Arbeitsplatz.
2. VBA-Editor (Alt+F11) > *Extras > Verweise* > **SmartTools E-Rechnung-Bibliothek** anhaken.
3. *Debuggen > Kompilieren VBAProject*. Meldet der Editor in `modAdapterSmartTools` „Benutzerdefinierter Typ nicht definiert“, steht der Verweis nicht (Schritt 2) oder ein Klassenname weicht ab. Die vier abgeleiteten Klassennamen stehen gesammelt unten im Modul (`NeueKennung`, `NeuesMerkmal`, `NeueNotiz`, `NeuerAbschlag`); den richtigen Namen zeigt der Objektkatalog (F2).
4. Blatt „Einstellungen“: „E-Rechnung-Adapter“ = `SMARTTOOLS`.
5. Optional: Blatt „Brauerei“, Zeile „GLN“ mit der eigenen GLN (von REWE gewünscht).

Ohne Verweis lässt sich das Projekt nicht kompilieren, auch nicht im Stub-Modus. Wer Smart Tools nicht nutzt, entfernt `modAdapterSmartTools` und die Zeile `Case "SMARTTOOLS"` in `modZugferd`.

## REWE-Empfänger

REWE (REWE-Zentralfinanz, Startpaket E-Rechnung 1.7 vom 14.07.2026) verlangt unter anderem RE-GLN und Lieferort-GLN mit Schema 0088, die USt-IdNr. von REWE, einen Hinweis auf Entgeltminderungen (Code AAK) und getrennte Positionen für Pfand-Lieferung und Pfand-Rücknahme.

- Beim Einlesen wird ein Rechnungsempfänger, dessen Name mit „REWE“ beginnt und der eine RE-GLN hat, im Blatt „Empfaenger“ in der Spalte **Vorgaben** mit `REWE` markiert. Von Hand geht das genauso.
- Die **USt-IdNr. von REWE** steht nicht auf der BSI-Rechnung. Sie muss einmal im Blatt „Empfaenger“ eingetragen werden, sonst meldet die Prüfung einen Fehler.
- Der Hinweistext steht in der Einstellung „REWE Hinweis AAK“.
- Der Mailtext an REWE darf keine buchungsrelevanten Inhalte enthalten. Beim E-Mail-Entwurf erscheint ein Hinweis, wenn die Vorlage Betrag, Fälligkeit oder Zahlungsbedingung enthält.
- Noch nicht umgesetzt: Bestellnummer (BT-13), Kostenstelle (BT-19), Konditionsschlüssel (OSI) und Profil EXTENDED. Die BSI-Rechnungen enthalten diese Angaben nicht.

## Versionsnummern

PDF2E verwendet das übliche Schema HAUPT.NEBEN.KORREKTUR (Semantic Versioning):

| Änderung | Beispiel |
| --- | --- |
| Fehlerbehebung | 1.0.0 → 1.0.1 |
| neue Funktion | 1.0.1 → 1.1.0 |
| inkompatible Änderung, z. B. neues Stammdatenformat | 1.4.2 → 2.0.0 |

Die Version steht als Konstante `TOOL_VERSION` in `modPDF2E`. Jede Änderung kommt zusätzlich ins Blatt „Versionshistorie“.

## E-Mail-Versand

Nach „E-Rechnung erzeugen“ öffnet die Schaltfläche **E-Mail erstellen** einen Outlook-Entwurf. Er geht an die E-Mail-Adresse aus dem Blatt „Empfaenger“ und hat die E-Rechnung als Anhang. Der Anwender prüft den Entwurf, ergänzt ihn bei Bedarf und sendet selbst.

Betreff und Text stehen im Blatt **Mailvorlage**. Platzhalter: `{Rechnungsnummer}`, `{Rechnungsdatum}`, `{Betrag}`, `{Faellig}`, `{Zahlungsbedingung}`, `{Kundennummer}`, `{Kunde}`, `{Lieferort}`, `{Lieferschein}`, `{Absender}`, `{Dateiname}`. Außerdem lassen sich dort CC, BCC, Signatur (Ja/Nein) und ein abweichendes Absenderkonto einstellen.

Voraussetzung ist das klassische Outlook für Windows. Das „neue Outlook“ bietet keine Schnittstelle für VBA.

## Test mit den Mustern

Die Auswertungslogik ist an allen vier Mustern gegengeprüft. Warenwert, Pfand, Nettosumme, MwSt und Rechnungsbetrag stimmen jeweils mit dem PDF überein:

| Rechnung | Empfänger | Positionen | Pfand | Brutto |
| --- | --- | --- | --- | --- |
| 2603976 | Markgrafen, Lieferort Stadtsteinach | 13 + 2 Aktionen | 201,50 + 15,50 | 1.029,53 |
| 2603894 | Markgrafen, Lieferort Herzogenaurach | 3 | +31,00 / −49,60 = −18,60 | 141,19 |
| 2603816 | REWE Markt GmbH (RE-GLN 4388096000002) | 3 | +124,00 / −99,20 = 24,80 | 489,74 |
| 2603587 | REWE Markt GmbH, Lieferort Forchheim | 4 | +62,00 / −111,60 = −49,60 | 176,30 |

Seit 1.5.0 steht Pfand getrennt nach Lieferung und Rücknahme in der Rechnung. Das Programm leitet beide Mengen aus dem Leergutkonto ab (Lieferung − Rücknahme = Differenz, alter Saldo + Differenz = neuer Saldo). Gelingt das nicht eindeutig, bleibt es bei der Differenz mit einer Warnung im Protokoll.

Im Ordner `tests/pdftotext` liegen die pdftotext-Ausgaben dieser Muster. Dateien mit `_windows` stammen aus pdftotext unter Windows (Xpdf), die übrigen aus Linux (Poppler). Die Windows-Version verteilt manche Leergutzeilen auf zwei Zeilen; der Parser fügt sie spaltengenau zusammen.

Im Stub-Modus legt **E-Rechnung erzeugen** zwei Dateien an: `<RechNr>_ZUGFeRD_STUB.pdf`, eine unveränderte Kopie, und daneben `.daten.xml` mit allen Übergabewerten.

## Noch nicht getestet

- Der Smart-Tools-Adapter (1.5.0) ist noch nicht in Excel gelaufen. Zu prüfen: Kompilieren (vier abgeleitete Klassennamen), Zahlenübergabe mit deutschem Dezimalkomma, Ergebnis im Prüfbericht und in einem externen Validator.

- Der VBA-Code ist noch nicht in Excel gelaufen. Gegengeprüft habe ich nur die Auswertungslogik, und zwar mit Poppler-pdftotext unter Linux.
- `VBScript.RegExp`: Microsoft kündigt die schrittweise Abschaltung von VBScript an. Ob und wann das RegExp in VBA betrifft, ist offen.

## Bekannte Grenzen

- Bei mehreren Lieferscheinen pro Rechnung wird nur der erste übernommen. Das Protokoll gibt dann eine Warnung aus.
- Bei mehreren MwSt-Sätzen muss der Satz je Position von Hand eingetragen werden.
- Als Einheit ist C62 (Stück) für Kasten und Fass eingestellt. Das lässt sich in den Einstellungen ändern.
