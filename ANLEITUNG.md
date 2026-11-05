# PDF2E 1.6.0 – PDF-Rechnungen als E-Rechnung (ZUGFeRD)

Kundenneutrales Werkzeug, Stand 04.11.2026. Erster Kunde ist die Krug-Bräu. Kundenspezifisch sind nur die Stammdaten-Blätter und das Rechnungslayout (Einstellung „Layout“, Einlese-Vorlage `modParserBSI` mit Kundenvariante als Parameter). Echte E-Rechnungen entstehen mit der Smart-Tools-Bibliothek (Adapter `SMARTTOOLS`). Im **Stub-Modus** (Adapter `STUB`) entsteht keine gültige E-Rechnung, er dient nur zum Testen ohne Bibliothek.

## Inhalt

PDF2E und XL2E teilen sich seit PDF2E 1.6.0 denselben Kern. Im ZIP liegen alle benötigten Module gemeinsam im Ordner `vba/`.

| Datei | Zweck |
| --- | --- |
| `PDF2E.xlsx` | Arbeitsmappe ohne Makros |
| `modProdukt.bas` | **nur PDF2E:** Name, Version, Dateiauswahl für PDF, Aufruf der Einlese-Vorlage |
| `modPdfText.bas` | **nur PDF2E:** ruft `pdftotext -layout` auf und liest den Text ein |
| `modParserBSI.bas` | **nur PDF2E:** Einlese-Vorlage für BSI-PDF-Rechnungen, Variante über die Einstellung „Layout-Variante“ |
| `modAnwendung.bas` | Kern: Splash beim Start, Info-Fenster, Schaltflächen anlegen (`PDF2E_Einrichten`) |
| `modHilfen.bas` | Kern: Einstellungen, Feldzugriff, Zahlen, Datum, RegExp, Protokoll, Positionen anlegen |
| `modEinlesen.bas` | Kern: Makro **RechnungEinlesen**, Empfänger anlegen |
| `modPruefen.bas` | Kern: Pflichtangaben, Menge × Preis, Summenabgleich, GLN, Platzhalter in Stammdaten |
| `modZugferd.bas` | Kern: Makro **ZugferdErzeugen**, ruft den Adapter auf |
| `modAdapterStub.bas` | Kern: Testlauf ohne Smart Tools, schreibt `<Ziel>.daten.xml` |
| `modAdapterSmartTools.bas` | Kern: ZUGFeRD EN16931 erzeugen, prüfen, in die PDF einbetten |
| `modMail.bas` | Kern: E-Mail-Entwurf in Outlook |
| `modEntwicklung.bas` | Kern, nur Entwicklung: `PDF2E_Exportieren` schreibt die Module nach `src/` |
| `frmInfo_Code.txt` | Code für das Splash- und Info-Fenster (UserForm) |
| `ThisWorkbook_Code.txt` | Code für „DieseArbeitsmappe“ (Splash beim Öffnen) |

## Update von 1.5.0 auf 1.6.0

Die Funktionen bleiben gleich. Geändert hat sich die Aufteilung der Module, damit PDF2E und XL2E denselben Code nutzen. Die Stammdaten in der vorhandenen Mappe bleiben erhalten.

1. Vorhandene `PDF2E.xlsm` sichern (Kopie anlegen).
2. VBA-Editor (Alt+F11): alle Module **entfernen** (Rechtsklick > *Entfernen*, Frage nach Export mit *Nein*). Auch `modPDF2E` muss weg, es heißt jetzt `modAnwendung`.
3. Alle `.bas`-Dateien aus `vba/` importieren.
4. Code in `frmInfo` und in *DieseArbeitsmappe* vollständig durch `frmInfo_Code.txt` bzw. `ThisWorkbook_Code.txt` ersetzen. *DieseArbeitsmappe* ruft jetzt `AnwendungStart` auf.
5. *Debuggen > Kompilieren VBAProject*.
6. Alt+F8 > `PDF2E_Einrichten` einmal ausführen. Die Schaltflächen werden neu angelegt.
7. Optional im Blatt „Einstellungen“ eine Zeile **Sammelkundennummern** ergänzen (siehe unten).

## Einrichtung (ca. 15 Minuten)

1. **pdftotext:** Lade auf https://www.xpdfreader.com/download.html die *Xpdf command line tools* für Windows und lege `pdftotext.exe` (64-Bit) z. B. unter `C:\Tools\xpdf\` ab.
2. **Mappe vorbereiten:** `PDF2E.xlsx` öffnen und als **Excel-Arbeitsmappe mit Makros (.xlsm)** speichern.
3. **Module importieren:** VBA-Editor mit Alt+F11 öffnen. Über *Datei > Datei importieren* alle `.bas`-Dateien aus `vba/` einlesen.
4. **Info-Fenster anlegen:** *Einfügen > UserForm*. Im Eigenschaftenfenster bei (Name) `frmInfo` eintragen. Dann mit Rechtsklick auf `frmInfo` *Code anzeigen* wählen und den gesamten Inhalt von `frmInfo_Code.txt` einfügen. Auf der Form selbst nichts platzieren, die Elemente entstehen zur Laufzeit.
5. **Start-Code:** Doppelklick auf *DieseArbeitsmappe* und den Inhalt von `ThisWorkbook_Code.txt` einfügen.
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
- Der Mailtext an REWE darf keine buchungsrelevanten Inhalte enthalten (Startpaket: „keine buchungsrelevanten Inhalte im E-Mail-Text“, ohne nähere Festlegung). Seit 1.6.0 bekommt REWE deshalb einen neutralen Text ohne Rechnungsdaten: Anrede, „anbei erhalten Sie unsere E-Rechnung im ZUGFeRD-Format“, Gruß. Der Text steht im Blatt „Mailvorlage“ in der Zeile **Text REWE**. Fehlt die Zeile (Mappen vor 1.6.0), gilt dieser Standard. Betrag, Fälligkeit und Zahlungsbedingung werden auch dort immer entfernt. Der Betreff bleibt wie in der Vorlage, z. B. „Rechnung 2603587 vom 27.09.2026“.
- Noch nicht umgesetzt: Bestellnummer (BT-13), Kostenstelle (BT-19), Konditionsschlüssel (OSI) und Profil EXTENDED. Die BSI-Rechnungen enthalten diese Angaben nicht.

## Versionsnummern

PDF2E verwendet das übliche Schema HAUPT.NEBEN.KORREKTUR (Semantic Versioning):

| Änderung | Beispiel |
| --- | --- |
| Fehlerbehebung | 1.0.0 → 1.0.1 |
| neue Funktion | 1.0.1 → 1.1.0 |
| inkompatible Änderung, z. B. neues Stammdatenformat | 1.4.2 → 2.0.0 |

Die Version steht als Konstante `TOOL_VERSION` in `modProdukt`. Jede Änderung kommt zusätzlich ins Blatt „Versionshistorie“.

## Sammelkundennummern

Manche Kunden werden unter einer gemeinsamen Kundennummer abgerechnet, etwa Festlichkeiten oder Barverkäufe. Solche Nummern stehen in der Einstellung **Sammelkundennummern**, mehrere durch Komma getrennt. Bei diesen Nummern legt das Programm im Blatt „Empfaenger“ je Kunde einen eigenen Eintrag an, mit dem Schlüssel `Kundennr. / Name / PLZ`, zum Beispiel `12100 / Gasthof Linde / 09481`. E-Mail-Adresse und Vorgaben gelten dann für genau diesen Kunden.

## Stammdaten mit Platzhaltern

Die Prüfung meldet Werte, die erkennbar Platzhalter sind (z. B. `DE123456789`, `HRB 12345`, BIC `TESTDEFFXXX`). Im Testmodus (STUB) ist das eine Warnung, mit Smart Tools ein Fehler. So entsteht keine echte E-Rechnung mit Musterdaten.

## E-Mail-Versand

Nach „E-Rechnung erzeugen“ öffnet die Schaltfläche **E-Mail erstellen** einen Outlook-Entwurf. Er geht an die E-Mail-Adresse aus dem Blatt „Empfaenger“ und hat die E-Rechnung als Anhang. Der Anwender prüft den Entwurf, ergänzt ihn bei Bedarf und sendet selbst.

Betreff und Text stehen im Blatt **Mailvorlage**. Platzhalter: `{Rechnungsnummer}`, `{Rechnungsdatum}`, `{Betrag}`, `{Faellig}`, `{Zahlungsbedingung}`, `{Kundennummer}`, `{Kunde}`, `{Lieferort}`, `{Lieferschein}`, `{Absender}`, `{Dateiname}`. Außerdem lassen sich dort CC, BCC, Signatur (Ja/Nein) und ein abweichendes Absenderkonto einstellen.

**Absender:** Im Blatt „Mailvorlage“ legt das Feld **Absenderkonto** fest, von welcher Adresse die Mail kommt. Leer bedeutet: das Standardkonto des angemeldeten Outlook-Benutzers.

- Beim ersten „E-Mail erstellen“ (oder mit dem Makro `OutlookAbsenderLesen`) schreibt das Programm alle Konten aus Outlook in Spalte E. Das Feld bietet sie dann als Auswahl an.
- Ein **gemeinsames Postfach** wie `rechnung@firma.de` lässt sich auch eintippen, wenn es kein eigenes Konto in Outlook ist. Die Mail geht dann im Namen dieses Postfachs. Dafür muss der IT-Dienstleister jedem Benutzer das Recht „Senden als“ (oder „Senden im Auftrag von“) auf dieses Postfach geben. Fehlt das Recht, meldet der Mailserver erst nach dem Senden einen Fehler.
- So kommen alle Rechnungen von derselben Adresse, egal an welchem Arbeitsplatz sie erstellt werden. Das ist bei REWE wichtig: Dort werden nur angemeldete Absenderadressen angenommen (RZD-Starterkit 1.09, bis zu 5 Adressen; für die RZF beim Onboarding klären).
- Der Entwurf zeigt im Feld „Von“, welcher Absender gilt. Bitte vor dem Senden kurz prüfen.

Voraussetzung ist das klassische Outlook für Windows. Das „neue Outlook“ bietet keine Schnittstelle für VBA. Ob später ein Versand ohne Outlook (SMTP oder Microsoft Graph) nötig ist, wird mit dem IT-Dienstleister geklärt.

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

- Die neue Modulaufteilung (1.6.0) ist noch nicht in Excel kompiliert worden. Die Logik ist unverändert gegenüber 1.5.0, die am 04.11.2026 mit RE 2603587 eine gültige E-Rechnung erzeugt hat.
- `VBScript.RegExp`: Microsoft kündigt die schrittweise Abschaltung von VBScript an. Ob und wann das RegExp in VBA betrifft, ist offen.

## Bekannte Grenzen

- Bei mehreren Lieferscheinen pro Rechnung wird nur der erste übernommen. Das Protokoll gibt dann eine Warnung aus.
- Bei mehreren MwSt-Sätzen muss der Satz je Position von Hand eingetragen werden.
- Als Einheit ist C62 (Stück) für Kasten und Fass eingestellt. Das lässt sich in den Einstellungen ändern.
