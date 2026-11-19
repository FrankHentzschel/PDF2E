# PDF2E 1.7.0 – PDF-Rechnungen als E-Rechnung (ZUGFeRD)

Kundenneutrales Werkzeug, Stand 04.11.2026. Erster Kunde ist die Krug-Bräu. Kundenspezifisch sind nur die Stammdaten-Blätter und das Rechnungslayout (Einstellung „Layout“, Einlese-Vorlage `modParserBSI` mit Kundenvariante als Parameter). Echte E-Rechnungen entstehen mit der Smart-Tools-Bibliothek (Adapter `SMARTTOOLS`). Im **Stub-Modus** (Adapter `STUB`) entsteht keine gültige E-Rechnung, er dient nur zum Testen ohne Bibliothek.

## Inhalt

PDF2E und XL2E teilen sich seit PDF2E 1.6.0 denselben Kern. Im ZIP liegen alle benötigten Module gemeinsam im Ordner `vba/`.

| Datei | Zweck |
| --- | --- |
| `PDF2E.xlsx` | Arbeitsmappe ohne Makros |
| `modProdukt.bas` | **nur PDF2E:** Name, Version, Dateiauswahl für PDF, Aufruf der Einlese-Vorlage |
| `modPdfText.bas` | **nur PDF2E:** ruft `pdftotext -layout` auf und liest den Text ein |
| `modParserBSI.bas` | **nur PDF2E:** Einlese-Vorlage für BSI-PDF-Rechnungen, Variante über die Einstellung „Layout-Variante“ |
| `Tabelle1.doccls` … | **nur PDF2E:** Blattmodule (leer bis auf `Option Explicit`) |
| `modAnwendung.bas` | gemeinsam: Splash beim Start, Info-Fenster, Schaltflächen anlegen (`PDF2E_Einrichten`) |
| `modHilfen.bas` | gemeinsam: Einstellungen, Feldzugriff, Zahlen, Datum, RegExp, Protokoll, Positionen anlegen |
| `modEinlesen.bas` | gemeinsam: Makro **RechnungEinlesen**, Empfänger und Pfandartikel anlegen |
| `modPruefen.bas` | gemeinsam: Pflichtangaben, Menge × Preis, Summenabgleich, GLN, GTIN, Platzhalter in Stammdaten |
| `modZugferd.bas` | gemeinsam: Makro **ZugferdErzeugen**, ruft den Adapter auf |
| `modAdapterStub.bas` | gemeinsam: Testlauf ohne Smart Tools, schreibt `<Ziel>.daten.xml` |
| `modAdapterSmartTools.bas` | gemeinsam: ZUGFeRD EN16931 erzeugen, prüfen, in die PDF einbetten |
| `modMail.bas` | gemeinsam: E-Mail-Entwurf in Outlook |
| `modDaten.bas` | gemeinsam: Backend (Access-Datei) für Einstellungen, Mailvorlage, Stammdaten und Verarbeitungsprotokoll |
| `modEntwicklung.bas` | gemeinsam, nur Entwicklung: älterer Export ohne VbaSync |
| `frmInfo.frm` + `frmInfo.frx` | gemeinsam: Splash- und Info-Fenster (UserForm) |
| `DieseArbeitsmappe.doccls` | gemeinsam: Code für „DieseArbeitsmappe“ (Start, Laden/Speichern der Stammdatenblätter, Versionszelle) |
| `starter/Start_PDF2E.xlsx` + `starter/vba/` | Starter-Mappe: holt neue Versionen, setzt platzbezogene Werte, startet PDF2E |

„Gemeinsam“ heißt: dieselbe Datei steht auch in XL2E.

## Mehrplatz- und Einzelplatzbetrieb (ab 1.7.0)

Ab 1.7.0 liegen alle Daten außerhalb der Programmmappe. Die Programmmappe enthält nur noch Code und Vorgabewerte und kann bei einem Update einfach ersetzt werden.

| Wo | Was |
| --- | --- |
| Backend `PDF2E_Daten.accdb` (zentral) | Einstellungen (auch abweichende Werte je Arbeitsplatz), Mailvorlage, Brauerei, Bankkonten, Zahlungskonditionen, Empfänger, Pfandartikel, Verarbeitungsprotokoll |
| Blatt „Arbeitsplatz“ der Programmmappe | platzbezogene Werte (Backend-Pfad, ggf. pdftotext-Pfad); setzt die Starter-Mappe bei jedem Start |
| Starter-Mappe `Start_PDF2E.xlsm` (lokal) | Produkt, Update-Ordner, lokaler Programmordner und die Werte für das Blatt „Arbeitsplatz“ |
| Update-Ordner | aktuelle `PDF2E.xlsm` (Mastermappe) |

Einzelplatz und Mehrplatz unterscheiden sich nur in den Pfaden. Beim Einzelplatz liegen Update-Ordner und Backend z. B. unter `C:\PDF2E\Update` und `C:\PDF2E\Daten`.

**So arbeitet es**
- Beim ersten Start legt PDF2E das Backend an und füllt es mit den Blättern der Mappe (Einstellungen, Stammdaten). Danach gilt das Backend.
- Wechselst du in ein Blatt wie „Empfaenger“, lädt PDF2E es frisch aus dem Backend. Vor jedem Einlesen ebenso.
- Eine Änderung in einer Zeile wird sofort gespeichert. Hat ein anderer Platz die Zeile inzwischen geändert, wird die Änderung abgelehnt und das Blatt neu geladen (Meldung mit Name und Zeit).
- Neue Zeilen unten anfügen. **Löschen nur über Rechtsklick > „Datensatz löschen (Datenbank)“.** Zeilen mit Excel zu löschen oder einzufügen lädt das Blatt neu.
- Hinweis: Nach einer Änderung in diesen Blättern gibt es kein Rückgängig (Strg+Z), weil das Programm in das Blatt schreibt.
- Wurde eine Rechnung schon erzeugt (an irgendeinem Platz), fragt PDF2E beim Einlesen nach und nennt wer und wann. Ebenso vor dem Überschreiben einer vorhandenen E-Rechnung.
- Nach einem Update übernimmt PDF2E neue Einstellungen und Mailvorlage-Zeilen der neuen Mappe ins Backend. Vorhandene Werte bleiben.
- Ist „Backend-Pfad“ leer, arbeitet PDF2E wie bis 1.6.0 nur mit den Blättern.

**Sicherung:** Einstellung „Backup-Ordner“ (darf auf einem anderen Rechner liegen) und „Backup behalten“ (Anzahl). Beim ersten Start des Tages kopiert PDF2E das Backend als `PDF2E_Daten_JJJJ-MM-TT_hhmmss.accdb` dorthin. Vor einer Änderung des Datenmodells sichert PDF2E immer (ohne Backup-Ordner in den Unterordner `Sicherung` neben dem Backend).

**Voraussetzungen**
- Access-Datenbank-Engine (ACE) auf jedem Platz, passend zur Office-Bitness. Test: Makro `DAO_Test` aus dem Konzept. Access selbst muss nicht installiert sein.
- Im Backend-Ordner brauchen alle Benutzer das Recht Ändern (Access legt dort die Sperrdatei `.laccdb` an).
- Backend nur im LAN betreiben, nicht über WLAN oder VPN (Abbrüche können die Datei beschädigen).
- Lokalen Programmordner (und Ordner der Starter-Mappe) im Trust Center als **vertrauenswürdigen Speicherort** eintragen.

### Einrichtung Mehrplatz (z. B. Krug)

1. Ordner auf dem Netzlaufwerk anlegen, z. B. `\\SERVER\PDF2E\Programm`, `\\SERVER\PDF2E\Daten`, `\\SERVER\PDF2E\E-Rechnungen`, `\\SERVER\PDF2E\Sicherung`.
2. Mastermappe `PDF2E.xlsm` (eingerichtet wie unten, Stammdaten gepflegt) nach `…\Programm` kopieren.
3. Starter-Mappe einmalig bauen: `starter/Start_PDF2E.xlsx` öffnen, als `Start_PDF2E.xlsm` speichern, Module aus `starter/vba` importieren (`modStarter.bas`, Code von `DieseArbeitsmappe.doccls`), kompilieren, Alt+F8 > `StarterEinrichten` (Schaltfläche „Jetzt starten“). Im Blatt „Starter“ Pfade eintragen. Diese Datei dient als Vorlage für alle Plätze.
4. Je Platz: `Start_PDF2E.xlsm` z. B. nach `C:\PDF2E\` kopieren, Verknüpfung auf den Desktop (Symbol `pdf2e.ico`), `C:\PDF2E` als vertrauenswürdigen Speicherort eintragen.
5. Ersten Platz starten: Der Starter kopiert die Programmmappe, setzt den Backend-Pfad, PDF2E legt das Backend an. Dann im Blatt „Einstellungen“ Ausgabeordner, Eingangsordner und Backup-Ordner prüfen (gelten ab jetzt für alle Plätze).
6. Vorhandene Stammdaten aus einer älteren Mappe übernehmen: Alt+F8 > `StammdatenAusMappeUebernehmen` (nur fehlende Einträge werden angelegt).

**Update:** neue `PDF2E.xlsm` in den Update-Ordner kopieren. Jeder Platz holt sie beim nächsten Start über den Starter (die bisherige lokale Mappe bleibt als `PDF2E_alt.xlsm`).

**Starter bearbeiten:** beim Öffnen Strg oder Umschalt gedrückt halten, dann startet nichts automatisch.

**Abweichende Einstellung an einem Platz** (z. B. eigener Eingangsordner): im Blatt „Einstellungen“ dieselbe Einstellung in einer weiteren Zeile eintragen und in Spalte **Arbeitsplatz** den Computernamen des Platzes (steht im Info-Fenster). Leer = gilt für alle Plätze. Ein Platz nimmt seine eigene Zeile, sonst die gemeinsame. Das wird zentral im Backend gespeichert, du kannst es also von jedem Platz aus pflegen. Reihenfolge: Blatt „Arbeitsplatz“ (Starter) vor Zeile mit Computername vor gemeinsamer Zeile.

## Update von 1.6.0 auf 1.7.0

1. VbaSync „Alle importieren“ aus `vba` (neu: `modDaten`; *DieseArbeitsmappe* bekommt neuen Code). Kompilieren.
2. Im Blatt „Einstellungen“ die Zeilen **Backup-Ordner** und **Backup behalten** ergänzen (sonst ohne Tagessicherung). Die Überschrift **Arbeitsplatz** in Spalte D ergänzt das Programm selbst.
3. Weiter mit der Einrichtung Mehrplatz oben. Ohne Starter: Blatt **Arbeitsplatz** anlegen (Spalten Einstellung / Wert / Hinweis) und dort `Backend-Pfad` eintragen; der Starter legt es sonst selbst an.

## Update von 1.5.0 auf 1.6.0

Die Module sind neu aufgeteilt, damit PDF2E und XL2E denselben Code nutzen. Die Stammdaten in der Mappe bleiben erhalten.

1. Vorhandene `PDF2E.xlsm` sichern (Kopie anlegen).
2. VBA-Editor (Alt+F11): das Modul **`modPDF2E` entfernen** (Rechtsklick > *Entfernen*, Export mit *Nein*). Es heißt jetzt `modAnwendung`; bleibt es stehen, meldet der Editor doppelte Namen.
3. Mit VbaSync „Alle importieren“ aus dem Ordner `vba` des ZIP (nicht aus einem alten Exportordner, der noch `modPDF2E.bas` enthält). Vorhandene Module werden ersetzt, neue (`modAnwendung`, `modProdukt`) hinzugefügt, *DieseArbeitsmappe* bekommt den neuen Code (`AnwendungStart`).
4. *Debuggen > Kompilieren VBAProject*.
5. Alt+F8 > `PDF2E_Einrichten` einmal ausführen. Die Schaltflächen werden neu angelegt.
6. Optional im Blatt „Einstellungen“ eine Zeile **Sammelkundennummern** ergänzen (siehe unten). Das Blatt „Pfandartikel“ entsteht beim nächsten Einlesen von selbst.

## Einrichtung (ca. 15 Minuten)

1. **pdftotext:** Lade auf https://www.xpdfreader.com/download.html die *Xpdf command line tools* für Windows und lege `pdftotext.exe` (64-Bit) z. B. unter `C:\Tools\xpdf\` ab.
2. **Mappe vorbereiten:** `PDF2E.xlsx` öffnen und als **Excel-Arbeitsmappe mit Makros (.xlsm)** speichern.
3. **Module importieren:** VBA-Editor mit Alt+F11 öffnen. Über *Datei > Datei importieren* alle `.bas`-Dateien aus `vba/` einlesen.
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
- Der Mailtext an REWE darf keine buchungsrelevanten Inhalte enthalten (Startpaket: „keine buchungsrelevanten Inhalte im E-Mail-Text“, ohne nähere Festlegung). Seit 1.6.0 bekommt REWE deshalb einen neutralen Text ohne Rechnungsdaten: Anrede, „anbei erhalten Sie unsere E-Rechnung im ZUGFeRD-Format“, Gruß. Der Text steht im Blatt „Mailvorlage“ in der Zeile **Text REWE**. Fehlt die Zeile (Mappen vor 1.6.0), gilt dieser Standard. Betrag, Fälligkeit und Zahlungsbedingung werden auch dort immer entfernt. Der Betreff bleibt wie in der Vorlage, z. B. „Rechnung 2603587 vom 27.09.2026“.
- Noch nicht umgesetzt: Bestellnummer (BT-13), Kostenstelle (BT-19), Konditionsschlüssel (OSI) und Profil EXTENDED. Die BSI-Rechnungen enthalten diese Angaben nicht.

## EDEKA-Empfänger und GTIN für Pfand

EDEKA verlangt laut Auskunft gegenüber der Brauerei (November 2026):

- Bei negativen Positionen sind Menge und Positionsbetrag negativ, der Preis positiv. So erzeugt PDF2E Pfand-Rücknahmen seit 1.5.0.
- Lieferung und Rückgabe stehen auch dann als zwei Positionen in der E-Rechnung, wenn sie gleich groß sind (Differenz 0). Im PDF darf die Zeile ohne Betrag bleiben.
- Pfandpositionen tragen eine **eigene GTIN** des Pfandartikels, nicht die des Warenartikels.

Umsetzung:

- Beim Einlesen wird ein Empfänger, dessen Name „EDEKA“ enthält, im Blatt „Empfaenger“ in der Spalte **Vorgaben** mit `EDEKA` markiert. Andere Händler, die über EDEKA abrechnen, dort von Hand markieren.
- Das Blatt **Pfandartikel** sammelt alle Leergut-Nummern aus den Rechnungen automatisch. Die GTIN trägt der Anwender einmal ein; sie steht nicht auf der BSI-Rechnung. Fehlt sie bei einem EDEKA-Empfänger, meldet die Prüfung einen Fehler.
- Weitere Vorgaben, etwa die EDEKA-Bestellnummer (BT-13), sind nicht umgesetzt: Sie stehen nicht auf der BSI-Rechnung.

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
