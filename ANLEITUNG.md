# PDF2E 1.12.0 – PDF-Rechnungen als E-Rechnung (ZUGFeRD)

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
- Beim Start lädt PDF2E zuerst Einstellungen und Brauerei-Stammdaten aus dem Backend, danach erscheint der Splash. Er zeigt also den Lizenznehmer aus dem Backend, auch wenn die Programmmappe gerade neu ist. Ist das Backend nicht erreichbar, kommt eine Meldung und der Splash zeigt „(Stammdaten nicht verfügbar)“ bzw. den zuletzt geladenen Stand.
- Kundenneutrale Mastermappe: Bleiben die Stammdatenblätter der Mastermappe leer, wird ein neues Backend leer angelegt; die Brauerei-Daten dann einmal im Blatt „Brauerei“ eintragen.

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

## Update auf 1.12.0 (GLN je Lieferort)

1. **Vorher an allen Plätzen PDF2E schließen.** Beim ersten Start stellt die neue Version das Backend auf Schema-Version 3 um (neue Tabelle Lieferorte), mit Sicherung vorher.
2. VbaSync „Alle importieren“ aus `vba`. Kompilieren.
3. Blatt „Brauerei“, Zeile **GLN**: eigene GLN der Brauerei eintragen, falls vorhanden. Das Blatt **Lieferorte** legt das Programm selbst an.

## Update auf 1.11.x (Rückware, Bruttobeträge)

1. VbaSync „Alle importieren“ aus `vba`. Kompilieren. Danach die Mappe speichern, schließen und neu öffnen (bzw. über die Starter-Mappe starten), damit der Programmstart läuft: Backend-Umstellung, neue Einstellungen, Versionsnummer im Startblatt. Ab 1.11.1 holt „PDF einlesen“ die Backend-Umstellung notfalls selbst nach.
2. Sonst nichts. Das neue Feld **Rundungsbetrag** hängt das Programm beim ersten Einlesen unten im Blatt „Rechnung“ an.

## Update auf 1.10.0 (G5-Zentralregulierung)

1. **Vorher an allen Plätzen PDF2E schließen.** Die neue Version stellt das Backend beim ersten Start auf Schema-Version 2 um (neue Spalte ZR-Nummer). Das geht nur, wenn kein anderer Platz das Backend offen hat; vorher legt das Programm eine Sicherung an. Plätze mit der alten Version melden danach, dass sie über die Starter-Mappe neu starten sollen.
2. VbaSync „Alle importieren“ aus `vba`. Kompilieren.
3. Im Blatt „Einstellungen“ die **G5 Lieferantennummer** der Brauerei eintragen (siehe „G5-Empfänger“ unten). Die Zeilen „G5 Lieferantennummer“ und „G5 E-Mail“ übernimmt das Programm beim ersten Start selbst ins Backend.

## Update von 1.7.x auf 1.8.x und 1.9.0

1. VbaSync „Alle importieren“ aus `vba`. Kompilieren.
2. Sonst nichts. Die neuen Felder **Mandatsreferenz**, **Gläubiger-ID** und **Einzugsdatum** hängt das Programm beim ersten Einlesen unten im Blatt „Rechnung“ an (Abschnitt „SEPA-Lastschrift“).

## Update von 1.6.0 auf 1.7.x

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
   - Anwendungslogo (oben) und Firmenlogo (unten), jeweils BMP, JPG oder GIF, kein PNG. Entwickler, Firma und Kontakt sind ab 1.7.0 fest im Programm; „Lizenziert für“ zeigt den Namen aus dem Blatt „Brauerei“.
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

## G5-Empfänger (ab 1.10.0)

Die G5 Getränkefachhandels-Kooperation lässt ihre Zentralregulierung von der SGH Service GmbH abwickeln. Laut Merkblatt (Stand 11/2026) und E-Rechnung-Guideline EN 16931 V1.1 gilt ab 01.12.2026:

- E-Rechnungen (ZUGFeRD ab EN16931) gehen ohne Onboarding an **g5.e-invoice@documents.sgh-net.de**, jeder Beleg nur einmal. Reine PDF-Belege dürfen noch an eine bestehende individuelle Adresse gehen.
- Pflicht im XML: **ZR-Nummer** des Händlers als Käuferreferenz (BT-10, 7 Ziffern, beginnt mit 27) und die **G5-Lieferantennummer** der Brauerei als Verkäuferkennung (BT-29, 7 Ziffern, beginnt mit 27, endet auf 0). Eine Angabe nur im PDF reicht nicht.
- Nicht konforme Belege lehnt das System automatisch ab, die Meldung geht an die Absenderadresse der Mail.

Umsetzung:

- Im Blatt „Empfaenger“ bei jedem G5-Anschlusshaus in der Spalte **Vorgaben** `G5` und in der Spalte **ZR-Nummer** dessen ZR-Nummer eintragen. Automatisch erkannt wird ein G5-Kunde nicht, weil das nicht auf der BSI-Rechnung steht.
- Im Blatt „Einstellungen“ die **G5 Lieferantennummer** eintragen. Die Einstellung **G5 E-Mail** enthält die zentrale Adresse; der E-Mail-Entwurf geht bei G5-Empfängern immer dorthin, auch wenn beim Empfänger eine andere Adresse steht.
- Die Prüfung meldet einen Fehler, wenn eine der beiden Nummern fehlt oder falsch aufgebaut ist.
- Nicht umgesetzt: Hinweis auf Entgeltminderung (AAK mit Code EGM), Skonto im Format `#SKONTO#…#` und Gutschriften (Rechnungsart 381). Das ist erst nötig, wenn die Brauerei mit der G5 Rabatte/Boni oder Skonto vereinbart hat bzw. eigene Gutschriftsbelege erzeugt.

## GLN von Brauerei und Lieferort (ab 1.12.0)

Händler mit automatischer Verarbeitung (z. B. Getränke Hoffmann, REWE) ordnen die E-Rechnung über GLNs zu:

- **Lieferanten-GLN** (eigene GLN der Brauerei): Blatt „Brauerei“, Zeile **GLN**. Geht als BT-29 mit Schema 0088 in die E-Rechnung. Einmal eintragen, gilt für alle Rechnungen.
- **Warenempfänger-GLN** (GLN des Lieferorts, z. B. der Filiale): steht nicht auf der BSI-Rechnung. Beim Einlesen legt PDF2E jeden Lieferort ohne GLN im Blatt **Lieferorte** an, Schlüssel `Kundennr. / PLZ Ort, Straße`. Dort einmal die GLN eintragen, dann „Prüfen“ (oder neu einlesen); sie geht als BT-71 mit Schema 0088 in die E-Rechnung. Steht eine GLN schon im Rechnungstext („GLN:“, „Wir lieferten an:“), gilt diese.
- Als Lieferort gelten „Wir lieferten an:“ und, seit 1.12.0, die Zeile „+ Lieferfiliale: PLZ Ort, Straße“ unter dem Nettowert.
- Die Prüfung warnt, wenn ein Lieferort keine GLN hat oder die Brauerei-GLN fehlt.

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
| 2604264 | Groh & Holzner, Meran (IT), innergemeinschaftlich | nur Leergutrückgabe | −1.810,70 | −1.810,70, ohne MwSt |
| 2604307 | PART Trading, Hamburg | 4 + EURO-Palette (ohne Hl., mit *) | 0,00 | 15.598,52 |
| 2604304 | SAGASSER Fachmarkt, Lieferort Ebermannstadt | 4 + Skonto 3 % (−8,99) | 34,20 | 386,70, SEPA-Lastschrift |
| 2603531 | Barverkauf ab Rampe (Telefonnummer im Anschriftenblock) | 4 | 43,40 (brutto) | 244,40, enthaltene MwSt |
| 2603475 | J & J Festbedarf, Minus-Rechnung | 9 × Rückware | −440,50 (brutto) | −772,50, enthaltene MwSt |

Seit 1.5.0 steht Pfand getrennt nach Lieferung und Rücknahme in der Rechnung. Das Programm leitet beide Mengen aus dem Leergutkonto ab (Lieferung − Rücknahme = Differenz, alter Saldo + Differenz = neuer Saldo). Gelingt das nicht eindeutig, bleibt es bei der Differenz mit einer Warnung im Protokoll.

Im Ordner `tests/pdftotext` liegen die pdftotext-Ausgaben dieser Muster. Dateien mit `_windows` stammen aus pdftotext unter Windows (Xpdf), die übrigen aus Linux (Poppler). Die Windows-Version verteilt manche Leergutzeilen auf zwei Zeilen; der Parser fügt sie spaltengenau zusammen.

Im Stub-Modus legt **E-Rechnung erzeugen** zwei Dateien an: `<RechNr>_ZUGFeRD_STUB.pdf`, eine unveränderte Kopie, und daneben `.daten.xml` mit allen Übergabewerten.

## USt-IdNr. (ab 1.8.1)

Die E-Rechnung verlangt bei jeder USt-IdNr. das Länderkennzeichen (EN 16931, Regel BR-CO-09), z. B. `DE812706034`. PDF2E entfernt Leerzeichen, Punkte, Bindestriche und Schrägstriche und schreibt Großbuchstaben. Fehlt das Länderkennzeichen oder passt die Form nicht (deutsch: DE + 9 Ziffern), meldet die Prüfung einen Fehler mit dem Hinweis, wo zu korrigieren ist (Blatt „Empfaenger“ mit Kundennummer bzw. Blatt „Brauerei“). Schon bei der Eingabe im Blatt „Empfaenger“ (Spalte USt-IdNr.) wird der Wert bereinigt und bei falscher Form ein Hinweis angezeigt. Eine Steuernummer gehört nicht in dieses Feld; ist keine USt-IdNr. bekannt, bleibt es leer (außer bei REWE, dort ist sie Pflicht).

Ab 1.8.2: Steht in der Spalte USt-IdNr. der Kundenname, leert das Programm das Feld beim nächsten Einlesen und meldet es im Protokoll. Andere ungültige Werte werden gemeldet. Neu angelegte Empfänger bekommen nur eine USt-IdNr., die wie eine aussieht. Außerdem prüft das Programm vor jedem Einlesen, ob jeder Feldname im Blatt „Rechnung“ (f_KName, f_KUstId …) auf eine eigene Zelle zeigt. Wenn nicht, bricht es mit einer Meldung ab, die die betroffenen Namen nennt.

## Innergemeinschaftliche Lieferung (ab 1.9.0)

Steht auf der Rechnung „Steuerfreie innergemeinschaftliche Lieferung“ oder „Steuerfreie Lieferung …“ (z. B. „nach Italien ab Brauereihof“) und liegt die Anschrift in einem anderen EU-Land, setzt PDF2E alle Positionen auf USt-Kategorie **K** mit 0 %. Die E-Rechnung bekommt den Befreiungsgrund „Steuerfreie innergemeinschaftliche Lieferung“ mit Code VATEX-EU-IC (BT-120/121). Die Zeile `USt-ID: IT…` (auch `UST - ID: IT…`) wird als USt-IdNr. des Empfängers (BT-48) übernommen. Ein Eintrag im Blatt „Empfaenger“ hat Vorrang. Ausländische Anschriften mit Länderkennzeichen vor der PLZ (z. B. `I- 39012 Meran`, `A-6020 Innsbruck`) werden erkannt, das Land kommt ins Feld „Land“. Fehlt „Wir lieferten an“, gilt die Anschrift des Empfängers als Lieferort, denn das Lieferland ist bei K Pflicht (BR-IC-12). Die Prüfung verlangt dann eine USt-IdNr. des Empfängers und ein Land ungleich DE.

Findet PDF2E weder eine MwSt-Zeile noch einen Hinweis auf Steuerbefreiung, steht das im Protokoll. USt-Satz und Kategorie müssen dann im Blatt „Positionen“ von Hand eingetragen werden.

## Rückware und Bruttobeträge (ab 1.11.0)

**Rückware:** Positionen mit negativem Betrag (`5 Bernstein Radler 0,5 / P 5  4  80  14,000  56,00-` mit der Zeile `** Rückware **` darunter) bekommen eine negative Menge. Die Zeilen unter „Leergutrückgabe:“ im Positionsteil (nur Mengen) werden ohne Meldung übergangen, das Pfand steht im LEERGUT-KONTO.

**Bruttobeträge:** Steht unter dem Rechnungsbetrag `ENTHALTENE MWSt. 19,00 % (aus 772,50-) 123,34-` statt `+ 19,00 % MwSt (auf …)`, sind alle Preise, das Pfand und die Summen brutto. Die E-Rechnung verlangt Nettopreise. PDF2E rechnet deshalb jeden Preis auf 4 Nachkommastellen netto um und setzt den Betrag auf Menge × Nettopreis. Weil jede Zeile für sich gerundet wird, weicht die Summe um wenige Cent von der Nettosumme der Rechnung (Rechnungsbetrag − enthaltene MwSt) ab (RE 2603475: −649,17 statt −649,16). Ab 1.11.4 verteilt PDF2E diese Cents auf die Positionen mit dem größten Rundungsrest und setzt deren Nettopreis passend (z. B. Pfand Gourmet 2,6000 statt 2,6050). Nettosumme, MwSt und Zahlbetrag der E-Rechnung entsprechen damit genau der Rechnung. Nur wenn sich ein Cent so nicht verteilen lässt, steht er im Feld **Rundungsbetrag** und geht als BT-114 in die E-Rechnung. Die Prüfung vergleicht Warenwert und Pfand brutto mit bis zu 1 Cent Toleranz je Position (Hinweis im Protokoll). Der Rechnungsbetrag muss genau stimmen. Bruttobeträge mit mehreren MwSt-Sätzen werden nicht unterstützt.

Den Summenblock `MINUS-RECHNUNGSBETRAG incl. 440,50- Pfand) EUR 772,50-` liest PDF2E für Pfand und Rechnungsbetrag. Die Rechnung bleibt Rechnungsart 380 mit negativen Beträgen.

## Skonto und SEPA-Lastschrift (ab 1.8.0)

**Skonto/Nachlass unter dem Warenwert:** Eine Zeile wie `Skonto  Gesamtes Vergütung  299,75 EU  3,00  8,99-` zwischen WARENWERT und NETTOWERT wird als Position vom Typ „Nachlass“ übernommen (Bezeichnung „Skonto 3,00 % auf 299,75 EUR“). In der E-Rechnung wird daraus ein Abschlag auf Belegebene (BG-20) mit dem USt-Satz der Rechnung, wie bei den Aktionen. Das Pfand bleibt unberührt. Der Abgleich „Warenwert“ nutzt dann den NETTOWERT aus dem PDF. Stimmen Prozent × Basis oder die Basis mit dem Warenwert nicht überein, steht eine Warnung im Protokoll.

**SEPA-Lastschrift:** Steht unter dem Rechnungsbetrag „… ziehen wir mit einer SEPA-Lastschrift auf Grund des Mandates … mit der Glaeubiger-ID … am … ein.“, übernimmt das Programm Mandatsreferenz, Gläubiger-ID und Einzugsdatum in den Abschnitt „SEPA-Lastschrift“ im Blatt „Rechnung“. Die E-Rechnung bekommt dann Zahlungsart 59 (SEPA-Lastschrift) mit Mandatsreferenz (BT-89) und Gläubiger-ID (BT-90), das Einzugsdatum als Fälligkeit (BT-9) und als Zahlungsbedingung (BT-20) einen Einzugshinweis statt des Textes der Zahlungskondition. Die Bankverbindung der Brauerei entfällt in diesem Fall. Die IBAN des Kunden (BT-91) steht nicht auf der Rechnung und wird nicht übergeben. Der Platzhalter `{Faellig}` und `{Zahlungsbedingung}` in der E-Mail richten sich ebenfalls nach der Lastschrift.

## Noch nicht getestet

- 1.11.0: Rückware und Bruttobeträge sind nur mit der Linux-Ausgabe von pdftotext (RE 2603475) geprüft und noch nicht in Excel kompiliert. `Totals.RoundingAmount` ist aus der Typbibliothek `SmartTools.Invoicing.tlb` gelesen, aber noch nicht mit Smart Tools ausprobiert.
- 1.8.0: Skonto und SEPA-Lastschrift sind nur mit der Linux-Ausgabe von pdftotext geprüft und noch nicht in Excel kompiliert. Die Namen in der Smart-Tools-Bibliothek (`PaymentMeansCode_SEPA_direct_debit`, `Payment.DirectDebit.MandateReferenceId`, `.CreditorId`) sind aus der installierten `SmartTools.Invoicing.dll` gelesen.

- Die neue Modulaufteilung (1.6.0) ist noch nicht in Excel kompiliert worden. Die Logik ist unverändert gegenüber 1.5.0, die am 04.11.2026 mit RE 2603587 eine gültige E-Rechnung erzeugt hat.
- `VBScript.RegExp`: Microsoft kündigt die schrittweise Abschaltung von VBScript an. Ob und wann das RegExp in VBA betrifft, ist offen.

## Bekannte Grenzen

- Bei mehreren Lieferscheinen pro Rechnung wird nur der erste übernommen. Das Protokoll gibt dann eine Warnung aus.
- Bei mehreren MwSt-Sätzen muss der Satz je Position von Hand eingetragen werden.
- Als Einheit ist C62 (Stück) für Kasten und Fass eingestellt. Das lässt sich in den Einstellungen ändern.
