Attribute VB_Name = "modProdukt"
'==============================================================================
' PDF2E - Produktangaben und Datenquelle (PDF-Rechnungen aus der BSI-Software)
'
' Gegenstück in XL2E: src/xl2e/modProdukt.bas. Der gemeinsame Code in src/kern
' ruft nur die hier definierten Konstanten und Prozeduren auf.
'==============================================================================
Option Explicit

Public Const TOOL_NAME As String = "PDF2E"
Public Const TOOL_VERSION As String = "1.14.1"
Public Const TOOL_STAND As String = "26.11.2026"
Public Const TOOL_UNTERTITEL As String = "PDF-Rechnungen als E-Rechnung (ZUGFeRD)"
Public Const QUELLE_SCHALTFLAECHE As String = "PDF einlesen"

' Makros mit Produktnamen (bisherige Namen bleiben gültig)
Public Sub PDF2E_Start()
    AnwendungStart
End Sub

Public Sub PDF2E_Einrichten()
    Einrichten
End Sub

Public Sub PDF2E_Exportieren()
    Quellcode_Exportieren
End Sub

' Dateiauswahl für "PDF einlesen"
Public Function QuelleAuswaehlen() As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "BSI-Rechnung (PDF) auswählen"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "PDF-Dateien", "*.pdf"
        If Len(Einstellung("Eingangsordner")) > 0 Then .InitialFileName = Einstellung("Eingangsordner") & "\"
        If .Show = -1 Then QuelleAuswaehlen = .SelectedItems(1)
    End With
End Function

' Rechnung auslesen und in die Blätter "Rechnung" und "Positionen" schreiben
Public Sub QuelleEinlesen(ByVal pfad As String)
    Select Case UCase$(Einstellung("Layout"))
        Case "BSI": ParseBSI PdfZuText(pfad), Einstellung("Layout-Variante")
        Case Else: Err.Raise vbObjectError + 20, , "Unbekanntes Layout '" & Einstellung("Layout") & "' (Blatt Einstellungen)."
    End Select
End Sub

' PDF, in das die E-Rechnung eingebettet wird: Kopie der Original-PDF
Public Sub QuellPdfErstellen(ByVal quelle As String, ByVal ziel As String)
    FileCopy quelle, ziel          ' Original-PDF bleibt unverändert
End Sub

' Versionshistorie, neueste zuerst. Wird beim Start ins Blatt "Versionshistorie"
' geschrieben (modAnwendung.VersionshistorieAktualisieren). Bei jeder neuen Version
' oben eine Zeile ergänzen und TOOL_VERSION anpassen.
Public Function Versionen() As Collection
    Dim c As New Collection
    c.Add Array("1.14.1", "26.11.2026", "Fehlerbehebung", "Kompilierfehler „Erwartet: Datenfeld“ in modPdfA: Variable mId verdeckte die VBA-Funktion Mid (umbenannt in mTrailerId); Strukturtest prüft jetzt solche Namenskonflikte")
    c.Add Array("1.14.0", "26.11.2026", "Neue Funktion", "PDF/A-Korrektur nach dem Speichern (neues Modul modPdfA): Briefpapier der BSI-Rechnung in CMYK mit Überdrucken machte die E-Rechnung zu einem ungültigen PDF/A-3 (Meldung Getränke Ziegler: Überdrucken-Modus OPM 1 bei ICC-basiertem CMYK); Überdrucken-Modus wird auf 0 gesetzt und ein CMYK-Farbprofil als DefaultCMYK eingetragen, abschaltbar mit Einstellung PDF/A-Korrektur = Nein. Registerangaben der Brauerei (Sitz, Handelsregister, Geschäftsführer) als Freitext mit Code REG (BT-22), Hinweis VD-Valitool-71; Text aus dem Blatt Brauerei (optional Zeilen Sitz und Registerangaben)")
    c.Add Array("1.13.0", "26.11.2026", "Neue Funktion", "Warenempfänger-GLN zusätzlich als Standard je Kunde (neue Spalte Warenempfänger-GLN im Blatt Empfaenger), Vorrang hat die GLN im Blatt Lieferorte; GLN nur bei Kunden, die sie verlangen (Vorgaben GLN, REWE, EDEKA oder Warenempfänger-GLN gepflegt): nur für sie werden Lieferorte angelegt und GLN-Hinweise gemeldet; Backend Schema-Version 4")
    c.Add Array("1.12.0", "26.11.2026", "Neue Funktion", "Lieferfiliale unter dem Nettowert (z. B. Getränke Hoffmann) als Lieferort; Warenempfänger-GLN je Lieferort: neues Blatt Lieferorte (wird beim Einlesen automatisch gefüllt, GLN einmal eintragen, Prüfen übernimmt sie auch nachträglich), Lieferort-GLN als BT-71; Hinweis bei fehlender Lieferort-GLN und fehlender eigener GLN der Brauerei (BT-29, Blatt Brauerei); Backend Schema-Version 3")
    c.Add Array("1.11.7", "25.11.2026", "Fehlerbehebung", "Prüfung übernimmt Angaben, die nach dem Einlesen im Blatt Empfaenger nachgetragen wurden (USt-IdNr., E-Mail, Konto, Kondition), in leere Felder des Blatts Rechnung; bisher war neues Einlesen nötig")
    c.Add Array("1.11.6", "25.11.2026", "Fehlerbehebung", "RE 2604315 (Barverkauf): Position mit negativer Menge und * (59 Rabatt 1- 15,00-*) erkannt; bei Bruttobeträgen werden nur echte Rundungsdifferenzen (max. 1 Cent je Position) verteilt, größere Differenzen bleiben als Fehler sichtbar")
    c.Add Array("1.11.5", "25.11.2026", "Fehlerbehebung", "RE 2604314 (Italien): Hinweis „Steuerfreie Lieferung nach Italien ab Brauereihof“ und „UST - ID:“ erkannt (bisher nur „Steuerfreie innergemeinschaftliche Lieferung“ / „USt-ID:“); Kategorie K nur bei EU-Land, sonst klare Meldung")
    c.Add Array("1.11.4", "25.11.2026", "Fehlerbehebung", "Bruttobeträge: Rundungsdifferenz wird centweise auf Positionen verteilt, damit Nettosumme, MwSt und Zahlbetrag genau der Rechnung entsprechen; Rundungsbetrag BT-114 nur noch als Rückfall")
    c.Add Array("1.11.3", "25.11.2026", "Fehlerbehebung", "RE 2603531 (Barverkauf ab Rampe): Summenzeile RECHNUNGSBETRAG (incl. ... Pfand) EUR ... mit Klammer erkannt; Telefonnummer im Anschriftenblock nicht mehr als Teil des Namens")
    c.Add Array("1.11.2", "25.11.2026", "Fehlerbehebung", "Bruttobeträge: Nettopreis wurde in Excel auf 2 Stellen gerundet (11,76 statt 11,7647), dadurch Fehler Menge x Preis; Preis jetzt mit 4 Stellen")
    c.Add Array("1.11.1", "25.11.2026", "Fehlerbehebung", "Einlesen nach Import neuer Module ohne Neustart der Mappe brach mit „Element in dieser Auflistung nicht gefunden“ ab (Backend noch auf altem Schema, Spalte ZRNummer fehlte). Schema-Prüfung/Migration und neue Einstellungen jetzt auch vor jedem Laden der Stammdaten")
    c.Add Array("1.11.0", "25.11.2026", "Neue Funktion", "Minus-Rechnung mit Rückware und Bruttobeträgen (ENTHALTENE MWSt., RE 2603475): Summenblock MINUS-RECHNUNGSBETRAG erkannt, Rückware mit negativer Menge, Preise und Pfand in netto umgerechnet, Rundungsdifferenz als Rundungsbetrag (BT-114); Leergutrückgabe-Zeilen ohne Meldung; Rechnungsbetrag im Verarbeitungsprotokoll mit Cent")
    c.Add Array("1.10.0", "25.11.2026", "Neue Funktion", "G5-Zentralregulierung (SGH Service): Vorgabe G5 je Empfänger, ZR-Nummer des Händlers als Käuferreferenz (BT-10, neue Spalte ZR-Nummer im Blatt Empfaenger), G5-Lieferantennummer der Brauerei als Verkäuferkennung (BT-29, Einstellung G5 Lieferantennummer), E-Mail an die zentrale Adresse (Einstellung G5 E-Mail), Prüfung beider Nummern; Backend Schema-Version 2")
    c.Add Array("1.9.1", "23.11.2026", "Fehlerbehebung", "Positionen ohne Hl./Fl.-Angabe und mit * am Zeilenende (z. B. EURO-Palette) wurden nicht übernommen; Ort ohne Zusatz (Deutschland)")
    c.Add Array("1.9.0", "23.11.2026", "Neue Funktion", "Steuerfreie innergemeinschaftliche Lieferung: USt-IdNr. des Empfängers aus der Rechnung, USt-Kategorie K mit 0 % und Befreiungsgrund VATEX-EU-IC, Lieferland aus der Empfängeranschrift; ausländische Anschriften mit Länderkennzeichen vor der PLZ (z. B. I- 39012 Meran); WARENWERT ohne Betrag (reine Leergutrückgabe); klare Meldung, wenn weder MwSt noch Steuerbefreiung erkannt wird")
    c.Add Array("1.8.3", "23.11.2026", "Fehlerbehebung", "Kompilierfehler: Konstante UST_PRAEFIXE stand zwischen den Prozeduren statt im Deklarationsteil von modHilfen")
    c.Add Array("1.8.2", "23.11.2026", "Fehlerbehebung", "Kundenname in der Spalte USt-IdNr. (Blatt Empfaenger): wird beim Einlesen erkannt und geleert, neue Empfänger bekommen nur eine gültige USt-IdNr.; Prüfung der Feldnamen im Blatt Rechnung (zwei Namen auf derselben Zelle) mit klarer Meldung")
    c.Add Array("1.8.1", "23.11.2026", "Fehlerbehebung", "USt-IdNr. von Brauerei und Rechnungsempfänger wird bereinigt (Leerzeichen, Punkte, Bindestriche, Großbuchstaben) und vor dem Erzeugen auf Länderkennzeichen geprüft (EN 16931 BR-CO-09) mit Hinweis, wo zu korrigieren ist; Prüfung schon bei der Eingabe im Blatt Empfaenger")
    c.Add Array("1.8.0", "23.11.2026", "Neue Funktion", "Skonto- bzw. Nachlasszeile unter dem Warenwert (z. B. Skonto 3 % auf den Warenwert) wird als Nachlass auf Belegebene übernommen, Abgleich mit NETTOWERT; SEPA-Lastschrift aus dem Rechnungstext (Mandat, Gläubiger-ID, Einzugsdatum) als Zahlungsart 59 in der E-Rechnung, Einzugsdatum als Fälligkeit, Einzugshinweis als Zahlungsbedingung (auch in der E-Mail)")
    c.Add Array("1.7.2", "19.11.2026", "Fehlerbehebung", "Blatt Versionshistorie wird beim Start aus dem Programm nachgeführt (fehlte ab 1.5.0, weil nur die Vorlage es befüllte); Datum als echtes Datum statt Text")
    c.Add Array("1.7.1", "19.11.2026", "Fehlerbehebung", "Änderungen im Blatt Einstellungen wurden von Access abgelehnt (Pflichtfeld Arbeitsplatz); Speichern geänderter Zeilen jetzt über das Recordset. Veraltete Einstellungen Entwickler/Firma/Kontakt werden entfernt")
    c.Add Array("1.7.0", "19.11.2026", "Neue Funktion", "Mehrplatzbetrieb: Einstellungen, Mailvorlage und Stammdaten im Access-Backend (wird automatisch angelegt und migriert), Laden beim Aktivieren eines Blatts, Speichern pro Zeile mit Konfliktprüfung, Datensatz löschen über das Kontextmenü; Verarbeitungsprotokoll mit Rückfrage bei bereits erzeugten Rechnungen; tägliche Sicherung des Backends; Übernahme von Stammdaten aus einer älteren Mappe; Blatt Arbeitsplatz für platzbezogene Werte; Version in Start!Z1 für die Starter-Mappe")
    c.Add Array("1.6.0", "17.11.2026", "Neue Funktion", "Pfand mit Differenz 0 als Lieferung und Rückgabe; Blatt Pfandartikel mit GTIN; EDEKA-Vorgaben; gemeinsame Module mit XL2E (Module neu aufgeteilt: modAnwendung, modProdukt); Sammelkundennummern; Prüfung warnt vor Platzhaltern in den Stammdaten; E-Mail an REWE mit neutralem Text ohne Rechnungsdaten; Absender der E-Mail wählbar, auch gemeinsames Postfach")
    c.Add Array("1.5.0", "04.11.2026", "Neue Funktion", "Smart-Tools-Adapter (ZUGFeRD EN16931, Prüfung vor dem Speichern); Pfand getrennt nach Lieferung und Rücknahme; REWE-Vorgaben je Empfänger (Hinweis AAK, Pflicht RE-GLN, WE-GLN, USt-IdNr.); GLN-Prüfziffer")
    c.Add Array("1.4.0", "26.10.2026", "Neue Funktion", "E-Mail-Entwurf in Outlook mit der E-Rechnung als Anhang, Text aus dem Blatt Mailvorlage (Platzhalter)")
    c.Add Array("1.3.0", "26.10.2026", "Neue Funktion", "Firmenlogo im Info-Fenster standardmäßig links, wahlweise rechts (Einstellung ""Firmenlogo-Position"")")
    c.Add Array("1.2.0", "26.10.2026", "Neue Funktion", "Neue Rechnungsempfänger werden automatisch im Blatt Empfaenger angelegt; Fehlerbehebung: umgebrochene Leergutzeilen der echten Windows-Ausgabe (Pfandteil vor der Artikelzeile) werden spaltengenau zusammengeführt")
    c.Add Array("1.1.0", "26.10.2026", "Neue Funktion", "Info-Fenster: Anwendungslogo oben, Firmenlogo des Entwicklers unten (Einstellungen ""Anwendungslogo"" und ""Firmenlogo"")")
    c.Add Array("1.0.3", "26.10.2026", "Fehlerbehebung", "Fehler beim Laden des Logos werden im Protokoll gemeldet")
    c.Add Array("1.0.2", "26.10.2026", "Fehlerbehebung", "Schaltflächen als Formen statt Formularsteuerelemente (Anzeigefehler); Versionszeile im Blatt Start wird automatisch nachgeführt")
    c.Add Array("1.0.1", "26.10.2026", "Fehlerbehebung", "Umgebrochene Leergutzeilen (pdftotext unter Windows) werden zusammengeführt; Option ""Textdatei behalten""")
    c.Add Array("1.0.0", "26.10.2026", "Erstversion", "Einlesen BSI-PDF, Prüfung, Stub-Adapter, Splash- und Info-Fenster, Schaltflächen")
    Set Versionen = c
End Function
