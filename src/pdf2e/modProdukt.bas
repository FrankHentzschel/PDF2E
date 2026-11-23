Attribute VB_Name = "modProdukt"
'==============================================================================
' PDF2E - Produktangaben und Datenquelle (PDF-Rechnungen aus der BSI-Software)
'
' Gegenstück in XL2E: src/xl2e/modProdukt.bas. Der gemeinsame Code in src/kern
' ruft nur die hier definierten Konstanten und Prozeduren auf.
'==============================================================================
Option Explicit

Public Const TOOL_NAME As String = "PDF2E"
Public Const TOOL_VERSION As String = "1.8.2"
Public Const TOOL_STAND As String = "23.11.2026"
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
