Attribute VB_Name = "modZugferd"
'==============================================================================
' ZUGFeRD erzeugen: PDF bereitstellen (modProdukt.QuellPdfErstellen),
' XML per Smart Tools einbetten und validieren
'
' Die eigentliche Erzeugung übernimmt ein Adapter (Blatt Einstellungen,
' "E-Rechnung-Adapter"):
'   STUB        modAdapterStub: lässt die PDF unverändert und schreibt die Daten als XML
'               daneben. KEINE gültige E-Rechnung, nur zum Testen des Ablaufs.
'   SMARTTOOLS  modAdapterSmartTools: Smart-Tools-Bibliothek, Profil EN16931
'==============================================================================
Option Explicit

Public Sub ZugferdErzeugen()
    Dim quelle As String, ziel As String, adapter As String, info As String
    Dim nr As Long, txt As String

    If Not RechnungPruefen() Then
        MsgBox "Die Prüfung meldet Fehler. Bitte das Blatt 'Protokoll' ansehen.", vbExclamation
        ThisWorkbook.Worksheets(BLATT_PROTOKOLL).Activate
        Exit Sub
    End If

    adapter = UCase$(Einstellung("E-Rechnung-Adapter"))
    quelle = Feld("f_Quelldatei").Value
    ziel = Einstellung("Ausgabeordner") & "\" & Feld("f_RechNr").Value & Einstellung("Dateinamenszusatz")
    If adapter = "STUB" Then ziel = ziel & "_STUB"
    ziel = ziel & ".pdf"
    On Error GoTo Fehler
    If Len(Dir$(ziel)) > 0 Then
        info = VerarbeitungInfo(CStr(Feld("f_RechNr").Value))      ' Mehrplatz: wer hat sie erzeugt?
        If Len(info) = 0 And DatenAktiv() Then info = "Herkunft der Datei unbekannt."
        If MsgBox("Die Datei existiert bereits:" & vbLf & ziel & vbLf & vbLf & IIf(Len(info) > 0, info & vbLf & vbLf, "") & _
                  "Überschreiben?", vbYesNo + vbQuestion + vbDefaultButton2, TOOL_NAME) <> vbYes Then Exit Sub
        DateiLoeschen ziel          ' wartet, falls die Datei auf dem Netzlaufwerk noch gesperrt ist
    End If

    QuellPdfErstellen quelle, ziel ' PDF2E: Kopie der Original-PDF, XL2E: PDF aus der Excel-Rechnung
    Select Case adapter
        Case "STUB": StubEinbetten ziel
        Case "SMARTTOOLS": SmartToolsEinbetten ziel
        Case Else: Err.Raise vbObjectError + 30, , "Unbekannter E-Rechnung-Adapter '" & adapter & "' (Blatt Einstellungen)."
    End Select
    ' Kontrolle: Datei muss danach noch da sein (Netzlaufwerk, Virenscanner)
    If Len(Dir$(ziel)) = 0 Then Err.Raise vbObjectError + 31, , "Die E-Rechnung ist nach dem Speichern nicht mehr im Zielordner. " & _
        "Vermutlich hat ein anderes Programm (Virenscanner, Synchronisation) sie entfernt oder gesperrt."
    Protokoll "Info", "Gespeichert: " & ziel & " (" & Format$(FileLen(ziel), "#,##0") & " Byte)"
    Feld("f_Ausgabedatei").Value = ziel
    Protokoll "Info", "ZUGFeRD erzeugt: " & ziel
    VerarbeitungErzeugt CStr(Feld("f_RechNr").Value), quelle, ziel, CStr(Feld("f_KName").Value), _
                        Zahlbetrag(), IIf(adapter = "STUB", "Test (STUB)", "erzeugt")
    MsgBox IIf(adapter = "STUB", "TESTLAUF (Stub) - keine gültige E-Rechnung:", "ZUGFeRD-Rechnung erzeugt:") & vbLf & ziel, vbInformation
    Exit Sub

Fehler:
    nr = Err.Number: txt = Err.Description
    Resume Aufraeumen               ' Fehlerbehandlung beenden, damit Aufräumen keinen zweiten Laufzeitfehler auslöst
Aufraeumen:
    If nr = 70 Or nr = 75 Then txt = txt & " (Fehler " & nr & "). Die Datei ist vermutlich noch geöffnet oder wird gerade gelöscht " & _
        "(PDF-Anzeige, Explorer-Vorschau, Virenscanner, Excel selbst). Bitte einige Sekunden warten und erneut versuchen; " & _
        "hilft das nicht, Excel ganz schließen und neu starten."
    Protokoll "Fehler", txt
    On Error Resume Next
    If Len(ziel) > 0 Then DateiLoeschen ziel ' umbenennen und löschen, damit kein gesperrter Rest den Namen blockiert
    On Error GoTo 0
    MsgBox "Erzeugen fehlgeschlagen:" & vbLf & txt, vbExclamation
End Sub
