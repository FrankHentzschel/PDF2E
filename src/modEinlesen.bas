Attribute VB_Name = "modEinlesen"
'==============================================================================
' Ablauf "PDF-Rechnung einlesen": Auswahl, Auslesen, Vorgaben, Prüfung
'==============================================================================
Option Explicit

' Eingabefelder, die vor jedem Einlesen geleert werden (Formelfelder fehlen bewusst)
Private Const FELDER_LEEREN As String = _
    "f_Quelldatei,f_RechNr,f_RechDatum,f_LSNr,f_LSDatum,f_KundenNr,f_KGLN,f_KName,f_KStrasse,f_KPLZ,f_KOrt," & _
    "f_KEmail,f_KUstId,f_LName,f_LZeile2,f_LPLZ,f_LOrt,f_LGLN,f_Marktnr,f_Kondition,f_Konto," & _
    "f_PDF_Warenwert,f_PDF_Pfand,f_PDF_Netto,f_PDF_MwSt,f_PDF_Brutto,f_Status,f_Ausgabedatei"

Public Sub RechnungEinlesen()
    Dim pfad As String
    pfad = PdfAuswaehlen()
    If Len(pfad) = 0 Then Exit Sub

    On Error GoTo Fehler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ArbeitsblaetterLeeren
    Feld("f_Quelldatei").Value = pfad
    Protokoll "Info", "Einlesen: " & pfad

    Select Case UCase$(Einstellung("Layout"))
        Case "BSI": ParseBSI PdfZuText(pfad), Einstellung("Layout-Variante")
        Case Else: Err.Raise vbObjectError + 20, , "Unbekanntes Layout '" & Einstellung("Layout") & "' (Blatt Einstellungen)."
    End Select

    VorgabenSetzen

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    RechnungPruefen
    ThisWorkbook.Worksheets(BLATT_RECHNUNG).Activate
    Exit Sub

Fehler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Protokoll "Fehler", Err.Description
    MsgBox "Einlesen fehlgeschlagen:" & vbLf & Err.Description, vbExclamation
End Sub

Private Function PdfAuswaehlen() As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "BSI-Rechnung (PDF) auswählen"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "PDF-Dateien", "*.pdf"
        If Len(Einstellung("Eingangsordner")) > 0 Then .InitialFileName = Einstellung("Eingangsordner") & "\"
        If .Show = -1 Then PdfAuswaehlen = .SelectedItems(1)
    End With
End Function

Private Sub ArbeitsblaetterLeeren()
    Dim n As Variant, ws As Worksheet
    For Each n In Split(FELDER_LEEREN, ",")
        Feld(CStr(n)).ClearContents
    Next n
    Feld("f_RechArt").Value = "380"
    Feld("f_Waehrung").Value = "EUR"
    Feld("f_KLand").Value = "DE"
    Feld("f_LLand").Value = "DE"

    Set ws = ThisWorkbook.Worksheets(BLATT_POSITIONEN)
    If ws.Cells(ws.Rows.Count, SP_TYP).End(xlUp).Row > 1 Then
        ws.Range(ws.Cells(2, SP_TYP), ws.Cells(ws.Rows.Count, SP_INFO)).ClearContents
    End If

    Set ws = ThisWorkbook.Worksheets(BLATT_PROTOKOLL)
    If ws.Cells(ws.Rows.Count, 1).End(xlUp).Row > 1 Then ws.Range("A2:C" & ws.Rows.Count).ClearContents
End Sub

' Bankkonto, Zahlungskondition und E-Mail aus dem Blatt "Empfaenger",
' sonst die mit "x" markierten Standard-Einträge
Private Sub VorgabenSetzen()
    Dim wsE As Worksheet, z As Long
    Set wsE = ThisWorkbook.Worksheets("Empfaenger")
    z = TabellenZeile("Empfaenger", CStr(Feld("f_KundenNr").Value))

    If z > 0 Then
        If Len(wsE.Cells(z, 3).Value) > 0 Then Feld("f_KEmail").Value = wsE.Cells(z, 3).Value
        If Len(wsE.Cells(z, 4).Value) > 0 Then Feld("f_KUstId").Value = wsE.Cells(z, 4).Value
        Feld("f_Konto").Value = wsE.Cells(z, 5).Value
        Feld("f_Kondition").Value = wsE.Cells(z, 6).Value
    End If

    If Len(Feld("f_Konto").Value) = 0 Then
        z = StandardZeile("Bankkonten", 5)
        If z > 0 Then Feld("f_Konto").Value = ThisWorkbook.Worksheets("Bankkonten").Cells(z, 1).Value
    End If
    If Len(Feld("f_Kondition").Value) = 0 Then
        z = StandardZeile("Zahlungskonditionen", 4)
        If z > 0 Then Feld("f_Kondition").Value = ThisWorkbook.Worksheets("Zahlungskonditionen").Cells(z, 1).Value
    End If

    If TabellenZeile("Empfaenger", CStr(Feld("f_KundenNr").Value)) = 0 Then EmpfaengerAnlegen
End Sub

' Neuen Rechnungsempfänger mit den aktuellen Vorgaben ins Blatt "Empfaenger"
' übernehmen, damit der Anwender nichts von Hand anlegen muss. Abweichende
' Vorgaben (Konto, Kondition, E-Mail) werden dort nur bei Bedarf eingetragen.
Private Sub EmpfaengerAnlegen()
    Dim wsE As Worksheet, z As Long
    If Len(CStr(Feld("f_KundenNr").Value)) = 0 Then Exit Sub
    Set wsE = ThisWorkbook.Worksheets("Empfaenger")
    z = wsE.Cells(wsE.Rows.Count, 1).End(xlUp).Row + 1
    wsE.Cells(z, 1).NumberFormat = "@"
    wsE.Cells(z, 1).Value = CStr(Feld("f_KundenNr").Value)
    wsE.Cells(z, 2).Value = Feld("f_KName").Value
    wsE.Cells(z, 3).Value = Feld("f_KEmail").Value
    wsE.Cells(z, 4).Value = Feld("f_KUstId").Value
    ' Konto und Kondition bleiben leer = es gilt der jeweilige Standard-Eintrag
    wsE.Cells(z, 7).Value = "Automatisch angelegt am " & Format$(Date, "dd.mm.yyyy") & _
                            " aus RE " & Feld("f_RechNr").Value & " - Konto/Kondition leer = Standard"
    wsE.Range(wsE.Cells(z, 1), wsE.Cells(z, 7)).Interior.Color = RGB(255, 248, 225)
    Protokoll "Info", "Neuer Empfänger " & Feld("f_KundenNr").Value & " (" & Feld("f_KName").Value & _
                      ") im Blatt 'Empfaenger' angelegt. Abweichendes Konto oder Kondition dort eintragen."
End Sub
