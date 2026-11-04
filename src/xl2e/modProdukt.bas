Attribute VB_Name = "modProdukt"
'==============================================================================
' XL2E - Produktangaben und Datenquelle (Excel-Rechnungen aus der BSI-Software)
'
' Gegenstück in PDF2E: src/pdf2e/modProdukt.bas. Der gemeinsame Code in src/kern
' ruft nur die hier definierten Konstanten und Prozeduren auf.
'
' Die Excel-Rechnung hat keinen Briefkopf (gedacht für vorgedrucktes Papier).
' Für die E-Rechnung erzeugt XL2E deshalb ein PDF aus einer Kopie des
' Rechnungsblatts und setzt Absender, Logo, Registerangaben und Bank aus den
' Stammdaten in Kopf- und Fußzeile. Die Originaldatei bleibt unverändert.
'==============================================================================
Option Explicit

Public Const TOOL_NAME As String = "XL2E"
Public Const TOOL_VERSION As String = "1.0.0"
Public Const TOOL_STAND As String = "04.11.2026"
Public Const TOOL_UNTERTITEL As String = "Excel-Rechnungen als E-Rechnung (ZUGFeRD)"
Public Const QUELLE_SCHALTFLAECHE As String = "Excel einlesen"

Public Sub XL2E_Start()
    AnwendungStart
End Sub

Public Sub XL2E_Einrichten()
    Einrichten
End Sub

Public Sub XL2E_Exportieren()
    Quellcode_Exportieren
End Sub

' Dateiauswahl für "Excel einlesen"
Public Function QuelleAuswaehlen() As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Excel-Rechnung auswählen"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel-Dateien", "*.xlsx;*.xlsm;*.xls"
        If Len(Einstellung("Eingangsordner")) > 0 Then .InitialFileName = Einstellung("Eingangsordner") & "\"
        If .Show = -1 Then QuelleAuswaehlen = .SelectedItems(1)
    End With
End Function

' Rechnung auslesen und in die Blätter "Rechnung" und "Positionen" schreiben
Public Sub QuelleEinlesen(ByVal pfad As String)
    Dim wb As Workbook
    Set wb = QuelleOeffnen(pfad)
    On Error GoTo Fehler
    ParseXL QuellBlatt(wb)
    wb.Close SaveChanges:=False
    Exit Sub
Fehler:
    Dim nr As Long, txt As String
    nr = Err.Number: txt = Err.Description
    wb.Close SaveChanges:=False
    Err.Raise nr, , txt
End Sub

' PDF der Rechnung mit Briefkopf aus den Stammdaten erzeugen
Public Sub QuellPdfErstellen(ByVal quelle As String, ByVal ziel As String)
    Dim wbQ As Workbook, wbN As Workbook, ws As Worksheet
    Dim nr As Long, txt As String

    Application.ScreenUpdating = False
    Set wbQ = QuelleOeffnen(quelle)
    On Error GoTo Fehler
    QuellBlatt(wbQ).Copy                     ' Kopie in eine neue, ungespeicherte Mappe
    Set wbN = ActiveWorkbook
    Set ws = wbN.Worksheets(1)
    If UCase$(EinstellungOderLeer("Briefkopf")) <> "NEIN" Then BriefkopfSetzen ws
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=ziel, Quality:=xlQualityStandard, _
                           IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
    wbN.Close SaveChanges:=False
    wbQ.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Exit Sub

Fehler:
    nr = Err.Number: txt = Err.Description
    On Error Resume Next
    If Not wbN Is Nothing Then wbN.Close SaveChanges:=False
    wbQ.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Err.Raise nr, , "PDF aus der Excel-Rechnung: " & txt
End Sub

Private Function QuelleOeffnen(ByVal pfad As String) As Workbook
    If Len(Dir$(pfad)) = 0 Then Err.Raise vbObjectError + 43, , "Datei nicht gefunden: " & pfad
    Set QuelleOeffnen = Workbooks.Open(Filename:=pfad, ReadOnly:=True, UpdateLinks:=0, AddToMru:=False)
End Function

' Blatt aus Einstellung "Quellblatt" (Standard PA400_00), sonst das erste Blatt
Private Function QuellBlatt(ByVal wb As Workbook) As Worksheet
    Dim name As String
    name = EinstellungOderLeer("Quellblatt")
    If Len(name) = 0 Then name = "PA400_00"
    On Error Resume Next
    Set QuellBlatt = wb.Worksheets(name)
    On Error GoTo 0
    If QuellBlatt Is Nothing Then Set QuellBlatt = wb.Worksheets(1)
End Function

' Absender und Logo in die Kopfzeile, Register, Steuer und Bank in die Fußzeile.
' Excel erlaubt je Abschnitt nur rund 255 Zeichen. Formatcodes ohne Schriftnamen
' (&B fett, &8 Größe), damit sie in jeder Excel-Sprache gelten.
Private Sub BriefkopfSetzen(ByVal ws As Worksheet)
    Dim logo As String, bz As Long, kontakt As String, bank As String, gf As String

    kontakt = "Tel. " & StammOderLeer("Telefon")
    If Len(StammOderLeer("E-Mail")) > 0 Then kontakt = kontakt & "  " & ChrW(183) & "  " & StammOderLeer("E-Mail")
    If Len(StammOderLeer("Internet")) > 0 Then kontakt = kontakt & "  " & ChrW(183) & "  " & StammOderLeer("Internet")

    bz = TabellenZeile("Bankkonten", CStr(Feld("f_Konto").Value))
    If bz > 0 Then
        With ThisWorkbook.Worksheets("Bankkonten")
            bank = .Cells(bz, 1).Value & vbLf & "IBAN " & IbanLesbar(CStr(.Cells(bz, 3).Value))
            If Len(.Cells(bz, 4).Value) > 0 Then bank = bank & vbLf & "BIC " & .Cells(bz, 4).Value
        End With
    End If
    gf = StammOderLeer("Geschäftsführer")
    If Len(gf) > 0 Then gf = "Geschäftsführer: " & gf & vbLf

    Application.PrintCommunication = False
    With ws.PageSetup
        .LeftHeader = "&B&12" & Stamm("Name") & "&B" & vbLf & _
                      "&8" & Stamm("Straße") & ", " & Stamm("PLZ") & " " & Stamm("Ort") & vbLf & kontakt
        logo = EinstellungOderLeer("Briefkopf-Logo")
        If Len(logo) > 0 And Len(Dir$(logo)) > 0 Then
            .RightHeaderPicture.Filename = logo
            .RightHeaderPicture.LockAspectRatio = msoTrue
            .RightHeaderPicture.Height = Application.CentimetersToPoints(2.2)
            .RightHeader = "&G"
        ElseIf Len(logo) > 0 Then
            Protokoll "Warnung", "Briefkopf-Logo nicht gefunden: " & logo
        End If
        .LeftFooter = "&7" & Stamm("Name") & vbLf & gf & StammOderLeer("Registernummer")
        .CenterFooter = "&7" & "USt-IdNr. " & Stamm("USt-IdNr") & _
                        IIf(Len(StammOderLeer("Steuernummer")) > 0, vbLf & "Steuernummer " & StammOderLeer("Steuernummer"), "") & _
                        vbLf & "Seite &P von &N"
        .RightFooter = "&7" & bank
        .HeaderMargin = Application.CentimetersToPoints(0.8)
        .TopMargin = Application.CentimetersToPoints(3.6)
        .FooterMargin = Application.CentimetersToPoints(0.8)
        .BottomMargin = Application.CentimetersToPoints(2.4)
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
    End With
    Application.PrintCommunication = True
End Sub

Private Function IbanLesbar(ByVal iban As String) As String
    Dim i As Long
    iban = Replace(iban, " ", "")
    For i = 1 To Len(iban) Step 4
        IbanLesbar = IbanLesbar & IIf(i > 1, " ", "") & Mid$(iban, i, 4)
    Next i
End Function
