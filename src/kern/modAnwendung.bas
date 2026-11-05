Attribute VB_Name = "modAnwendung"
'==============================================================================
' Gemeinsamer Rahmen für PDF2E und XL2E: Start (Splash), Info, Prüfen,
' Einrichtung der Schaltflächen. Name, Version und Quelle stehen je Produkt im
' Modul modProdukt (src/pdf2e bzw. src/xl2e).
'
' Versionierung nach Semantic Versioning (HAUPT.NEBEN.KORREKTUR):
'   Fehlerbehebung          -> dritte Stelle +1          (1.0.0 -> 1.0.1)
'   neue Funktion           -> zweite Stelle +1, dritte 0 (1.0.1 -> 1.1.0)
'   inkompatible Änderung   -> erste Stelle +1, Rest 0   (1.4.2 -> 2.0.0)
' Jede Änderung zusätzlich im Blatt "Versionshistorie" eintragen.
'==============================================================================
Option Explicit

Private Const SPLASH_SEKUNDEN As Long = 3

' Aufruf aus ThisWorkbook.Workbook_Open
Public Sub AnwendungStart()
    On Error Resume Next            ' Ein fehlender Splash darf den Start nie verhindern
    StartblattAktualisieren
    frmInfo.Anzeigen True
    Application.OnTime Now + TimeSerial(0, 0, SPLASH_SEKUNDEN), "SplashSchliessen"
End Sub

Public Sub SplashSchliessen()
    Dim f As Object
    On Error Resume Next
    For Each f In VBA.UserForms
        If f.Name = "frmInfo" Then
            If f.AlsSplash Then Unload f
        End If
    Next f
End Sub

' Makro für die Schaltfläche "Info"
Public Sub InfoAnzeigen()
    SplashSchliessen
    frmInfo.Anzeigen False
End Sub

' Makro für die Schaltfläche "Prüfen"
Public Sub PruefenStarten()
    If RechnungPruefen() Then
        MsgBox "Prüfung ohne Fehler. Die E-Rechnung kann erzeugt werden.", vbInformation, TOOL_NAME
    Else
        MsgBox "Die Prüfung meldet Fehler. Details im Blatt 'Protokoll'.", vbExclamation, TOOL_NAME
        ThisWorkbook.Worksheets(BLATT_PROTOKOLL).Activate
    End If
End Sub

Public Function VersionText() As String
    VersionText = "Version " & TOOL_VERSION & "  |  Stand " & TOOL_STAND
End Function

' Systemangaben für Support-Anfragen
Public Function SystemText() As String
    Dim bit As String
    #If Win64 Then
        bit = "64-Bit"
    #Else
        bit = "32-Bit"
    #End If
    SystemText = "Excel " & Application.Version & " (" & bit & ")  |  Adapter: " & EinstellungOderLeer("E-Rechnung-Adapter")
End Function

' Versionszeile im Blatt "Start" aus TOOL_VERSION nachführen (nur bei Abweichung,
' damit die Mappe nicht bei jedem Öffnen als geändert gilt)
Private Sub StartblattAktualisieren()
    Dim c As Range, soll As String
    soll = TOOL_UNTERTITEL & "  |  Version " & TOOL_VERSION
    Set c = ThisWorkbook.Worksheets("Start").UsedRange.Find(What:="|  Version", LookIn:=xlValues, LookAt:=xlPart)
    If Not c Is Nothing Then
        If CStr(c.Value) <> soll Then c.Value = soll
    End If
End Sub

Public Function EinstellungOderLeer(ByVal schluessel As String) As String
    On Error Resume Next
    EinstellungOderLeer = Einstellung(schluessel)
End Function

'------------------------------------------------------------------------------
' Einmalig ausführen (über PDF2E_Einrichten bzw. XL2E_Einrichten): legt die
' Schaltflächen auf den Blättern "Start" und "Rechnung" an (vorhandene werden ersetzt).
'------------------------------------------------------------------------------
Public Sub Einrichten()
    SchaltflaechenAnlegen ThisWorkbook.Worksheets("Start"), "B2"
    SchaltflaechenAnlegen ThisWorkbook.Worksheets(BLATT_RECHNUNG), "E2"
    StartblattAktualisieren
    MailvorlageSicherstellen
    MsgBox "Schaltflächen angelegt.", vbInformation, TOOL_NAME
End Sub

' Schaltflächen als abgerundete Formen statt Formularsteuerelementen: Diese werden
' in neueren Excel-Versionen teils als "Das Bild kann nicht angezeigt werden" dargestellt.
Private Sub SchaltflaechenAnlegen(ByVal ws As Worksheet, ByVal startZelle As String)
    Dim beschriftung As Variant, makro As Variant, i As Long
    Dim l As Double, t As Double, shp As Shape
    beschriftung = Array(QUELLE_SCHALTFLAECHE, "Prüfen", "E-Rechnung erzeugen", "E-Mail erstellen", "Info")
    makro = Array("RechnungEinlesen", "PruefenStarten", "ZugferdErzeugen", "EMailErstellen", "InfoAnzeigen")

    ' Vorhandene Schaltflächen entfernen (auch die Formularsteuerelemente von PDF2E bis 1.0.1)
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, Len(TOOL_NAME) + 1) = TOOL_NAME & "_" Then ws.Shapes(i).Delete
    Next i

    l = ws.Range(startZelle).Left
    t = ws.Range(startZelle).Top
    For i = 0 To UBound(beschriftung)
        Set shp = ws.Shapes.AddShape(5, l, t + i * 32, 160, 26)     ' 5 = msoShapeRoundedRectangle
        With shp
            .Name = TOOL_NAME & "_" & makro(i)
            .OnAction = makro(i)
            .Placement = xlFreeFloating
            .Fill.ForeColor.RGB = RGB(122, 74, 18)                  ' Braun wie die Tabellenköpfe
            .Line.Visible = msoFalse
            With .TextFrame2
                .VerticalAnchor = 3                                 ' msoAnchorMiddle
                .TextRange.Text = beschriftung(i)
                .TextRange.ParagraphFormat.Alignment = 2            ' msoAlignCenter
                .TextRange.Font.Size = 10
                .TextRange.Font.Bold = msoTrue
                .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            End With
        End With
    Next i
End Sub
