Attribute VB_Name = "modStarter"
'==============================================================================
' Starter für PDF2E und XL2E
'
' Die Starter-Mappe liegt lokal (z. B. Desktop-Verknüpfung) und ändert sich
' praktisch nie. Beim Öffnen:
'   1. Einstellungen und Pfade prüfen (Blatt "Starter")
'   2. Version im Update-Ordner und lokal lesen (aus der geschlossenen Mappe,
'      Blatt Start, Zelle Z1; Rückfall: schreibgeschützt öffnen)
'   3. Ist der Update-Ordner neuer oder fehlt die lokale Mappe: kopieren
'      (die bisherige lokale Mappe wird als <Produkt>_alt.xlsm gesichert)
'   4. Lokale Mappe ohne Ereignisse öffnen, Werte aus Bereich 2 mit dem Blatt
'      Einstellungen abgleichen (nur Abweichungen schreiben und speichern)
'   5. Programm starten (AnwendungStart), Starter schließt sich
'
' Zum Bearbeiten der Einstellungen die Mappe mit gedrückter Umschalttaste
' öffnen oder beim Start Strg gedrückt halten: dann startet nichts automatisch.
'==============================================================================
Option Explicit

Private Const BLATT As String = "Starter"
Private Const BEREICH2 As String = "Einstellung für die Programmmappe"
Private Const STARTER_VERSION As String = "1.0.0"

#If VBA7 Then
    Private Declare PtrSafe Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
#Else
    Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
#End If

' Aufruf aus Workbook_Open
Public Sub StarterOeffnen()
    If GetKeyState(&H11) < 0 Or GetKeyState(&H10) < 0 Then       ' Strg oder Umschalt gedrückt
        ThisWorkbook.Worksheets(BLATT).Activate
        Exit Sub
    End If
    Application.OnTime Now, "'" & ThisWorkbook.Name & "'!Starten"    ' erst nach dem vollständigen Öffnen
End Sub

' Makro für die Schaltfläche "Jetzt starten"
Public Sub Starten()
    Dim produkt As String, quelle As String, lokal As String, mappe As String
    Dim netz As String, lok As String, vNetz As String, vLok As String, wb As Workbook
    Dim geaendert As Long, ereignisse As Boolean, kopiert As Boolean

    On Error GoTo Fehler
    produkt = UCase$(Wert("Produkt"))
    quelle = OhneBackslash(Wert("Update-Ordner"))
    lokal = OhneBackslash(Wert("Programmordner lokal"))
    If produkt <> "PDF2E" And produkt <> "XL2E" Then Abbruch "Produkt muss PDF2E oder XL2E sein."
    If Len(lokal) = 0 Then Abbruch "Bitte 'Programmordner lokal' eintragen."
    mappe = produkt & ".xlsm"
    lok = lokal & "\" & mappe
    If Len(quelle) > 0 Then netz = quelle & "\" & mappe

    ' Programm schon offen: nur nach vorne holen
    If MappeOffen(mappe) Then
        Workbooks(mappe).Activate
        ThisWorkbook.Close False
        Exit Sub
    End If

    If Not OrdnerExistiert(lokal) Then OrdnerAnlegen lokal
    If Len(netz) > 0 Then
        If Not OrdnerExistiert(quelle) Then
            If Len(Dir$(lok)) = 0 Then Abbruch "Der Update-Ordner ist nicht erreichbar und lokal gibt es noch keine Programmmappe:" & vbLf & quelle
            MsgBox "Der Update-Ordner ist nicht erreichbar:" & vbLf & quelle & vbLf & vbLf & _
                   "Es wird ohne Update-Prüfung gestartet.", vbExclamation, "Starter"
            netz = ""
        ElseIf Len(Dir$(netz)) = 0 Then
            If Len(Dir$(lok)) = 0 Then Abbruch "Im Update-Ordner liegt keine " & mappe & ":" & vbLf & quelle
            netz = ""
        End If
    End If
    If Len(netz) = 0 And Len(Dir$(lok)) = 0 Then Abbruch "Keine Programmmappe gefunden:" & vbLf & lok

    ' Version vergleichen und ggf. kopieren
    If Len(netz) > 0 Then
        vNetz = VersionLesen(netz)
        If Len(Dir$(lok)) > 0 Then vLok = VersionLesen(lok)
        If Len(Dir$(lok)) = 0 Or VersionVergleich(vNetz, vLok) > 0 Then
            Application.StatusBar = "Starter: neue Version " & vNetz & " wird geladen ..."
            ' FileSystemObject statt FileCopy: kopiert auch, wenn die Quelle anderswo geöffnet ist
            If Len(Dir$(lok)) > 0 Then CreateObject("Scripting.FileSystemObject").CopyFile lok, lokal & "\" & produkt & "_alt.xlsm", True
            CreateObject("Scripting.FileSystemObject").CopyFile netz, lok, True
            kopiert = True
        End If
    End If

    ' Lokale Mappe ohne Workbook_Open öffnen, Einstellungen abgleichen
    ereignisse = Application.EnableEvents
    Application.EnableEvents = False
    Set wb = Workbooks.Open(lok, UpdateLinks:=0)
    Application.EnableEvents = ereignisse
    geaendert = EinstellungenAbgleichen(wb)
    If geaendert > 0 Then wb.Save
    Application.StatusBar = False

    If kopiert Then
        MsgBox produkt & " wurde auf Version " & vNetz & " aktualisiert" & _
               IIf(Len(vLok) > 0, " (bisher " & vLok & ").", "."), vbInformation, "Starter"
    End If

    Application.Run "'" & wb.Name & "'!AnwendungStart"
    ThisWorkbook.Close False
    Exit Sub

Fehler:
    Application.EnableEvents = True
    Application.StatusBar = False
    ThisWorkbook.Worksheets(BLATT).Activate
    MsgBox "Start nicht möglich:" & vbLf & Err.Description & vbLf & vbLf & _
           "Einstellungen im Blatt '" & BLATT & "' prüfen und 'Jetzt starten' wählen.", vbExclamation, "Starter " & STARTER_VERSION
End Sub

' Einmalig ausführen (Alt+F8): legt die Schaltfläche "Jetzt starten" an
Public Sub StarterEinrichten()
    Dim ws As Worksheet, shp As Shape, i As Long
    Set ws = ThisWorkbook.Worksheets(BLATT)
    For i = ws.Shapes.Count To 1 Step -1
        If ws.Shapes(i).Name = "Starter_Starten" Then ws.Shapes(i).Delete
    Next i
    Set shp = ws.Shapes.AddShape(5, ws.Range("D5").Left + 8, ws.Range("D5").Top, 140, 28)   ' abgerundetes Rechteck
    With shp
        .Name = "Starter_Starten"
        .OnAction = "Starten"
        .Fill.ForeColor.RGB = RGB(122, 74, 18)
        .Line.Visible = msoFalse
        .TextFrame2.TextRange.Text = "Jetzt starten"
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.TextRange.ParagraphFormat.Alignment = 2
        .TextFrame2.VerticalAnchor = 3
    End With
End Sub

Private Sub Abbruch(ByVal text As String)
    Err.Raise vbObjectError + 1, , text
End Sub

'------------------------------------------------------------------------------
' Version lesen: zuerst aus der geschlossenen Mappe (ExecuteExcel4Macro),
' sonst schreibgeschützt öffnen. Leer = unbekannt (ältere Mappe).
'------------------------------------------------------------------------------
Private Function VersionLesen(ByVal pfad As String) As String
    Dim ordner As String, datei As String, v As Variant, wb As Workbook, ereignisse As Boolean
    ordner = Left$(pfad, InStrRev(pfad, "\") - 1)
    datei = Mid$(pfad, InStrRev(pfad, "\") + 1)
    On Error Resume Next
    v = Application.ExecuteExcel4Macro("'" & ordner & "\[" & datei & "]Start'!R1C26")
    If Err.Number = 0 And Not IsError(v) Then
        If Len(CStr(v)) > 0 And CStr(v) <> "0" Then
            VersionLesen = CStr(v)
            Exit Function
        End If
    End If
    Err.Clear
    ' Rückfall
    ereignisse = Application.EnableEvents
    Application.EnableEvents = False
    Set wb = Workbooks.Open(pfad, UpdateLinks:=0, ReadOnly:=True)
    Application.EnableEvents = ereignisse
    If Not wb Is Nothing Then
        VersionLesen = Trim$(CStr(wb.Worksheets("Start").Range("Z1").Value))
        wb.Close False
    End If
End Function

' > 0: a ist neuer, 0: gleich, < 0: b ist neuer. Leer gilt als ältester Stand.
Private Function VersionVergleich(ByVal a As String, ByVal b As String) As Long
    Dim ta() As String, tb() As String, i As Long, x As Long, y As Long
    If a = b Then Exit Function
    If Len(b) = 0 Then VersionVergleich = 1: Exit Function
    If Len(a) = 0 Then VersionVergleich = -1: Exit Function
    ta = Split(a, ".")
    tb = Split(b, ".")
    For i = 0 To 2
        x = 0: y = 0
        If i <= UBound(ta) Then x = CLng(Val(ta(i)))
        If i <= UBound(tb) Then y = CLng(Val(tb(i)))
        If x <> y Then VersionVergleich = Sgn(x - y): Exit Function
    Next i
End Function

'------------------------------------------------------------------------------
' Bereich 2 des Starters in das Blatt "Arbeitsplatz" der Programmmappe
' übernehmen (platzbezogene Werte, sie gehen bei einem Update der Mappe nicht
' verloren, weil der Starter sie bei jedem Start setzt). Nur abweichende Werte
' werden geschrieben, fehlende Einträge angehängt. Rückgabe: Anzahl Änderungen.
'------------------------------------------------------------------------------
Private Function EinstellungenAbgleichen(ByVal wb As Workbook) As Long
    Dim ws As Worksheet, wsE As Worksheet, start As Range, z As Long, k As String, v As String, c As Range, n As Long
    Set ws = ThisWorkbook.Worksheets(BLATT)
    Set wsE = ArbeitsplatzBlatt(wb)
    Set start = ws.Columns(1).Find(What:=BEREICH2, LookIn:=xlValues, LookAt:=xlWhole)
    If start Is Nothing Then Exit Function
    z = start.Row + 1
    Do While Len(Trim$(CStr(ws.Cells(z, 1).Value))) > 0
        k = Trim$(CStr(ws.Cells(z, 1).Value))
        v = Trim$(CStr(ws.Cells(z, 2).Value))
        Set c = wsE.Columns(1).Find(What:=k, LookIn:=xlValues, LookAt:=xlWhole)
        If c Is Nothing Then
            Set c = wsE.Cells(wsE.Rows.Count, 1).End(xlUp).Offset(1, 0)
            c.Value = k
            c.Offset(0, 2).Value = "vom Starter angelegt"
        End If
        If CStr(c.Offset(0, 1).Value) <> v Then
            c.Offset(0, 1).NumberFormat = "@"
            c.Offset(0, 1).Value = v
            n = n + 1
        End If
        z = z + 1
    Loop
    EinstellungenAbgleichen = n
End Function

'------------------------------------------------------------------------------
' Hilfen
'------------------------------------------------------------------------------
Private Function ArbeitsplatzBlatt(ByVal wb As Workbook) As Worksheet
    On Error Resume Next
    Set ArbeitsplatzBlatt = wb.Worksheets("Arbeitsplatz")
    On Error GoTo 0
    If Not ArbeitsplatzBlatt Is Nothing Then Exit Function
    Set ArbeitsplatzBlatt = wb.Worksheets.Add(After:=wb.Worksheets("Einstellungen"))
    With ArbeitsplatzBlatt
        .Name = "Arbeitsplatz"
        .Range("A1:C1").Value = Array("Einstellung", "Wert", "Hinweis")
        .Range("A1:C1").Font.Bold = True
        .Columns(1).ColumnWidth = 22
        .Columns(2).ColumnWidth = 60
        .Columns(3).ColumnWidth = 60
    End With
End Function

Private Function Wert(ByVal schluessel As String) As String
    Dim c As Range
    Set c = ThisWorkbook.Worksheets(BLATT).Columns(1).Find(What:=schluessel, LookIn:=xlValues, LookAt:=xlWhole)
    If c Is Nothing Then Err.Raise vbObjectError + 2, , "Eintrag '" & schluessel & "' fehlt im Blatt '" & BLATT & "'."
    Wert = Trim$(CStr(c.Offset(0, 1).Value))
End Function

Private Function OhneBackslash(ByVal s As String) As String
    s = Trim$(s)
    Do While Len(s) > 3 And Right$(s, 1) = "\"
        s = Left$(s, Len(s) - 1)
    Loop
    OhneBackslash = s
End Function

Private Function MappeOffen(ByVal name As String) As Boolean
    Dim wb As Workbook
    On Error Resume Next
    Set wb = Workbooks(name)
    MappeOffen = Not wb Is Nothing
End Function

Private Function OrdnerExistiert(ByVal ordner As String) As Boolean
    If Len(ordner) = 0 Then Exit Function
    OrdnerExistiert = CreateObject("Scripting.FileSystemObject").FolderExists(ordner)
End Function

' Legt auch fehlende übergeordnete Ordner an
Private Sub OrdnerAnlegen(ByVal ordner As String)
    Dim fso As Object, eltern As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(ordner) Then Exit Sub
    eltern = fso.GetParentFolderName(ordner)
    If Len(eltern) > 0 And Not fso.FolderExists(eltern) Then OrdnerAnlegen eltern
    fso.CreateFolder ordner
End Sub
