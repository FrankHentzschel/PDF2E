Attribute VB_Name = "modHilfen"
'==============================================================================
' Hilfsfunktionen: Einstellungen, Feldzugriff, Zahlen/Datumswerte, RegExp
'==============================================================================
Option Explicit

Public Const BLATT_RECHNUNG As String = "Rechnung"
Public Const BLATT_POSITIONEN As String = "Positionen"
Public Const BLATT_PROTOKOLL As String = "Protokoll"

' Spalten im Blatt "Positionen"
Public Const SP_TYP As Long = 1
Public Const SP_ARTNR As Long = 2
Public Const SP_BEZ As Long = 3
Public Const SP_MENGE As Long = 4
Public Const SP_EINHEIT As Long = 5
Public Const SP_PREIS As Long = 6
Public Const SP_BETRAG As Long = 7
Public Const SP_USTSATZ As Long = 8
Public Const SP_USTKAT As Long = 9
Public Const SP_INFO As Long = 10

' Liest einen Wert aus dem Blatt "Einstellungen" (Spalte A = Schlüssel, B = Wert)
Public Function Einstellung(ByVal schluessel As String) As String
    Einstellung = SchluesselWert("Einstellungen", schluessel)
End Function

' Liest einen Wert aus dem Blatt "Brauerei" (Spalte A = Schlüssel, B = Wert)
Public Function Stamm(ByVal schluessel As String) As String
    Stamm = SchluesselWert("Brauerei", schluessel)
End Function

Private Function SchluesselWert(ByVal blatt As String, ByVal schluessel As String) As String
    Dim c As Range
    Set c = ThisWorkbook.Worksheets(blatt).Columns(1).Find(What:=schluessel, LookIn:=xlValues, LookAt:=xlWhole)
    If c Is Nothing Then Err.Raise vbObjectError + 1, , "Eintrag '" & schluessel & "' fehlt im Blatt '" & blatt & "'."
    SchluesselWert = Trim$(CStr(c.Offset(0, 1).Value))
End Function

' Zelle eines benannten Feldes im Blatt "Rechnung" (Namen beginnen mit f_)
Public Function Feld(ByVal name As String) As Range
    Set Feld = ThisWorkbook.Names(name).RefersToRange
End Function

' Deutsche Zahl in Currency umwandeln: "1.029,53" -> 1029,53 / "25,00-" -> -25
Public Function ZahlDE(ByVal s As String) As Currency
    Dim neg As Boolean
    s = Trim$(s)
    If Right$(s, 1) = "-" Then neg = True: s = Left$(s, Len(s) - 1)
    If Left$(s, 1) = "-" Then neg = True: s = Mid$(s, 2)
    s = Replace(s, ".", "")
    s = Replace(s, ",", ".")
    ZahlDE = CCur(Val(s))   ' Val ist unabhängig von den Ländereinstellungen
    If neg Then ZahlDE = -ZahlDE
End Function

' Kaufmännisch auf 2 Stellen runden (VBA-Round rundet "zur geraden Zahl")
Public Function RundenKfm(ByVal x As Currency) As Currency
    RundenKfm = Sgn(x) * Int(Abs(x) * 100 + 0.5) / 100
End Function

' "25.10.2026" oder "5.10.2026" -> Datum
Public Function DatumDE(ByVal s As String) As Date
    Dim t() As String
    t = Split(Trim$(s), ".")     ' auch "5.10.2026"
    DatumDE = DateSerial(CInt(t(2)), CInt(t(1)), CInt(t(0)))
End Function

' Neues RegExp-Objekt (VBScript-Engine, Late Binding)
Public Function NeuesRegex(ByVal muster As String) As Object
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = muster
    re.IgnoreCase = False
    re.Global = False
    re.MultiLine = False
    Set NeuesRegex = re
End Function

' Liefert das erste Match oder Nothing
Public Function Treffer(ByVal re As Object, ByVal text As String) As Object
    Dim m As Object
    Set m = re.Execute(text)
    If m.Count > 0 Then Set Treffer = m(0) Else Set Treffer = Nothing
End Function

' Meldung ins Blatt "Protokoll" schreiben
Public Sub Protokoll(ByVal art As String, ByVal meldung As String)
    Dim ws As Worksheet, z As Long
    Set ws = ThisWorkbook.Worksheets(BLATT_PROTOKOLL)
    z = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(z, 1).Value = Now
    ws.Cells(z, 2).Value = art
    ws.Cells(z, 3).Value = meldung
End Sub

' Zeile in einer Stammdatentabelle suchen (Spalte 1 = Schlüssel). 0 = nicht gefunden
Public Function TabellenZeile(ByVal blatt As String, ByVal schluessel As String) As Long
    Dim c As Range
    If Len(schluessel) = 0 Then Exit Function
    Set c = ThisWorkbook.Worksheets(blatt).Columns(1).Find(What:=schluessel, LookIn:=xlValues, LookAt:=xlWhole)
    If Not c Is Nothing Then TabellenZeile = c.Row
End Function

' Zeile mit "x" in der Standard-Spalte. 0 = keine
Public Function StandardZeile(ByVal blatt As String, ByVal standardSpalte As Long) As Long
    Dim ws As Worksheet, z As Long
    Set ws = ThisWorkbook.Worksheets(blatt)
    For z = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If LCase$(Trim$(CStr(ws.Cells(z, standardSpalte).Value))) = "x" Then StandardZeile = z: Exit Function
    Next z
End Function
