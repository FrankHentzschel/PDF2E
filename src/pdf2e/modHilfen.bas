Attribute VB_Name = "modHilfen"
'==============================================================================
' Hilfsfunktionen: Einstellungen, Feldzugriff, Zahlen/Datumswerte, RegExp
'==============================================================================
Option Explicit

Public Const BLATT_RECHNUNG As String = "Rechnung"
Public Const BLATT_POSITIONEN As String = "Positionen"
Public Const BLATT_PROTOKOLL As String = "Protokoll"
Public Const BLATT_ARBEITSPLATZ As String = "Arbeitsplatz"

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

' Länderkennzeichen für USt-IdNr. (UStIdFehler)
Private Const UST_PRAEFIXE As String = " AT BE BG CY CZ DE DK EE EL ES FI FR HR HU IE IT LT LU LV MT NL PL PT RO SE SI SK XI CH GB NO LI "

' Liest eine Einstellung. Reihenfolge:
'   1. Blatt "Arbeitsplatz" (von der Starter-Mappe gesetzt, z. B. Backend-Pfad)
'   2. Blatt "Einstellungen", Zeile mit Spalte D (Arbeitsplatz) = Computername
'   3. Blatt "Einstellungen", Zeile mit leerer Spalte D (gilt für alle Plätze)
Public Function Einstellung(ByVal schluessel As String) As String
    Dim v As String, ws As Worksheet, z As Long, platz As String, zAlle As Long, zPlatz As Long
    v = PlatzWert(schluessel)
    If v <> vbNullChar And Len(v) > 0 Then Einstellung = v: Exit Function

    Set ws = ThisWorkbook.Worksheets("Einstellungen")
    platz = UCase$(Environ$("COMPUTERNAME"))
    For z = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If Trim$(CStr(ws.Cells(z, 1).Value)) = schluessel Then
            v = UCase$(Trim$(CStr(ws.Cells(z, 4).Value)))
            If Len(v) = 0 Then
                If zAlle = 0 Then zAlle = z
            ElseIf v = platz Then
                zPlatz = z
                Exit For
            End If
        End If
    Next z
    If zPlatz > 0 Then z = zPlatz Else z = zAlle
    If z = 0 Then Err.Raise vbObjectError + 1, , "Eintrag '" & schluessel & "' fehlt im Blatt 'Einstellungen'."
    Einstellung = Trim$(CStr(ws.Cells(z, 2).Value))
End Function

' Wert aus dem Blatt "Arbeitsplatz"; vbNullChar, wenn Blatt oder Eintrag fehlt
Public Function PlatzWert(ByVal schluessel As String) As String
    Dim ws As Worksheet, c As Range
    PlatzWert = vbNullChar
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(BLATT_ARBEITSPLATZ)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    Set c = ws.Columns(1).Find(What:=schluessel, LookIn:=xlValues, LookAt:=xlWhole)
    If Not c Is Nothing Then PlatzWert = Trim$(CStr(c.Offset(0, 1).Value))
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

' Gibt es den Namen (Feld) in der Mappe? Ältere Mappen haben neuere Felder evtl. noch nicht.
Public Function NameVorhanden(ByVal name As String) As Boolean
    Dim n As Object
    On Error Resume Next
    Set n = ThisWorkbook.Names(name)
    NameVorhanden = (Err.Number = 0)
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' USt-IdNr. (BT-31, BT-48): EN 16931 verlangt ein Länderkennzeichen (BR-CO-09)
'------------------------------------------------------------------------------

' Leerzeichen, Punkte, Bindestriche und Schrägstriche entfernen, Großbuchstaben
Public Function UStIdBereinigt(ByVal ustId As Variant) As String
    Dim s As String, z As Variant
    If IsError(ustId) Or IsNull(ustId) Then Exit Function
    s = UCase$(CStr(ustId))
    For Each z In Array(" ", ".", "-", "/", Chr$(160), vbTab)
        s = Replace(s, CStr(z), "")
    Next z
    UStIdBereinigt = s
End Function

' Leer = in Ordnung (oder keine USt-IdNr.), sonst verständliche Fehlermeldung.
' Erwartet die bereinigte Form (UStIdBereinigt).
Public Function UStIdFehler(ByVal ustId As String) As String
    If Len(ustId) = 0 Then Exit Function
    If Not ustId Like "[A-Z][A-Z]*" Then
        UStIdFehler = "'" & ustId & "' hat kein Länderkennzeichen. Eine USt-IdNr. beginnt mit dem Land, z. B. DE123456789. " & _
                      "Eine Steuernummer gehört nicht in dieses Feld. Ist keine USt-IdNr. bekannt, das Feld leer lassen."
    ElseIf InStr(UST_PRAEFIXE, " " & Left$(ustId, 2) & " ") = 0 Then
        UStIdFehler = "'" & ustId & "': Länderkennzeichen '" & Left$(ustId, 2) & "' ist unbekannt (Griechenland: EL statt GR)."
    ElseIf Left$(ustId, 2) = "DE" And Not ustId Like "DE#########" Then
        UStIdFehler = "'" & ustId & "': Eine deutsche USt-IdNr. besteht aus DE und 9 Ziffern."
    ElseIf Len(ustId) < 4 Or Len(ustId) > 14 Then
        UStIdFehler = "'" & ustId & "' hat nicht die Länge einer USt-IdNr."
    End If
End Function

' Eingabe im Blatt "Empfaenger", Spalte D (USt-IdNr.): sofort bereinigen und prüfen
Public Sub UStIdEingabePruefen(ByVal sh As Object, ByVal target As Range)
    Dim bereich As Range, c As Range, neu As String, meldung As String
    If sh.Name <> "Empfaenger" Then Exit Sub
    Set bereich = Intersect(target, sh.Range(sh.Cells(2, 4), sh.Cells(sh.Rows.Count, 4)))
    If bereich Is Nothing Then Exit Sub
    If bereich.Cells.Count > 50 Then Exit Sub            ' großes Einfügen: Prüfung beim Einlesen
    For Each c In bereich.Cells
        If Not IsError(c.Value) Then
            neu = UStIdBereinigt(c.Value)
            If neu <> CStr(c.Value) Then
                Application.EnableEvents = False
                c.NumberFormat = "@"
                c.Value = neu
                Application.EnableEvents = True
            End If
            meldung = UStIdFehler(neu)
            If Len(meldung) > 0 Then
                MsgBox "Empfänger " & sh.Cells(c.Row, 1).Value & " (Zeile " & c.Row & "), USt-IdNr.:" & vbLf & meldung, _
                       vbExclamation, TOOL_NAME
            End If
        End If
    Next c
End Sub

' SEPA-Lastschrift laut Quellrechnung (Felder f_Mandat, f_GlaeubigerId, f_Einzug)
Public Function IstLastschrift() As Boolean
    If NameVorhanden("f_Mandat") Then IstLastschrift = Len(Trim$(CStr(Feld("f_Mandat").Value))) > 0
End Function

' Rundungsbetrag (BT-114, Feld f_Rundung): nur bei Rechnungen mit Bruttobeträgen
' ("ENTHALTENE MWSt."), sonst leer. Ältere Mappen haben das Feld evtl. nicht.
Public Function BruttoBetraege() As Boolean
    If NameVorhanden("f_Rundung") Then BruttoBetraege = Len(Trim$(CStr(Feld("f_Rundung").Value))) > 0
End Function

Public Function Rundung() As Currency
    If BruttoBetraege() Then Rundung = CCur(Feld("f_Rundung").Value)
End Function

' Zahlbetrag (BT-115) = Rechnungsbetrag (BT-112) + Rundungsbetrag (BT-114)
Public Function Zahlbetrag() As Currency
    Zahlbetrag = CCur(Feld("f_Brutto").Value) + Rundung()
End Function

' Fälligkeit (BT-9): bei Lastschrift das Einzugsdatum, sonst Rechnungsdatum + Zahlungsziel
Public Function FaelligAm() As Variant
    FaelligAm = Feld("f_Faellig").Value
    If IstLastschrift() Then
        If IsDate(Feld("f_Einzug").Value) Then FaelligAm = Feld("f_Einzug").Value
    End If
End Function

' Zahlungsbedingung (BT-20): bei Lastschrift der Einzugshinweis, sonst der Text
' der Zahlungskondition aus dem Blatt "Zahlungskonditionen"
Public Function Zahlungsbedingung() As String
    Dim kz As Long, einzug As String
    If IstLastschrift() Then
        If IsDate(Feld("f_Einzug").Value) Then einzug = " am " & Format$(Feld("f_Einzug").Value, "dd.mm.yyyy")
        Zahlungsbedingung = "Der Betrag wird per SEPA-Lastschrift" & einzug & " eingezogen (Mandatsreferenz " & _
            Trim$(CStr(Feld("f_Mandat").Value)) & ", Gläubiger-ID " & Trim$(CStr(Feld("f_GlaeubigerId").Value)) & _
            "). Fällt der Termin auf ein Wochenende oder einen Feiertag, wird am nächstmöglichen Termin eingezogen."
        Exit Function
    End If
    kz = TabellenZeile("Zahlungskonditionen", CStr(Feld("f_Kondition").Value))
    If kz > 0 Then Zahlungsbedingung = CStr(ThisWorkbook.Worksheets("Zahlungskonditionen").Cells(kz, 2).Value)
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

' Optionaler Eintrag im Blatt "Brauerei": leer, wenn er fehlt
Public Function StammOderLeer(ByVal schluessel As String) As String
    On Error Resume Next
    StammOderLeer = Stamm(schluessel)
End Function

' Registerangaben der Brauerei für die E-Rechnung (BT-22 mit Code REG), z. B.
' "Konrad Krug Brauerei und Tanzsaal GmbH; Sitz: Waischenfeld; Handelsregister:
' HRB 1273, Amtsgericht Bayreuth; Geschäftsführer: Konrad Krug".
' Blatt "Brauerei": Zeile "Registerangaben" (optional) ersetzt den ganzen Text;
' sonst aus Name, "Sitz" (optional, leer = Ort), "Registernummer" und
' "Geschäftsführer". Leer, wenn weder Registernummer noch Geschäftsführer da sind.
Public Function Registerangaben() As String
    Dim t As String, sitz As String, reg As String, gf As String
    t = StammOderLeer("Registerangaben")
    If Len(t) > 0 Then Registerangaben = t: Exit Function
    reg = StammOderLeer("Registernummer")
    gf = StammOderLeer("Geschäftsführer")
    If Len(reg) = 0 And Len(gf) = 0 Then Exit Function
    sitz = StammOderLeer("Sitz")
    If Len(sitz) = 0 Then sitz = StammOderLeer("Ort")
    t = StammOderLeer("Name")
    If Len(sitz) > 0 Then t = t & "; Sitz: " & sitz
    If Len(reg) > 0 Then t = t & "; Handelsregister: " & reg
    If Len(gf) > 0 Then t = t & "; Geschäftsführer: " & gf
    If Left$(t, 2) = "; " Then t = Mid$(t, 3)
    Registerangaben = t
End Function

'------------------------------------------------------------------------------
' Positionen (von den Einlese-Vorlagen beider Produkte genutzt)
'------------------------------------------------------------------------------
Public Sub PositionAnhaengen(ByVal typ As String, ByVal artNr As String, ByVal bez As String, _
                              ByVal menge As Double, ByVal einheit As String, ByVal preis As Currency, _
                              ByVal betrag As Currency, ByVal info As String)
    Dim ws As Worksheet, z As Long
    Set ws = ThisWorkbook.Worksheets(BLATT_POSITIONEN)
    z = ws.Cells(ws.Rows.Count, SP_TYP).End(xlUp).Row + 1
    ws.Cells(z, SP_TYP).Value = typ
    ws.Cells(z, SP_ARTNR).Value = "'" & artNr
    ws.Cells(z, SP_BEZ).Value = bez
    ws.Cells(z, SP_MENGE).Value = menge
    ws.Cells(z, SP_EINHEIT).Value = einheit
    ws.Cells(z, SP_PREIS).Value = preis
    ws.Cells(z, SP_BETRAG).Value = betrag
    ws.Cells(z, SP_INFO).Value = "'" & info
End Sub

' Steuerbefreite Rechnung (z. B. K = innergemeinschaftliche Lieferung): 0 % für alle Zeilen
Public Sub SteuerfreiSetzen(ByVal kategorie As String)
    Dim ws As Worksheet, z As Long
    Set ws = ThisWorkbook.Worksheets(BLATT_POSITIONEN)
    For z = 2 To ws.Cells(ws.Rows.Count, SP_TYP).End(xlUp).Row
        ws.Cells(z, SP_USTSATZ).Value = 0
        ws.Cells(z, SP_USTKAT).Value = kategorie
    Next z
End Sub

' Kommt in den Positionen die USt-Kategorie vor (z. B. "K")?
Public Function KategorieVorhanden(ByVal kategorie As String) As Boolean
    Dim ws As Worksheet, z As Long
    Set ws = ThisWorkbook.Worksheets(BLATT_POSITIONEN)
    For z = 2 To ws.Cells(ws.Rows.Count, SP_TYP).End(xlUp).Row
        If UCase$(Trim$(CStr(ws.Cells(z, SP_USTKAT).Value))) = UCase$(kategorie) Then KategorieVorhanden = True: Exit Function
    Next z
End Function

' Bei genau einem MwSt-Satz gilt er für alle Zeilen. Bei mehreren muss der Anwender zuordnen.
Public Sub SteuersatzSetzen(saetze As Collection)
    Dim ws As Worksheet, z As Long, letzte As Long
    Set ws = ThisWorkbook.Worksheets(BLATT_POSITIONEN)
    letzte = ws.Cells(ws.Rows.Count, SP_TYP).End(xlUp).Row
    If saetze.Count = 1 Then
        For z = 2 To letzte
            ws.Cells(z, SP_USTSATZ).Value = saetze(1)
            ws.Cells(z, SP_USTKAT).Value = "S"
        Next z
    Else
        Protokoll "Warnung", saetze.Count & " MwSt-Sätze auf der Rechnung. Bitte Spalte 'USt-Satz' je Position ausfüllen."
    End If
End Sub

'------------------------------------------------------------------------------
' Rechnungsempfänger
'------------------------------------------------------------------------------
' Schlüssel im Blatt "Empfaenger" (Spalte A). Normal die Kundennummer. Bei einer
' Sammelkundennummer (Einstellung "Sammelkundennummern", z. B. "12100" für
' Festlichkeiten) teilen sich viele Kunden eine Nummer; dann gilt
' Kundennummer / Name / PLZ, damit E-Mail und Vorgaben nicht beim falschen
' Kunden landen.
Public Function EmpfaengerSchluessel() As String
    Dim nr As String, liste As Variant, x As Variant
    nr = Trim$(CStr(Feld("f_KundenNr").Value))
    EmpfaengerSchluessel = nr
    If Len(nr) = 0 Then Exit Function
    liste = Split(Replace(EinstellungOderLeer("Sammelkundennummern"), ";", ","), ",")
    For Each x In liste
        If Trim$(CStr(x)) = nr Then
            EmpfaengerSchluessel = nr & " / " & Trim$(CStr(Feld("f_KName").Value)) & " / " & Trim$(CStr(Feld("f_KPLZ").Value))
            Exit Function
        End If
    Next x
End Function

'------------------------------------------------------------------------------
' Dateien auf Netzlaufwerken (ab 1.15.1)
'------------------------------------------------------------------------------
' Wartet die angegebene Zeit, ohne Excel zu blockieren (kein Declare nötig)
Public Sub Warten(ByVal sekunden As Double)
    Dim start As Single
    start = Timer
    Do While Timer - start < sekunden And Timer >= start   ' Timer springt um Mitternacht auf 0
        DoEvents
    Loop
End Sub

' Löscht eine Datei und wartet, bis sie wirklich weg ist. Auf Netzlaufwerken
' bleibt eine Datei, die ein anderes Programm (PDF-Anzeige, Explorer-Vorschau,
' Virenscanner) noch offen hält, nach dem Löschen kurz "zum Löschen vorgemerkt"
' sichtbar; jeder Zugriff darauf meldet "Zugriff verweigert". Bis etwa 10
' Sekunden wird gewartet, danach kommt eine verständliche Fehlermeldung.
' Ab 1.15.2 wird die Datei zuerst in einen eindeutigen Namen umbenannt
' ("…pdf.alt-JJJJMMTThhmmss") und dann gelöscht. Damit ist der Name sofort frei,
' auch wenn Excel/Smart Tools die alte Datei noch offen hält; die umbenannte
' Datei verschwindet, sobald das andere Programm sie freigibt.
Public Sub DateiLoeschen(ByVal pfad As String)
    Dim versuch As Long, nr As Long, txt As String, alt As String
    For versuch = 1 To 20
        If Len(Dir$(pfad)) = 0 Then Exit Sub
        On Error Resume Next
        SetAttr pfad, vbNormal
        alt = pfad & ".alt-" & Format$(Now, "yyyymmddhhnnss") & "-" & versuch
        Err.Clear
        Name pfad As alt
        If Err.Number = 0 Then
            Kill alt                ' darf scheitern: dann bleibt nur die umbenannte Datei
            Err.Clear
            On Error GoTo 0
            Exit Sub
        End If
        Err.Clear
        Kill pfad
        If Err.Number <> 0 Then nr = Err.Number: txt = Err.Description
        On Error GoTo 0
        If Len(Dir$(pfad)) = 0 Then Exit Sub
        Warten 0.5
    Next versuch
    Err.Raise vbObjectError + 40, , "Die Datei lässt sich nicht löschen oder wird noch gelöscht:" & vbLf & pfad & vbLf & _
        IIf(nr <> 0, "(" & txt & ", Fehler " & nr & ")" & vbLf, "") & _
        "Vermutlich ist sie noch geöffnet (PDF-Anzeige, Explorer-Vorschau, Virenscanner, anderer Arbeitsplatz). " & _
        "Bitte schließen und in einigen Sekunden erneut versuchen."
End Sub
