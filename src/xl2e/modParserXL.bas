Attribute VB_Name = "modParserXL"
'==============================================================================
' Einlese-Vorlage für Excel-Rechnungen aus der BSI-Software (Blatt "PA400_00").
'
' Aufbau (alle Werte als Text mit Leerzeichen, deutsches Zahlenformat):
'   D4..D7   Anschrift, letzte Zeile "PLZ Ort"
'   I.. "Kun-Nr.    RechNr", darunter "11104   2602271"; K darunter: Datum
'   Ab "Art-Nr": B Art-Nr, D Bezeichnung, H Ka/Fa, I Flaschen, J Liter,
'                K Preis, L Betrag. Fässer: Preis je 100 Liter.
'   "LEERGUT-KONTO:": A Stock-Nr, B Bezeichnung, F alter Saldo, G geliefert,
'                H zurück, I neuer Saldo, J Differenz, K Pfandsatz, L Betrag
'   Summen in G/I/L: Zwischensumme (incl. ... Pfand), "+ 19,00 % MwSt",
'                RECHNUNGSBETRAG
' Python-Nachbau zum Testen: tests/xl_parser_test.py
'==============================================================================
Option Explicit

Private Enum Abschnitt
    abKopf = 0
    abPositionen = 1
    abLeergutInfo = 2
    abLeergut = 3
    abSummen = 4
End Enum

Public Sub ParseXL(ByVal ws As Worksheet)
    Dim z As Long, letzte As Long, modus As Abschnitt, m As Object, kopfZeile As Long
    Dim a As String, b As String, d As String, g As String, i As String, k As String, l As String
    Dim reLS As Object, reInkl As Object, reSatz As Object
    Dim steuerSaetze As New Collection, mwstSumme As Currency

    Set reLS = NeuesRegex("Lieferschein:\s+(\d+)\s+vom\s+(\d{1,2}\.\d{1,2}\.\d{4})")
    Set reInkl = NeuesRegex("\(incl\.\s+([\d.,]+-?)\s+Pfand\)")
    Set reSatz = NeuesRegex("([\d,]+)\s*%\s*MwSt")

    letzte = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    z = 1
    Do While z <= letzte
        a = ZellText(ws, z, 1): b = ZellText(ws, z, 2): d = ZellText(ws, z, 4)
        g = ZellText(ws, z, 7): i = ZellText(ws, z, 9): k = ZellText(ws, z, 11): l = ZellText(ws, z, 12)

        If InStr(i, "Kun-Nr") > 0 Then
            KopfLesen ws, z
            kopfZeile = z
            z = z + 2
            GoTo naechste
        End If

        If b = "Art-Nr" Then
            modus = abPositionen
        ElseIf Left$(b, 13) = "LEERGUT-KONTO" Then
            modus = abLeergut
            z = z + 2                                    ' zweite Überschriftenzeile überspringen
            GoTo naechste
        ElseIf Left$(b, 1) = ChrW(&H2500) Then          ' Trennlinie
            If modus = abLeergut Then modus = abSummen
            If modus = abLeergutInfo Then modus = abPositionen
        ElseIf d = "WARENWERT:" Then
            Feld("f_PDF_Warenwert").Value = ZahlDE(l)
            modus = abSummen
        ElseIf g = "Zwischensumme" Then
            Set m = Treffer(reInkl, i)
            If Not m Is Nothing Then Feld("f_PDF_Pfand").Value = ZahlDE(m.SubMatches(0))
            Feld("f_PDF_Netto").Value = ZahlDE(l)
        ElseIf Left$(g, 1) = "+" And InStr(g, "MwSt") > 0 Then
            Set m = Treffer(reSatz, g)
            If Not m Is Nothing Then steuerSaetze.Add CDbl(ZahlDE(m.SubMatches(0)))
            mwstSumme = mwstSumme + ZahlDE(l)
        ElseIf g = "RECHNUNGSBETRAG" Then
            Feld("f_PDF_Brutto").Value = ZahlDE(l)
        ElseIf modus = abPositionen Or modus = abLeergutInfo Then
            Set m = Treffer(reLS, d)
            If Not m Is Nothing Then
                If Len(Feld("f_LSNr").Value) = 0 Then
                    Feld("f_LSNr").Value = "'" & m.SubMatches(0)
                    Feld("f_LSDatum").Value = DatumDE(m.SubMatches(1))
                Else
                    Protokoll "Warnung", "Weiterer Lieferschein " & m.SubMatches(0) & ". In der E-Rechnung steht nur der erste."
                End If
            ElseIf Left$(d, 10) = "Lieferort:" Then
                Feld("f_LName").Value = Trim$(Mid$(d, 11))
            ElseIf d = "Leergutrückgabe:" Then
                modus = abLeergutInfo
            ElseIf modus = abPositionen And Len(k) > 0 And Len(l) > 0 Then
                PositionLesen ws, z
            ElseIf Len(b) > 0 Or Len(d) > 0 Then
                Protokoll "Info", IIf(modus = abLeergutInfo, "Leergutrückgabe (nur Information): ", "Nicht übernommen: ") & _
                                  b & " " & Application.WorksheetFunction.Trim(d)
            End If
        ElseIf modus = abLeergut Then
            If Len(k) > 0 Then PfandLesen ws, z
        ElseIf modus = abSummen And Len(d) > 0 Then
            Protokoll "Info", "Hinweis auf der Rechnung: " & d
        End If
naechste:
        z = z + 1
    Loop

    If kopfZeile = 0 Then Err.Raise vbObjectError + 40, , "Rechnungskopf (Kun-Nr./RechNr) nicht gefunden. Ist das eine BSI-Excel-Rechnung?"
    AnschriftLesen ws, kopfZeile
    Feld("f_PDF_MwSt").Value = mwstSumme
    SteuersatzSetzen steuerSaetze
End Sub

' "Kun-Nr.    RechNr" / "Datum" in Zeile z, Werte in der Zeile darunter
Private Sub KopfLesen(ByVal ws As Worksheet, ByVal z As Long)
    Dim t() As String, s As String
    s = Application.WorksheetFunction.Trim(ZellText(ws, z + 1, 9))
    t = Split(s, " ")
    If UBound(t) < 1 Then Err.Raise vbObjectError + 41, , "Kundennummer und Rechnungsnummer nicht lesbar: '" & s & "'"
    Feld("f_KundenNr").Value = "'" & t(0)
    Feld("f_RechNr").Value = "'" & t(1)
    Feld("f_RechDatum").Value = DatumDE(ZellText(ws, z + 1, 11))
End Sub

' Anschrift in Spalte D oberhalb des Rechnungskopfs: Name (eine oder mehrere
' Zeilen), Straße, "PLZ Ort"
Private Sub AnschriftLesen(ByVal ws As Worksheet, ByVal bisZeile As Long)
    Dim z As Long, zeilen As New Collection, n As Long, name As String, m As Object
    For z = 1 To bisZeile - 1
        If Len(ZellText(ws, z, 4)) > 0 Then zeilen.Add Application.WorksheetFunction.Trim(ZellText(ws, z, 4))
    Next z
    For n = zeilen.Count To 1 Step -1
        Set m = Treffer(NeuesRegex("^(\d{5})\s+(.+)$"), zeilen(n))
        If Not m Is Nothing Then Exit For
    Next n
    If m Is Nothing Or n < 3 Then Err.Raise vbObjectError + 42, , "Anschrift (Name, Straße, PLZ Ort) in Spalte D nicht erkannt."
    For z = 1 To n - 2
        name = name & IIf(Len(name) > 0, " ", "") & zeilen(z)
    Next z
    Feld("f_KName").Value = name
    Feld("f_KStrasse").Value = zeilen(n - 1)
    Feld("f_KPLZ").Value = "'" & m.SubMatches(0)
    Feld("f_KOrt").Value = Trim$(m.SubMatches(1))
End Sub

' Warenposition. Menge = Ka/Fa, wenn Ka/Fa x Preis den Betrag ergibt; sonst
' Hektoliter (Liter / 100), weil Fassbier je 100 Liter berechnet wird.
Private Sub PositionLesen(ByVal ws As Worksheet, ByVal z As Long)
    Dim artNr As String, bez As String, ka As Double, liter As Double, fl As String
    Dim preis As Currency, betrag As Currency

    artNr = ZellText(ws, z, 2)
    bez = Application.WorksheetFunction.Trim(ZellText(ws, z, 4))
    ka = CDbl(ZahlDE(ZellText(ws, z, 8)))
    fl = ZellText(ws, z, 9)
    liter = CDbl(ZahlDE(ZellText(ws, z, 10)))
    preis = ZahlDE(ZellText(ws, z, 11))
    betrag = ZahlDE(ZellText(ws, z, 12))

    If betrag < 0 And Len(artNr) = 0 Then
        PositionAnhaengen "Nachlass", "", bez, 1, "", -betrag, betrag, ""
    ElseIf ka <> 0 And RundenKfm(CCur(ka) * preis) = betrag Then
        PositionAnhaengen "Ware", artNr, bez, ka, Einstellung("Einheit Ware"), preis, betrag, IIf(Len(fl) > 0, fl & " Fl.", "")
    ElseIf liter <> 0 And RundenKfm(CCur(liter) / 100 * preis) = betrag Then
        PositionAnhaengen "Ware", artNr, bez, liter / 100, "HLT", preis, betrag, ka & " Fass, Preis je hl"
    Else
        PositionAnhaengen "Ware", artNr, bez, ka, Einstellung("Einheit Ware"), preis, betrag, "Menge prüfen"
        Protokoll "Warnung", "Position " & artNr & " " & bez & ": Weder Ka/Fa noch Liter ergeben mit dem Preis den Betrag."
    End If
End Sub

' Zeile des Leergutkontos. Lieferung und Rücknahme als getrennte Positionen
' (Einstellung "Pfand aufteilen" = Ja, Standard), sonst die Differenz.
Private Sub PfandLesen(ByVal ws As Worksheet, ByVal z As Long)
    Dim artNr As String, bez As String, einheit As String
    Dim lief As Double, zur As Double, satz As Currency, betrag As Currency

    artNr = ZellText(ws, z, 1)
    bez = "Pfand " & Application.WorksheetFunction.Trim(ZellText(ws, z, 2))
    lief = CDbl(ZahlDE(ZellText(ws, z, 7)))
    zur = CDbl(ZahlDE(ZellText(ws, z, 8)))
    satz = ZahlDE(ZellText(ws, z, 11))
    betrag = ZahlDE(ZellText(ws, z, 12))
    einheit = Einstellung("Einheit Pfand")
    If satz = 0 Or (lief = 0 And zur = 0) Then Exit Sub

    If RundenKfm(CCur(lief - zur) * satz) <> betrag Then
        Protokoll "Warnung", "Leergut " & bez & ": (geliefert - zurück) x Pfandsatz ergibt nicht den Betrag " & Format$(betrag, "#,##0.00")
    End If
    If UCase$(EinstellungOderLeer("Pfand aufteilen")) = "NEIN" Then
        If lief <> zur Then PositionAnhaengen "Pfand", artNr, bez, lief - zur, einheit, satz, betrag, ""
    Else
        If lief <> 0 Then PositionAnhaengen "Pfand", artNr, bez & " Lieferung", lief, einheit, satz, RundenKfm(CCur(lief) * satz), ""
        If zur <> 0 Then PositionAnhaengen "Pfand", artNr, bez & " Rücknahme", -zur, einheit, satz, RundenKfm(CCur(-zur) * satz), ""
    End If
End Sub

Private Function ZellText(ByVal ws As Worksheet, ByVal z As Long, ByVal sp As Long) As String
    ZellText = Trim$(CStr(ws.Cells(z, sp).Value))
End Function
