Attribute VB_Name = "modParserBSI"
'==============================================================================
' Einlese-Vorlage für Rechnungen, die mit der BSI-Software gedruckt werden.
' Erwartet den Text aus PdfZuText (pdftotext -layout).
'
' variante: Kennung aus Einstellungen "Layout-Variante". Weichen die
' BSI-Rechnungen eines Kunden im Detail ab, wird die Abweichung hier per
' Select Case variante behandelt, statt ein eigenes Modul anzulegen.
' "STANDARD" = Layout wie in den Mustern RE 2603976, 2603894, 2603816, 2603587.
'==============================================================================
Option Explicit

Private Enum Abschnitt
    abPositionen = 1
    abLeergut = 2
    abSummen = 3
End Enum

Public Sub ParseBSI(ByVal text As String, ByVal variante As String)
    Dim seiten() As String, zeilen() As String
    Dim i As Long, p As Long, start As Long
    Dim l As String, s As String, modus As Abschnitt
    Dim m As Object
    Dim anzahlLS As Long, lsListe As String, lgRest As String
    Dim reArtNr As Object
    Dim steuerSaetze As New Collection, mwstSumme As Currency

    ' Muster (VBScript-RegExp)
    Dim reKopf As Object, reLS As Object, rePos As Object, reAkt As Object, reLG As Object
    Dim reWW As Object, reZS As Object, reSt As Object, reRB As Object
    Dim reMarkt As Object, reILN As Object, reMail As Object, rePLZ As Object, reREGLN As Object
    Set reKopf = NeuesRegex("(\d+)\s+(\d{5,})\s+(\d{2}\.\d{2}\.\d{4})\s*$")
    Set reLS = NeuesRegex("^\s*\d+\s+Lieferschein:\s+(\d+)\s+vom\s+(\d{1,2}\.\d{1,2}\.\d{4})")
    Set rePos = NeuesRegex("^\s*(\d+)\s+(.+?)\s+(\d+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+-?)\s*$")
    Set reAkt = NeuesRegex("^\s*Aktion\s+(.+?)\s+(\d+)\s+Ka/F\s+([\d.,]+)\s+([\d.,]+)-\s*$")
    Set reLG = NeuesRegex("^\s*(\d{5})\s+(\S+(?: \S+)*)\s{2,}(?:.*\s)?(-?\d+-?)\s+(\d+,\d{2})\s+([\d.]*\d,\d{2}-?)\s*$")
    Set reWW = NeuesRegex("WARENWERT:\s+([\d.,]+-?)")
    Set reZS = NeuesRegex("Zwischensumme\s+\(incl\.\s+([\d.,]+-?)\s+Pfand\)\s+([\d.,]+-?)")
    Set reSt = NeuesRegex("([\d,]+)\s*%\s*MwSt\s+\(auf\s+([\d.,]+-?)\s*\)\s+([\d.,]+-?)")
    Set reRB = NeuesRegex("RECHNUNGSBETRAG\s+EUR\s+([\d.,]+-?)")
    Set reMarkt = NeuesRegex("Marktnummer:\s*(\d+)")
    Set reILN = NeuesRegex("(?:ILN|GLN):\s*(\d{13})")
    Set reREGLN = NeuesRegex("^RE-GLN:\s*([\d ]+)$")
    Set reMail = NeuesRegex("([\w.+-]+@[\w-]+(?:\.[\w-]+)+)")
    Set rePLZ = NeuesRegex("^(\d{5})\s+(.+)$")
    Set reArtNr = NeuesRegex("^\d{5}\s")

    Select Case UCase$(variante)
        Case "STANDARD"
        Case Else
            Protokoll "Warnung", "Layout-Variante '" & variante & "' unbekannt, verwende STANDARD."
    End Select

    seiten = Split(text, Chr$(12))

    ' ---- Kopf, Rechnungsempfänger und Lieferort (Seite 1) ----
    zeilen = Split(seiten(0), vbLf)
    Dim links As New Collection, einzug As Long
    For i = 0 To UBound(zeilen)
        l = zeilen(i)
        Set m = Treffer(reKopf, l)
        If Not m Is Nothing And Len(Feld("f_RechNr").Value) = 0 Then
            Feld("f_KundenNr").Value = m.SubMatches(0)
            Feld("f_RechNr").Value = m.SubMatches(1)
            Feld("f_RechDatum").Value = DatumDE(m.SubMatches(2))
        End If
        If InStr(l, "Wir lieferten an:") > 0 Then
            LieferortLesen zeilen, i
            Exit For
        End If
        ' Linke Spalte: Text mit geringem Einzug, bis zur ersten großen Lücke
        s = LTrim$(l)
        einzug = Len(l) - Len(s)
        If Len(s) > 0 And einzug < 30 Then
            If InStr(s, "     ") > 0 Then s = Left$(s, InStr(s, "     ") - 1)
            Set m = Treffer(reREGLN, Trim$(s))
            If Not m Is Nothing Then
                Feld("f_KGLN").Value = "'" & Replace(m.SubMatches(0), " ", "")
            ElseIf LCase$(Left$(s, 4)) <> "per " Then
                links.Add Trim$(s)
            End If
        End If
    Next i
    EmpfaengerSetzen links, rePLZ

    ' ---- Positionen, Leergut, Summen (alle Seiten) ----
    modus = abPositionen
    For p = 0 To UBound(seiten)
        zeilen = Split(seiten(p), vbLf)
        start = -1
        For i = 0 To UBound(zeilen)
            If InStr(zeilen(i), "ARTNR") > 0 And InStr(zeilen(i), "BEZEICHNUNG") > 0 Then start = i + 1: Exit For
        Next i
        If start >= 0 Then
            For i = start To UBound(zeilen)
                l = zeilen(i)
                s = Trim$(l)
                If Len(s) = 0 Then GoTo weiter
                If Left$(s, 1) = ChrW(&H2500) Or Left$(s, 1) = ChrW(&H2550) Then lgRest = "": GoTo weiter
                If Left$(s, 8) = "ÜBERTRAG" Then GoTo weiter
                If Left$(s, 13) = "LEERGUT-KONTO" Then modus = abLeergut: GoTo weiter
                If modus = abLeergut And Left$(s, 5) = "SALDO" Then GoTo weiter

                Set m = Treffer(reLS, l)
                If Not m Is Nothing Then
                    anzahlLS = anzahlLS + 1
                    lsListe = lsListe & IIf(anzahlLS > 1, ", ", "") & m.SubMatches(0) & " vom " & m.SubMatches(1)
                    If anzahlLS = 1 Then
                        Feld("f_LSNr").Value = m.SubMatches(0)
                        Feld("f_LSDatum").Value = DatumDE(m.SubMatches(1))
                    End If
                    GoTo weiter
                End If

                Set m = Treffer(reAkt, l)
                If Not m Is Nothing Then
                    PositionAnhaengen "Nachlass", "", "Aktion " & Trim$(m.SubMatches(0)), CDbl(m.SubMatches(1)), "", _
                        ZahlDE(m.SubMatches(2)), -ZahlDE(m.SubMatches(3)), ""
                    GoTo weiter
                End If

                Set m = Treffer(reWW, l)
                If Not m Is Nothing Then Feld("f_PDF_Warenwert").Value = ZahlDE(m.SubMatches(0)): modus = abSummen: GoTo weiter
                Set m = Treffer(reZS, l)
                If Not m Is Nothing Then
                    Feld("f_PDF_Pfand").Value = ZahlDE(m.SubMatches(0))
                    Feld("f_PDF_Netto").Value = ZahlDE(m.SubMatches(1))
                    GoTo weiter
                End If
                Set m = Treffer(reSt, l)
                If Not m Is Nothing Then
                    steuerSaetze.Add CDbl(ZahlDE(m.SubMatches(0)))
                    mwstSumme = mwstSumme + ZahlDE(m.SubMatches(2))
                    GoTo weiter
                End If
                Set m = Treffer(reRB, l)
                If Not m Is Nothing Then Feld("f_PDF_Brutto").Value = ZahlDE(m.SubMatches(0)): GoTo weiter

                If modus = abPositionen Then
                    Set m = Treffer(rePos, l)
                    If Not m Is Nothing Then
                        PositionAnhaengen "Ware", m.SubMatches(0), Trim$(m.SubMatches(1)), CDbl(m.SubMatches(2)), _
                            Einstellung("Einheit Ware"), ZahlDE(m.SubMatches(4)), ZahlDE(m.SubMatches(5)), m.SubMatches(3)
                        GoTo weiter
                    End If
                ElseIf modus = abLeergut Then
                    ' pdftotext (Xpdf, Windows) bricht Leergutzeilen teils in zwei Zeilen um:
                    ' "99001 Euro 0,5   1443-   20" / "36   1459-   16-  3,10   49,60-"
                    If Len(lgRest) > 0 Then
                        If Treffer(reArtNr, s) Is Nothing Then
                            Set m = Treffer(reLG, lgRest & "   " & s)
                            lgRest = ""
                            If Not m Is Nothing Then PfandAnhaengen m: GoTo weiter
                        End If
                        lgRest = ""
                    End If
                    Set m = Treffer(reLG, l)
                    If Not m Is Nothing Then PfandAnhaengen m: GoTo weiter
                    ' Zeile mit Artikelnummer ohne Pfandbetrag: Kontostand (nicht rechnungsrelevant)
                    ' oder erster Teil einer umgebrochenen Zeile - für die nächste Zeile merken
                    If Not Treffer(reArtNr, s) Is Nothing Then lgRest = s: GoTo weiter
                End If

                ' Alles andere: Zusatzangaben auswerten, sonst als Hinweis protokollieren
                Set m = Treffer(reMarkt, l)
                If Not m Is Nothing Then Feld("f_Marktnr").Value = m.SubMatches(0): GoTo weiter
                Set m = Treffer(reILN, l)
                If Not m Is Nothing Then
                    If Len(Feld("f_LGLN").Value) = 0 Then
                        Feld("f_LGLN").Value = "'" & m.SubMatches(0)
                    ElseIf Feld("f_LGLN").Value <> m.SubMatches(0) Then
                        Protokoll "Warnung", "Abweichende GLN im Text: " & s
                    End If
                    GoTo weiter
                End If
                Set m = Treffer(reMail, l)
                If Not m Is Nothing Then Feld("f_KEmail").Value = m.SubMatches(0): GoTo weiter
                Protokoll "Info", "Nicht übernommen: " & s
weiter:
            Next i
        End If
    Next p

    If anzahlLS > 1 Then Protokoll "Warnung", "Mehrere Lieferscheine: " & lsListe & ". In der E-Rechnung steht nur der erste."
    Feld("f_PDF_MwSt").Value = mwstSumme
    SteuersatzSetzen steuerSaetze
End Sub

Private Sub PfandAnhaengen(ByVal m As Object)
    PositionAnhaengen "Pfand", m.SubMatches(0), "Pfand " & Trim$(m.SubMatches(1)), _
        CDbl(ZahlDE(m.SubMatches(2))), Einstellung("Einheit Pfand"), ZahlDE(m.SubMatches(3)), _
        ZahlDE(m.SubMatches(4)), ""
End Sub

Private Sub LieferortLesen(zeilen() As String, ByVal i As Long)
    Dim z As Collection, s As String, m As Object
    Set z = New Collection
    z.Add Trim$(Mid$(zeilen(i), InStr(zeilen(i), "Wir lieferten an:") + 17))
    i = i + 1
    Do While i <= UBound(zeilen)
        s = Trim$(zeilen(i))
        If Len(s) = 0 Then Exit Do
        z.Add s
        i = i + 1
    Loop
    Feld("f_LName").Value = z(1)
    If z.Count >= 3 Then
        ' z. B. "GLN. 4388096598622 Henkerstegstraße 2a"
        Set m = Treffer(NeuesRegex("^GLN\.?:?\s*(\d{13})\s*(.*)$"), z(2))
        If m Is Nothing Then
            Feld("f_LZeile2").Value = z(2)
        Else
            Feld("f_LGLN").Value = "'" & m.SubMatches(0)
            Feld("f_LZeile2").Value = Trim$(m.SubMatches(1))
        End If
        Set m = Treffer(NeuesRegex("^(\d{5})\s+(.+)$"), z(z.Count))
        If Not m Is Nothing Then
            Feld("f_LPLZ").Value = "'" & m.SubMatches(0)
            Feld("f_LOrt").Value = Trim$(m.SubMatches(1))
        End If
    End If
End Sub

Private Sub EmpfaengerSetzen(links As Collection, rePLZ As Object)
    ' Erwartet: Name (1-n Zeilen), Straße, "PLZ Ort"
    Dim k As Long, j As Long, name As String, m As Object
    For k = 1 To links.Count
        If Not Treffer(rePLZ, links(k)) Is Nothing Then Exit For
    Next k
    If k > links.Count Or k < 3 Then
        Protokoll "Fehler", "Anschrift des Rechnungsempfängers nicht erkannt."
        Exit Sub
    End If
    For j = 1 To k - 2
        name = name & IIf(j > 1, " ", "") & links(j)
    Next j
    Set m = Treffer(rePLZ, links(k))
    Feld("f_KName").Value = name
    Feld("f_KStrasse").Value = links(k - 1)
    Feld("f_KPLZ").Value = "'" & m.SubMatches(0)
    Feld("f_KOrt").Value = Trim$(m.SubMatches(1))
End Sub

Private Sub PositionAnhaengen(ByVal typ As String, ByVal artNr As String, ByVal bez As String, _
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

' Bei genau einem MwSt-Satz gilt er für alle Zeilen. Bei mehreren muss der Anwender zuordnen.
Private Sub SteuersatzSetzen(saetze As Collection)
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
