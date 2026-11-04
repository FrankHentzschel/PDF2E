Attribute VB_Name = "modPruefen"
'==============================================================================
' Prüfung vor dem Erzeugen: Pflichtangaben und Summenabgleich mit der Quellrechnung
'==============================================================================
Option Explicit

Private mFehler As Long

Public Function RechnungPruefen() As Boolean
    Dim ws As Worksheet, z As Long, letzte As Long
    Dim ware As Currency, pfand As Currency, typ As String, erwartet As Currency

    Application.Calculate
    mFehler = 0
    Protokoll "Prüfung", "---- Prüfung gestartet ----"

    ' Pflichtangaben Rechnung
    Pflicht "f_RechNr", "Rechnungsnummer"
    Pflicht "f_RechDatum", "Rechnungsdatum"
    Pflicht "f_KName", "Name Rechnungsempfänger"
    Pflicht "f_KPLZ", "PLZ Rechnungsempfänger"
    Pflicht "f_KOrt", "Ort Rechnungsempfänger"
    Pflicht "f_Kondition", "Zahlungskondition"
    Pflicht "f_Konto", "Bankkonto"
    If Len(Feld("f_Konto").Value) > 0 And TabellenZeile("Bankkonten", CStr(Feld("f_Konto").Value)) = 0 Then
        Fehler "Bankkonto '" & Feld("f_Konto").Value & "' gibt es im Blatt 'Bankkonten' nicht."
    End If

    ' GLN (13 Stellen mit Prüfziffer), sofern angegeben
    GlnPruefen "f_KGLN", "GLN Rechnungsempfänger"
    GlnPruefen "f_LGLN", "GLN Lieferort"

    ' Sondervorgaben des Empfängers (Blatt "Empfaenger", Spalte "Vorgaben")
    If EmpfaengerVorgabe() = "REWE" Then
        Pflicht "f_KGLN", "REWE: GLN Rechnungsempfänger (RE-GLN)"
        Pflicht "f_LGLN", "REWE: GLN Lieferort (WE-GLN)"
        If Len(Trim$(CStr(Feld("f_KUstId").Value))) = 0 Then
            Fehler "REWE: USt-IdNr. des Rechnungsempfängers fehlt. Bitte im Blatt 'Empfaenger' (Spalte USt-IdNr.) eintragen."
        End If
        If UCase$(EinstellungOderLeer("Pfand aufteilen")) = "NEIN" Then
            Fehler "REWE: Pfand muss in Lieferung und Rücknahme aufgeteilt sein (Einstellung 'Pfand aufteilen' = Ja)."
        End If
    End If

    ' Pflichtangaben Brauerei
    Dim k As Variant
    For Each k In Array("Name", "Straße", "PLZ", "Ort", "Land", "USt-IdNr")
        If Len(Stamm(CStr(k))) = 0 Then Fehler "Brauerei-Stammdaten: '" & k & "' fehlt."
    Next k

    PlatzhalterPruefen

    ' Positionen
    Set ws = ThisWorkbook.Worksheets(BLATT_POSITIONEN)
    letzte = ws.Cells(ws.Rows.Count, SP_TYP).End(xlUp).Row
    If letzte < 2 Then Fehler "Keine Positionen vorhanden."
    For z = 2 To letzte
        typ = ws.Cells(z, SP_TYP).Value
        If Len(ws.Cells(z, SP_USTSATZ).Value) = 0 Then Fehler "Zeile " & z & ": USt-Satz fehlt."
        If typ = "Ware" Or typ = "Pfand" Then
            erwartet = RundenKfm(CCur(ws.Cells(z, SP_MENGE).Value) * CCur(ws.Cells(z, SP_PREIS).Value))
            If erwartet <> CCur(ws.Cells(z, SP_BETRAG).Value) Then
                Fehler "Zeile " & z & ": Menge x Preis = " & Format$(erwartet, "#,##0.00") & _
                       ", Betrag laut Rechnung = " & Format$(ws.Cells(z, SP_BETRAG).Value, "#,##0.00")
            End If
        End If
        If typ = "Pfand" Then pfand = pfand + ws.Cells(z, SP_BETRAG).Value Else ware = ware + ws.Cells(z, SP_BETRAG).Value
    Next z

    ' Summenabgleich mit den Werten aus der Quellrechnung (Felder f_PDF_*)
    Abgleich "Warenwert (Ware - Aktionen)", ware, Feld("f_PDF_Warenwert").Value
    Abgleich "Pfand", pfand, Feld("f_PDF_Pfand").Value
    Abgleich "Nettosumme", Feld("f_Netto").Value, Feld("f_PDF_Netto").Value
    Abgleich "MwSt", Feld("f_MwSt").Value, Feld("f_PDF_MwSt").Value
    Abgleich "Rechnungsbetrag", Feld("f_Brutto").Value, Feld("f_PDF_Brutto").Value

    RechnungPruefen = (mFehler = 0)
    Feld("f_Status").Value = IIf(mFehler = 0, "OK - bereit zum Erzeugen", mFehler & " Fehler - siehe Protokoll")
    Protokoll "Prüfung", "---- Ergebnis: " & Feld("f_Status").Value & " ----"
End Function

Private Sub Pflicht(ByVal feldName As String, ByVal bezeichnung As String)
    If Len(Trim$(CStr(Feld(feldName).Value))) = 0 Then Fehler bezeichnung & " fehlt."
End Sub

' Platzhalter in den Stammdaten (z. B. bei neuen Kunden, solange echte Daten fehlen).
' Mit dem Adapter SMARTTOOLS ist das ein Fehler, damit keine echte Rechnung mit
' Fantasiedaten entsteht; im Testmodus (STUB) nur eine Warnung.
Private Sub PlatzhalterPruefen()
    Dim muster As Variant, blatt As Variant, c As Range, m As Variant, meldung As String
    muster = Array("DE123456789", "TESTDEFFXXX", "HRB 12345", "123/456/78901", "DE87123456781234567890", "DE58876543211234567890")
    For Each blatt In Array("Brauerei", "Bankkonten")
        For Each c In ThisWorkbook.Worksheets(blatt).UsedRange
            For Each m In muster
                If InStr(1, Replace(CStr(c.Value), " ", ""), Replace(m, " ", ""), vbTextCompare) > 0 Then
                    meldung = "Platzhalter im Blatt '" & blatt & "', Zelle " & c.Address(False, False) & ": " & c.Value
                    If UCase$(EinstellungOderLeer("E-Rechnung-Adapter")) = "STUB" Then
                        Protokoll "Warnung", meldung
                    Else
                        Fehler meldung & " - bitte die echten Daten eintragen."
                    End If
                End If
            Next m
        Next c
    Next blatt
End Sub

Private Sub GlnPruefen(ByVal feldName As String, ByVal bezeichnung As String)
    Dim s As String, i As Long, summe As Long
    s = Trim$(CStr(Feld(feldName).Value))
    If Len(s) = 0 Then Exit Sub
    If Len(s) <> 13 Or Not s Like String$(13, "#") Then Fehler bezeichnung & " '" & s & "' hat nicht 13 Ziffern.": Exit Sub
    For i = 1 To 12
        summe = summe + CLng(Mid$(s, i, 1)) * IIf(i Mod 2 = 0, 3, 1)
    Next i
    If (10 - summe Mod 10) Mod 10 <> CLng(Right$(s, 1)) Then Fehler bezeichnung & " '" & s & "': Prüfziffer stimmt nicht."
End Sub

Private Sub Abgleich(ByVal was As String, ByVal berechnet As Variant, ByVal laut As Variant)
    If Len(CStr(laut)) = 0 Then
        Fehler was & ": Wert aus der Rechnung fehlt."
    ElseIf RundenKfm(CCur(berechnet)) <> RundenKfm(CCur(laut)) Then
        Fehler was & ": berechnet " & Format$(berechnet, "#,##0.00") & ", laut Rechnung " & Format$(laut, "#,##0.00")
    End If
End Sub

Private Sub Fehler(ByVal meldung As String)
    mFehler = mFehler + 1
    Protokoll "Fehler", meldung
End Sub
