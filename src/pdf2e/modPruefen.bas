Attribute VB_Name = "modPruefen"
'==============================================================================
' Prüfung vor dem Erzeugen: Pflichtangaben und Summenabgleich mit der Quellrechnung
'==============================================================================
Option Explicit

Private mFehler As Long

Public Function RechnungPruefen() As Boolean
    Dim ws As Worksheet, z As Long, letzte As Long
    Dim ware As Currency, pfand As Currency, typ As String, erwartet As Currency
    Dim wareB As Currency, pfandB As Currency, brutto As Currency, toleranz As Currency

    Application.Calculate
    mFehler = 0
    Protokoll "Prüfung", "---- Prüfung gestartet ----"
    EmpfaengerNachtragen            ' z. B. USt-IdNr. erst nach dem Einlesen im Blatt Empfaenger eingetragen

    ' Pflichtangaben Rechnung
    Pflicht "f_RechNr", "Rechnungsnummer"
    Pflicht "f_RechDatum", "Rechnungsdatum"
    Pflicht "f_KName", "Name Rechnungsempfänger"
    Pflicht "f_KPLZ", "PLZ Rechnungsempfänger"
    Pflicht "f_KOrt", "Ort Rechnungsempfänger"
    Pflicht "f_Kondition", "Zahlungskondition"
    Pflicht "f_Konto", "Bankkonto"
    If IstLastschrift() Then
        If Not IsDate(Feld("f_Einzug").Value) Then Fehler "SEPA-Lastschrift: Einzugsdatum fehlt."
        Pflicht "f_GlaeubigerId", "SEPA-Lastschrift: Gläubiger-ID"
    End If
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

    If EmpfaengerVorgabe() = "EDEKA" Then
        If UCase$(EinstellungOderLeer("Pfand aufteilen")) = "NEIN" Then
            Fehler "EDEKA: Pfand muss in Lieferung und Rückgabe aufgeteilt sein (Einstellung 'Pfand aufteilen' = Ja)."
        End If
    End If
    PfandGtinPruefen EmpfaengerVorgabe() = "EDEKA"
    If EmpfaengerVorgabe() = "G5" Then G5Pruefen

    ' Innergemeinschaftliche Lieferung (USt-Kategorie K, EN 16931 BR-IC)
    If KategorieVorhanden("K") Then
        If Len(UStIdBereinigt(Feld("f_KUstId").Value)) = 0 Then
            Fehler "Innergemeinschaftliche Lieferung: USt-IdNr. des Empfängers fehlt (Pflicht). Bitte im Blatt 'Empfaenger' beim Kunden " & _
                   EmpfaengerSchluessel() & " eintragen."
        End If
        If UCase$(Trim$(CStr(Feld("f_KLand").Value))) = "DE" Then
            Fehler "Innergemeinschaftliche Lieferung, aber Land des Empfängers ist DE. Bitte Feld 'Land' im Blatt 'Rechnung' prüfen."
        End If
        If Len(Trim$(CStr(Feld("f_LLand").Value))) = 0 Then Fehler "Innergemeinschaftliche Lieferung: Lieferland fehlt."
    End If
    UStIdPruefen Feld("f_KUstId").Value, "USt-IdNr. des Rechnungsempfängers: im Blatt 'Empfaenger' beim Kunden " & _
                 EmpfaengerSchluessel() & " (Spalte USt-IdNr.) korrigieren oder leeren und neu einlesen."

    ' Pflichtangaben Brauerei
    Dim k As Variant
    For Each k In Array("Name", "Straße", "PLZ", "Ort", "Land", "USt-IdNr")
        If Len(Stamm(CStr(k))) = 0 Then Fehler "Brauerei-Stammdaten: '" & k & "' fehlt."
    Next k
    UStIdPruefen Stamm("USt-IdNr"), "USt-IdNr. der Brauerei: im Blatt 'Brauerei' korrigieren."

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
        ' Bruttobeträge: Netto zurückgerechnet, für den Abgleich mit den Bruttowerten laut Rechnung
        If IsNumeric(ws.Cells(z, SP_USTSATZ).Value) Then
            brutto = RundenKfm(CCur(ws.Cells(z, SP_BETRAG).Value) * (1 + CCur(ws.Cells(z, SP_USTSATZ).Value) / 100))
            If typ = "Pfand" Then pfandB = pfandB + brutto Else wareB = wareB + brutto
        End If
    Next z

    ' Summenabgleich mit den Werten aus der Quellrechnung (Felder f_PDF_*)
    ' Bei Bruttobeträgen ("ENTHALTENE MWSt.") sind Warenwert und Pfand laut Rechnung
    ' brutto; je Position darf die Umrechnung um bis zu 1 Cent abweichen. Der
    ' Rechnungsbetrag muss mit Rundungsbetrag (BT-114) genau stimmen.
    If BruttoBetraege() Then
        toleranz = 0.01 * (letzte - 1)
        Abgleich "Warenwert brutto (Ware - Nachlässe)", wareB, Feld("f_PDF_Warenwert").Value, toleranz
        Abgleich "Pfand brutto", pfandB, Feld("f_PDF_Pfand").Value, toleranz
        Abgleich "Nettosumme", Feld("f_Netto").Value, Feld("f_PDF_Netto").Value, toleranz
        Abgleich "MwSt", Feld("f_MwSt").Value, Feld("f_PDF_MwSt").Value, toleranz
    Else
        Abgleich "Warenwert (Ware - Nachlässe)", ware, Feld("f_PDF_Warenwert").Value
        Abgleich "Pfand", pfand, Feld("f_PDF_Pfand").Value
        Abgleich "Nettosumme", Feld("f_Netto").Value, Feld("f_PDF_Netto").Value
        Abgleich "MwSt", Feld("f_MwSt").Value, Feld("f_PDF_MwSt").Value
    End If
    Abgleich "Rechnungsbetrag", Zahlbetrag(), Feld("f_PDF_Brutto").Value

    RechnungPruefen = (mFehler = 0)
    Feld("f_Status").Value = IIf(mFehler = 0, "OK - bereit zum Erzeugen", mFehler & " Fehler - siehe Protokoll")
    Protokoll "Prüfung", "---- Ergebnis: " & Feld("f_Status").Value & " ----"
End Function

' USt-IdNr. mit Länderkennzeichen (EN 16931, BR-CO-09), sonst Fehler mit Hinweis, wo korrigieren
Private Sub UStIdPruefen(ByVal ustId As Variant, ByVal wo As String)
    Dim meldung As String
    meldung = UStIdFehler(UStIdBereinigt(ustId))
    If Len(meldung) > 0 Then Fehler meldung & " " & wo
End Sub

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

' GTIN der Pfandartikel: Prüfziffer, sofern eingetragen; bei Pflicht (EDEKA)
' fehlende GTIN als Fehler. Jeder Pfandartikel wird nur einmal gemeldet.
Private Sub PfandGtinPruefen(ByVal pflicht As Boolean)
    Dim ws As Worksheet, z As Long, nr As String, gtin As String, gemeldet As String
    Set ws = ThisWorkbook.Worksheets(BLATT_POSITIONEN)
    For z = 2 To ws.Cells(ws.Rows.Count, SP_TYP).End(xlUp).Row
        nr = Trim$(CStr(ws.Cells(z, SP_ARTNR).Value))
        If ws.Cells(z, SP_TYP).Value = "Pfand" And InStr(gemeldet, "|" & nr & "|") = 0 Then
            gemeldet = gemeldet & "|" & nr & "|"
            gtin = PfandGtin(nr)
            If Len(gtin) = 0 Then
                If pflicht Then Fehler "EDEKA: GTIN für Pfandartikel " & nr & " fehlt (Blatt '" & BLATT_PFANDARTIKEL & "')."
            ElseIf Not GtinGueltig(gtin) Then
                Fehler "GTIN '" & gtin & "' für Pfandartikel " & nr & " ist ungültig (Länge oder Prüfziffer)."
            End If
        End If
    Next z
End Sub

' GS1-Prüfziffer für GTIN-8, -12, -13 und -14
Private Function GtinGueltig(ByVal s As String) As Boolean
    Dim i As Long, summe As Long, n As Long
    n = Len(s)
    If n <> 8 And n <> 12 And n <> 13 And n <> 14 Then Exit Function
    If Not s Like String$(n, "#") Then Exit Function
    For i = 1 To n - 1                            ' von rechts: Gewicht 3, 1, 3, ...
        summe = summe + CLng(Mid$(s, n - i, 1)) * IIf(i Mod 2 = 1, 3, 1)
    Next i
    GtinGueltig = ((10 - summe Mod 10) Mod 10 = CLng(Right$(s, 1)))
End Function

' toleranz > 0 (nur bei Bruttobeträgen): kleinere Abweichungen sind Rundung, nur Info
Private Sub Abgleich(ByVal was As String, ByVal berechnet As Variant, ByVal laut As Variant, _
                     Optional ByVal toleranz As Currency = 0)
    Dim diff As Currency
    If Len(CStr(laut)) = 0 Then
        Fehler was & ": Wert aus der Rechnung fehlt."
        Exit Sub
    End If
    diff = Abs(RundenKfm(CCur(berechnet)) - RundenKfm(CCur(laut)))
    If diff = 0 Then Exit Sub
    If diff <= toleranz Then
        Protokoll "Info", was & ": berechnet " & Format$(berechnet, "#,##0.00") & ", laut Rechnung " & Format$(laut, "#,##0.00") & _
                          " (Rundung aus der Umrechnung brutto in netto)"
    Else
        Fehler was & ": berechnet " & Format$(berechnet, "#,##0.00") & ", laut Rechnung " & Format$(laut, "#,##0.00")
    End If
End Sub

Private Sub Fehler(ByVal meldung As String)
    mFehler = mFehler + 1
    Protokoll "Fehler", meldung
End Sub

' G5-Zentralregulierung (SGH Service, Merkblatt 11/2026): ZR-Nummer des Händlers
' (BT-10) und G5-Lieferantennummer der Brauerei (BT-29) sind Pflicht im XML
Private Sub G5Pruefen()
    Dim zr As String, lief As String
    zr = EmpfaengerZRNummer()
    If Len(zr) = 0 Then
        Fehler "G5: ZR-Nummer des Rechnungsempfängers fehlt. Bitte im Blatt 'Empfaenger' beim Kunden " & _
               EmpfaengerSchluessel() & " (Spalte ZR-Nummer) eintragen."
    ElseIf Not G5NummerOk(zr, False) Then
        Fehler "G5: ZR-Nummer '" & zr & "' ist ungültig (7 Ziffern, beginnt mit 27). Blatt 'Empfaenger', Kunde " & _
               EmpfaengerSchluessel() & "."
    End If
    lief = Replace(Trim$(EinstellungOderLeer("G5 Lieferantennummer")), " ", "")
    If Len(lief) = 0 Then
        Fehler "G5: Lieferantennummer der Brauerei fehlt (Blatt 'Einstellungen', Zeile 'G5 Lieferantennummer')."
    ElseIf Not G5NummerOk(lief, True) Then
        Fehler "G5: Lieferantennummer '" & lief & "' ist ungültig (7 Ziffern, beginnt mit 27, endet auf 0). Blatt 'Einstellungen'."
    End If
End Sub

' 7 Ziffern, beginnt mit 27; Lieferantennummer endet zusätzlich auf 0
Private Function G5NummerOk(ByVal nr As String, ByVal endetAufNull As Boolean) As Boolean
    If Not (nr Like "27#####") Then Exit Function
    If endetAufNull And Right$(nr, 1) <> "0" Then Exit Function
    G5NummerOk = True
End Function
