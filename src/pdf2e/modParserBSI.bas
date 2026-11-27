Attribute VB_Name = "modParserBSI"
'==============================================================================
' Einlese-Vorlage für Rechnungen, die mit der BSI-Software gedruckt werden.
' Erwartet den Text aus PdfZuText (pdftotext -layout).
'
' variante: Kennung aus Einstellungen "Layout-Variante". Weichen die
' BSI-Rechnungen eines Kunden im Detail ab, wird die Abweichung hier per
' Select Case variante behandelt, statt ein eigenes Modul anzulegen.
' "STANDARD" = Layout wie in den Mustern RE 2603976, 2603894, 2603816, 2603587.
' Minus-Rechnung mit Rückware und Bruttobeträgen ("ENTHALTENE MWSt.") wie RE 2603475:
' Mengen der Rückware negativ, Preise und Pfand werden in netto umgerechnet.
' Tabellen-Layout (Spalte Pos., MwSt je Zeile, RE 2604269): ParseBSITabelle unten,
' Text mit pdftotext -raw; Auswahl in modProdukt.QuelleEinlesen (IstBSITabelle).
'==============================================================================
Option Explicit

' EU-Mitgliedstaaten außer DE (ISO 3166, Griechenland = GR) für die steuerfreie
' innergemeinschaftliche Lieferung (USt-Kategorie K)
Private Const EU_LAENDER As String = "|AT|BE|BG|CY|CZ|DK|EE|ES|FI|FR|GR|HR|HU|IE|IT|LT|LU|LV|MT|NL|PL|PT|RO|SE|SI|SK|"

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
    Dim anzahlLS As Long, lsListe As String, lgOffen As String, z As String
    Dim reArtNr As Object, reLGRest As Object
    Dim steuerSaetze As New Collection, mwstSumme As Currency

    ' Muster (VBScript-RegExp)
    Dim reKopf As Object, reLS As Object, rePos As Object, reAkt As Object, reLG As Object
    Dim reWW As Object, reZS As Object, reSt As Object, reRB As Object
    Dim reMarkt As Object, reILN As Object, reMail As Object, rePLZ As Object, reREGLN As Object
    Dim reNachlass As Object, reNW As Object, reSepaBetrag As Object, reSepaMandat As Object, reSepaHinweis As Object
    Dim reWWLeer As Object, reIG As Object, reUStId As Object, igLieferung As Boolean, reFiliale As Object
    Dim reMinusRB As Object, reEnthalten As Object, reRueckware As Object, reTelefon As Object
    Dim bruttoBetraege As Boolean, lgRueckgabe As Boolean, menge As Double, betrag As Currency
    Set reKopf = NeuesRegex("(\d+)\s+(\d{5,})\s+(\d{1,2}\.\d{1,2}\.\d{4})\s*$")
    Set reLS = NeuesRegex("^\s*\d+\s+Lieferschein:\s+(\d+)\s+vom\s+(\d{1,2}\.\d{1,2}\.\d{4})")
    ' Position: ArtNr, Bezeichnung, Ka/Fa, Hl./Fl. (fehlt z. B. bei Paletten), Preis, Betrag,
    ' am Ende optional ein "*" (z. B. "136 EURO-Palette  29  12,000  348,00 *")
    ' Menge auch negativ, z. B. "59 Rabatt   1-   15,000   15,00-*" (RE 2604315)
    Set rePos = NeuesRegex("^\s*(\d+)\s+(.+?)\s+(\d+-?)\s+(?:([\d.,]+)\s+)?([\d.,]+)\s+([\d.,]+-?)(?:\s*\*)?\s*$")
    Set reAkt = NeuesRegex("^\s*Aktion\s+(.+?)\s+(\d+)\s+Ka/F\s+([\d.,]+)\s+([\d.,]+)-\s*$")
    Set reLG = NeuesRegex("^\s*(\d{5})\s+(\S+(?: \S+)*)\s{2,}(?:(.*\S)\s+)?(-?\d+-?)\s+(\d+,\d{2})\s+([\d.]*\d,\d{2}-?)\s*$")
    Set reWW = NeuesRegex("WARENWERT:\s+([\d.,]+-?)")
    Set reZS = NeuesRegex("Zwischensumme\s+\(incl\.\s+([\d.,]+-?)\s+Pfand\)\s+([\d.,]+-?)")
    Set reSt = NeuesRegex("([\d,]+)\s*%\s*MwSt\s+\(auf\s+([\d.,]+-?)\s*\)\s+([\d.,]+-?)")
    Set reRB = NeuesRegex("RECHNUNGSBETRAG\s+EUR\s+([\d.,]+-?)")
    ' Minus-Rechnung mit Bruttobeträgen, z. B.
    ' "MINUS-RECHNUNGSBETRAG    incl.   440,50- Pfand) EUR    772,50-"
    ' "ENTHALTENE MWSt. 19,00 % (aus   772,50-)    123,34-"
    ' auch ohne MINUS mit Klammer: "RECHNUNGSBETRAG (incl.   43,40 Pfand) EUR   244,40" (RE 2603531)
    Set reMinusRB = NeuesRegex("RECHNUNGSBETRAG\s+\(?incl\.\s+([\d.,]+-?)\s+Pfand\)\s+EUR\s+([\d.,]+-?)")
    Set reEnthalten = NeuesRegex("ENTHALTENE\s+MWSt\.?\s+([\d,]+)\s*%\s*\(aus\s+([\d.,]+-?)\s*\)\s+([\d.,]+-?)")
    Set reRueckware = NeuesRegex("^\*+\s*R" & ChrW(252) & "ckware\s*\*+$")
    ' Nachlass auf den Warenwert, z. B.
    ' "Skonto     Gesamtes Vergütung     299,75 EU     3,00     8,99-"
    ' Gruppen: 0 Art, 1 Text, 2 Basis, 3 Prozent, 4 Betrag
    Set reNachlass = NeuesRegex("^\s*(\S+)\s+(?:(.*?)\s+)?([\d.]*\d,\d{2})\s+EUR?\s+(\d+,\d{2})\s+([\d.]*\d,\d{2})-\s*$")
    Set reNW = NeuesRegex("NETTOWERT:\s+([\d.,]+-?)")
    Set reWWLeer = NeuesRegex("WARENWERT:\s*$")             ' z. B. reine Leergutrückgabe
    ' Steuerfreie innergemeinschaftliche Lieferung mit USt-IdNr. des Empfängers
    ' auch "** Steuerfreie Lieferung nach Italien ab Brauereihof **" (RE 2604314)
    Set reIG = NeuesRegex("[Ss]teuerfreie\s+(?:innergemeinschaftliche\s+)?Lieferung")
    ' "USt-ID: IT…", auch "UST - ID: IT…" (RE 2604314)
    Set reUStId = NeuesRegex("^\s*U[Ss][Tt]\s*-?\s*ID(?:\s*-?\s*Nr\.?)?\s*:?\s+([A-Za-z]{2}[\w .-]*\w)\s*$")
    ' SEPA-Lastschrift (Text unter dem Rechnungsbetrag)
    Set reSepaBetrag = NeuesRegex("Forderung von\s+([\d.,]+)\s+ziehen wir mit einer SEPA-Lastschrift")
    Set reSepaMandat = NeuesRegex("^\s*(\S+)\s+und mit der Gl(?:ae|" & ChrW(228) & ")ubiger-ID\s+(\S+)\s+am\s+(\d{1,2}\.\d{1,2}\.\d{4})\s+ein")
    Set reSepaHinweis = NeuesRegex("^\s*F(?:ae|" & ChrW(228) & ")llt der Termin auf ein Wochenende")
    Set reMarkt = NeuesRegex("Marktnummer:\s*(\d+)")
    Set reILN = NeuesRegex("(?:ILN|GLN):\s*(\d{13})")
    ' Abweichende Lieferanschrift unter dem Nettowert, z. B.
    ' "+ Lieferfiliale: 95233 Helmbrechts, Frankenstr. 13 *"   Gruppen: 0 PLZ, 1 Ort, 2 Straße
    Set reFiliale = NeuesRegex("Lieferfiliale:\s*(\d{5})\s+([^,]+?)\s*,\s*(.+?)(?:\s*\*)?\s*$")
    Set reREGLN = NeuesRegex("^RE-GLN:\s*([\d ]+)$")
    Set reMail = NeuesRegex("([\w.+-]+@[\w-]+(?:\.[\w-]+)+)")
    ' "96450 Coburg" oder mit Länderkennzeichen "I- 39012 Meran", "A-6020 Innsbruck"
    ' Gruppen: 0 Kennzeichen, 1 PLZ mit Kennzeichen, 2 PLZ ohne, 3 Ort
    Set rePLZ = NeuesRegex("^(?:([A-Z]{1,3})\s*-\s*(\d{4,5})|(\d{5}))\s+(.+)$")
    Set reArtNr = NeuesRegex("^\d{5}\s")
    ' Telefonnummer im Anschriftenblock, z. B. "0170-4383514" (Barverkauf ab Rampe, RE 2603531)
    Set reTelefon = NeuesRegex("^(?:Tel\.?:?\s*)?\+?\d[\d \-/()]{5,}\d$")
    Set reLGRest = NeuesRegex("\d+,\d{2}\s+[\d.]*\d,\d{2}-?\s*$")

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
            ElseIf Not Treffer(reTelefon, Trim$(s)) Is Nothing Then
                Protokoll "Info", "Telefonnummer im Anschriftenblock nicht als Name übernommen: " & Trim$(s)
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
                If Left$(s, 1) = ChrW(&H2500) Or Left$(s, 1) = ChrW(&H2550) Then
                    LeergutAbschliessen lgOffen, reArtNr
                    lgRueckgabe = False
                    GoTo weiter
                End If
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
                If Not Treffer(reWWLeer, l) Is Nothing Then Feld("f_PDF_Warenwert").Value = 0: modus = abSummen: GoTo weiter
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
                Set m = Treffer(reMinusRB, l)
                If Not m Is Nothing Then
                    Feld("f_PDF_Pfand").Value = ZahlDE(m.SubMatches(0))
                    Feld("f_PDF_Brutto").Value = ZahlDE(m.SubMatches(1))
                    GoTo weiter
                End If
                Set m = Treffer(reEnthalten, l)
                If Not m Is Nothing Then
                    steuerSaetze.Add CDbl(ZahlDE(m.SubMatches(0)))
                    mwstSumme = mwstSumme + ZahlDE(m.SubMatches(2))
                    bruttoBetraege = True
                    GoTo weiter
                End If
                Set m = Treffer(reRB, l)
                If Not m Is Nothing Then Feld("f_PDF_Brutto").Value = ZahlDE(m.SubMatches(0)): GoTo weiter

                If modus = abSummen Then
                    Set m = Treffer(reNachlass, l)
                    If Not m Is Nothing Then NachlassAnhaengen m: GoTo weiter
                    ' Warenwert nach Nachlässen: ersetzt den WARENWERT für den Abgleich
                    Set m = Treffer(reNW, l)
                    If Not m Is Nothing Then Feld("f_PDF_Warenwert").Value = ZahlDE(m.SubMatches(0)): GoTo weiter
                End If

                Set m = Treffer(reSepaBetrag, l)
                If Not m Is Nothing Then
                    If Len(Feld("f_PDF_Brutto").Value) > 0 Then
                        If ZahlDE(m.SubMatches(0)) <> CCur(Feld("f_PDF_Brutto").Value) Then
                            Protokoll "Warnung", "SEPA-Lastschrift über " & m.SubMatches(0) & " weicht vom Rechnungsbetrag ab."
                        End If
                    End If
                    GoTo weiter
                End If
                Set m = Treffer(reSepaMandat, l)
                If Not m Is Nothing Then
                    Feld("f_Mandat").Value = "'" & m.SubMatches(0)
                    Feld("f_GlaeubigerId").Value = "'" & m.SubMatches(1)
                    Feld("f_Einzug").Value = DatumDE(m.SubMatches(2))
                    Protokoll "Info", "SEPA-Lastschrift erkannt: Mandat " & m.SubMatches(0) & ", Gläubiger-ID " & _
                                      m.SubMatches(1) & ", Einzug am " & m.SubMatches(2) & ". Zahlungsart in der E-Rechnung: Lastschrift (59)."
                    GoTo weiter
                End If
                If Not Treffer(reSepaHinweis, l) Is Nothing Then GoTo weiter
                ' "** Rückware **" unter einer Position: Menge ist dort schon negativ übernommen
                If Not Treffer(reRueckware, s) Is Nothing Then GoTo weiter
                ' "Leergutrückgabe:" mit Mengen ohne Betrag; das Pfand steht im LEERGUT-KONTO
                If modus = abPositionen And s = "Leergutr" & ChrW(252) & "ckgabe:" Then lgRueckgabe = True: GoTo weiter

                If Not Treffer(reIG, l) Is Nothing Then
                    igLieferung = True
                    Protokoll "Info", "Steuerfreie innergemeinschaftliche Lieferung: USt-Kategorie K, 0 %, Befreiungsgrund VATEX-EU-IC."
                    GoTo weiter
                End If
                Set m = Treffer(reUStId, l)
                If Not m Is Nothing Then
                    Feld("f_KUstId").Value = "'" & UStIdBereinigt(m.SubMatches(0))
                    Protokoll "Info", "USt-IdNr. des Empfängers aus der Rechnung übernommen: " & UStIdBereinigt(m.SubMatches(0))
                    GoTo weiter
                End If

                If modus = abPositionen Then
                    Set m = Treffer(rePos, l)
                    If Not m Is Nothing Then
                        menge = CDbl(ZahlDE(m.SubMatches(2)))    ' "1-" = -1 (Rabatt)
                        betrag = ZahlDE(m.SubMatches(5))
                        If betrag < 0 And menge > 0 Then menge = -menge   ' Rückware: "4  80  14,000  56,00-"
                        PositionAnhaengen "Ware", m.SubMatches(0), Trim$(m.SubMatches(1)), menge, _
                            Einstellung("Einheit Ware"), ZahlDE(m.SubMatches(4)), betrag, m.SubMatches(3) & ""
                        GoTo weiter
                    End If
                    If lgRueckgabe And Not Treffer(reArtNr, s) Is Nothing Then GoTo weiter
                ElseIf modus = abLeergut Then
                    Set m = Treffer(reLG, l)
                    If Not m Is Nothing Then LeergutAbschliessen lgOffen, reArtNr: PfandAnhaengen m: GoTo weiter
                    ' pdftotext (Xpdf, Windows) verteilt eine Leergutzeile teils auf zwei Zeilen,
                    ' der Teil mit dem Pfandbetrag steht dabei VOR der Artikelnummer:
                    '   "                    36    1459-    16-  3,10    49,60-"
                    '   "99001 Euro 0,5   1443-  20"
                    ' Beide Teile werden spaltengenau übereinandergelegt.
                    If Len(lgOffen) > 0 Then
                        z = Ueberlagern(lgOffen, l)
                        If Len(z) > 0 Then
                            Set m = Treffer(reLG, z)
                            If Not m Is Nothing Then lgOffen = "": PfandAnhaengen m: GoTo weiter
                        End If
                    End If
                    ' Artikelzeile ohne Pfandbetrag (Kontostand) oder Teilzeile: für die nächste Zeile merken
                    If Not Treffer(reArtNr, s) Is Nothing Or Not Treffer(reLGRest, s) Is Nothing Then
                        LeergutAbschliessen lgOffen, reArtNr
                        lgOffen = l
                        GoTo weiter
                    End If
                End If

                ' Alles andere: Zusatzangaben auswerten, sonst als Hinweis protokollieren
                Set m = Treffer(reMarkt, l)
                If Not m Is Nothing Then Feld("f_Marktnr").Value = m.SubMatches(0): GoTo weiter
                Set m = Treffer(reFiliale, l)
                If Not m Is Nothing Then
                    If Len(Feld("f_LName").Value) = 0 Then Feld("f_LName").Value = Feld("f_KName").Value
                    Feld("f_LZeile2").Value = Trim$(m.SubMatches(2))
                    Feld("f_LPLZ").Value = "'" & m.SubMatches(0)
                    Feld("f_LOrt").Value = Trim$(m.SubMatches(1))
                    Protokoll "Info", "Lieferfiliale als Lieferort übernommen: " & m.SubMatches(0) & " " & _
                                      Trim$(m.SubMatches(1)) & ", " & Trim$(m.SubMatches(2))
                    GoTo weiter
                End If
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

    LeergutAbschliessen lgOffen, reArtNr
    If anzahlLS > 1 Then Protokoll "Warnung", "Mehrere Lieferscheine: " & lsListe & ". In der E-Rechnung steht nur der erste."
    Feld("f_PDF_MwSt").Value = mwstSumme
    SteuernAbschliessen igLieferung, steuerSaetze, mwstSumme, bruttoBetraege
End Sub

' MwSt-Sätze bzw. Steuerbefreiung auf die Positionen anwenden (beide Layouts)
Private Sub SteuernAbschliessen(ByVal igLieferung As Boolean, steuerSaetze As Collection, _
                                ByVal mwstSumme As Currency, ByVal bruttoBetraege As Boolean)
    If igLieferung And InStr(EU_LAENDER, "|" & UCase$(Trim$(CStr(Feld("f_KLand").Value))) & "|") = 0 Then
        Protokoll "Fehler", "Steuerfreie Lieferung, aber Land des Empfängers '" & Feld("f_KLand").Value & _
                         "' ist kein anderes EU-Land. Ausfuhrlieferungen in Drittländer werden noch nicht unterstützt; bitte 'USt-Satz' und 'USt-Kat.' im Blatt Positionen prüfen."
    ElseIf igLieferung Then
        If steuerSaetze.Count > 0 Then Protokoll "Warnung", "Innergemeinschaftliche Lieferung, aber MwSt auf der Rechnung. Bitte prüfen."
        SteuerfreiSetzen "K"
        LieferortVomEmpfaenger
    ElseIf steuerSaetze.Count = 0 Then
        Protokoll "Fehler", "Kein MwSt-Satz und kein Hinweis auf eine Steuerbefreiung gefunden. Bitte Spalten 'USt-Satz' und 'USt-Kat.' im Blatt Positionen ausfüllen."
    Else
        SteuersatzSetzen steuerSaetze
        If bruttoBetraege Then BruttoInNetto steuerSaetze, mwstSumme
    End If
End Sub

' Rechnung mit Bruttobeträgen ("ENTHALTENE MWSt."): Die E-Rechnung verlangt
' Nettopreise. Preise werden auf 4 Stellen netto umgerechnet, Beträge = Menge x
' Nettopreis. Damit Nettosumme, MwSt und Zahlbetrag genau der Rechnung
' entsprechen (Vorgabe Frank 25.11.2026), wird die Rundungsdifferenz zur
' Nettosumme laut Rechnung (Rechnungsbetrag - enthaltene MwSt) centweise auf die
' Positionen mit dem größten Rundungsrest verteilt; deren Nettopreis wird passend
' neu berechnet. Nur was so nicht aufgeht (selten), wird Rundungsbetrag (BT-114).
' Mit tests/parser_test.py (brutto_in_netto) abgleichen.
Private Sub BruttoInNetto(saetze As Collection, ByVal mwstPdf As Currency)
    Dim ws As Worksheet, z As Long, letzte As Long, faktor As Double, i As Long, beste As Long
    Dim preis As Currency, betrag As Currency, summe As Currency, mwst As Currency, rundung As Currency
    Dim ziel As Currency, schritt As Currency, n As Long, verteilt As Long, menge As Currency
    Dim rest() As Double, gesperrt() As Boolean

    If saetze.Count <> 1 Then
        Protokoll "Fehler", "Bruttobeträge mit mehreren MwSt-Sätzen werden nicht unterstützt. Bitte Preise und Beträge im Blatt Positionen netto eintragen."
        Exit Sub
    End If
    faktor = 1 + saetze(1) / 100
    Set ws = ThisWorkbook.Worksheets(BLATT_POSITIONEN)
    letzte = ws.Cells(ws.Rows.Count, SP_TYP).End(xlUp).Row
    If letzte < 2 Then Exit Sub
    ReDim rest(2 To letzte)
    ReDim gesperrt(2 To letzte)
    For z = 2 To letzte
        preis = Runden4(CDbl(ws.Cells(z, SP_PREIS).Value) / faktor)
        If ws.Cells(z, SP_TYP).Value = "Nachlass" Then
            betrag = RundenKfm(CDbl(ws.Cells(z, SP_BETRAG).Value) / faktor)
            gesperrt(z) = True
        Else
            betrag = RundenKfm(CCur(ws.Cells(z, SP_MENGE).Value) * preis)
        End If
        rest(z) = CDbl(ws.Cells(z, SP_BETRAG).Value) / faktor - CDbl(betrag)   ' genauer Nettobetrag - gerundeter
        ' Nettopreis mit 4 Stellen: als Double schreiben und 4 Stellen anzeigen. Ein
        ' Currency-Wert bzw. das Format mit 2 Stellen führte in Excel zu 11,76 statt
        ' 11,7647 und damit zu Fehlern "Menge x Preis" (Test 1.11.1).
        ws.Cells(z, SP_PREIS).NumberFormat = "#,##0.0000"
        ws.Cells(z, SP_PREIS).Value = CDbl(preis)
        ws.Cells(z, SP_BETRAG).Value = betrag
        summe = summe + betrag
    Next z

    If Len(CStr(Feld("f_PDF_Brutto").Value)) > 0 Then
        ziel = CCur(Feld("f_PDF_Brutto").Value) - mwstPdf
        Feld("f_PDF_Netto").Value = ziel
        ' Differenz centweise verteilen: jeweils auf die Position, deren Rundungsrest
        ' am weitesten in die nötige Richtung zeigt und deren Preis sich passend setzen lässt
        schritt = IIf(ziel > summe, 0.01, -0.01)
        n = CLng(Abs(ziel - summe) * 100)
        ' Nur echte Rundungsdifferenzen verteilen (höchstens 1 Cent je Position). Eine
        ' größere Differenz ist ein Lesefehler (z. B. fehlende Position) und darf nicht
        ' über die Preise kaschiert werden: dann bleibt sie stehen und die Prüfung meldet sie.
        If n > letzte - 1 Then
            Protokoll "Warnung", "Differenz zur Nettosumme der Rechnung " & Format$(ziel - summe, "0.00") & _
                                 " EUR ist keine Rundungsdifferenz und wird nicht verteilt. Fehlt eine Position?"
            n = 0
        End If
        Do While verteilt < n
            beste = 0
            For i = 2 To letzte
                If Not gesperrt(i) Then
                    If beste = 0 Then
                        beste = i
                    ElseIf rest(i) * Sgn(schritt) > rest(beste) * Sgn(schritt) Then
                        beste = i
                    End If
                End If
            Next i
            If beste = 0 Then Exit Do
            gesperrt(beste) = True
            menge = CCur(ws.Cells(beste, SP_MENGE).Value)
            If menge <> 0 Then
                betrag = CCur(ws.Cells(beste, SP_BETRAG).Value) + schritt
                preis = Runden4(CDbl(betrag) / CDbl(menge))
                If RundenKfm(menge * preis) = betrag Then
                    ws.Cells(beste, SP_PREIS).Value = CDbl(preis)
                    ws.Cells(beste, SP_BETRAG).Value = betrag
                    summe = summe + schritt
                    verteilt = verteilt + 1
                End If
            End If
        Loop
        mwst = RundenKfm(summe * saetze(1) / 100)
        rundung = CCur(Feld("f_PDF_Brutto").Value) - (summe + mwst)
    End If
    Feld("f_Rundung").Value = rundung
    Protokoll "Info", "Rechnung mit Bruttobeträgen (enthaltene MwSt " & saetze(1) & " %): Preise, Pfand und Nachlässe in netto umgerechnet" & _
                      IIf(verteilt > 0, ", Rundungsdifferenz " & Format$(verteilt * schritt, "0.00") & " EUR auf " & verteilt & " Position(en) verteilt", "") & "."
    If rundung <> 0 Then
        Protokoll "Info", "Restliche Rundungsdifferenz " & Format$(rundung, "0.00") & " EUR als Rundungsbetrag (BT-114)."
    End If
    If Abs(rundung) > 0.01 * (letzte - 1) Then
        Protokoll "Warnung", "Rundungsdifferenz " & Format$(rundung, "0.00") & " EUR ist größer als erwartet. Bitte Positionen prüfen."
    End If
End Sub

' Kaufmännisch auf 4 Stellen runden (Nettopreis)
Private Function Runden4(ByVal x As Double) As Currency
    Runden4 = Sgn(x) * Int(Abs(x) * 10000 + 0.5) / 10000
End Function

' Innergemeinschaftliche Lieferung (BR-IC-12): Lieferland ist Pflicht. Ohne
' eigenen Lieferort ("Wir lieferten an") gilt die Anschrift des Empfängers.
Private Sub LieferortVomEmpfaenger()
    If Len(Feld("f_LName").Value) > 0 Or Len(Feld("f_LOrt").Value) > 0 Then Exit Sub
    Feld("f_LName").Value = Feld("f_KName").Value
    Feld("f_LZeile2").Value = Feld("f_KStrasse").Value
    Feld("f_LPLZ").Value = "'" & Feld("f_KPLZ").Value
    Feld("f_LOrt").Value = Feld("f_KOrt").Value
    Feld("f_LLand").Value = Feld("f_KLand").Value
    Protokoll "Info", "Kein Lieferort auf der Rechnung: Anschrift des Empfängers als Lieferort übernommen (Lieferland " & Feld("f_KLand").Value & ")."
End Sub

' Nachlass auf den Warenwert (z. B. Skonto, bereits abgezogen) als Position vom
' Typ "Nachlass". Der Adapter macht daraus einen Abschlag auf Belegebene (BG-20)
' mit USt-Satz der Rechnung; das Pfand bleibt unberührt.
' Gruppen: 0 Art, 1 Text, 2 Basis, 3 Prozent, 4 Betrag
Private Sub NachlassAnhaengen(ByVal m As Object)
    Dim basis As Currency, prozent As Currency, betrag As Currency, bez As String

    basis = ZahlDE(m.SubMatches(2))
    prozent = ZahlDE(m.SubMatches(3))
    betrag = ZahlDE(m.SubMatches(4))
    bez = Trim$(m.SubMatches(0)) & " " & m.SubMatches(3) & " % auf " & m.SubMatches(2) & " EUR"

    PositionAnhaengen "Nachlass", "", bez, 1, "", betrag, -betrag, Trim$(m.SubMatches(1) & "")
    Protokoll "Info", bez & " als Nachlass auf Belegebene übernommen: " & m.SubMatches(4) & "-"
    If RundenKfm(basis * prozent / 100) <> betrag Then
        Protokoll "Warnung", bez & ": berechnet " & Format$(RundenKfm(basis * prozent / 100), "#,##0.00") & _
                             ", laut Rechnung " & m.SubMatches(4) & "."
    End If
    If Len(Feld("f_PDF_Warenwert").Value) > 0 Then
        If basis <> CCur(Feld("f_PDF_Warenwert").Value) Then
            Protokoll "Warnung", bez & ": Basis weicht vom Warenwert " & Format$(Feld("f_PDF_Warenwert").Value, "#,##0.00") & " ab."
        End If
    End If
End Sub

' Eine offene Teilzeile ohne Artikelnummer, die keinem Artikel zugeordnet werden
' konnte, wäre ein fehlender Pfandbetrag - deshalb als Warnung protokollieren.
' Eine vollständige Artikelzeile ohne Pfandbetrag (Differenz 0) wird hier
' ausgewertet: Lieferung und Rückgabe gleich groß, z. B. 5 geliefert, 5 zurück.
Private Sub LeergutAbschliessen(ByRef lgOffen As String, ByVal reArtNr As Object)
    Dim m As Object
    If Len(lgOffen) > 0 Then
        If Treffer(reArtNr, Trim$(lgOffen)) Is Nothing Then
            Protokoll "Warnung", "Leergutzeile keinem Artikel zugeordnet: " & Trim$(lgOffen)
        Else
            Set m = Treffer(NeuesRegex("^\s*(\d{5})\s+(\S+(?: \S+)*)\s{2,}(.*\S)\s+(\d+,\d{2})\s*$"), lgOffen)
            If Not m Is Nothing Then PfandOhneDifferenz m
        End If
    End If
    lgOffen = ""
End Sub

' Leergutzeile ohne Pfandbetrag: Gruppen 0 ArtNr, 1 Bezeichnung, 2 Mengen
' (ALTER SALDO, LIEFERUNG, ZURÜCK, NEUER SALDO; leere fehlen), 3 Pfandsatz.
' Sind Lieferung und Rückgabe gleich (Differenz 0), entstehen trotzdem zwei
' Positionen: +Menge und -Menge zum Pfandsatz. Saldiert wäre das 0,00 EUR, was
' z. B. EDEKA und REWE nicht annehmen. Zeilen ohne Bewegung (nur Saldo) entfallen.
Private Sub PfandOhneDifferenz(ByVal m As Object)
    Dim artNr As String, bez As String, einheit As String
    Dim satz As Currency, lief As Double, zur As Double

    If UCase$(EinstellungOderLeer("Pfand aufteilen")) = "NEIN" Then Exit Sub
    If Not LeergutAufteilen(CStr(m.SubMatches(2)), 0, lief, zur) Then
        Protokoll "Warnung", "Leergutzeile ohne Pfandbetrag nicht auswertbar: " & Trim$(m.Value)
        Exit Sub
    End If
    If lief = 0 And zur = 0 Then Exit Sub         ' keine Bewegung, nur Kontostand

    artNr = m.SubMatches(0)
    bez = "Pfand " & Trim$(m.SubMatches(1))
    einheit = Einstellung("Einheit Pfand")
    satz = ZahlDE(m.SubMatches(3))
    If lief <> 0 Then PositionAnhaengen "Pfand", artNr, bez & " Lieferung", lief, einheit, satz, RundenKfm(CCur(lief) * satz), ""
    If zur <> 0 Then PositionAnhaengen "Pfand", artNr, bez & " Rücknahme", -zur, einheit, satz, RundenKfm(CCur(-zur) * satz), ""
    Protokoll "Info", "Leergut mit Differenz 0 als Lieferung und Rücknahme übernommen: " & bez & " " & lief & " / " & zur
End Sub

' Legt zwei Textzeilen spaltengenau übereinander. Leer, wenn sich Zeichen überschneiden.
Private Function Ueberlagern(ByVal a As String, ByVal b As String) As String
    Dim n As Long, i As Long, ca As String, cb As String, erg As String
    n = Len(a)
    If Len(b) > n Then n = Len(b)
    a = a & Space$(n - Len(a))
    b = b & Space$(n - Len(b))
    For i = 1 To n
        ca = Mid$(a, i, 1)
        cb = Mid$(b, i, 1)
        If ca <> " " And cb <> " " Then Exit Function
        If ca <> " " Then erg = erg & ca Else erg = erg & cb
    Next i
    Ueberlagern = erg
End Function

' Leergutzeile übernehmen. Gruppen: 0 ArtNr, 1 Bezeichnung, 2 Mengen vor MENGE DIFF.,
' 3 MENGE DIFF., 4 Pfandsatz, 5 Pfandbetrag.
' Standard (Einstellung "Pfand aufteilen" = Ja): Lieferung und Rücknahme als zwei
' Positionen mit positiver bzw. negativer Menge, wie es z. B. REWE verlangt
' (keine Saldierung). Lässt sich die Zeile nicht eindeutig aufteilen, bleibt es
' bei einer Position mit der Differenz und einer Warnung im Protokoll.
Private Sub PfandAnhaengen(ByVal m As Object)
    Dim artNr As String, bez As String, einheit As String
    Dim diff As Double, satz As Currency, betrag As Currency, lief As Double, zur As Double

    artNr = m.SubMatches(0)
    bez = "Pfand " & Trim$(m.SubMatches(1))
    einheit = Einstellung("Einheit Pfand")
    diff = CDbl(ZahlDE(m.SubMatches(3)))
    satz = ZahlDE(m.SubMatches(4))
    betrag = ZahlDE(m.SubMatches(5))

    If UCase$(EinstellungOderLeer("Pfand aufteilen")) <> "NEIN" Then
        If LeergutAufteilen(CStr(m.SubMatches(2)), diff, lief, zur) Then
            If lief <> 0 Then PositionAnhaengen "Pfand", artNr, bez & " Lieferung", lief, einheit, satz, RundenKfm(CCur(lief) * satz), ""
            If zur <> 0 Then PositionAnhaengen "Pfand", artNr, bez & " Rücknahme", -zur, einheit, satz, RundenKfm(CCur(-zur) * satz), ""
            Exit Sub
        End If
        Protokoll "Warnung", "Pfand nicht in Lieferung und Rücknahme aufteilbar, Differenz übernommen: " & Trim$(m.Value)
    End If
    PositionAnhaengen "Pfand", artNr, bez, diff, einheit, satz, betrag, ""
End Sub

' Lieferung und Rücknahme aus den Mengen vor MENGE DIFF. bestimmen. "mitte" enthält
' eine Auswahl aus ALTER SALDO, LIEFERUNG, ZURÜCK (leere Spalten fehlen im Text)
' und als letzte Zahl NEUER SALDO. Es gilt LIEFERUNG - ZURÜCK = DIFF und
' ALT + DIFF = NEU. Gesucht ist die eindeutige Zuordnung, die beides erfüllt.
' Mit tests/parser_test.py (leergut_aufteilen) abgleichen.
Private Function LeergutAufteilen(ByVal mitte As String, ByVal diff As Double, _
                                  ByRef lief As Double, ByRef zur As Double) As Boolean
    Dim t() As String, w() As Double, n As Long, i As Long, k As Long, maske As Long, bits As Long
    Dim v(0 To 2) As Double, neu As Double, anzahl As Long, tL As Double, tZ As Double

    mitte = Trim$(mitte)
    If Len(mitte) = 0 Then Exit Function
    Do While InStr(mitte, "  ") > 0: mitte = Replace(mitte, "  ", " "): Loop
    t = Split(mitte, " ")
    n = UBound(t)                       ' Anzahl Werte vor NEUER SALDO
    If n > 3 Then Exit Function
    ReDim w(0 To n)
    For i = 0 To n
        If Not t(i) Like "*#*" Then Exit Function
        w(i) = CDbl(ZahlDE(t(i)))
    Next i
    neu = w(n)

    For maske = 0 To 7                  ' Bit 0 = ALT, 1 = LIEFERUNG, 2 = ZURÜCK
        bits = (maske And 1) + ((maske And 2) \ 2) + ((maske And 4) \ 4)
        If bits = n Then
            i = 0
            For k = 0 To 2
                If maske And (2 ^ k) Then v(k) = w(i): i = i + 1 Else v(k) = 0
            Next k
            If v(1) - v(2) = diff And v(0) + diff = neu Then
                If anzahl = 0 Or v(1) <> tL Or v(2) <> tZ Then anzahl = anzahl + 1
                tL = v(1): tZ = v(2)
            End If
        End If
    Next maske

    If anzahl = 1 Then lief = tL: zur = tZ: LeergutAufteilen = True
End Function

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
        Set m = Treffer(NeuesRegex("^(?:([A-Z]{1,3})\s*-\s*(\d{4,5})|(\d{5}))\s+(.+)$"), z(z.Count))
        If Not m Is Nothing Then
            Feld("f_LPLZ").Value = "'" & m.SubMatches(1) & m.SubMatches(2)
            Feld("f_LOrt").Value = Trim$(m.SubMatches(3))
            If Len(m.SubMatches(0) & "") > 0 Then Feld("f_LLand").Value = LandAusKennzeichen(m.SubMatches(0))
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
    Feld("f_KPLZ").Value = "'" & m.SubMatches(1) & m.SubMatches(2)
    Feld("f_KOrt").Value = Trim$(Replace(m.SubMatches(3), "(Deutschland)", ""))   ' "22305 Hamburg (Deutschland)"
    If Len(m.SubMatches(0) & "") > 0 Then Feld("f_KLand").Value = LandAusKennzeichen(m.SubMatches(0))
End Sub

' Altes Länderkennzeichen vor der PLZ ("I-", "A-", "CH-") -> ISO-Code
Private Function LandAusKennzeichen(ByVal k As String) As String
    Select Case UCase$(Trim$(k))
        Case "D": LandAusKennzeichen = "DE"
        Case "A": LandAusKennzeichen = "AT"
        Case "I": LandAusKennzeichen = "IT"
        Case "F": LandAusKennzeichen = "FR"
        Case "B": LandAusKennzeichen = "BE"
        Case "L": LandAusKennzeichen = "LU"
        Case "S": LandAusKennzeichen = "SE"
        Case "E": LandAusKennzeichen = "ES"
        Case "P": LandAusKennzeichen = "PT"
        Case "H": LandAusKennzeichen = "HU"
        Case "FL": LandAusKennzeichen = "LI"
        Case "SLO": LandAusKennzeichen = "SI"
        Case Else: LandAusKennzeichen = UCase$(Trim$(k))     ' CH, NL, DK, CZ, PL ...
    End Select
    Protokoll "Info", "Land des Empfängers aus der Anschrift: " & LandAusKennzeichen
End Function

'==============================================================================
' Tabellen-Layout der BSI-Rechnung (RE 2604269): Tabellenkopf
' "Pos. ArtNr. Bezeichnung Ka/Fa Hl./Fl. Preis Betrag zzgl. MwSt.", Beträge mit €,
' Bezeichnungen über mehrere Zeilen, Summenblock "Summe / Warenwert", "Pfand",
' "Rundungsdifferenz MwSt.", "Summe 19% MwSt.", "RECHNUNGSBETRAG".
' Mit pdftotext -layout verschiebt Xpdf den Summenblock (Beträge neben falschen
' Bezeichnungen). Deshalb liest PDF2E dieses Layout mit pdftotext -raw: Jede
' Angabe steht dort in einer eigenen Zeile in der Reihenfolge der PDF.
' Mit tests/parser_test.py (parse_tabelle) abgleichen.
'==============================================================================

' Erkennung in der -layout-Ausgabe am Tabellenkopf (das alte Layout hat "ARTNR BEZEICHNUNG")
Public Function IstBSITabelle(ByVal text As String) As Boolean
    IstBSITabelle = Not Treffer(NeuesRegex("(^|\n)\s*Pos\.\s+ArtNr\.\s+Bezeichnung"), text) Is Nothing
End Function

Public Sub ParseBSITabelle(ByVal text As String)
    Dim roh() As String, z() As String, n As Long, i As Long, kopf As Long
    Dim l As String, m As Object, eur As String, betragT As String
    Dim reKopf As Object, reKopfzeile As Object, rePosStart As Object, reWare As Object, reLeergut As Object
    Dim reWW As Object, rePfand As Object, reSteuer As Object, reRundung As Object, reRB As Object
    Dim reGLN As Object, reZiffer As Object, reNurZahl As Object, rePLZ As Object, reTelefon As Object
    Dim reILN As Object, reMail As Object, reIG As Object, reUStId As Object
    Dim links As New Collection, steuerSaetze As New Collection
    Dim posErwartet As Long, imSummenteil As Boolean, sammler As Long, puf As String, posNr As String, artNr As String
    Dim anzahlLS As Long, lsListe As String, rueckgaben As String
    Dim mwstSumme As Currency, pfand As Currency, warenwert As Currency, hatWW As Boolean, hatPfand As Boolean
    Dim igLieferung As Boolean, s As String

    eur = ChrW(&H20AC)
    betragT = "(-?[\d.]*\d,\d{2}-?)\s*" & eur
    Set reKopf = NeuesRegex("(\d+)\s+(\d{5,})\s+(\d{1,2}\.\d{1,2}\.\d{4})\s*$")
    Set reKopfzeile = NeuesRegex("^Pos\.\s+ArtNr\.\s+Bezeichnung")
    ' Beginn einer Position: "1 1", "4 100 Palette(n) 1 0"
    Set rePosStart = NeuesRegex("^(\d+)\s+(\d+)(?:\s+(.*))?$")
    ' Ware: Bezeichnung, Ka/Fa, Hl./Fl. (bis zu zwei Werte), Preis €, Betrag €, MwSt €
    Set reWare = NeuesRegex("^(.+?)\s+(-?\d+-?)\s+(?:([\d.,]+)\s+)?(?:([\d.,]+)\s+)?(-?[\d.,]+-?)\s*" & eur & _
                            "\s+" & betragT & "\s+" & betragT & "$")
    Set reLeergut = NeuesRegex("^Leergutr" & ChrW(252) & "ckgabe:\s*(.*?)\s+(-?\d+-?)(?:\s+(-?\d+-?))?$")
    Set reWW = NeuesRegex("^Summen?\s*/\s*Warenwert\b.*?" & betragT & "\s+" & betragT & "$")
    Set rePfand = NeuesRegex("^Pfand\s+" & betragT & "(?:\s+" & betragT & ")?$")
    Set reSteuer = NeuesRegex("^Summe\s+(\d+(?:,\d+)?)\s*%\s*MwSt\.?\s+" & betragT & "$")
    Set reRundung = NeuesRegex("^Rundungsdifferenz\s+MwSt\.?\s+" & betragT & "$")
    Set reRB = NeuesRegex("^RECHNUNGSBETRAG\s+" & betragT & "$")
    ' GLN des Rechnungsempfängers in eigener Zeile: "GLN: 43 88 096 00000 2"
    Set reGLN = NeuesRegex("^(?:RE-)?GLN:?\s*(\d[\d ]{11,}\d)$")
    Set reZiffer = NeuesRegex("^\d")
    Set reNurZahl = NeuesRegex("^\d+$")
    Set rePLZ = NeuesRegex("^(?:([A-Z]{1,3})\s*-\s*(\d{4,5})|(\d{5}))\s+(.+)$")
    Set reTelefon = NeuesRegex("^(?:Tel\.?:?\s*)?\+?\d[\d \-/()]{5,}\d$")
    Set reILN = NeuesRegex("(?:ILN|GLN):\s*(\d{13})")
    Set reMail = NeuesRegex("([\w.+-]+@[\w-]+(?:\.[\w-]+)+)")
    Set reIG = NeuesRegex("[Ss]teuerfreie\s+(?:innergemeinschaftliche\s+)?Lieferung")
    Set reUStId = NeuesRegex("^\s*U[Ss][Tt]\s*-?\s*ID(?:\s*-?\s*Nr\.?)?\s*:?\s+([A-Za-z]{2}[\w .-]*\w)\s*$")

    ' Zeilen ohne Leerzeilen und ohne Wasserzeichen (einzelne Zeichen wie "T", "*"; Ziffern bleiben)
    text = Replace(Replace(text, vbCr, ""), Chr$(12), vbLf)
    roh = Split(text, vbLf)
    ReDim z(0 To UBound(roh) + 1)
    n = 0
    For i = 0 To UBound(roh)
        l = Trim$(roh(i))
        If Len(l) > 1 Or l Like "#" Then z(n) = l: n = n + 1
    Next i
    z(n) = ""                                         ' Ende-Markierung für den Blick auf die Folgezeile
    kopf = -1
    For i = 0 To n - 1
        If Not Treffer(reKopfzeile, z(i)) Is Nothing Then kopf = i: Exit For
    Next i
    If kopf < 0 Then
        Protokoll "Fehler", "Tabellenkopf 'Pos. ArtNr. Bezeichnung' nicht gefunden."
        Exit Sub
    End If

    ' ---- Kopf und Rechnungsempfänger (vor dem Tabellenkopf) ----
    For i = 0 To kopf - 1
        l = z(i)
        Set m = Treffer(reKopf, l)
        If Not m Is Nothing Then
            If Len(Feld("f_RechNr").Value) = 0 Then
                Feld("f_KundenNr").Value = m.SubMatches(0)
                Feld("f_RechNr").Value = m.SubMatches(1)
                Feld("f_RechDatum").Value = DatumDE(m.SubMatches(2))
            End If
        ElseIf Replace(l, " ", "") = "RECHNUNG" Or InStr(l, "Kunden-Nr.") > 0 Then
            ' Überschriften
        Else
            Set m = Treffer(reGLN, l)
            If Not m Is Nothing Then
                ' "REWE MARKT GmbH Zweig NL. Süd RE-" + "GLN: 43 88 096 00000 2": RE- gehört zur GLN
                Feld("f_KGLN").Value = "'" & Replace(m.SubMatches(0), " ", "")
                If links.Count > 0 Then
                    s = links(links.Count)
                    If Right$(s, 4) = " RE-" Then
                        links.Remove links.Count
                        links.Add RTrim$(Left$(s, Len(s) - 4))
                    End If
                End If
                Protokoll "Info", "GLN des Rechnungsempfängers aus der Rechnung übernommen: " & Replace(m.SubMatches(0), " ", "")
            ElseIf Not Treffer(reTelefon, l) Is Nothing Then
                Protokoll "Info", "Telefonnummer im Anschriftenblock nicht als Name übernommen: " & l
            ElseIf LCase$(Left$(l, 4)) <> "per " Then
                links.Add l
            End If
        End If
    Next i
    EmpfaengerSetzen links, rePLZ

    ' ---- Tabelle und Summenblock ----
    ' Der Tabellenkopf endet mit "zzgl. MwSt."
    i = kopf
    Do While i < n
        If InStr(z(i), "MwSt") > 0 Then Exit Do
        i = i + 1
    Loop
    i = i + 1
    posErwartet = 1
    ' sammler: 0 nichts, 1 Lieferschein, 2 Lieferort ("Wir lieferten an:"), 3 Position
    Do While i < n
        l = z(i)
        If Not imSummenteil Then
            Set m = Treffer(rePosStart, l)
            If Not m Is Nothing Then
                If CLng(m.SubMatches(0)) <> posErwartet Then
                    Set m = Nothing
                ElseIf Len(m.SubMatches(2) & "") = 0 And Not Treffer(reZiffer, z(i + 1)) Is Nothing Then
                    Set m = Nothing                   ' Mengenzeile ("25 500"), keine neue Position
                End If
            End If
            If Not m Is Nothing Then
                TabelleAbschliessen sammler, puf, posNr, artNr, reWare, reLeergut, rueckgaben, anzahlLS, lsListe
                sammler = 3
                posNr = m.SubMatches(0)
                artNr = m.SubMatches(1)
                puf = m.SubMatches(2) & ""
                posErwartet = posErwartet + 1
                GoTo weiter
            End If
            If Not Treffer(reNurZahl, l) Is Nothing And Left$(z(i + 1), 12) = "geliefert am" Then
                TabelleAbschliessen sammler, puf, posNr, artNr, reWare, reLeergut, rueckgaben, anzahlLS, lsListe
                sammler = 1
                puf = l
                GoTo weiter
            End If
            If Left$(l, 17) = "Wir lieferten an:" Then
                TabelleAbschliessen sammler, puf, posNr, artNr, reWare, reLeergut, rueckgaben, anzahlLS, lsListe
                sammler = 2
                puf = Trim$(Mid$(l, 18))
                GoTo weiter
            End If
            If Not Treffer(reWW, l) Is Nothing Then
                TabelleAbschliessen sammler, puf, posNr, artNr, reWare, reLeergut, rueckgaben, anzahlLS, lsListe
                imSummenteil = True
            ElseIf sammler <> 0 Then
                puf = Trim$(puf & " " & l)
                GoTo weiter
            End If
        End If

        ' Summenblock und Zusatzangaben
        Set m = Treffer(reWW, l)
        If Not m Is Nothing Then
            ' steht in der Tabelle und im Summenblock
            If hatWW And ZahlDE(m.SubMatches(0)) <> warenwert Then
                Protokoll "Warnung", "Warenwert in Tabelle und Summenblock verschieden: " & l
            End If
            warenwert = ZahlDE(m.SubMatches(0))
            hatWW = True
            GoTo weiter
        End If
        Set m = Treffer(rePfand, l)
        If Not m Is Nothing Then pfand = ZahlDE(m.SubMatches(0)): hatPfand = True: GoTo weiter
        Set m = Treffer(reSteuer, l)
        If Not m Is Nothing Then
            steuerSaetze.Add CDbl(ZahlDE(m.SubMatches(0)))
            mwstSumme = mwstSumme + ZahlDE(m.SubMatches(1))
            GoTo weiter
        End If
        Set m = Treffer(reRundung, l)
        If Not m Is Nothing Then
            Protokoll "Info", "Rundungsdifferenz MwSt. laut Rechnung " & m.SubMatches(0) & " (in der MwSt-Summe enthalten)."
            GoTo weiter
        End If
        Set m = Treffer(reRB, l)
        If Not m Is Nothing Then Feld("f_PDF_Brutto").Value = ZahlDE(m.SubMatches(0)): GoTo weiter
        If Left$(l, 11) = "Betrag zzgl" Then GoTo weiter
        If Not Treffer(reIG, l) Is Nothing Then
            igLieferung = True
            Protokoll "Info", "Steuerfreie innergemeinschaftliche Lieferung: USt-Kategorie K, 0 %, Befreiungsgrund VATEX-EU-IC."
            GoTo weiter
        End If
        Set m = Treffer(reUStId, l)
        If Not m Is Nothing Then
            Feld("f_KUstId").Value = "'" & UStIdBereinigt(m.SubMatches(0))
            Protokoll "Info", "USt-IdNr. des Empfängers aus der Rechnung übernommen: " & UStIdBereinigt(m.SubMatches(0))
            GoTo weiter
        End If
        Set m = Treffer(reILN, l)
        If Not m Is Nothing Then
            If Len(Feld("f_LGLN").Value) = 0 Then
                Feld("f_LGLN").Value = "'" & m.SubMatches(0)
            ElseIf Feld("f_LGLN").Value <> m.SubMatches(0) Then
                Protokoll "Warnung", "Abweichende GLN im Text: " & l
            End If
            GoTo weiter
        End If
        Set m = Treffer(reMail, l)
        If Not m Is Nothing Then Feld("f_KEmail").Value = m.SubMatches(0): GoTo weiter
        Protokoll "Info", "Nicht übernommen: " & l
weiter:
        i = i + 1
    Loop
    TabelleAbschliessen sammler, puf, posNr, artNr, reWare, reLeergut, rueckgaben, anzahlLS, lsListe

    If anzahlLS = 0 Then Protokoll "Warnung", "Lieferschein (geliefert am … Nr:) nicht erkannt."
    If anzahlLS > 1 Then Protokoll "Warnung", "Mehrere Lieferscheine: " & lsListe & ". In der E-Rechnung steht nur der erste."

    ' Pfand: Dieses Layout hat kein Leergut-Konto mit Pfandsätzen, nur den Saldo.
    ' Er wird als eine Pfand-Position übernommen (Preis positiv, Menge -1 bei Gutschrift).
    If hatPfand And pfand <> 0 Then
        PositionAnhaengen "Pfand", "", "Pfand (Saldo laut Rechnung)", Sgn(pfand), Einstellung("Einheit Pfand"), Abs(pfand), pfand, ""
    End If
    If hatPfand Or Len(rueckgaben) > 0 Then
        Protokoll "Info", "Pfand als Saldo laut Rechnung übernommen (" & Format$(pfand, "#,##0.00") & _
                          " EUR), das Layout enthält kein Leergut-Konto mit Lieferung und Rücknahme je Pfandartikel." & _
                          IIf(Len(rueckgaben) > 0, " Leergutrückgabe laut Rechnung: " & rueckgaben & ".", "")
    End If

    If hatWW Then Feld("f_PDF_Warenwert").Value = warenwert
    Feld("f_PDF_Pfand").Value = pfand
    If hatWW Then Feld("f_PDF_Netto").Value = warenwert + pfand
    Feld("f_PDF_MwSt").Value = mwstSumme
    SteuernAbschliessen igLieferung, steuerSaetze, mwstSumme, False
End Sub

' Gesammelten Block des Tabellen-Layouts auswerten (siehe ParseBSITabelle)
Private Sub TabelleAbschliessen(ByRef sammler As Long, ByRef puf As String, ByVal posNr As String, ByVal artNr As String, _
                                ByVal reWare As Object, ByVal reLeergut As Object, ByRef rueckgaben As String, _
                                ByRef anzahlLS As Long, ByRef lsListe As String)
    Dim m As Object, menge As Double, betrag As Currency
    Select Case sammler
        Case 1: TabelleLieferschein puf, anzahlLS, lsListe
        Case 2: TabelleLieferort puf
        Case 3
            Set m = Treffer(reLeergut, puf)
            If Not m Is Nothing Then
                ' Rückgabe ohne Betrag; das Pfand steht nur als Saldo im Summenblock
                rueckgaben = rueckgaben & IIf(Len(rueckgaben) > 0, ", ", "") & m.SubMatches(1) & " x " & Trim$(m.SubMatches(0)) & " (" & artNr & ")"
            Else
                Set m = Treffer(reWare, puf)
                If m Is Nothing Then
                    Protokoll "Info", "Nicht übernommen: Pos. " & posNr & " " & artNr & " " & puf
                Else
                    menge = CDbl(ZahlDE(m.SubMatches(1)))
                    betrag = ZahlDE(m.SubMatches(5))
                    If betrag < 0 And menge > 0 Then menge = -menge     ' Rückware
                    PositionAnhaengen "Ware", artNr, Trim$(m.SubMatches(0)), menge, _
                        Einstellung("Einheit Ware"), ZahlDE(m.SubMatches(4)), betrag, m.SubMatches(2) & ""
                End If
            End If
    End Select
    sammler = 0
    puf = ""
End Sub

' "150425 geliefert am Nr: 25.11.2026 5392": Datum und Lieferscheinnummer (letzte Zahl)
Private Sub TabelleLieferschein(ByVal block As String, ByRef anzahlLS As Long, ByRef lsListe As String)
    Dim m As Object, t() As String, k As Long, nr As String, zahlen As Long
    Set m = Treffer(NeuesRegex("\d{1,2}\.\d{1,2}\.\d{4}"), block)
    t = Split(block, " ")
    For k = 0 To UBound(t)
        If Len(t(k)) > 0 And Not t(k) Like "*[!0-9]*" Then nr = t(k): zahlen = zahlen + 1
    Next k
    If m Is Nothing Or zahlen < 2 Then
        Protokoll "Warnung", "Lieferschein nicht erkannt: " & block
        Exit Sub
    End If
    anzahlLS = anzahlLS + 1
    lsListe = lsListe & IIf(anzahlLS > 1, ", ", "") & nr & " vom " & m.Value
    If anzahlLS = 1 Then
        Feld("f_LSNr").Value = nr
        Feld("f_LSDatum").Value = DatumDE(m.Value)
    End If
End Sub

' "REWE Dominic Mayer oHG - Nr. 46653518 GLN. 4388096535184, Oberes Tor 7, 91320 Ebermannstadt"
' Gruppen: 0 Name, 1 Marktnummer, 2 GLN, 3 Straße, 4 Länderkennzeichen, 5 PLZ, 6 Ort
Private Sub TabelleLieferort(ByVal s As String)
    Dim m As Object
    Set m = Treffer(NeuesRegex("^(.+?)(?:\s+-\s+Nr\.?:?\s*(\d[\d ]*\d))?\s+GLN\.?:?\s*(\d{13})\s*,\s*(.+?)\s*,\s*" & _
                               "(?:([A-Z]{1,3})\s*-\s*)?(\d{4,5})\s+(.+)$"), s)
    If m Is Nothing Then
        ' ohne GLN: "Name, Straße, PLZ Ort"
        Set m = Treffer(NeuesRegex("^(.+?)\s*,\s*(.+?)\s*,\s*(?:([A-Z]{1,3})\s*-\s*)?(\d{4,5})\s+(.+)$"), s)
        If m Is Nothing Then
            Feld("f_LName").Value = s
            Protokoll "Warnung", "Lieferort nicht in Name, Straße, PLZ und Ort zerlegt: " & s
            Exit Sub
        End If
        Feld("f_LName").Value = m.SubMatches(0)
        Feld("f_LZeile2").Value = m.SubMatches(1)
        Feld("f_LPLZ").Value = "'" & m.SubMatches(3)
        Feld("f_LOrt").Value = Trim$(m.SubMatches(4))
        If Len(m.SubMatches(2) & "") > 0 Then Feld("f_LLand").Value = LandAusKennzeichen(m.SubMatches(2))
        Exit Sub
    End If
    Feld("f_LName").Value = m.SubMatches(0)
    If Len(m.SubMatches(1) & "") > 0 Then Feld("f_Marktnr").Value = "'" & Replace(m.SubMatches(1), " ", "")
    Feld("f_LGLN").Value = "'" & m.SubMatches(2)
    Feld("f_LZeile2").Value = m.SubMatches(3)
    Feld("f_LPLZ").Value = "'" & m.SubMatches(5)
    Feld("f_LOrt").Value = Trim$(m.SubMatches(6))
    If Len(m.SubMatches(4) & "") > 0 Then Feld("f_LLand").Value = LandAusKennzeichen(m.SubMatches(4))
End Sub
