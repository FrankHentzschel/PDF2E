Attribute VB_Name = "modEinlesen"
'==============================================================================
' Ablauf "Rechnung einlesen": Auswahl, Auslesen, Vorgaben, Prüfung.
' Auswahl und Auslesen der Quelle liefert das Produkt (modProdukt:
' QuelleAuswaehlen, QuelleEinlesen), der Rest ist für PDF2E und XL2E gleich.
'==============================================================================
Option Explicit

' Spalte "Vorgaben" im Blatt "Empfaenger" (z. B. REWE)
Public Const SP_E_VORGABEN As Long = 8
' Spalte "ZR-Nummer" im Blatt "Empfaenger" (G5-Zentralregulierung, BT-10)
Public Const SP_E_ZRNR As Long = 9
Private Const G5_EMAIL_STANDARD As String = "g5.e-invoice@documents.sgh-net.de"
Public Const BLATT_PFANDARTIKEL As String = "Pfandartikel"
' GLN je Lieferort (Warenempfänger), Schlüssel = Kunde / PLZ Ort, Straße
Public Const BLATT_LIEFERORTE As String = "Lieferorte"

' Eingabefelder, die vor jedem Einlesen geleert werden (Formelfelder fehlen bewusst)
Private Const FELDER_LEEREN As String = _
    "f_Quelldatei,f_RechNr,f_RechDatum,f_LSNr,f_LSDatum,f_KundenNr,f_KGLN,f_KName,f_KStrasse,f_KPLZ,f_KOrt," & _
    "f_KEmail,f_KUstId,f_LName,f_LZeile2,f_LPLZ,f_LOrt,f_LGLN,f_Marktnr,f_Kondition,f_Konto," & _
    "f_PDF_Warenwert,f_PDF_Pfand,f_PDF_Netto,f_PDF_MwSt,f_PDF_Brutto,f_Status,f_Ausgabedatei," & _
    "f_Mandat,f_GlaeubigerId,f_Einzug"

Public Sub RechnungEinlesen()
    Dim pfad As String
    pfad = QuelleAuswaehlen()
    If Len(pfad) = 0 Then Exit Sub

    Dim info As String
    On Error GoTo Fehler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    LastschriftFelderSicherstellen
    FeldnamenPruefen
    ArbeitsblaetterLeeren
    Feld("f_Quelldatei").Value = pfad
    Protokoll "Info", "Einlesen: " & pfad

    StammdatenLaden                     ' Mehrplatz: aktueller Stand aus dem Backend (modDaten)
    QuelleEinlesen pfad

    ' Wurde die Rechnung schon erzeugt (an diesem oder einem anderen Platz)?
    info = VerarbeitungInfo(CStr(Feld("f_RechNr").Value))
    If Len(info) > 0 Then
        Application.ScreenUpdating = True
        If MsgBox(info & vbLf & vbLf & "Trotzdem fortfahren?", vbYesNo + vbQuestion + vbDefaultButton2, TOOL_NAME) <> vbYes Then
            ArbeitsblaetterLeeren
            Protokoll "Info", "Abgebrochen: " & info
            Application.Calculation = xlCalculationAutomatic
            Exit Sub
        End If
        Protokoll "Warnung", info & " Erneut eingelesen."
        Application.ScreenUpdating = False
    End If

    Application.EnableEvents = False    ' Stammdaten schreibt VorgabenSetzen selbst ins Backend
    VorgabenSetzen
    Application.EnableEvents = True

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    RechnungPruefen
    ThisWorkbook.Worksheets(BLATT_RECHNUNG).Activate
    Exit Sub

Fehler:
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Protokoll "Fehler", Err.Description
    MsgBox "Einlesen fehlgeschlagen:" & vbLf & Err.Description, vbExclamation
End Sub

Private Sub ArbeitsblaetterLeeren()
    Dim n As Variant, ws As Worksheet
    For Each n In Split(FELDER_LEEREN, ",")
        Feld(CStr(n)).ClearContents
    Next n
    Feld("f_RechArt").Value = "380"
    Feld("f_Waehrung").Value = "EUR"
    Feld("f_KLand").Value = "DE"
    Feld("f_LLand").Value = "DE"

    Set ws = ThisWorkbook.Worksheets(BLATT_POSITIONEN)
    If ws.Cells(ws.Rows.Count, SP_TYP).End(xlUp).Row > 1 Then
        ws.Range(ws.Cells(2, SP_TYP), ws.Cells(ws.Rows.Count, SP_INFO)).ClearContents
    End If

    Set ws = ThisWorkbook.Worksheets(BLATT_PROTOKOLL)
    If ws.Cells(ws.Rows.Count, 1).End(xlUp).Row > 1 Then ws.Range("A2:C" & ws.Rows.Count).ClearContents
End Sub

' Felder für die SEPA-Lastschrift (ab PDF2E 1.8.0). Fehlen sie in einer älteren
' Mappe, werden sie unten im Blatt "Rechnung" angehängt.
Private Sub LastschriftFelderSicherstellen()
    Dim ws As Worksheet, z As Long, i As Long, namen As Variant, texte As Variant, hinweise As Variant

    namen = Array("f_Mandat", "f_GlaeubigerId", "f_Einzug")
    If NameVorhanden(namen(0)) And NameVorhanden(namen(1)) And NameVorhanden(namen(2)) Then Exit Sub
    texte = Array("Mandatsreferenz", "Gläubiger-ID", "Einzugsdatum")
    hinweise = Array("BT-89, aus der Rechnung", "BT-90, aus der Rechnung", "BT-9 bei Lastschrift (statt Fälligkeitsdatum)")

    Set ws = ThisWorkbook.Worksheets(BLATT_RECHNUNG)
    z = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 2
    ws.Cells(z, 1).Value = "SEPA-Lastschrift"
    ws.Cells(z, 1).Font.Bold = True
    For i = 0 To 2
        z = z + 1
        If Not NameVorhanden(CStr(namen(i))) Then
            ws.Cells(z, 1).Value = texte(i)
            ws.Cells(z, 2).Interior.Color = RGB(238, 238, 238)
            ws.Cells(z, 2).NumberFormat = IIf(i = 2, "DD.MM.YYYY", "@")
            ws.Cells(z, 3).Value = hinweise(i)
            ws.Cells(z, 3).Font.Italic = True
            ws.Cells(z, 3).Font.Color = RGB(119, 119, 119)
            ThisWorkbook.Names.Add name:=CStr(namen(i)), RefersTo:="='" & ws.name & "'!$B$" & z
        End If
    Next i
End Sub

' Jeder Feldname (f_...) muss auf eine eigene Zelle zeigen. Zeigen zwei Namen auf
' dieselbe Zelle (z. B. nach Umbau der Mappe), landen Werte im falschen Feld,
' etwa der Kundenname in der USt-IdNr. Dann lieber mit klarer Meldung abbrechen.
Private Sub FeldnamenPruefen()
    Dim n As Variant, r As Range, adr As Object, meldung As String, k As String
    Set adr = CreateObject("Scripting.Dictionary")
    For Each n In Split(FELDER_LEEREN & ",f_RechArt,f_Waehrung,f_KLand,f_LLand,f_Faellig,f_Netto,f_MwSt,f_Brutto", ",")
        Set r = Nothing
        On Error Resume Next
        Set r = ThisWorkbook.Names(CStr(n)).RefersToRange
        On Error GoTo 0
        If r Is Nothing Then
            meldung = meldung & vbLf & n & ": fehlt oder zeigt ins Leere (#BEZUG!)"
        Else
            k = r.Worksheet.Name & "!" & r.Address(False, False)
            If adr.Exists(k) Then
                meldung = meldung & vbLf & n & " und " & adr(k) & " zeigen beide auf " & k
            Else
                adr.Add k, CStr(n)
            End If
        End If
    Next n
    If Len(meldung) > 0 Then
        Err.Raise vbObjectError + 30, , "Die Feldnamen im Blatt 'Rechnung' sind fehlerhaft (Formeln > Namens-Manager):" & _
                  meldung & vbLf & "Bitte die Mappe aus PDF2E.xlsx im ZIP neu aufbauen."
    End If
End Sub

' Blatt "Empfaenger": USt-IdNr. in Spalte D bereinigen. Steht dort der
' Kundenname (Spalte B), wird das Feld geleert; andere ungültige Werte werden
' gemeldet (die Prüfung verhindert dann das Erzeugen für diesen Kunden).
Private Sub EmpfaengerUStIdBereinigen()
    Dim wsE As Worksheet, z As Long, roh As String, neu As String, meldung As String
    Set wsE = ThisWorkbook.Worksheets("Empfaenger")
    For z = 2 To wsE.Cells(wsE.Rows.Count, 1).End(xlUp).Row
        If Not IsError(wsE.Cells(z, 4).Value) Then
            roh = CStr(wsE.Cells(z, 4).Value)
            If Len(Trim$(roh)) > 0 Then
                neu = UStIdBereinigt(roh)
                meldung = UStIdFehler(neu)
                If Len(meldung) > 0 And neu = UStIdBereinigt(wsE.Cells(z, 2).Value) Then
                    wsE.Cells(z, 4).ClearContents
                    ZeileSpeichern "Empfaenger", z
                    Protokoll "Warnung", "Empfänger " & wsE.Cells(z, 1).Value & ": In der Spalte USt-IdNr. stand der Kundenname '" & _
                                         roh & "'. Feld geleert, bei Bedarf die richtige USt-IdNr. eintragen."
                ElseIf Len(meldung) > 0 Then
                    Protokoll "Warnung", "Empfänger " & wsE.Cells(z, 1).Value & ", USt-IdNr.: " & meldung
                ElseIf neu <> roh Then
                    wsE.Cells(z, 4).NumberFormat = "@"
                    wsE.Cells(z, 4).Value = neu
                    ZeileSpeichern "Empfaenger", z
                End If
            End If
        End If
    Next z
End Sub

' Bankkonto, Zahlungskondition und E-Mail aus dem Blatt "Empfaenger",
' sonst die mit "x" markierten Standard-Einträge
Private Sub VorgabenSetzen()
    Dim wsE As Worksheet, z As Long
    Set wsE = ThisWorkbook.Worksheets("Empfaenger")
    EmpfaengerUStIdBereinigen
    z = TabellenZeile("Empfaenger", EmpfaengerSchluessel())

    If z > 0 Then
        If Len(wsE.Cells(z, 3).Value) > 0 Then Feld("f_KEmail").Value = wsE.Cells(z, 3).Value
        If Len(wsE.Cells(z, 4).Value) > 0 Then Feld("f_KUstId").Value = "'" & UStIdBereinigt(wsE.Cells(z, 4).Value)
        Feld("f_Konto").Value = wsE.Cells(z, 5).Value
        Feld("f_Kondition").Value = wsE.Cells(z, 6).Value
    End If

    If Len(Feld("f_Konto").Value) = 0 Then
        z = StandardZeile("Bankkonten", 5)
        If z > 0 Then Feld("f_Konto").Value = ThisWorkbook.Worksheets("Bankkonten").Cells(z, 1).Value
    End If
    If Len(Feld("f_Kondition").Value) = 0 Then
        z = StandardZeile("Zahlungskonditionen", 4)
        If z > 0 Then Feld("f_Kondition").Value = ThisWorkbook.Worksheets("Zahlungskonditionen").Cells(z, 1).Value
    End If

    If TabellenZeile("Empfaenger", EmpfaengerSchluessel()) = 0 Then EmpfaengerAnlegen
    ReweErkennen
    EdekaErkennen
    G5VorgabenSetzen
    PfandartikelAnlegen
    LieferortGlnSetzen
End Sub

' Neuen Rechnungsempfänger mit den aktuellen Vorgaben ins Blatt "Empfaenger"
' übernehmen, damit der Anwender nichts von Hand anlegen muss. Abweichende
' Vorgaben (Konto, Kondition, E-Mail) werden dort nur bei Bedarf eingetragen.
Private Sub EmpfaengerAnlegen()
    Dim wsE As Worksheet, z As Long
    If Len(CStr(Feld("f_KundenNr").Value)) = 0 Then Exit Sub
    Set wsE = ThisWorkbook.Worksheets("Empfaenger")
    z = wsE.Cells(wsE.Rows.Count, 1).End(xlUp).Row + 1
    wsE.Cells(z, 1).NumberFormat = "@"
    wsE.Cells(z, 1).Value = EmpfaengerSchluessel()
    wsE.Cells(z, 2).Value = Feld("f_KName").Value
    wsE.Cells(z, 3).Value = Feld("f_KEmail").Value
    ' USt-IdNr. nur, wenn sie wie eine aussieht (nie Name o. Ä. aus einem falsch verknüpften Feld)
    If Len(UStIdFehler(UStIdBereinigt(Feld("f_KUstId").Value))) = 0 Then
        wsE.Cells(z, 4).NumberFormat = "@"
        wsE.Cells(z, 4).Value = UStIdBereinigt(Feld("f_KUstId").Value)
    End If
    ' Konto und Kondition bleiben leer = es gilt der jeweilige Standard-Eintrag
    wsE.Cells(z, 7).Value = "Automatisch angelegt am " & Format$(Date, "dd.mm.yyyy") & _
                            " aus RE " & Feld("f_RechNr").Value & " - Konto/Kondition leer = Standard"
    wsE.Range(wsE.Cells(z, 1), wsE.Cells(z, SP_E_VORGABEN)).Interior.Color = RGB(255, 248, 225)
    ZeileSpeichern "Empfaenger", z
    Protokoll "Info", "Neuer Empfänger " & EmpfaengerSchluessel() & " (" & Feld("f_KName").Value & _
                      ") im Blatt 'Empfaenger' angelegt. Abweichendes Konto oder Kondition dort eintragen."
End Sub

' REWE-Zentralregulierung erkennen (Rechnungsempfänger mit RE-GLN, Name beginnt
' mit "REWE") und in der Spalte "Vorgaben" vermerken, sofern dort nichts steht.
Private Sub ReweErkennen()
    Dim wsE As Worksheet, z As Long
    VorgabenSpalteSicherstellen
    z = TabellenZeile("Empfaenger", EmpfaengerSchluessel())
    If z = 0 Then Exit Sub
    Set wsE = ThisWorkbook.Worksheets("Empfaenger")
    If Len(Trim$(CStr(wsE.Cells(z, SP_E_VORGABEN).Value))) > 0 Then Exit Sub
    If UCase$(Left$(Trim$(CStr(Feld("f_KName").Value)), 4)) = "REWE" And Len(Feld("f_KGLN").Value) > 0 Then
        wsE.Cells(z, SP_E_VORGABEN).Value = "REWE"
        ZeileSpeichern "Empfaenger", z
        Protokoll "Info", "Empfänger " & Feld("f_KundenNr").Value & " als REWE erkannt (Blatt 'Empfaenger', Spalte 'Vorgaben')."
    End If
End Sub

' EDEKA erkennen (Name des Rechnungsempfängers enthält "EDEKA") und in der
' Spalte "Vorgaben" vermerken, sofern dort nichts steht. Händler, die nicht
' "EDEKA" im Namen tragen, dort von Hand mit EDEKA markieren.
Private Sub EdekaErkennen()
    Dim wsE As Worksheet, z As Long
    z = TabellenZeile("Empfaenger", EmpfaengerSchluessel())
    If z = 0 Then Exit Sub
    Set wsE = ThisWorkbook.Worksheets("Empfaenger")
    If Len(Trim$(CStr(wsE.Cells(z, SP_E_VORGABEN).Value))) > 0 Then Exit Sub
    If InStr(1, CStr(Feld("f_KName").Value), "EDEKA", vbTextCompare) > 0 Then
        wsE.Cells(z, SP_E_VORGABEN).Value = "EDEKA"
        ZeileSpeichern "Empfaenger", z
        Protokoll "Info", "Empfänger " & EmpfaengerSchluessel() & " als EDEKA erkannt (Blatt 'Empfaenger', Spalte 'Vorgaben')."
    End If
End Sub

' G5-Zentralregulierung (SGH Service GmbH): Empfänger werden von Hand in der
' Spalte "Vorgaben" mit G5 markiert. Die E-Rechnung geht dann immer an die
' zentrale Adresse (Einstellung "G5 E-Mail"), nicht an die Adresse des Händlers.
Private Sub G5VorgabenSetzen()
    Dim adr As String
    ZRSpalteSicherstellen
    If EmpfaengerVorgabe() <> "G5" Then Exit Sub
    adr = Trim$(EinstellungOderLeer("G5 E-Mail"))
    If Len(adr) = 0 Then adr = G5_EMAIL_STANDARD
    If StrComp(adr, Trim$(CStr(Feld("f_KEmail").Value)), vbTextCompare) <> 0 Then
        Feld("f_KEmail").Value = adr
        Protokoll "Info", "G5: E-Mail an die zentrale Adresse " & adr & " (Einstellung 'G5 E-Mail')."
    End If
End Sub

' ZR-Nummer des aktuellen Rechnungsempfängers (Blatt "Empfaenger", Spalte
' "ZR-Nummer"), ohne Leerzeichen. Leer = keine.
Public Function EmpfaengerZRNummer() As String
    Dim z As Long
    z = TabellenZeile("Empfaenger", EmpfaengerSchluessel())
    If z > 0 Then EmpfaengerZRNummer = Replace(Trim$(CStr(ThisWorkbook.Worksheets("Empfaenger").Cells(z, SP_E_ZRNR).Value)), " ", "")
End Function

' Ältere Arbeitsmappen: Spalte "ZR-Nummer" ergänzen (vor den Verwaltungsspalten)
Private Sub ZRSpalteSicherstellen()
    Dim wsE As Worksheet
    Set wsE = ThisWorkbook.Worksheets("Empfaenger")
    If CStr(wsE.Cells(1, SP_E_ZRNR).Value) = "ZR-Nummer" Then Exit Sub
    If CStr(wsE.Cells(1, SP_E_ZRNR).Value) = "Geändert am" Then wsE.Columns(SP_E_ZRNR).Insert
    wsE.Columns(SP_E_ZRNR).NumberFormat = "@"
    wsE.Columns(SP_E_ZRNR).ColumnWidth = 12
    With wsE.Cells(1, SP_E_ZRNR)
        .Value = "ZR-Nummer"
        .Font.Bold = True
        .Interior.Color = wsE.Cells(1, 1).Interior.Color
        .Font.Color = wsE.Cells(1, 1).Font.Color
        If .Comment Is Nothing Then
            .AddComment "Nur bei Vorgaben = G5: ZR-Nummer (Mitgliedsnummer) des Händlers bei der G5, " & _
                        "7 Ziffern, beginnt mit 27. Wird als Käuferreferenz (BT-10) übergeben."
        End If
    End With
End Sub

' Pfandartikel der Rechnung, die noch nicht im Blatt "Pfandartikel" stehen,
' dort anlegen. Der Anwender trägt nur noch die GTIN des Pfandartikels ein
' (eigene GTIN des Leerguts, nicht die des Warenartikels).
Private Sub PfandartikelAnlegen()
    Dim ws As Worksheet, wsP As Worksheet, z As Long, n As Long, nr As String, bez As String
    PfandartikelSicherstellen
    Set ws = ThisWorkbook.Worksheets(BLATT_POSITIONEN)
    Set wsP = ThisWorkbook.Worksheets(BLATT_PFANDARTIKEL)
    For z = 2 To ws.Cells(ws.Rows.Count, SP_TYP).End(xlUp).Row
        nr = Trim$(CStr(ws.Cells(z, SP_ARTNR).Value))
        If ws.Cells(z, SP_TYP).Value = "Pfand" And Len(nr) > 0 Then
            If TabellenZeile(BLATT_PFANDARTIKEL, nr) = 0 Then
                bez = CStr(ws.Cells(z, SP_BEZ).Value)
                If Left$(bez, 6) = "Pfand " Then bez = Mid$(bez, 7)
                bez = Replace(Replace(bez, " Lieferung", ""), " Rücknahme", "")
                n = wsP.Cells(wsP.Rows.Count, 1).End(xlUp).Row + 1
                wsP.Cells(n, 1).NumberFormat = "@"
                wsP.Cells(n, 1).Value = nr
                wsP.Cells(n, 2).Value = bez
                wsP.Cells(n, 3).NumberFormat = "@"
                wsP.Cells(n, 4).Value = "Automatisch angelegt am " & Format$(Date, "dd.mm.yyyy") & " aus RE " & Feld("f_RechNr").Value
                wsP.Range(wsP.Cells(n, 1), wsP.Cells(n, 4)).Interior.Color = RGB(255, 248, 225)
                ZeileSpeichern BLATT_PFANDARTIKEL, n
                Protokoll "Info", "Pfandartikel " & nr & " (" & bez & ") im Blatt '" & BLATT_PFANDARTIKEL & "' angelegt. GTIN dort eintragen."
            End If
        End If
    Next z
End Sub

' Legt das Blatt "Pfandartikel" an, falls es fehlt (ältere Arbeitsmappen).
Public Sub PfandartikelSicherstellen()
    Dim ws As Worksheet, vorher As Object
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(BLATT_PFANDARTIKEL)
    On Error GoTo 0
    If Not ws Is Nothing Then Exit Sub
    Set vorher = ActiveSheet
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets("Empfaenger"))
    ws.Name = BLATT_PFANDARTIKEL
    ws.Range("A1:D1").Value = Array("Leergut-Nr.", "Bezeichnung", "GTIN", "Bemerkung")
    With ws.Range("A1:D1")
        .Font.Bold = True
        .Font.Color = ThisWorkbook.Worksheets("Empfaenger").Cells(1, 1).Font.Color
        .Interior.Color = ThisWorkbook.Worksheets("Empfaenger").Cells(1, 1).Interior.Color
    End With
    ws.Range("C1").AddComment "GTIN des Pfandartikels (Leergut), 8, 12, 13 oder 14 Ziffern. " & _
                               "Nicht die GTIN des Warenartikels. Von EDEKA verlangt."
    ws.Columns(1).ColumnWidth = 12
    ws.Columns(2).ColumnWidth = 34
    ws.Columns(3).ColumnWidth = 18
    ws.Columns(3).NumberFormat = "@"
    ws.Columns(4).ColumnWidth = 50
    If Not vorher Is Nothing Then vorher.Activate
    Protokoll "Info", "Blatt '" & BLATT_PFANDARTIKEL & "' angelegt."
End Sub

' Warenempfänger-GLN (BT-71) aus dem Blatt "Lieferorte", wenn die Rechnung einen
' Lieferort ohne GLN hat (z. B. "Lieferfiliale" bei Getränke Hoffmann). Unbekannte
' Lieferorte werden dort angelegt; der Anwender trägt die GLN einmal ein.
Private Sub LieferortGlnSetzen()
    Dim wsL As Worksheet, z As Long, schluessel As String, gln As String
    If Len(Trim$(CStr(Feld("f_LGLN").Value))) > 0 Then Exit Sub
    If Len(Trim$(CStr(Feld("f_LOrt").Value))) = 0 Then Exit Sub
    LieferorteSicherstellen
    Set wsL = ThisWorkbook.Worksheets(BLATT_LIEFERORTE)
    schluessel = LieferortSchluessel()
    z = TabellenZeile(BLATT_LIEFERORTE, schluessel)
    If z > 0 Then
        gln = Replace(Trim$(CStr(wsL.Cells(z, 3).Value)), " ", "")
        If Len(gln) > 0 Then
            Feld("f_LGLN").Value = "'" & gln
            Protokoll "Info", "GLN des Lieferorts aus dem Blatt '" & BLATT_LIEFERORTE & "': " & gln
        End If
        Exit Sub
    End If
    z = wsL.Cells(wsL.Rows.Count, 1).End(xlUp).Row + 1
    wsL.Cells(z, 1).NumberFormat = "@"
    wsL.Cells(z, 1).Value = schluessel
    wsL.Cells(z, 2).Value = Feld("f_LName").Value
    wsL.Cells(z, 3).NumberFormat = "@"
    wsL.Cells(z, 4).Value = "Automatisch angelegt am " & Format$(Date, "dd.mm.yyyy") & " aus RE " & Feld("f_RechNr").Value
    wsL.Range(wsL.Cells(z, 1), wsL.Cells(z, 4)).Interior.Color = RGB(255, 248, 225)
    ZeileSpeichern BLATT_LIEFERORTE, z
    Protokoll "Info", "Lieferort " & schluessel & " im Blatt '" & BLATT_LIEFERORTE & "' angelegt. GLN dort eintragen, falls der Kunde sie verlangt."
End Sub

' Schlüssel eines Lieferorts: Empfänger / PLZ Ort, Straße
Public Function LieferortSchluessel() As String
    LieferortSchluessel = EmpfaengerSchluessel() & " / " & Trim$(CStr(Feld("f_LPLZ").Value)) & " " & _
                          Trim$(CStr(Feld("f_LOrt").Value)) & ", " & Trim$(CStr(Feld("f_LZeile2").Value))
End Function

' Legt das Blatt "Lieferorte" an, falls es fehlt.
Public Sub LieferorteSicherstellen()
    Dim ws As Worksheet, vorher As Object
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(BLATT_LIEFERORTE)
    On Error GoTo 0
    If Not ws Is Nothing Then Exit Sub
    Set vorher = ActiveSheet
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets("Empfaenger"))
    ws.Name = BLATT_LIEFERORTE
    ws.Range("A1:D1").Value = Array("Lieferort", "Name", "GLN", "Bemerkung")
    With ws.Range("A1:D1")
        .Font.Bold = True
        .Font.Color = ThisWorkbook.Worksheets("Empfaenger").Cells(1, 1).Font.Color
        .Interior.Color = ThisWorkbook.Worksheets("Empfaenger").Cells(1, 1).Interior.Color
    End With
    ws.Range("A1").AddComment "Kundennr. / PLZ Ort, Straße des Lieferorts. Wird beim Einlesen automatisch angelegt, nicht ändern."
    ws.Range("C1").AddComment "GLN des Warenempfängers (13 Ziffern), z. B. der Filiale. Wird als Lieferort-GLN (BT-71) übergeben. " & _
                               "Steht nicht auf der BSI-Rechnung; beim Kunden erfragen."
    ws.Columns(1).ColumnWidth = 50
    ws.Columns(2).ColumnWidth = 36
    ws.Columns(3).ColumnWidth = 16
    ws.Columns(3).NumberFormat = "@"
    ws.Columns(4).ColumnWidth = 45
    If Not vorher Is Nothing Then vorher.Activate
    Protokoll "Info", "Blatt '" & BLATT_LIEFERORTE & "' angelegt."
End Sub

' GTIN eines Pfandartikels aus dem Blatt "Pfandartikel", leer wenn unbekannt
Public Function PfandGtin(ByVal artNr As String) As String
    Dim z As Long
    On Error Resume Next
    z = TabellenZeile(BLATT_PFANDARTIKEL, artNr)
    If z > 0 Then PfandGtin = Replace(Trim$(CStr(ThisWorkbook.Worksheets(BLATT_PFANDARTIKEL).Cells(z, 3).Value)), " ", "")
End Function

' Sondervorgaben des aktuellen Rechnungsempfängers (Blatt "Empfaenger", Spalte
' "Vorgaben"), z. B. REWE. Leer = keine.
Public Function EmpfaengerVorgabe() As String
    Dim z As Long
    z = TabellenZeile("Empfaenger", EmpfaengerSchluessel())
    If z > 0 Then EmpfaengerVorgabe = UCase$(Trim$(CStr(ThisWorkbook.Worksheets("Empfaenger").Cells(z, SP_E_VORGABEN).Value)))
End Function

' Ältere Arbeitsmappen: Überschrift der Spalte "Vorgaben" ergänzen
Private Sub VorgabenSpalteSicherstellen()
    With ThisWorkbook.Worksheets("Empfaenger").Cells(1, SP_E_VORGABEN)
        If Len(.Value) = 0 Then
            .Value = "Vorgaben"
            .Font.Bold = True
            .Interior.Color = ThisWorkbook.Worksheets("Empfaenger").Cells(1, 1).Interior.Color
            .Font.Color = ThisWorkbook.Worksheets("Empfaenger").Cells(1, 1).Font.Color
            .AddComment "Sondervorgaben des Empfängers: REWE = Vorgaben der REWE-Zentralregulierung " & _
                        "(Hinweis Entgeltminderungen, Pflicht: RE-GLN, Lieferort-GLN, USt-IdNr.); " & _
                        "EDEKA = Pfand getrennt, GTIN je Pfandartikel Pflicht; " & _
                        "G5 = G5-Zentralregulierung (ZR-Nummer, G5-Lieferantennummer, zentrale E-Mail-Adresse)"
            .ColumnWidth = 14
        End If
    End With
End Sub
