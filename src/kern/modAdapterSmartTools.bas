Attribute VB_Name = "modAdapterSmartTools"
'==============================================================================
' Adapter für die SmartTools E-Rechnung VBA-Bibliothek (Profil EN16931)
'
' Voraussetzungen:
'   - Bibliothek installiert, Lizenzdatei "Lizenz ... xl-e-rechnung-vba.txt" in
'     %APPDATA%\SmartTools\Excel E-Rechnung VBA
'   - VBA-Editor > Extras > Verweise: "SmartTools E-Rechnung-Bibliothek" angehakt
'
' Grundlage: Musterlösung "SmartTools E-Rechnung.xlsm" (modZUGFeRD, modLizenzCheck)
' und die Online-Dokumentation (Objekt-Katalog). Die Klassen, deren Namen nur aus
' dem Namensschema abgeleitet sind, stehen gesammelt unten in "Neu...". Meldet der
' VBA-Editor dort "Benutzerdefinierter Typ nicht definiert", im Objektkatalog (F2)
' nach dem Objekt suchen und den Namen dort anpassen.
'
' Codes (Einheit, Kennungsschema, Textqualifizierung, USt-Kategorie) werden als
' Zahlen übergeben. Die Werte stammen aus Codelisten.xlsx der Bibliothek.
'==============================================================================
Option Explicit

Private Const LIZENZ_ORDNER As String = "SmartTools\Excel E-Rechnung VBA"

' Codelisten.xlsx: IdentifierCode, TextSubjectQualifierCode
Private Const ID_GLN As Long = 86             ' 0088 EAN Location Code (GLN)
Private Const TEXT_AAK As Long = 9            ' AAK Preiskonditionen
Private Const WAEHRUNG_EUR As Long = 48

Private Const REWE_AAK_STANDARD As String = _
    "Bezüglich der Entgeltminderungen verweisen wir auf die aktuellen Zahlungs- und Konditionsvereinbarungen"

Private mCore As SmartTools_Invoicing_CoreSettings

Public Sub SmartToolsEinbetten(ByVal zielPdf As String)
    Dim inv As SmartTools_Invoicing_EN16931_ZUGFeRD_ZUGFeRDEN16931Invoice
    Dim verstoesse() As SmartTools_Invoicing_RuleViolation
    Dim profil As String, i As Long

    profil = UCase$(EinstellungOderLeer("Profil"))
    If Len(profil) = 0 Then profil = "EN16931"
    If profil <> "EN16931" Then
        Err.Raise vbObjectError + 101, , "Profil '" & profil & "' wird noch nicht unterstützt. Bitte in den Einstellungen 'Profil' = EN16931 eintragen."
    End If

    LizenzPruefen
    Set inv = New SmartTools_Invoicing_EN16931_ZUGFeRD_ZUGFeRDEN16931Invoice

    KopfSetzen inv
    VerkaeuferSetzen inv
    KaeuferSetzen inv
    LieferungSetzen inv
    ZahlungSetzen inv
    PositionenUndSummenSetzen inv
    If EmpfaengerVorgabe() = "REWE" Then ReweVorgabenSetzen inv

    verstoesse = inv.Validate()
    If UBound(verstoesse) >= 0 Then
        For i = LBound(verstoesse) To UBound(verstoesse)
            Protokoll "Fehler", "E-Rechnung " & verstoesse(i).Id & ": " & verstoesse(i).Description & _
                                IIf(Len(verstoesse(i).Location) > 0, " (" & verstoesse(i).Location & ")", "")
        Next i
        If UCase$(EinstellungOderLeer("Prüfbericht anzeigen")) <> "NEIN" Then inv.ShowValidationReport
        Err.Raise vbObjectError + 102, , (UBound(verstoesse) - LBound(verstoesse) + 1) & _
            " Regelverletzung(en) laut Smart-Tools-Prüfung, siehe Blatt 'Protokoll'. Es wurde keine Datei erzeugt."
    End If

    inv.Save zielPdf        ' bettet factur-x.xml in die PDF-Kopie ein (PDF/A-3)
    Protokoll "Info", "Smart Tools: Profil EN16931, Prüfung ohne Fehler"
End Sub

'------------------------------------------------------------------------------
' Lizenz (nach modLizenzCheck der Musterlösung)
'------------------------------------------------------------------------------
Private Sub LizenzPruefen()
    Dim ordner As String, schluessel As String, grund As String

    If mCore Is Nothing Then
        ordner = Environ$("AppData") & "\" & LIZENZ_ORDNER
        Set mCore = New SmartTools_Invoicing_CoreSettings
        mCore.Init ordner
        schluessel = LizenzSchluessel(ordner)
        If Len(schluessel) = 0 Then
            Set mCore = Nothing
            Err.Raise vbObjectError + 103, , "Smart-Tools-Lizenzdatei nicht gefunden. Erwartet: " & ordner & _
                "\Lizenz xl-e-rechnung-vba.txt (oder 'Lizenz <Nr> xl-e-rechnung-vba.txt')."
        End If
        mCore.LicenseControl.LicenseKey = schluessel
    End If

    Select Case mCore.LicenseControl.LicenseStatus
        Case LicenseStatus.LicenseStatus_Valid: Exit Sub
        Case LicenseStatus.LicenseStatus_Blocked: grund = "blockiert"
        Case LicenseStatus.LicenseStatus_Deactivated: grund = "deaktiviert"
        Case LicenseStatus.LicenseStatus_Exhausted: grund = "aufgebraucht"
        Case LicenseStatus.LicenseStatus_Expired: grund = "abgelaufen"
        Case LicenseStatus.LicenseStatus_Invalid: grund = "ungültig"
        Case Else: grund = "aus unbekannten Gründen nicht gültig"
    End Select
    Set mCore = Nothing
    Err.Raise vbObjectError + 104, , "Die Smart-Tools-Lizenz ist " & grund & "."
End Sub

Private Function LizenzSchluessel(ByVal ordner As String) As String
    Dim datei As String, st As Object, bom() As Byte, s As String

    datei = Dir$(ordner & "\Lizenz*xl-e-rechnung-vba.txt")
    If Len(datei) = 0 Then Exit Function

    Set st = CreateObject("ADODB.Stream")
    st.Type = 1                         ' binär, um die Kodierung zu erkennen
    st.Open
    st.LoadFromFile ordner & "\" & datei
    bom = st.Read(2)
    st.Position = 0
    st.Type = 2                         ' Text
    If bom(0) = &HFF And bom(1) = &HFE Then st.Charset = "Unicode" Else st.Charset = "UTF-8"
    s = st.ReadText(-1)
    st.Close
    LizenzSchluessel = Replace(Replace(s, vbCr, ""), vbLf, "")
End Function

'------------------------------------------------------------------------------
' Rechnungsdaten
'------------------------------------------------------------------------------
Private Sub KopfSetzen(ByVal inv As SmartTools_Invoicing_EN16931_ZUGFeRD_ZUGFeRDEN16931Invoice)
    If UCase$(Feld("f_Waehrung").Value) <> "EUR" Then
        Err.Raise vbObjectError + 105, , "Währung '" & Feld("f_Waehrung").Value & "' wird nicht unterstützt (nur EUR)."
    End If
    inv.TypeCode = CLng(Feld("f_RechArt").Value)                ' BT-3, 380 Handelsrechnung
    inv.ID = CStr(Feld("f_RechNr").Value)                       ' BT-1
    inv.IssueDate = Feld("f_RechDatum").Value                   ' BT-2
    inv.CurrencyCode = WAEHRUNG_EUR                             ' BT-5
End Sub

Private Sub VerkaeuferSetzen(ByVal inv As SmartTools_Invoicing_EN16931_ZUGFeRD_ZUGFeRDEN16931Invoice)
    Dim gln As String
    With inv.Seller
        .Name = Stamm("Name")                                   ' BT-27
        .PostalAddress.AddressLine1 = Stamm("Straße")           ' BT-35
        .PostalAddress.PostCode = Stamm("PLZ")                  ' BT-38
        .PostalAddress.City = Stamm("Ort")                      ' BT-37
        .PostalAddress.CountryCode = LandCode(Stamm("Land"))    ' BT-40
        .VatId = Stamm("USt-IdNr")                              ' BT-31
        If Len(Stamm("Steuernummer")) > 0 Then .TaxRegistrationId = Stamm("Steuernummer")   ' BT-32
        If Len(Stamm("Registernummer")) > 0 Then .LegalId.Value = Stamm("Registernummer")   ' BT-30
        If Len(Stamm("E-Mail")) > 0 Then
            .ElectronicAddress.Value = Stamm("E-Mail")          ' BT-34
            .ElectronicAddress.SchemeCode = ElectronicAddressCode_EM
        End If
        If Len(Stamm("Kontakt Name")) > 0 Then .Contact.Name = Stamm("Kontakt Name")        ' BT-41
        If Len(Stamm("Telefon")) > 0 Then .Contact.TelephoneNumber = Stamm("Telefon")       ' BT-42
        If Len(Stamm("E-Mail")) > 0 Then .Contact.EmailAddress = Stamm("E-Mail")            ' BT-43
        gln = StammOderLeer("GLN")
        If Len(gln) > 0 Then .AddId NeueKennung(gln, ID_GLN)    ' BT-29, Schema 0088
    End With
End Sub

Private Sub KaeuferSetzen(ByVal inv As SmartTools_Invoicing_EN16931_ZUGFeRD_ZUGFeRDEN16931Invoice)
    With inv.Buyer
        .Name = CStr(Feld("f_KName").Value)                     ' BT-44
        .PostalAddress.AddressLine1 = CStr(Feld("f_KStrasse").Value)   ' BT-50
        .PostalAddress.PostCode = CStr(Feld("f_KPLZ").Value)    ' BT-53
        .PostalAddress.City = CStr(Feld("f_KOrt").Value)        ' BT-52
        .PostalAddress.CountryCode = LandCode(CStr(Feld("f_KLand").Value))   ' BT-55
        If Len(Feld("f_KGLN").Value) > 0 Then                   ' BT-46: GLN mit Schema 0088
            .Id.Value = GlnBereinigt(Feld("f_KGLN").Value)
            .Id.SchemeCode = ID_GLN
        ElseIf Len(Feld("f_KundenNr").Value) > 0 Then           ' sonst Kundennummer ohne Schema
            .Id.Value = CStr(Feld("f_KundenNr").Value)
        End If
        If Len(Feld("f_KUstId").Value) > 0 Then .VatId = CStr(Feld("f_KUstId").Value)        ' BT-48
        If Len(Feld("f_KEmail").Value) > 0 Then
            .ElectronicAddress.Value = CStr(Feld("f_KEmail").Value)   ' BT-49
            .ElectronicAddress.SchemeCode = ElectronicAddressCode_EM
        End If
    End With
End Sub

Private Sub LieferungSetzen(ByVal inv As SmartTools_Invoicing_EN16931_ZUGFeRD_ZUGFeRDEN16931Invoice)
    With inv.Delivery
        If IsDate(Feld("f_LSDatum").Value) Then .ActualDeliveryDate = Feld("f_LSDatum").Value   ' BT-72
        If Len(Feld("f_LSNr").Value) > 0 Then .DespatchAdviceId = CStr(Feld("f_LSNr").Value)    ' BT-16
        If Len(Feld("f_LName").Value) > 0 Or Len(Feld("f_LGLN").Value) > 0 Then
            .ShipTo.Name = CStr(Feld("f_LName").Value)          ' BT-70
            If Len(Feld("f_LGLN").Value) > 0 Then               ' BT-71: GLN mit Schema 0088
                .ShipTo.Id.Value = GlnBereinigt(Feld("f_LGLN").Value)
                .ShipTo.Id.SchemeCode = ID_GLN
            End If
            If Len(Feld("f_LPLZ").Value) > 0 Or Len(Feld("f_LOrt").Value) > 0 Then
                .ShipTo.PostalAddress.AddressLine1 = CStr(Feld("f_LZeile2").Value)   ' BT-75
                .ShipTo.PostalAddress.PostCode = CStr(Feld("f_LPLZ").Value)        ' BT-78
                .ShipTo.PostalAddress.City = CStr(Feld("f_LOrt").Value)            ' BT-77
                .ShipTo.PostalAddress.CountryCode = LandCode(CStr(Feld("f_LLand").Value))   ' BT-80
            End If
        End If
    End With
End Sub

Private Sub ZahlungSetzen(ByVal inv As SmartTools_Invoicing_EN16931_ZUGFeRD_ZUGFeRDEN16931Invoice)
    Dim konto As SmartTools_Invoicing_EN16931_ZUGFeRD_EN16931_CreditTransfer
    Dim kz As Long, bz As Long

    kz = TabellenZeile("Zahlungskonditionen", CStr(Feld("f_Kondition").Value))
    If kz > 0 Then inv.PaymentTerms.Description = ThisWorkbook.Worksheets("Zahlungskonditionen").Cells(kz, 2).Value   ' BT-20
    If IsDate(Feld("f_Faellig").Value) Then inv.PaymentTerms.DueDate = Feld("f_Faellig").Value                       ' BT-9

    bz = TabellenZeile("Bankkonten", CStr(Feld("f_Konto").Value))
    If bz = 0 Then Err.Raise vbObjectError + 106, , "Bankkonto '" & Feld("f_Konto").Value & "' fehlt im Blatt 'Bankkonten'."
    Set konto = New SmartTools_Invoicing_EN16931_ZUGFeRD_EN16931_CreditTransfer
    With ThisWorkbook.Worksheets("Bankkonten")
        konto.AccountId = Replace(CStr(.Cells(bz, 3).Value), " ", "")          ' BT-84 IBAN
        konto.AccountName = CStr(.Cells(bz, 2).Value)                          ' BT-85
        If Len(.Cells(bz, 4).Value) > 0 Then konto.ServiceProviderId = CStr(.Cells(bz, 4).Value)   ' BT-86 BIC
    End With
    inv.Payment.AddCreditTransfer konto
    inv.Payment.PaymentMeansCode = PaymentMeansCode_SEPA_credit_transfer      ' BT-81 = 58
End Sub

' Positionen (Ware, Pfand), Nachlässe als Abschlag auf Belegebene,
' USt-Aufschlüsselung je Satz und Summen
Private Sub PositionenUndSummenSetzen(ByVal inv As SmartTools_Invoicing_EN16931_ZUGFeRD_ZUGFeRDEN16931Invoice)
    Dim ws As Worksheet, z As Long, nr As Long, typ As String, satzKey As String
    Dim pos As SmartTools_Invoicing_EN16931_ZUGFeRD_EN16931_InvoiceLine
    Dim ust As SmartTools_Invoicing_EN16931_ZUGFeRD_EN16931_InvoiceApplicableTax
    Dim positionen As Currency, nachlaesse As Currency, ustSumme As Currency, ustBetrag As Currency
    Dim basis As Object, kat As Object, k As Variant

    Set basis = CreateObject("Scripting.Dictionary")     ' Schlüssel "Kategorie|Satz" -> Bemessungsgrundlage
    Set kat = CreateObject("Scripting.Dictionary")
    Set ws = ThisWorkbook.Worksheets(BLATT_POSITIONEN)

    For z = 2 To ws.Cells(ws.Rows.Count, SP_TYP).End(xlUp).Row
        typ = CStr(ws.Cells(z, SP_TYP).Value)
        satzKey = UCase$(Trim$(CStr(ws.Cells(z, SP_USTKAT).Value))) & "|" & CStr(ws.Cells(z, SP_USTSATZ).Value)
        basis(satzKey) = basis(satzKey) + CCur(ws.Cells(z, SP_BETRAG).Value)

        If typ = "Nachlass" Then
            nachlaesse = nachlaesse - CCur(ws.Cells(z, SP_BETRAG).Value)
            inv.AddAllowance NeuerAbschlag(Abs(CCur(ws.Cells(z, SP_BETRAG).Value)), CStr(ws.Cells(z, SP_BEZ).Value), _
                                           CStr(ws.Cells(z, SP_USTKAT).Value), ws.Cells(z, SP_USTSATZ).Value)
        Else
            nr = nr + 1
            positionen = positionen + CCur(ws.Cells(z, SP_BETRAG).Value)
            Set pos = New SmartTools_Invoicing_EN16931_ZUGFeRD_EN16931_InvoiceLine
            With pos
                .ID = CStr(nr)                                                      ' BT-126
                .Item.Name = CStr(ws.Cells(z, SP_BEZ).Value)                        ' BT-153
                If Len(ws.Cells(z, SP_ARTNR).Value) > 0 Then .Item.SellerAssignedId = CStr(ws.Cells(z, SP_ARTNR).Value)   ' BT-155
                .InvoicedQuantity = ws.Cells(z, SP_MENGE).Value                     ' BT-129 (Rücknahme negativ)
                .InvoicedQuantityUnitCode = EinheitCode(CStr(ws.Cells(z, SP_EINHEIT).Value))   ' BT-130
                .PriceDetails.NetAmount = ws.Cells(z, SP_PREIS).Value               ' BT-146
                .Vat.Rate = ws.Cells(z, SP_USTSATZ).Value                           ' BT-152
                .Vat.CategoryCode = KategorieCode(CStr(ws.Cells(z, SP_USTKAT).Value))   ' BT-151
                .NetAmount = ws.Cells(z, SP_BETRAG).Value                           ' BT-131
                If typ = "Pfand" Then                                               ' BT-160/161
                    .Item.AddAttribute NeuesMerkmal("DEPOSIT_SYSTEM", PfandArt())
                End If
            End With
            inv.AddLine pos
        End If
    Next z

    For Each k In basis.Keys
        Set ust = New SmartTools_Invoicing_EN16931_ZUGFeRD_EN16931_InvoiceApplicableTax
        ustBetrag = RundenKfm(basis(k) * CCur(Split(k, "|")(1)) / 100)
        With ust
            .TaxableAmount = basis(k)                                   ' BT-116
            .Rate = CCur(Split(k, "|")(1))                              ' BT-119
            .CategoryCode = KategorieCode(CStr(Split(k, "|")(0)))       ' BT-118
            .Amount = ustBetrag                                         ' BT-117
        End With
        inv.AddVat ust
        ustSumme = ustSumme + ustBetrag
    Next k
    If ustSumme <> CCur(Feld("f_MwSt").Value) Then
        Err.Raise vbObjectError + 107, , "USt-Summe je Steuersatz (" & Format$(ustSumme, "#,##0.00") & _
            ") weicht von der MwSt der Rechnung (" & Format$(Feld("f_MwSt").Value, "#,##0.00") & ") ab."
    End If

    With inv.Totals
        .LinesNetAmount = positionen                                    ' BT-106
        If nachlaesse <> 0 Then .AllowancesAmount = nachlaesse          ' BT-107
        .NetAmount = positionen - nachlaesse                            ' BT-109
        .VatAmount = ustSumme                                           ' BT-110
        .GrossAmount = positionen - nachlaesse + ustSumme               ' BT-112
        .PayableAmount = positionen - nachlaesse + ustSumme             ' BT-115
    End With
End Sub

' REWE-Zentralregulierung (Startpaket 1.7 RZF): Hinweis auf Entgeltminderungen (AAK)
Private Sub ReweVorgabenSetzen(ByVal inv As SmartTools_Invoicing_EN16931_ZUGFeRD_ZUGFeRDEN16931Invoice)
    Dim text As String
    text = EinstellungOderLeer("REWE Hinweis AAK")
    If Len(text) = 0 Then text = REWE_AAK_STANDARD
    inv.AddNote NeueNotiz(text, TEXT_AAK)                               ' BT-21 AAK, BT-22
End Sub

'------------------------------------------------------------------------------
' Objekte, deren Klassennamen aus dem Namensschema der Bibliothek abgeleitet sind
' (Dokumentation: EN16931.GlobalIdentifier, .ItemAttribute, .InvoiceNote,
' .InvoiceAllowance). Bei einem Übersetzungsfehler hier den Namen anpassen.
'------------------------------------------------------------------------------
Private Function NeueKennung(ByVal wert As String, ByVal schema As Long) As SmartTools_Invoicing_EN16931_ZUGFeRD_EN16931_GlobalIdentifier
    Set NeueKennung = New SmartTools_Invoicing_EN16931_ZUGFeRD_EN16931_GlobalIdentifier
    NeueKennung.Value = GlnBereinigt(wert)
    NeueKennung.SchemeCode = schema
End Function

Private Function NeuesMerkmal(ByVal name As String, ByVal wert As String) As SmartTools_Invoicing_EN16931_ZUGFeRD_EN16931_ItemAttribute
    Set NeuesMerkmal = New SmartTools_Invoicing_EN16931_ZUGFeRD_EN16931_ItemAttribute
    NeuesMerkmal.Name = name
    NeuesMerkmal.Value = wert
End Function

Private Function NeueNotiz(ByVal text As String, ByVal qualifizierung As Long) As SmartTools_Invoicing_EN16931_ZUGFeRD_EN16931_InvoiceNote
    Set NeueNotiz = New SmartTools_Invoicing_EN16931_ZUGFeRD_EN16931_InvoiceNote
    NeueNotiz.Content = text
    NeueNotiz.SubjectCode = qualifizierung
End Function

Private Function NeuerAbschlag(ByVal betrag As Currency, ByVal grund As String, ByVal kategorie As String, _
                               ByVal satz As Variant) As SmartTools_Invoicing_EN16931_ZUGFeRD_EN16931_InvoiceAllowance
    Set NeuerAbschlag = New SmartTools_Invoicing_EN16931_ZUGFeRD_EN16931_InvoiceAllowance
    NeuerAbschlag.Amount = betrag                                       ' BT-92
    NeuerAbschlag.Reason = grund                                        ' BT-97
    NeuerAbschlag.Vat.CategoryCode = KategorieCode(kategorie)           ' BT-95
    NeuerAbschlag.Vat.Rate = satz                                       ' BT-96
End Function

'------------------------------------------------------------------------------
' Codes
'------------------------------------------------------------------------------
' UN/ECE-Einheit (Einstellungen "Einheit Ware"/"Einheit Pfand") -> UnitCode der Bibliothek
Private Function EinheitCode(ByVal einheit As String) As Long
    Select Case UCase$(Trim$(einheit))
        Case "C62": EinheitCode = 325       ' Stück (one)
        Case "H87": EinheitCode = 895       ' Stück (piece)
        Case "XCR": EinheitCode = 1843      ' Kasten (crate)
        Case "XCS": EinheitCode = 1844      ' Kiste (case)
        Case "XBO": EinheitCode = 1814      ' Flasche
        Case "XKG": EinheitCode = 1923      ' Fass (keg)
        Case "XPX": EinheitCode = 2005      ' Palette
        Case "HLT": EinheitCode = 918       ' Hektoliter
        Case "LTR": EinheitCode = 1267      ' Liter
        Case "KGM": EinheitCode = 1120      ' Kilogramm
        Case "LS": EinheitCode = 1265       ' pauschal
        Case Else
            Err.Raise vbObjectError + 108, , "Einheit '" & einheit & "' ist im Adapter nicht hinterlegt " & _
                "(modAdapterSmartTools, EinheitCode; Wert aus Codelisten.xlsx, Blatt UnitCode, Spalte ID)."
    End Select
End Function

Private Function KategorieCode(ByVal kategorie As String) As Long
    Select Case UCase$(Trim$(kategorie))
        Case "S", "": KategorieCode = DutyOrTaxOrFeeCategoryCode_S   ' Normalsatz/ermäßigt
        Case "Z": KategorieCode = 1
        Case "E": KategorieCode = 2
        Case "AE": KategorieCode = 3
        Case "K": KategorieCode = 4
        Case "G": KategorieCode = 5
        Case "O": KategorieCode = 6
        Case Else: Err.Raise vbObjectError + 109, , "USt-Kategorie '" & kategorie & "' ist im Adapter nicht hinterlegt."
    End Select
End Function

Private Function LandCode(ByVal land As String) As Long
    Select Case UCase$(Trim$(land))
        Case "DE", "": LandCode = CountryCode_DE
        Case "AT": LandCode = 11
        Case Else: Err.Raise vbObjectError + 110, , "Ländercode '" & land & "' ist im Adapter nicht hinterlegt (modAdapterSmartTools, LandCode)."
    End Select
End Function

' Mehrweg (RETURNABLE, Standard) oder Einweg (DISPOSABLE)
Private Function PfandArt() As String
    PfandArt = UCase$(EinstellungOderLeer("Pfandart"))
    If PfandArt <> "DISPOSABLE" Then PfandArt = "RETURNABLE"
End Function

Private Function GlnBereinigt(ByVal gln As Variant) As String
    Dim s As String, i As Long
    s = CStr(gln)
    For i = 1 To Len(s)
        If Mid$(s, i, 1) Like "#" Then GlnBereinigt = GlnBereinigt & Mid$(s, i, 1)
    Next i
End Function
