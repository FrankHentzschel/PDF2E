Attribute VB_Name = "modAdapterStub"
'==============================================================================
' Stub-Adapter: Platzhalter für Smart Tools, bis die Bibliothek vorliegt.
' Schreibt alle Daten, die später an Smart Tools übergeben werden, als
' lesbare XML-Datei neben die PDF (<Ziel>.daten.xml). Die PDF bleibt eine
' unveränderte Kopie. Das Ergebnis ist KEINE gültige E-Rechnung.
'==============================================================================
Option Explicit

Private mDoc As Object

Public Sub StubEinbetten(ByVal zielPdf As String)
    Dim root As Object, grp As Object, pos As Object
    Dim ws As Worksheet, z As Long, nr As Long, kz As Long, bz As Long

    Set mDoc = CreateObject("MSXML2.DOMDocument.6.0")
    mDoc.appendChild mDoc.createProcessingInstruction("xml", "version=""1.0"" encoding=""UTF-8""")
    Set root = mDoc.createElement("ERechnungDaten")
    root.setAttribute "adapter", "STUB"
    root.setAttribute "erzeugt", Format$(Now, "yyyy-mm-dd\Thh:nn:ss")
    root.setAttribute "hinweis", "Testausgabe - keine gueltige E-Rechnung"
    root.setAttribute "profil", Einstellung("Profil")
    mDoc.appendChild root

    Set grp = Gruppe(root, "Kopf")
    Wert grp, "BT-1", "Rechnungsnummer", Feld("f_RechNr").Value
    Wert grp, "BT-2", "Rechnungsdatum", DatumXml(Feld("f_RechDatum").Value)
    Wert grp, "BT-3", "Rechnungsart", Feld("f_RechArt").Value
    Wert grp, "BT-5", "Waehrung", Feld("f_Waehrung").Value
    Wert grp, "BT-16", "Lieferschein", Feld("f_LSNr").Value
    Wert grp, "BT-72", "Leistungsdatum", DatumXml(Feld("f_LSDatum").Value)
    Wert grp, "BT-9", "Faelligkeitsdatum", DatumXml(Feld("f_Faellig").Value)
    kz = TabellenZeile("Zahlungskonditionen", CStr(Feld("f_Kondition").Value))
    If kz > 0 Then Wert grp, "BT-20", "Zahlungsbedingungen", ThisWorkbook.Worksheets("Zahlungskonditionen").Cells(kz, 2).Value

    Set grp = Gruppe(root, "Verkaeufer")
    Wert grp, "BT-27", "Name", Stamm("Name")
    Wert grp, "BT-35", "Strasse", Stamm("Straße")
    Wert grp, "BT-38", "PLZ", Stamm("PLZ")
    Wert grp, "BT-37", "Ort", Stamm("Ort")
    Wert grp, "BT-40", "Land", Stamm("Land")
    Wert grp, "BT-31", "UStIdNr", Stamm("USt-IdNr")
    Wert grp, "BT-32", "Steuernummer", Stamm("Steuernummer")
    Wert grp, "BT-30", "Registernummer", Stamm("Registernummer")
    Wert grp, "BT-41", "Kontakt", Stamm("Kontakt Name")
    Wert grp, "BT-42", "Telefon", Stamm("Telefon")
    Wert grp, "BT-43", "EMail", Stamm("E-Mail")

    Set grp = Gruppe(root, "Kaeufer")
    Wert grp, "BT-44", "Name", Feld("f_KName").Value
    Wert grp, "BT-46", "Kundennummer", Feld("f_KundenNr").Value
    Wert grp, "BT-46", "GLN", Feld("f_KGLN").Value
    Wert grp, "BT-48", "UStIdNr", Feld("f_KUstId").Value
    Wert grp, "BT-49", "EMail", Feld("f_KEmail").Value
    Wert grp, "BT-50", "Strasse", Feld("f_KStrasse").Value
    Wert grp, "BT-53", "PLZ", Feld("f_KPLZ").Value
    Wert grp, "BT-52", "Ort", Feld("f_KOrt").Value
    Wert grp, "BT-55", "Land", Feld("f_KLand").Value

    Set grp = Gruppe(root, "Lieferort")
    Wert grp, "BT-70", "Name", Feld("f_LName").Value
    Wert grp, "BT-71", "GLN", Feld("f_LGLN").Value
    Wert grp, "BT-75", "Zeile", Feld("f_LZeile2").Value
    Wert grp, "BT-78", "PLZ", Feld("f_LPLZ").Value
    Wert grp, "BT-77", "Ort", Feld("f_LOrt").Value
    Wert grp, "BT-80", "Land", Feld("f_LLand").Value

    Set grp = Gruppe(root, "Zahlung")
    Wert grp, "BT-81", "Zahlungsart", "58"
    bz = TabellenZeile("Bankkonten", CStr(Feld("f_Konto").Value))
    If bz > 0 Then
        With ThisWorkbook.Worksheets("Bankkonten")
            Wert grp, "BT-84", "IBAN", .Cells(bz, 3).Value
            Wert grp, "BT-85", "Kontoinhaber", .Cells(bz, 2).Value
            Wert grp, "BT-86", "BIC", .Cells(bz, 4).Value
        End With
    End If

    Set ws = ThisWorkbook.Worksheets(BLATT_POSITIONEN)
    Dim posGrp As Object, nlGrp As Object
    Set posGrp = Gruppe(root, "Positionen")
    Set nlGrp = Gruppe(root, "Nachlaesse")
    For z = 2 To ws.Cells(ws.Rows.Count, SP_TYP).End(xlUp).Row
        If ws.Cells(z, SP_TYP).Value = "Nachlass" Then
            Set pos = Gruppe(nlGrp, "Nachlass")
            Wert pos, "BT-92", "Betrag", ZahlXml(Abs(ws.Cells(z, SP_BETRAG).Value), 2)
            Wert pos, "BT-97", "Grund", ws.Cells(z, SP_BEZ).Value
            Wert pos, "BT-95", "UStKategorie", ws.Cells(z, SP_USTKAT).Value
            Wert pos, "BT-96", "UStSatz", ZahlXml(ws.Cells(z, SP_USTSATZ).Value, 2)
        Else
            nr = nr + 1
            Set pos = Gruppe(posGrp, "Position")
            pos.setAttribute "typ", ws.Cells(z, SP_TYP).Value
            Wert pos, "BT-126", "Nr", CStr(nr)
            Wert pos, "BT-155", "ArtikelNr", ws.Cells(z, SP_ARTNR).Value
            Wert pos, "BT-153", "Bezeichnung", ws.Cells(z, SP_BEZ).Value
            Wert pos, "BT-129", "Menge", ZahlXml(ws.Cells(z, SP_MENGE).Value, 4)
            Wert pos, "BT-130", "Einheit", ws.Cells(z, SP_EINHEIT).Value
            Wert pos, "BT-146", "Nettopreis", ZahlXml(ws.Cells(z, SP_PREIS).Value, 4)
            Wert pos, "BT-131", "Nettobetrag", ZahlXml(ws.Cells(z, SP_BETRAG).Value, 2)
            Wert pos, "BT-151", "UStKategorie", ws.Cells(z, SP_USTKAT).Value
            Wert pos, "BT-152", "UStSatz", ZahlXml(ws.Cells(z, SP_USTSATZ).Value, 2)
        End If
    Next z

    Set grp = Gruppe(root, "Summen")
    Wert grp, "BT-109", "Nettosumme", ZahlXml(Feld("f_Netto").Value, 2)
    Wert grp, "BT-110", "MwSt", ZahlXml(Feld("f_MwSt").Value, 2)
    Wert grp, "BT-112", "Bruttosumme", ZahlXml(Feld("f_Brutto").Value, 2)

    mDoc.Save zielPdf & ".daten.xml"
    Protokoll "Warnung", "STUB: Daten geschrieben nach " & zielPdf & ".daten.xml - PDF ist unverändert, keine E-Rechnung."
End Sub

Private Function Gruppe(ByVal eltern As Object, ByVal name As String) As Object
    Dim e As Object
    Set e = mDoc.createElement(name)
    eltern.appendChild e
    Set Gruppe = e
End Function

Private Sub Wert(ByVal eltern As Object, ByVal bt As String, ByVal name As String, ByVal inhalt As Variant)
    Dim e As Object
    Set e = mDoc.createElement(name)
    e.setAttribute "bt", bt
    e.Text = CStr(inhalt)
    eltern.appendChild e
End Sub

Private Function DatumXml(ByVal d As Variant) As String
    If IsDate(d) Then DatumXml = Format$(d, "yyyy-mm-dd")
End Function

' Zahl mit Punkt als Dezimaltrennzeichen, unabhängig von den Ländereinstellungen
Private Function ZahlXml(ByVal x As Variant, ByVal stellen As Long) As String
    If Len(CStr(x)) = 0 Then Exit Function
    ZahlXml = Replace(Format$(CDbl(x), "0." & String$(stellen, "0")), ",", ".")   ' Format ohne Tausendertrennzeichen
End Function
