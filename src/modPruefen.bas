Attribute VB_Name = "modPruefen"
'==============================================================================
' Prüfung vor dem Erzeugen: Pflichtangaben und Summenabgleich mit dem PDF
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

    ' Pflichtangaben Brauerei
    Dim k As Variant
    For Each k In Array("Name", "Straße", "PLZ", "Ort", "Land", "USt-IdNr")
        If Len(Stamm(CStr(k))) = 0 Then Fehler "Brauerei-Stammdaten: '" & k & "' fehlt."
    Next k

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
                       ", Betrag laut PDF = " & Format$(ws.Cells(z, SP_BETRAG).Value, "#,##0.00")
            End If
        End If
        If typ = "Pfand" Then pfand = pfand + ws.Cells(z, SP_BETRAG).Value Else ware = ware + ws.Cells(z, SP_BETRAG).Value
    Next z

    ' Summenabgleich mit den Werten aus dem PDF
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

Private Sub Abgleich(ByVal was As String, ByVal berechnet As Variant, ByVal laut As Variant)
    If Len(CStr(laut)) = 0 Then
        Fehler was & ": Wert aus dem PDF fehlt."
    ElseIf RundenKfm(CCur(berechnet)) <> RundenKfm(CCur(laut)) Then
        Fehler was & ": berechnet " & Format$(berechnet, "#,##0.00") & ", laut PDF " & Format$(laut, "#,##0.00")
    End If
End Sub

Private Sub Fehler(ByVal meldung As String)
    mFehler = mFehler + 1
    Protokoll "Fehler", meldung
End Sub
