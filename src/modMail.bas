Attribute VB_Name = "modMail"
'==============================================================================
' E-Mail mit der E-Rechnung als Outlook-Entwurf erstellen
'
' Betreff und Text stehen im Blatt "Mailvorlage" und werden vom Anwender
' gepflegt. Platzhalter in geschweiften Klammern, z. B. {Rechnungsnummer},
' werden beim Erstellen ersetzt. Fehlt das Blatt, wird es mit einer
' Standardvorlage angelegt.
'
' Die Mail wird nur angezeigt, nicht gesendet: Der Anwender prüft, ergänzt
' bei Bedarf und sendet selbst.
'
' Voraussetzung: klassisches Outlook für Windows. Das "neue Outlook" hat keine
' Programmierschnittstelle für VBA.
'==============================================================================
Option Explicit

Public Const BLATT_MAILVORLAGE As String = "Mailvorlage"

' Makro für die Schaltfläche "E-Mail erstellen"
Public Sub EMailErstellen()
    Dim datei As String, an As String, betreff As String, text As String
    Dim ol As Object, mail As Object, konto As String

    datei = CStr(Feld("f_Ausgabedatei").Value)
    If Len(datei) = 0 Then
        MsgBox "Bitte zuerst die E-Rechnung erzeugen.", vbExclamation, TOOL_NAME
        Exit Sub
    End If
    If Len(Dir$(datei)) = 0 Then
        MsgBox "Die E-Rechnung wurde nicht gefunden:" & vbLf & datei & vbLf & vbLf & _
               "Bitte erneut erzeugen.", vbExclamation, TOOL_NAME
        Exit Sub
    End If

    MailvorlageSicherstellen
    an = Trim$(CStr(Feld("f_KEmail").Value))
    betreff = PlatzhalterErsetzen(VorlageWert("Betreff"))
    text = PlatzhalterErsetzen(VorlageWert("Text"))
    If UCase$(EinstellungOderLeer("E-Rechnung-Adapter")) = "STUB" Then
        betreff = "TEST - keine gültige E-Rechnung - " & betreff
    End If

    On Error Resume Next
    Set ol = GetObject(, "Outlook.Application")
    If ol Is Nothing Then Set ol = CreateObject("Outlook.Application")
    On Error GoTo Fehler
    If ol Is Nothing Then
        Err.Raise vbObjectError + 40, , "Outlook konnte nicht gestartet werden. Erforderlich ist das " & _
                                       "klassische Outlook für Windows (nicht das 'neue Outlook')."
    End If

    Set mail = ol.CreateItem(0)                 ' 0 = olMailItem
    With mail
        .To = an
        .CC = VorlageWert("CC")
        .BCC = VorlageWert("BCC")
        .Subject = betreff
        konto = VorlageWert("Absenderkonto")
        If Len(konto) > 0 Then KontoSetzen ol, mail, konto
        .Attachments.Add datei
        .Display                                ' Outlook fügt dabei die Standardsignatur ein
        If UCase$(VorlageWert("Signatur einfügen")) = "NEIN" Then
            .HTMLBody = "<html><body>" & TextAlsHtml(text) & "</body></html>"
        Else
            .HTMLBody = VorBodyEinfuegen(.HTMLBody, TextAlsHtml(text))
        End If
    End With

    Protokoll "Info", "E-Mail-Entwurf erstellt" & IIf(Len(an) > 0, " an " & an, "") & ": " & betreff
    If Len(an) = 0 Then
        Protokoll "Warnung", "Keine E-Mail-Adresse für Kunde " & Feld("f_KundenNr").Value & _
                             ". Bitte im Entwurf ergänzen und im Blatt 'Empfaenger' eintragen."
        MsgBox "Für diesen Kunden ist keine E-Mail-Adresse hinterlegt." & vbLf & _
               "Bitte im Entwurf ergänzen und im Blatt 'Empfaenger' eintragen.", vbInformation, TOOL_NAME
    End If
    Exit Sub

Fehler:
    Protokoll "Fehler", "E-Mail: " & Err.Description
    MsgBox "E-Mail konnte nicht erstellt werden:" & vbLf & Err.Description, vbExclamation, TOOL_NAME
End Sub

' Legt das Blatt "Mailvorlage" mit einer Standardvorlage an, falls es fehlt.
Public Sub MailvorlageSicherstellen()
    Dim ws As Worksheet, i As Long, zeilen As Variant, vorher As Object
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(BLATT_MAILVORLAGE)
    On Error GoTo 0
    If Not ws Is Nothing Then Exit Sub

    Set vorher = ActiveSheet
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = BLATT_MAILVORLAGE
    zeilen = Array( _
        Array("Feld", "Inhalt", "Hinweis"), _
        Array("Betreff", "Rechnung {Rechnungsnummer} vom {Rechnungsdatum}", "Platzhalter siehe unten"), _
        Array("Text", "", "Zeilenumbruch in der Zelle mit Alt+Eingabe"), _
        Array("CC", "", "optional, mehrere Adressen mit ; trennen"), _
        Array("BCC", "", "optional, z. B. eigene Adresse für die Ablage"), _
        Array("Signatur einfügen", "Ja", "Ja = Outlook-Standardsignatur unter dem Text, Nein = ohne"), _
        Array("Absenderkonto", "", "optional: E-Mail-Adresse des Outlook-Kontos, leer = Standardkonto"), _
        Array("", "", ""), _
        Array("Platzhalter", "{Rechnungsnummer} {Rechnungsdatum} {Betrag} {Faellig} {Zahlungsbedingung} " & _
              "{Kundennummer} {Kunde} {Lieferort} {Lieferschein} {Absender} {Dateiname}", ""))
    For i = 0 To UBound(zeilen)
        ws.Cells(i + 1, 1).Resize(1, 3).Value = zeilen(i)
    Next i
    ws.Cells(3, 2).Value = StandardText()      ' einzeln, da länger als 255 Zeichen
    With ws.Range("A1:C1")
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(122, 74, 18)
    End With
    ws.Columns(1).ColumnWidth = 20
    ws.Columns(2).ColumnWidth = 90
    ws.Columns(3).ColumnWidth = 55
    ws.Range("B2:B7").Interior.Color = RGB(255, 248, 225)
    ws.Range("B2:B9").WrapText = True
    ws.Range("A1:C9").VerticalAlignment = xlTop
    ws.Range("C2:C9").Font.Italic = True
    ws.Range("C2:C9").Font.Color = RGB(119, 119, 119)
    ws.Rows(3).AutoFit
    If Not vorher Is Nothing Then vorher.Activate
    Protokoll "Info", "Blatt '" & BLATT_MAILVORLAGE & "' mit Standardvorlage angelegt."
End Sub

Private Function StandardText() As String
    StandardText = "Sehr geehrte Damen und Herren," & vbLf & vbLf & _
        "anbei erhalten Sie unsere Rechnung Nr. {Rechnungsnummer} vom {Rechnungsdatum} über {Betrag} EUR." & vbLf & vbLf & _
        "Die Rechnung ist eine E-Rechnung im ZUGFeRD-Format: Die PDF-Datei enthält zusätzlich die " & _
        "maschinenlesbaren Rechnungsdaten und kann direkt in Ihre Buchhaltung übernommen werden." & vbLf & vbLf & _
        "Zahlungsbedingung: {Zahlungsbedingung}" & vbLf & _
        "Fällig am: {Faellig}" & vbLf & _
        "Lieferung an: {Lieferort}" & vbLf & _
        "Ihre Kundennummer: {Kundennummer}" & vbLf & vbLf & _
        "Mit freundlichen Grüßen" & vbLf & _
        "{Absender}"
End Function

Private Function VorlageWert(ByVal schluessel As String) As String
    Dim c As Range
    Set c = ThisWorkbook.Worksheets(BLATT_MAILVORLAGE).Columns(1).Find( _
        What:=schluessel, LookIn:=xlValues, LookAt:=xlWhole)
    If Not c Is Nothing Then VorlageWert = Trim$(CStr(c.Offset(0, 1).Value))
End Function

Private Function PlatzhalterErsetzen(ByVal s As String) As String
    Dim kz As Long, bedingung As String
    kz = TabellenZeile("Zahlungskonditionen", CStr(Feld("f_Kondition").Value))
    If kz > 0 Then bedingung = CStr(ThisWorkbook.Worksheets("Zahlungskonditionen").Cells(kz, 2).Value)

    s = Replace(s, "{Rechnungsnummer}", CStr(Feld("f_RechNr").Value))
    s = Replace(s, "{Rechnungsdatum}", DatumText(Feld("f_RechDatum").Value))
    s = Replace(s, "{Betrag}", Format$(Feld("f_Brutto").Value, "#,##0.00"))
    s = Replace(s, "{Faellig}", DatumText(Feld("f_Faellig").Value))
    s = Replace(s, "{Zahlungsbedingung}", bedingung)
    s = Replace(s, "{Kundennummer}", CStr(Feld("f_KundenNr").Value))
    s = Replace(s, "{Kunde}", CStr(Feld("f_KName").Value))
    s = Replace(s, "{Lieferort}", CStr(Feld("f_LName").Value))
    s = Replace(s, "{Lieferschein}", CStr(Feld("f_LSNr").Value))
    s = Replace(s, "{Absender}", Stamm("Name"))
    s = Replace(s, "{Dateiname}", Mid$(CStr(Feld("f_Ausgabedatei").Value), InStrRev(CStr(Feld("f_Ausgabedatei").Value), "\") + 1))
    PlatzhalterErsetzen = s
End Function

Private Function DatumText(ByVal v As Variant) As String
    If IsDate(v) Then DatumText = Format$(v, "dd.mm.yyyy")
End Function

' Klartext in HTML umwandeln (Sonderzeichen maskieren, Zeilenumbrüche erhalten)
Private Function TextAlsHtml(ByVal s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, vbCrLf, vbLf)
    s = Replace(s, vbCr, vbLf)
    s = Replace(s, vbLf, "<br>")
    TextAlsHtml = "<div style=""font-family:Calibri,Arial,sans-serif;font-size:11pt"">" & s & "</div><br>"
End Function

' Text direkt hinter dem <body>-Tag einfügen, damit die Signatur darunter bleibt
Private Function VorBodyEinfuegen(ByVal html As String, ByVal einfuegen As String) As String
    Dim p As Long
    p = InStr(1, html, "<body", vbTextCompare)
    If p > 0 Then p = InStr(p, html, ">")
    If p > 0 Then
        VorBodyEinfuegen = Left$(html, p) & einfuegen & Mid$(html, p + 1)
    Else
        VorBodyEinfuegen = einfuegen & html
    End If
End Function

Private Sub KontoSetzen(ByVal ol As Object, ByVal mail As Object, ByVal adresse As String)
    Dim acc As Object
    For Each acc In ol.Session.Accounts
        If LCase$(acc.SmtpAddress) = LCase$(adresse) Then
            Set mail.SendUsingAccount = acc
            Exit Sub
        End If
    Next acc
    Protokoll "Warnung", "Absenderkonto '" & adresse & "' nicht in Outlook gefunden. Standardkonto verwendet."
End Sub
