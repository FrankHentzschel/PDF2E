Attribute VB_Name = "modZugferd"
'==============================================================================
' ZUGFeRD erzeugen: PDF bereitstellen (modProdukt.QuellPdfErstellen),
' XML per Smart Tools einbetten und validieren
'
' Die eigentliche Erzeugung übernimmt ein Adapter (Blatt Einstellungen,
' "E-Rechnung-Adapter"):
'   STUB        modAdapterStub: lässt die PDF unverändert und schreibt die Daten als XML
'               daneben. KEINE gültige E-Rechnung, nur zum Testen des Ablaufs.
'   SMARTTOOLS  modAdapterSmartTools: Smart-Tools-Bibliothek, Profil EN16931
'==============================================================================
Option Explicit

Public Sub ZugferdErzeugen()
    Dim quelle As String, ziel As String, adapter As String, info As String

    If Not RechnungPruefen() Then
        MsgBox "Die Prüfung meldet Fehler. Bitte das Blatt 'Protokoll' ansehen.", vbExclamation
        ThisWorkbook.Worksheets(BLATT_PROTOKOLL).Activate
        Exit Sub
    End If

    adapter = UCase$(Einstellung("E-Rechnung-Adapter"))
    quelle = Feld("f_Quelldatei").Value
    ziel = Einstellung("Ausgabeordner") & "\" & Feld("f_RechNr").Value & Einstellung("Dateinamenszusatz")
    If adapter = "STUB" Then ziel = ziel & "_STUB"
    ziel = ziel & ".pdf"
    If Len(Dir$(ziel)) > 0 Then
        info = VerarbeitungInfo(CStr(Feld("f_RechNr").Value))      ' Mehrplatz: wer hat sie erzeugt?
        If Len(info) = 0 And DatenAktiv() Then info = "Herkunft der Datei unbekannt."
        If MsgBox("Die Datei existiert bereits:" & vbLf & ziel & vbLf & vbLf & IIf(Len(info) > 0, info & vbLf & vbLf, "") & _
                  "Überschreiben?", vbYesNo + vbQuestion + vbDefaultButton2, TOOL_NAME) <> vbYes Then Exit Sub
        Kill ziel
    End If

    On Error GoTo Fehler
    QuellPdfErstellen quelle, ziel ' PDF2E: Kopie der Original-PDF, XL2E: PDF aus der Excel-Rechnung
    Select Case adapter
        Case "STUB": StubEinbetten ziel
        Case "SMARTTOOLS": SmartToolsEinbetten ziel
        Case Else: Err.Raise vbObjectError + 30, , "Unbekannter E-Rechnung-Adapter '" & adapter & "' (Blatt Einstellungen)."
    End Select
    Feld("f_Ausgabedatei").Value = ziel
    Protokoll "Info", "ZUGFeRD erzeugt: " & ziel
    VerarbeitungErzeugt CStr(Feld("f_RechNr").Value), quelle, ziel, CStr(Feld("f_KName").Value), _
                        CCur(Val(Feld("f_Brutto").Value)), IIf(adapter = "STUB", "Test (STUB)", "erzeugt")
    MsgBox IIf(adapter = "STUB", "TESTLAUF (Stub) - keine gültige E-Rechnung:", "ZUGFeRD-Rechnung erzeugt:") & vbLf & ziel, vbInformation
    Exit Sub

Fehler:
    Protokoll "Fehler", Err.Description
    If Len(Dir$(ziel)) > 0 Then Kill ziel
    MsgBox "Erzeugen fehlgeschlagen:" & vbLf & Err.Description, vbExclamation
End Sub
