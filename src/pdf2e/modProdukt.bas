Attribute VB_Name = "modProdukt"
'==============================================================================
' PDF2E - Produktangaben und Datenquelle (PDF-Rechnungen aus der BSI-Software)
'
' Gegenstück in XL2E: src/xl2e/modProdukt.bas. Der gemeinsame Code in src/kern
' ruft nur die hier definierten Konstanten und Prozeduren auf.
'==============================================================================
Option Explicit

Public Const TOOL_NAME As String = "PDF2E"
Public Const TOOL_VERSION As String = "1.6.0"
Public Const TOOL_STAND As String = "04.11.2026"
Public Const TOOL_UNTERTITEL As String = "PDF-Rechnungen als E-Rechnung (ZUGFeRD)"
Public Const QUELLE_SCHALTFLAECHE As String = "PDF einlesen"

' Makros mit Produktnamen (bisherige Namen bleiben gültig)
Public Sub PDF2E_Start()
    AnwendungStart
End Sub

Public Sub PDF2E_Einrichten()
    Einrichten
End Sub

Public Sub PDF2E_Exportieren()
    Quellcode_Exportieren
End Sub

' Dateiauswahl für "PDF einlesen"
Public Function QuelleAuswaehlen() As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "BSI-Rechnung (PDF) auswählen"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "PDF-Dateien", "*.pdf"
        If Len(Einstellung("Eingangsordner")) > 0 Then .InitialFileName = Einstellung("Eingangsordner") & "\"
        If .Show = -1 Then QuelleAuswaehlen = .SelectedItems(1)
    End With
End Function

' Rechnung auslesen und in die Blätter "Rechnung" und "Positionen" schreiben
Public Sub QuelleEinlesen(ByVal pfad As String)
    Select Case UCase$(Einstellung("Layout"))
        Case "BSI": ParseBSI PdfZuText(pfad), Einstellung("Layout-Variante")
        Case Else: Err.Raise vbObjectError + 20, , "Unbekanntes Layout '" & Einstellung("Layout") & "' (Blatt Einstellungen)."
    End Select
End Sub

' PDF, in das die E-Rechnung eingebettet wird: Kopie der Original-PDF
Public Sub QuellPdfErstellen(ByVal quelle As String, ByVal ziel As String)
    FileCopy quelle, ziel          ' Original-PDF bleibt unverändert
End Sub
