Attribute VB_Name = "modEntwicklung"
'==============================================================================
' Nur für die Entwicklung: Quellcode für Git exportieren
'
' Voraussetzung (einmalig): Datei > Optionen > Trust Center > Einstellungen
' für das Trust Center > Makroeinstellungen > "Zugriff auf das
' VBA-Projektobjektmodell vertrauen" aktivieren.
'
' Zielordner: Einstellung "Quellcode-Ordner" (z. B. C:\Work\PDF2E\src),
' sonst Unterordner "src" neben der Arbeitsmappe.
'==============================================================================
Option Explicit

Public Sub PDF2E_Exportieren()
    Dim ziel As String, comp As Object, anzahl As Long

    ziel = EinstellungOderLeer("Quellcode-Ordner")
    If Len(ziel) = 0 Then ziel = ThisWorkbook.Path & "\src"
    If Len(Dir$(ziel, vbDirectory)) = 0 Then MkDir ziel

    On Error GoTo Fehler
    For Each comp In ThisWorkbook.VBProject.VBComponents
        Select Case comp.Type
            Case 1          ' Standardmodul
                comp.Export ziel & "\" & comp.Name & ".bas"
                anzahl = anzahl + 1
            Case 2          ' Klassenmodul
                comp.Export ziel & "\" & comp.Name & ".cls"
                anzahl = anzahl + 1
            Case 3, 100     ' UserForm, DieseArbeitsmappe/Tabellen: nur der Code als Text
                If comp.CodeModule.CountOfLines > 0 Then
                    CodeAlsText comp, ziel & "\" & DateiName(comp) & "_Code.txt"
                    anzahl = anzahl + 1
                End If
        End Select
    Next comp

    MsgBox anzahl & " Dateien exportiert nach" & vbLf & ziel, vbInformation, TOOL_NAME
    Exit Sub

Fehler:
    MsgBox "Export fehlgeschlagen:" & vbLf & Err.Description & vbLf & vbLf & _
           "Ist der Zugriff auf das VBA-Projektobjektmodell erlaubt (Trust Center)?", vbExclamation, TOOL_NAME
End Sub

' "DieseArbeitsmappe" heißt intern je nach Sprache unterschiedlich - einheitlich exportieren
Private Function DateiName(ByVal comp As Object) As String
    If comp.Type = 100 And comp.Name = ThisWorkbook.CodeName Then
        DateiName = "ThisWorkbook"
    Else
        DateiName = comp.Name
    End If
End Function

Private Sub CodeAlsText(ByVal comp As Object, ByVal pfad As String)
    Dim f As Integer
    f = FreeFile
    Open pfad For Output As #f
    Print #f, comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines)
    Close #f
End Sub
