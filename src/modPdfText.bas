Attribute VB_Name = "modPdfText"
'==============================================================================
' PDF-Text auslesen mit pdftotext (Xpdf-Kommandozeilenwerkzeug)
'==============================================================================
Option Explicit

' Liefert den Text der PDF im Layout-Modus (Spalten bleiben erhalten).
' Seiten sind durch Seitenvorschub (Chr(12)) getrennt, Zeilenenden sind vbLf.
Public Function PdfZuText(ByVal pdfPfad As String) As String
    Dim exe As String, txtPfad As String, cmd As String, rc As Long, s As String

    exe = Einstellung("pdftotext-Pfad")
    If Len(Dir$(exe)) = 0 Then Err.Raise vbObjectError + 10, , "pdftotext nicht gefunden: " & exe
    If Len(Dir$(pdfPfad)) = 0 Then Err.Raise vbObjectError + 11, , "PDF nicht gefunden: " & pdfPfad

    txtPfad = Environ$("TEMP") & "\zugferd_" & Format$(Now, "yyyymmdd_hhnnss") & ".txt"
    cmd = """" & exe & """ -layout -enc UTF-8 """ & pdfPfad & """ """ & txtPfad & """"
    rc = CreateObject("WScript.Shell").Run(cmd, 0, True)
    If rc <> 0 Or Len(Dir$(txtPfad)) = 0 Then
        Err.Raise vbObjectError + 12, , "pdftotext meldet Fehler (Code " & rc & ") für " & pdfPfad
    End If

    s = DateiLesenUtf8(txtPfad)
    If UCase$(EinstellungOderLeer("Textdatei behalten")) = "JA" Then
        Protokoll "Info", "pdftotext-Ausgabe gespeichert: " & txtPfad
    Else
        Kill txtPfad
    End If

    s = Replace(s, vbCrLf, vbLf)
    s = Replace(s, vbCr, vbLf)
    PdfZuText = s
End Function

Private Function DateiLesenUtf8(ByVal pfad As String) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2            ' Text
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile pfad
    DateiLesenUtf8 = st.ReadText(-1)
    st.Close
End Function
