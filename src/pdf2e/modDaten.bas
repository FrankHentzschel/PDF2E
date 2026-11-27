Attribute VB_Name = "modDaten"
'==============================================================================
' Datenschicht: Stammdaten und Verarbeitungsprotokoll im Access-Backend (.accdb)
'
' Grundsatz: Die Blätter (Einstellungen, Mailvorlage, Brauerei, Bankkonten,
' Zahlungskonditionen, Empfaenger, Pfandartikel) bleiben die Oberfläche, das
' Backend ist der Speicher. Die Programmmappe kann daher bei einem Update
' einfach ersetzt werden. Platzbezogene Werte (Backend-Pfad, pdftotext-Pfad)
' stehen im Blatt "Arbeitsplatz"; die Starter-Mappe schreibt sie bei jedem
' Start. Ist "Backend-Pfad" leer, arbeitet das Programm wie bis 1.6.0 nur mit
' den Blättern (Rückfall).
'
'   Laden     beim Aktivieren eines Stammdatenblatts und vor jedem Einlesen
'   Ändern    sofort pro Zeile (Workbook_SheetChange), Konfliktprüfung über
'             den Änderungszähler Rev: UPDATE ... WHERE ID = x AND Rev = y
'   Löschen   nur über "Datensatz löschen" (Kontextmenü der Zelle)
'
' Rechts neben den Datenspalten stehen je Blatt vier Verwaltungsspalten:
' "Geändert am", "Geändert von", ID und Rev (ID und Rev ausgeblendet).
'
' Das Backend wird beim ersten Start angelegt (mit dem Inhalt der Blätter
' vorbelegt) und bei einer neuen Schema-Version automatisch migriert. Die
' Schema-Version steht in der Tabelle Systeminfo und ist unabhängig von der
' Programmversion.
'
' Zugriff per DAO mit später Bindung (kein Verweis nötig). Voraussetzung ist
' die Access-Datenbank-Engine (ACE) in der Bitness von Office.
'==============================================================================
Option Explicit

Public Const SCHEMA_VERSION As Long = 4                ' 2 = Empfaenger.ZRNummer (1.10.0), 3 = Tabelle Lieferorte (1.12.0),
                                                       ' 4 = Empfaenger.WEGLN (1.13.0)
Private Const DB_SPRACHE As String = ";LANGID=0x0409;CP=1252;COUNTRY=0"   ' dbLangGeneral
Private Const DB_VERSION120 As Long = 128                                ' dbVersion120 (.accdb)
Private Const DB_OPEN_DYNASET As Long = 2
Private Const DB_FAIL_ON_ERROR As Long = 128
Private Const DATENSATZ_MENUE As String = "Datensatz löschen (Datenbank)"

Private mEngine As Object
Private mIntern As Boolean              ' True = Blatt wird vom Programm geschrieben

'------------------------------------------------------------------------------
' Tabellenbeschreibung: Blatt, Tabelle, Felder (Name:Typ, erste = Schlüssel)
' Typen: T = Text (255), M = Langtext, N = Zahl,
'        A = Arbeitsplatz (Computername; leer im Blatt = "*" im Backend = alle Plätze)
'------------------------------------------------------------------------------
Private Function Tabellen() As Variant
    Tabellen = Array("Einstellungen", "Mailvorlage", "Brauerei", "Bankkonten", "Zahlungskonditionen", "Empfaenger", "Pfandartikel", _
                     BLATT_LIEFERORTE)
End Function

Private Function Felder(ByVal tabelle As String) As Variant
    Select Case tabelle
        Case "Einstellungen":       Felder = Array("Einstellung:T", "Wert:T", "Hinweis:M", "Arbeitsplatz:A")
        Case "Mailvorlage":         Felder = Array("Feld:T", "Inhalt:M", "Hinweis:M")
        Case "Brauerei":            Felder = Array("Angabe:T", "Wert:T", "Hinweis:M")
        Case "Bankkonten":          Felder = Array("Bezeichnung:T", "Kontoinhaber:T", "IBAN:T", "BIC:T", "Standard:T")
        Case "Zahlungskonditionen": Felder = Array("Bezeichnung:T", "Zahlungstext:M", "ZielTage:N", "Standard:T")
        Case "Empfaenger":          Felder = Array("Schluessel:T", "Kundenname:T", "EMail:T", "UStId:T", "Konto:T", _
                                                   "Kondition:T", "Bemerkung:M", "Vorgaben:T", "ZRNummer:T", "WEGLN:T")
        Case "Pfandartikel":        Felder = Array("LeergutNr:T", "Bezeichnung:T", "GTIN:T", "Bemerkung:M")
        Case BLATT_LIEFERORTE:      Felder = Array("Lieferort:T", "Name:T", "GLN:T", "Bemerkung:M")
        Case Else:                  Felder = Empty
    End Select
End Function

' Erste Verwaltungsspalte (Geändert am). Mailvorlage: Spalte E ist die
' Absender-Auswahlliste aus Outlook, daher dort erst ab Spalte H.
Private Function MetaSpalte(ByVal tabelle As String) As Long
    If tabelle = "Mailvorlage" Then MetaSpalte = 8 Else MetaSpalte = UBound(Felder(tabelle)) + 2
End Function

' Felder des eindeutigen Index: Schlüssel, bei Einstellungen zusätzlich der Arbeitsplatz
Private Function IndexFelder(ByVal tabelle As String) As String
    IndexFelder = "[" & FeldName(Felder(tabelle)(0)) & "]"
    If tabelle = "Einstellungen" Then IndexFelder = IndexFelder & ", [Arbeitsplatz]"
End Function

Private Function FeldName(ByVal f As Variant) As String
    FeldName = Split(CStr(f), ":")(0)
End Function

Private Function FeldTyp(ByVal f As Variant) As String
    FeldTyp = Split(CStr(f), ":")(1)
End Function

Private Function IstDatenblatt(ByVal blatt As String) As Boolean
    Dim t As Variant
    For Each t In Tabellen()
        If t = blatt Then IstDatenblatt = True: Exit Function
    Next t
End Function

'------------------------------------------------------------------------------
' Öffentliche Abfragen
'------------------------------------------------------------------------------
Public Function DatenAktiv() As Boolean
    DatenAktiv = Len(BackendPfad()) > 0
End Function

Public Function BackendPfad() As String
    BackendPfad = Trim$(EinstellungOderLeer("Backend-Pfad"))
End Function

Public Function Benutzer() As String
    Benutzer = Environ$("USERNAME")
End Function

Public Function Rechner() As String
    Rechner = Environ$("COMPUTERNAME")
End Function

'------------------------------------------------------------------------------
' Verbindung
'------------------------------------------------------------------------------
Private Function Engine() As Object
    If mEngine Is Nothing Then
        On Error Resume Next
        Set mEngine = CreateObject("DAO.DBEngine.120")
        On Error GoTo 0
        If mEngine Is Nothing Then
            Err.Raise vbObjectError + 60, , "Die Access-Datenbank-Engine (DAO.DBEngine.120) ist auf diesem Rechner " & _
                      "nicht verfügbar. Ohne sie kann das Backend nicht genutzt werden " & _
                      "(Einstellung 'Backend-Pfad' leeren = Betrieb nur mit den Blättern)."
        End If
    End If
    Set Engine = mEngine
End Function

Private Function DbOeffnen(Optional ByVal exklusiv As Boolean = False) As Object
    Set DbOeffnen = Engine().OpenDatabase(BackendPfad(), exklusiv, False)
End Function

'------------------------------------------------------------------------------
' Programmstart: Backend prüfen, anlegen, migrieren, sichern
' Aufruf aus AnwendungStart. Fehler werden gemeldet, der Start geht weiter.
'------------------------------------------------------------------------------
Public Sub DatenStart()
    Dim pfad As String, schritt As String, txt As String
    If Not DatenAktiv() Then Exit Sub
    pfad = BackendPfad()
    On Error GoTo Fehler

    schritt = "Backend-Ordner prüfen"
    If Not OrdnerExistiert(OrdnerVon(pfad)) Then
        Err.Raise vbObjectError + 61, , "Der Ordner des Backends ist nicht erreichbar: " & OrdnerVon(pfad)
    End If

    If Len(Dir$(pfad)) = 0 Then
        schritt = "Backend anlegen"
        BackendAnlegen
    Else
        schritt = "Schema-Version prüfen und Backend umstellen"
        SchemaSicherstellen
    End If

    schritt = "Neue Einstellungen der Programmversion übernehmen"
    NachUpdateErgaenzen
    VeralteteEinstellungenEntfernen
    schritt = "Einstellungen und Brauerei laden"
    BlattLaden "Einstellungen"              ' gemeinsame Einstellungen aus dem Backend
    BlattLaden "Brauerei"                   ' Lizenznehmer im Splash
    TagessicherungPruefen
    Exit Sub

Fehler:
    txt = Err.Description
    If Err.Number = 76 Then txt = txt & " (Fehler 76: ein Ordner fehlt oder ist nicht erreichbar)"
    Protokoll "Fehler", "Backend (" & schritt & "): " & txt
    MsgBox "Problem mit dem Backend:" & vbLf & txt & vbLf & vbLf & _
           "Schritt: " & schritt & vbLf & _
           "Backend: " & pfad, vbExclamation, TOOL_NAME
End Sub

' Schema-Version des Backends prüfen und bei Bedarf migrieren. Läuft beim Start
' und vor jedem Laden der Stammdaten: Wurden neue Module importiert, ohne die
' Mappe neu zu öffnen, fehlt sonst z. B. die Spalte ZRNummer (DAO-Fehler 3265
' "Element in dieser Auflistung nicht gefunden").
Private Sub SchemaSicherstellen()
    Dim db As Object, ver As Long
    Set db = DbOeffnen()
    ver = CLng(Val(SysWert(db, "SchemaVersion")))
    db.Close
    Set db = Nothing
    If ver > SCHEMA_VERSION Then
        Err.Raise vbObjectError + 62, , "Das Backend hat die Schema-Version " & ver & ", dieses Programm kennt nur " & _
                  SCHEMA_VERSION & ". Bitte PDF2E über die Starter-Mappe neu starten, damit die aktuelle Version geladen wird."
    ElseIf ver < SCHEMA_VERSION Then
        BackendMigrieren ver
    End If
End Sub

Private Sub BackendAnlegen()
    Dim db As Object
    On Error GoTo Fehler
    Set db = Engine().CreateDatabase(BackendPfad(), DB_SPRACHE, DB_VERSION120)
    db.Execute "CREATE TABLE Systeminfo ([Schluessel] TEXT(50) CONSTRAINT pk_Systeminfo PRIMARY KEY, [Wert] TEXT(255))", DB_FAIL_ON_ERROR
    SchemaSchritt db, 1
    SysSetzen db, "SchemaVersion", CStr(SCHEMA_VERSION)
    SysSetzen db, "Stand", Format$(Now, "yyyy-mm-dd hh:nn:ss")
    SysSetzen db, "Programmversion", TOOL_NAME & " " & TOOL_VERSION
    SysSetzen db, "Angelegt", Format$(Now, "yyyy-mm-dd hh:nn:ss") & " von " & Benutzer() & " (" & Rechner() & ")"
    db.Close
    Set db = Nothing
    ImportAusMappe ThisWorkbook, True, Tabellen()
    Protokoll "Info", "Backend angelegt und mit den Daten der Arbeitsmappe vorbelegt: " & BackendPfad()
    MsgBox "Das Backend wurde neu angelegt und mit den Stammdaten dieser Arbeitsmappe vorbelegt:" & vbLf & _
           BackendPfad(), vbInformation, TOOL_NAME
    Exit Sub
Fehler:
    Dim nr As Long, txt As String
    nr = Err.Number: txt = Err.Description
    If Not db Is Nothing Then db.Close
    ' Zwei Plätze starten gleichzeitig: der zweite findet die Datei schon vor
    If Len(Dir$(BackendPfad())) > 0 And nr = 3204 Then Exit Sub
    Err.Raise nr, , "Backend konnte nicht angelegt werden: " & txt
End Sub

' Nach einem Programmupdate: Einstellungen und Mailvorlage-Einträge, die die
' neue Programmmappe mitbringt und die im Backend noch fehlen, übernehmen.
' Vorhandene Werte bleiben unverändert (die Werte des Kunden gelten).
Private Sub NachUpdateErgaenzen()
    Dim db As Object, v As String
    Set db = DbOeffnen()
    v = SysWert(db, "Programmversion")
    db.Close
    If v = TOOL_NAME & " " & TOOL_VERSION Then Exit Sub
    ImportAusMappe ThisWorkbook, False, Array("Einstellungen", "Mailvorlage")
    Set db = DbOeffnen()
    SysSetzen db, "Programmversion", TOOL_NAME & " " & TOOL_VERSION
    db.Close
End Sub

' Einstellungen, die das Programm nicht mehr nutzt (ab 1.7.0 fest im Code)
Private Sub VeralteteEinstellungenEntfernen()
    Dim db As Object
    On Error Resume Next
    Set db = DbOeffnen()
    db.Execute "DELETE FROM Einstellungen WHERE [Einstellung] IN ('Entwickler', 'Firma', 'Kontakt')"
    db.Close
End Sub

' Migration von Schema-Version "von" auf SCHEMA_VERSION, vorher Sicherung
Private Sub BackendMigrieren(ByVal von As Long)
    Dim db As Object, v As Long
    BackendSichern "vor_Migration"
    On Error Resume Next
    Set db = DbOeffnen(True)
    On Error GoTo 0
    If db Is Nothing Then
        Err.Raise vbObjectError + 63, , "Das Backend muss auf Schema-Version " & SCHEMA_VERSION & " aktualisiert werden. " & _
                  "Dazu dürfen keine anderen Plätze PDF2E geöffnet haben. Bitte dort schließen und erneut starten."
    End If
    For v = von + 1 To SCHEMA_VERSION
        SchemaSchritt db, v
        SysSetzen db, "SchemaVersion", CStr(v)
    Next v
    SysSetzen db, "Stand", Format$(Now, "yyyy-mm-dd hh:nn:ss")
    SysSetzen db, "Programmversion", TOOL_NAME & " " & TOOL_VERSION
    db.Close
    Protokoll "Info", "Backend von Schema-Version " & von & " auf " & SCHEMA_VERSION & " aktualisiert."
End Sub

' Ein Schritt des Datenmodells. Neue Schritte nur anhängen, nie ändern.
Private Sub SchemaSchritt(ByVal db As Object, ByVal v As Long)
    Dim t As Variant
    Select Case v
        Case 1
            For Each t In Tabellen()
                db.Execute TabelleDDL(CStr(t)), DB_FAIL_ON_ERROR
                db.Execute "CREATE UNIQUE INDEX ux_" & t & " ON [" & t & "] (" & IndexFelder(CStr(t)) & ")", DB_FAIL_ON_ERROR
            Next t
            db.Execute "CREATE TABLE Verarbeitung (ID COUNTER CONSTRAINT pk_Verarbeitung PRIMARY KEY, " & _
                       "RechNr TEXT(50) NOT NULL, Quelldatei TEXT(255), Ausgabedatei TEXT(255), Empfaenger TEXT(255), " & _
                       "Brutto CURRENCY, Status TEXT(30), ErzeugtAm DATETIME, ErzeugtVon TEXT(100), Rechner TEXT(100), " & _
                       "MailAm DATETIME, MailVon TEXT(100))", DB_FAIL_ON_ERROR
            db.Execute "CREATE UNIQUE INDEX ux_Verarbeitung ON Verarbeitung (RechNr)", DB_FAIL_ON_ERROR
        Case 2                                  ' ZR-Nummer des Händlers (G5-Zentralregulierung, BT-10)
            db.Execute "ALTER TABLE [Empfaenger] ADD COLUMN [ZRNummer] TEXT(255)", DB_FAIL_ON_ERROR
        Case 3                                  ' GLN je Lieferort (Warenempfänger-GLN, BT-71)
            db.Execute TabelleDDL(BLATT_LIEFERORTE), DB_FAIL_ON_ERROR
            db.Execute "CREATE UNIQUE INDEX ux_" & BLATT_LIEFERORTE & " ON [" & BLATT_LIEFERORTE & "] (" & _
                       IndexFelder(BLATT_LIEFERORTE) & ")", DB_FAIL_ON_ERROR
        Case 4                                  ' Standard-Warenempfänger-GLN je Empfänger (BT-71)
            db.Execute "ALTER TABLE [Empfaenger] ADD COLUMN [WEGLN] TEXT(255)", DB_FAIL_ON_ERROR
        ' Case 5: künftige Änderungen nur anhängen
    End Select
End Sub

Private Function TabelleDDL(ByVal tabelle As String) As String
    Dim s As String, f As Variant
    s = "CREATE TABLE [" & tabelle & "] ([ID] COUNTER CONSTRAINT pk_" & tabelle & " PRIMARY KEY"
    For Each f In Felder(tabelle)
        s = s & ", [" & FeldName(f) & "]"
        Select Case FeldTyp(f)
            Case "T": s = s & " TEXT(255)"
            Case "A": s = s & " TEXT(100) NOT NULL"
            Case "M": s = s & " LONGTEXT"
            Case "N": s = s & " DOUBLE"
        End Select
    Next f
    TabelleDDL = s & ", [Rev] LONG, [GeaendertAm] DATETIME, [GeaendertVon] TEXT(100))"
End Function

Private Function SysWert(ByVal db As Object, ByVal name As String) As String
    Dim rs As Object
    Set rs = db.OpenRecordset("SELECT [Wert] FROM Systeminfo WHERE [Schluessel] = '" & Replace(name, "'", "''") & "'")
    If Not rs.EOF Then SysWert = Nz0(rs.Fields("Wert").Value)
    rs.Close
End Function

Private Sub SysSetzen(ByVal db As Object, ByVal name As String, ByVal wert As String)
    Dim qd As Object
    Set qd = db.CreateQueryDef("", "UPDATE Systeminfo SET [Wert] = [pW] WHERE [Schluessel] = [pN]")
    qd.Parameters("pW").Value = wert
    qd.Parameters("pN").Value = name
    qd.Execute DB_FAIL_ON_ERROR
    If qd.RecordsAffected = 0 Then
        Set qd = db.CreateQueryDef("", "INSERT INTO Systeminfo ([Schluessel], [Wert]) VALUES ([pN], [pW])")
        qd.Parameters("pW").Value = wert
        qd.Parameters("pN").Value = name
        qd.Execute DB_FAIL_ON_ERROR
    End If
End Sub

'------------------------------------------------------------------------------
' Sicherung
'------------------------------------------------------------------------------
' Einmal am Tag (erster Start) eine Kopie in den Backup-Ordner, die ältesten
' über "Backup behalten" (Standard 30) werden gelöscht.
Private Sub TagessicherungPruefen()
    Dim db As Object, letzte As String
    If Len(SicherungsEinstellung("Backup-Ordner")) = 0 Then Exit Sub
    On Error GoTo Fehler
    Set db = DbOeffnen()
    letzte = SysWert(db, "LetzteSicherung")
    db.Close
    Set db = Nothing
    If Left$(letzte, 10) = Format$(Date, "yyyy-mm-dd") Then Exit Sub
    BackendSichern ""
    Exit Sub
Fehler:
    Dim txt As String
    txt = Err.Description
    On Error Resume Next
    If Not db Is Nothing Then db.Close
    Protokoll "Warnung", "Tagessicherung des Backends fehlgeschlagen: " & txt
    MsgBox "Die Tagessicherung des Backends ist fehlgeschlagen:" & vbLf & txt & vbLf & vbLf & _
           "Bitte die Einstellung 'Backup-Ordner' prüfen. Das Arbeiten mit PDF2E ist trotzdem möglich.", _
           vbExclamation, TOOL_NAME
End Sub

' Kopie des Backends in den Backup-Ordner (leer = neben dem Backend, Unterordner Sicherung)
' Ist der Backup-Ordner nicht erreichbar, wird bei einer Sicherung vor der
' Migration (zusatz nicht leer) in den Unterordner Sicherung neben dem Backend
' ausgewichen, damit die Umstellung nicht an der Sicherung scheitert.
Public Sub BackendSichern(ByVal zusatz As String)
    Dim ordner As String, ziel As String, db As Object, fehlerText As String, rueckfall As String
    If Not DatenAktiv() Then Exit Sub
    If Len(Dir$(BackendPfad())) = 0 Then Exit Sub
    rueckfall = OrdnerVon(BackendPfad()) & "\Sicherung"
    ordner = SicherungsEinstellung("Backup-Ordner")
    If Len(ordner) = 0 Then ordner = rueckfall
    If Right$(ordner, 1) = "\" Then ordner = Left$(ordner, Len(ordner) - 1)

    ziel = SicherungKopieren(ordner, zusatz, fehlerText)
    If Len(ziel) = 0 And Len(zusatz) > 0 And StrComp(ordner, rueckfall, vbTextCompare) <> 0 Then
        Protokoll "Warnung", "Backup-Ordner nicht nutzbar (" & fehlerText & "), Sicherung vor der Umstellung in " & rueckfall
        ziel = SicherungKopieren(rueckfall, zusatz, fehlerText)
        If Len(ziel) > 0 Then ordner = rueckfall
    End If
    If Len(ziel) = 0 Then Err.Raise vbObjectError + 64, , fehlerText

    Set db = DbOeffnen()
    SysSetzen db, "LetzteSicherung", Format$(Now, "yyyy-mm-dd hh:nn:ss") & " " & ziel
    db.Close
    Protokoll "Info", "Backend gesichert: " & ziel
    AlteSicherungenLoeschen ordner, DateiStamm(BackendPfad())
End Sub

' Einstellung für die Sicherung direkt aus dem Backend lesen (Zeile dieses
' Arbeitsplatzes vor gemeinsamer Zeile), sonst aus dem Blatt. Grund: Beim Start
' wird vor dem Laden des Blatts "Einstellungen" migriert und gesichert. Das Blatt
' enthält dann noch die Werte der ausgelieferten Programmmappe (z. B. einen
' Backup-Ordner aus der Entwicklungsumgebung), nicht die des Kunden.
Private Function SicherungsEinstellung(ByVal schluessel As String) As String
    Dim db As Object, rs As Object, v As String, gefunden As Boolean
    On Error GoTo Fehler
    Set db = DbOeffnen()
    Set rs = db.OpenRecordset("SELECT [Wert], [Arbeitsplatz] FROM Einstellungen WHERE [Einstellung] = '" & _
                              Replace(schluessel, "'", "''") & "' AND ([Arbeitsplatz] = '*' OR [Arbeitsplatz] = '" & _
                              Replace(Rechner(), "'", "''") & "')")
    Do While Not rs.EOF
        If Not gefunden Or StrComp(CStr(rs.Fields("Arbeitsplatz").Value), "*") <> 0 Then
            v = Trim$(CStr(Nz0(rs.Fields("Wert").Value)))
            gefunden = True
        End If
        rs.MoveNext
    Loop
    rs.Close
    db.Close
    Set db = Nothing
    If gefunden Then SicherungsEinstellung = v: Exit Function
Blatt:
    SicherungsEinstellung = Trim$(EinstellungOderLeer(schluessel))
    Exit Function
Fehler:
    Resume Weiter
Weiter:
    On Error Resume Next
    If Not db Is Nothing Then db.Close
    GoTo Blatt
End Function

' Kopie des Backends in "ordner" anlegen. Rückgabe: Zieldatei, leer bei Fehler (Text in fehlerText)
Private Function SicherungKopieren(ByVal ordner As String, ByVal zusatz As String, ByRef fehlerText As String) As String
    Dim ziel As String
    If Not OrdnerAnlegen(ordner) Then
        fehlerText = "Der Sicherungsordner ist nicht erreichbar oder kann nicht angelegt werden: " & ordner
        Exit Function
    End If
    ziel = ordner & "\" & DateiStamm(BackendPfad()) & "_" & Format$(Now, "yyyy-mm-dd_hhnnss") & _
           IIf(Len(zusatz) > 0, "_" & zusatz, "") & ".accdb"
    ' FileSystemObject statt FileCopy: kopiert auch, wenn andere Plätze die Datei geöffnet haben
    On Error Resume Next
    CreateObject("Scripting.FileSystemObject").CopyFile BackendPfad(), ziel, True
    If Err.Number <> 0 Then
        fehlerText = "Die Sicherungskopie konnte nicht geschrieben werden (" & Err.Description & "): " & ziel
        Exit Function
    End If
    SicherungKopieren = ziel
End Function

' Ordner samt fehlender übergeordneter Ordner anlegen (MkDir legt nur eine Ebene an).
' False, wenn das Laufwerk bzw. die Freigabe nicht erreichbar ist oder Rechte fehlen.
Private Function OrdnerAnlegen(ByVal ordner As String) As Boolean
    Dim fso As Object, eltern As String
    If Len(ordner) = 0 Then Exit Function
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(ordner) Then OrdnerAnlegen = True: Exit Function
    eltern = fso.GetParentFolderName(ordner)
    If Len(eltern) = 0 Or StrComp(eltern, ordner, vbTextCompare) = 0 Then Exit Function
    If Not OrdnerAnlegen(eltern) Then Exit Function
    On Error Resume Next
    fso.CreateFolder ordner
    OrdnerAnlegen = fso.FolderExists(ordner)
End Function

Private Sub AlteSicherungenLoeschen(ByVal ordner As String, ByVal stamm As String)
    Dim behalten As Long, namen() As String, n As Long, f As String, i As Long, j As Long, t As String
    behalten = CLng(Val(SicherungsEinstellung("Backup behalten")))
    If behalten <= 0 Then behalten = 30
    ReDim namen(0 To 0)
    f = Dir$(ordner & "\" & stamm & "_????-??-??_??????.accdb")      ' nur Tagessicherungen
    Do While Len(f) > 0
        ReDim Preserve namen(0 To n)
        namen(n) = f
        n = n + 1
        f = Dir$()
    Loop
    If n <= behalten Then Exit Sub
    For i = 0 To n - 2                       ' Namen enthalten Datum/Zeit: alphabetisch = zeitlich
        For j = i + 1 To n - 1
            If namen(j) < namen(i) Then t = namen(i): namen(i) = namen(j): namen(j) = t
        Next j
    Next i
    On Error Resume Next
    For i = 0 To n - behalten - 1
        Kill ordner & "\" & namen(i)
    Next i
End Sub

'------------------------------------------------------------------------------
' Blätter laden
'------------------------------------------------------------------------------
' Alle Stammdatenblätter aus dem Backend laden (vor dem Einlesen einer Rechnung)
Public Sub StammdatenLaden()
    Dim t As Variant
    If Not DatenAktiv() Then Exit Sub
    SchemaSicherstellen                 ' z. B. nach Modulimport ohne Neustart
    NachUpdateErgaenzen                 ' neue Einstellungen der Programmversion ins Backend
    For Each t In Tabellen()
        BlattLaden CStr(t)
    Next t
End Sub

Public Sub BlattLaden(ByVal blatt As String)
    Dim ws As Worksheet, db As Object, rs As Object, f As Variant, n As Long
    Dim daten() As Variant, meta() As Variant, z As Long, s As Long, alt As Long, letzte As Long, ereignisse As Boolean, m As Long
    If Not DatenAktiv() Then Exit Sub
    Set ws = DatenBlatt(blatt)
    If ws Is Nothing Then Exit Sub
    f = Felder(blatt)
    n = UBound(f) + 1
    m = MetaSpalte(blatt) - 1               ' Verwaltungsspalten m+1 .. m+4

    Set db = DbOeffnen()
    Set rs = db.OpenRecordset("SELECT * FROM [" & blatt & "] ORDER BY [ID]", DB_OPEN_DYNASET)
    For s = 1 To n                          ' fehlt ein Feld, ist die Schema-Umstellung nicht gelaufen
        If Not FeldVorhanden(rs, FeldName(f(s - 1))) Then
            rs.Close
            db.Close
            Err.Raise vbObjectError + 65, , "Im Backend fehlt das Feld '" & FeldName(f(s - 1)) & "' in der Tabelle " & blatt & _
                      ". Die Umstellung des Backends auf Schema-Version " & SCHEMA_VERSION & " ist noch nicht gelaufen. " & _
                      "Bitte alle PDF2E-Plätze schließen und PDF2E neu starten; die Meldung beim Start nennt die Ursache."
        End If
    Next s
    Dim zeilen As Collection, zeile As Variant
    Set zeilen = New Collection
    Do While Not rs.EOF
        zeile = Empty
        ReDim zeile(1 To n + 4)
        For s = 1 To n
            zeile(s) = ZellWert(rs.Fields(FeldName(f(s - 1))).Value, FeldTyp(f(s - 1)))
        Next s
        zeile(n + 1) = Nz0(rs.Fields("GeaendertAm").Value)
        zeile(n + 2) = Nz0(rs.Fields("GeaendertVon").Value)
        zeile(n + 3) = rs.Fields("ID").Value
        zeile(n + 4) = Nz0(rs.Fields("Rev").Value)
        zeilen.Add zeile
        rs.MoveNext
    Loop
    rs.Close
    db.Close

    ereignisse = Application.EnableEvents
    Application.EnableEvents = False
    mIntern = True
    On Error GoTo Ende
    ' Ältere Mappen: neue Datenspalten vor den Verwaltungsspalten einfügen (z. B. ZR-Nummer ab Schema 2,
    ' Warenempfänger-GLN ab Schema 4); auch mehrere auf einmal, wenn Versionen übersprungen wurden
    For s = 2 To n
        If CStr(ws.Cells(1, s).Value) = "Geändert am" Then
            ws.Range(ws.Cells(1, s), ws.Cells(1, n)).EntireColumn.Insert
            Exit For
        End If
    Next s
    VerwaltungsspaltenSicherstellen ws, m
    For s = 2 To n                              ' ältere Mappen: fehlende Spaltenüberschrift ergänzen
        If Len(CStr(ws.Cells(1, s).Value)) = 0 Then
            ws.Cells(1, s).Value = FeldName(f(s - 1))
            ws.Cells(1, s).Font.Bold = True
            ws.Cells(1, s).Font.Color = ws.Cells(1, 1).Font.Color
            ws.Cells(1, s).Interior.Color = ws.Cells(1, 1).Interior.Color
            ws.Columns(s).ColumnWidth = 16
        End If
    Next s
    alt = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    letzte = ws.Cells(ws.Rows.Count, m + 3).End(xlUp).Row
    If letzte > alt Then alt = letzte
    If alt >= 2 Then
        ws.Range(ws.Cells(2, 1), ws.Cells(alt, n)).ClearContents
        ws.Range(ws.Cells(2, m + 1), ws.Cells(alt, m + 4)).ClearContents
    End If
    If zeilen.Count > 0 Then
        ReDim daten(1 To zeilen.Count, 1 To n)
        ReDim meta(1 To zeilen.Count, 1 To 4)
        For z = 1 To zeilen.Count
            zeile = zeilen(z)
            For s = 1 To n
                daten(z, s) = zeile(s)
            Next s
            For s = 1 To 4
                meta(z, s) = zeile(n + s)
            Next s
        Next z
        ' Textspalten als Text formatieren (Kundennr., IBAN, GTIN, PLZ ...)
        For s = 1 To n
            If FeldTyp(f(s - 1)) = "T" Then ws.Range(ws.Cells(2, s), ws.Cells(zeilen.Count + 1, s)).NumberFormat = "@"
        Next s
        ws.Range(ws.Cells(2, 1), ws.Cells(zeilen.Count + 1, n)).Value = daten
        ws.Range(ws.Cells(2, m + 1), ws.Cells(zeilen.Count + 1, m + 4)).Value = meta
    End If
Ende:
    mIntern = False
    Application.EnableEvents = ereignisse
    If Err.Number <> 0 Then Err.Raise Err.Number, , Err.Description
End Sub

Private Sub VerwaltungsspaltenSicherstellen(ByVal ws As Worksheet, ByVal m As Long)
    If CStr(ws.Cells(1, m + 3).Value) = "ID" Then Exit Sub
    ws.Cells(1, m + 1).Value = "Geändert am"
    ws.Cells(1, m + 2).Value = "Geändert von"
    ws.Cells(1, m + 3).Value = "ID"
    ws.Cells(1, m + 4).Value = "Rev"
    With ws.Range(ws.Cells(1, m + 1), ws.Cells(1, m + 4))
        .Font.Bold = True
        .Font.Color = ws.Cells(1, 1).Font.Color
        .Interior.Color = ws.Cells(1, 1).Interior.Color
    End With
    ws.Columns(m + 1).ColumnWidth = 17
    ws.Columns(m + 1).NumberFormat = "dd.mm.yyyy hh:mm"
    ws.Columns(m + 2).ColumnWidth = 14
    ws.Columns(m + 3).Hidden = True
    ws.Columns(m + 4).Hidden = True
End Sub

'------------------------------------------------------------------------------
' Ereignisse (aus DieseArbeitsmappe)
'------------------------------------------------------------------------------
Public Sub DatenBlattAktiviert(ByVal sh As Object)
    If Not DatenAktiv() Then Exit Sub
    If Not IstDatenblatt(sh.Name) Then
        KontextmenueEntfernen
        Exit Sub
    End If
    On Error GoTo Fehler
    SchemaSicherstellen                 ' Umstellung nachholen, falls sie beim Start gescheitert ist
    BlattLaden sh.Name
    KontextmenueAnlegen
    Exit Sub
Fehler:
    MsgBox "Die Daten konnten nicht aus dem Backend geladen werden:" & vbLf & Err.Description & vbLf & vbLf & _
           "Angezeigt wird der zuletzt geladene Stand. Änderungen werden derzeit nicht gespeichert.", vbExclamation, TOOL_NAME
End Sub

Public Sub DatenBlattGeaendert(ByVal sh As Object, ByVal target As Range)
    Dim n As Long, bereich As Range, z As Variant, zeilen As Collection, r As Range, a As Range
    If mIntern Then Exit Sub
    UStIdEingabePruefen sh, target              ' Blatt Empfaenger: USt-IdNr. bereinigen und prüfen
    If Not IstDatenblatt(sh.Name) Then Exit Sub
    If Not DatenAktiv() Then Exit Sub
    n = UBound(Felder(sh.Name)) + 1

    ' Ganze Zeilen oder Spalten (Zeile gelöscht/eingefügt, Spalte geleert): neu laden
    If target.Columns.Count >= sh.Columns.Count Or target.Rows.Count >= sh.Rows.Count Then
        NeuLadenMitHinweis sh.Name, "Zeilen und Spalten bitte nicht mit Excel löschen oder einfügen." & vbLf & _
            "Neue Datensätze unten anfügen, löschen über Rechtsklick > '" & DATENSATZ_MENUE & "'."
        Exit Sub
    End If
    Set bereich = Intersect(target, sh.Range(sh.Cells(2, 1), sh.Cells(sh.Rows.Count, n)))
    If bereich Is Nothing Then Exit Sub

    Set zeilen = New Collection
    On Error Resume Next
    For Each a In bereich.Areas                 ' Rows deckt nur den ersten Bereich ab
        For Each r In a.Rows
            zeilen.Add r.Row, CStr(r.Row)       ' doppelte Zeilen ergeben einen Fehler und fallen weg
        Next r
    Next a
    On Error GoTo 0
    For Each z In zeilen
        If Not ZeileSpeichern(sh.Name, CLng(z)) Then Exit For
    Next z
End Sub

Private Sub NeuLadenMitHinweis(ByVal blatt As String, ByVal hinweis As String)
    On Error Resume Next
    BlattLaden blatt
    MsgBox hinweis, vbInformation, TOOL_NAME
End Sub

'------------------------------------------------------------------------------
' Zeile speichern (Einfügen oder Ändern). Rückgabe False = nicht gespeichert,
' Blatt wurde neu geladen.
' Auch für Neuanlagen durch das Programm (EmpfaengerAnlegen usw.).
'------------------------------------------------------------------------------
Public Function ZeileSpeichern(ByVal blatt As String, ByVal z As Long) As Boolean
    Dim ws As Worksheet, f As Variant, n As Long, schluessel As String, id As Long, rev As Long
    Dim db As Object, rs As Object, s As Long, info As String, m As Long
    If Not DatenAktiv() Then ZeileSpeichern = True: Exit Function
    Set ws = DatenBlatt(blatt)
    f = Felder(blatt)
    n = UBound(f) + 1
    m = MetaSpalte(blatt) - 1
    schluessel = Trim$(CStr(ws.Cells(z, 1).Value))
    id = CLng(Val(CStr(ws.Cells(z, m + 3).Value)))
    rev = CLng(Val(CStr(ws.Cells(z, m + 4).Value)))

    If id = 0 And Len(schluessel) = 0 Then ZeileSpeichern = True: Exit Function   ' leere Zeile
    If id > 0 And Len(schluessel) = 0 Then
        NeuLadenMitHinweis blatt, "Die erste Spalte darf nicht leer sein. Die Änderung wurde nicht gespeichert." & vbLf & _
            "Löschen über Rechtsklick > '" & DATENSATZ_MENUE & "'."
        Exit Function
    End If

    On Error GoTo Fehler
    Set db = DbOeffnen()
    If id = 0 Then
        Set rs = db.OpenRecordset(blatt, DB_OPEN_DYNASET)
        rs.AddNew
        For s = 1 To n
            rs.Fields(FeldName(f(s - 1))).Value = DbWert(ws.Cells(z, s).Value, FeldTyp(f(s - 1)))
        Next s
        rs.Fields("Rev").Value = 1
        rs.Fields("GeaendertAm").Value = Now
        rs.Fields("GeaendertVon").Value = Benutzer()
        rs.Update
        rs.Bookmark = rs.LastModified
        id = rs.Fields("ID").Value
        rs.Close
        rev = 1
    Else
        ' Über das Recordset statt UPDATE mit Parametern (Parameter lieferten im Test
        ' für ein Pflichtfeld Null). Edit sperrt den Datensatz (pessimistisch);
        ' danach wird der Änderungszähler geprüft.
        Set rs = db.OpenRecordset("SELECT * FROM [" & blatt & "] WHERE [ID] = " & id, DB_OPEN_DYNASET)
        If Not rs.EOF Then
            rs.Edit
            If CLng(Val(CStr(Nz0(rs.Fields("Rev").Value)))) <> rev Then
                rs.CancelUpdate
                rs.Close
                Set rs = Nothing
            End If
        Else
            rs.Close
            Set rs = Nothing
        End If
        If rs Is Nothing Then
            info = AenderungsInfo(db, blatt, id)
            db.Close
            Set db = Nothing
            NeuLadenMitHinweis blatt, "Die Änderung wurde NICHT gespeichert." & vbLf & vbLf & info & vbLf & vbLf & _
                "Das Blatt zeigt jetzt den aktuellen Stand. Bitte die Änderung bei Bedarf erneut eingeben."
            Exit Function
        End If
        For s = 1 To n
            rs.Fields(FeldName(f(s - 1))).Value = DbWert(ws.Cells(z, s).Value, FeldTyp(f(s - 1)))
        Next s
        rs.Fields("Rev").Value = rev + 1
        rs.Fields("GeaendertAm").Value = Now
        rs.Fields("GeaendertVon").Value = Benutzer()
        rs.Update
        rs.Close
        rev = rev + 1
    End If
    db.Close
    Set db = Nothing
    VerwaltungSchreiben ws, z, m, id, rev
    ZeileSpeichern = True
    Exit Function

Fehler:
    Dim nr As Long, txt As String
    nr = Err.Number: txt = Err.Description
    On Error Resume Next
    If Not db Is Nothing Then db.Close
    If nr = 3022 Then
        If blatt = "Einstellungen" Then
            txt = "Die Einstellung '" & schluessel & "' gibt es für diesen Arbeitsplatz (Spalte Arbeitsplatz, leer = alle) schon."
        Else
            txt = "'" & schluessel & "' gibt es schon (erste Spalte muss eindeutig sein)."
        End If
    End If
    Protokoll "Fehler", "Speichern " & blatt & " Zeile " & z & ": " & txt
    NeuLadenMitHinweis blatt, "Die Änderung wurde NICHT gespeichert:" & vbLf & txt
End Function

' Wer hat den Datensatz zuletzt geändert (oder gelöscht)?
Private Function AenderungsInfo(ByVal db As Object, ByVal blatt As String, ByVal id As Long) As String
    Dim rs As Object
    Set rs = db.OpenRecordset("SELECT [GeaendertAm], [GeaendertVon] FROM [" & blatt & "] WHERE [ID] = " & id)
    If rs.EOF Then
        AenderungsInfo = "Der Datensatz wurde inzwischen an einem anderen Platz gelöscht."
    Else
        AenderungsInfo = "Der Datensatz wurde inzwischen von " & Nz0(rs.Fields("GeaendertVon").Value) & _
                         " am " & Format$(rs.Fields("GeaendertAm").Value, "dd.mm.yyyy hh:nn:ss") & " geändert."
    End If
    rs.Close
End Function

Private Sub VerwaltungSchreiben(ByVal ws As Worksheet, ByVal z As Long, ByVal m As Long, ByVal id As Long, ByVal rev As Long)
    Dim ereignisse As Boolean
    ereignisse = Application.EnableEvents
    Application.EnableEvents = False
    mIntern = True
    VerwaltungsspaltenSicherstellen ws, m
    ws.Cells(z, m + 1).Value = Now
    ws.Cells(z, m + 2).Value = Benutzer()
    ws.Cells(z, m + 3).Value = id
    ws.Cells(z, m + 4).Value = rev
    mIntern = False
    Application.EnableEvents = ereignisse
End Sub

'------------------------------------------------------------------------------
' Datensatz löschen (Kontextmenü der Zelle)
'------------------------------------------------------------------------------
Public Sub DatensatzLoeschen()
    Dim ws As Worksheet, n As Long, z As Long, id As Long, rev As Long, db As Object, info As String
    Set ws = ActiveSheet
    If Not IstDatenblatt(ws.Name) Then Exit Sub
    n = MetaSpalte(ws.Name) - 1
    z = ActiveCell.Row
    If z < 2 Then Exit Sub
    id = CLng(Val(CStr(ws.Cells(z, n + 3).Value)))
    rev = CLng(Val(CStr(ws.Cells(z, n + 4).Value)))
    If id = 0 Then
        MsgBox "In dieser Zeile steht kein gespeicherter Datensatz.", vbInformation, TOOL_NAME
        Exit Sub
    End If
    If MsgBox("Datensatz '" & ws.Cells(z, 1).Value & "' im Blatt '" & ws.Name & "' für alle Plätze löschen?", _
              vbYesNo + vbQuestion + vbDefaultButton2, TOOL_NAME) <> vbYes Then Exit Sub
    On Error GoTo Fehler
    Set db = DbOeffnen()
    db.Execute "DELETE FROM [" & ws.Name & "] WHERE [ID] = " & id & " AND [Rev] = " & rev, DB_FAIL_ON_ERROR
    If db.RecordsAffected = 0 Then
        info = AenderungsInfo(db, ws.Name, id)
        db.Close
        NeuLadenMitHinweis ws.Name, "Nicht gelöscht. " & info
        Exit Sub
    End If
    db.Close
    Protokoll "Info", "Datensatz '" & ws.Cells(z, 1).Value & "' im Blatt '" & ws.Name & "' gelöscht."
    BlattLaden ws.Name
    Exit Sub
Fehler:
    If Not db Is Nothing Then db.Close
    MsgBox "Löschen fehlgeschlagen:" & vbLf & Err.Description, vbExclamation, TOOL_NAME
End Sub

Private Sub KontextmenueAnlegen()
    Dim c As Object
    KontextmenueEntfernen
    On Error Resume Next
    Set c = Application.CommandBars("Cell").Controls.Add(Type:=1, Temporary:=True)   ' 1 = msoControlButton
    c.Caption = DATENSATZ_MENUE
    c.OnAction = "'" & ThisWorkbook.Name & "'!DatensatzLoeschen"
    c.BeginGroup = True
    c.Tag = "PDF2E_Datensatz"
End Sub

' Beim Wechsel zurück in diese Mappe (Workbook_Activate)
Public Sub DatenMappeAktiviert()
    If Not DatenAktiv() Then Exit Sub
    If IstDatenblatt(ActiveSheet.Name) Then KontextmenueAnlegen
End Sub

Public Sub KontextmenueEntfernen()
    Dim c As Object
    On Error Resume Next
    Do
        Set c = Nothing
        Set c = Application.CommandBars("Cell").FindControl(Tag:="PDF2E_Datensatz")
        If c Is Nothing Then Exit Do
        c.Delete
    Loop
End Sub

'------------------------------------------------------------------------------
' Verarbeitungsprotokoll (Tabelle Verarbeitung)
'------------------------------------------------------------------------------
' Text "erzeugt am ... von ... (Rechner)" oder leer, wenn unbekannt
Public Function VerarbeitungInfo(ByVal rechNr As String) As String
    Dim db As Object, rs As Object
    If Not DatenAktiv() Or Len(rechNr) = 0 Then Exit Function
    On Error GoTo Fehler
    Set db = DbOeffnen()
    Set rs = db.OpenRecordset("SELECT * FROM Verarbeitung WHERE RechNr = '" & Replace(rechNr, "'", "''") & "'")
    If Not rs.EOF Then
        VerarbeitungInfo = "Rechnung " & rechNr & " wurde am " & Format$(rs.Fields("ErzeugtAm").Value, "dd.mm.yyyy hh:nn") & _
            " von " & Nz0(rs.Fields("ErzeugtVon").Value) & " (Platz " & Nz0(rs.Fields("Rechner").Value) & ") erzeugt"
        If Not IsNull(rs.Fields("MailAm").Value) Then
            VerarbeitungInfo = VerarbeitungInfo & ", E-Mail erstellt am " & Format$(rs.Fields("MailAm").Value, "dd.mm.yyyy hh:nn") & _
                " von " & Nz0(rs.Fields("MailVon").Value)
        End If
        VerarbeitungInfo = VerarbeitungInfo & "."
    End If
    rs.Close
    db.Close
    Exit Function
Fehler:
    If Not db Is Nothing Then db.Close
    Protokoll "Warnung", "Verarbeitungsprotokoll nicht lesbar: " & Err.Description
End Function

' Nach dem Erzeugen: Datensatz anlegen oder aktualisieren
Public Sub VerarbeitungErzeugt(ByVal rechNr As String, ByVal quelle As String, ByVal ziel As String, _
                               ByVal empfaenger As String, ByVal brutto As Currency, ByVal status As String)
    Dim db As Object, rs As Object
    If Not DatenAktiv() Then Exit Sub
    On Error GoTo Fehler
    Set db = DbOeffnen()
    Set rs = db.OpenRecordset("SELECT * FROM Verarbeitung WHERE RechNr = '" & Replace(rechNr, "'", "''") & "'", DB_OPEN_DYNASET)
    If rs.EOF Then
        rs.AddNew
        rs.Fields("RechNr").Value = rechNr
    Else
        rs.Edit
        rs.Fields("MailAm").Value = Null
        rs.Fields("MailVon").Value = Null
    End If
    rs.Fields("Quelldatei").Value = Left$(quelle, 255)
    rs.Fields("Ausgabedatei").Value = Left$(ziel, 255)
    rs.Fields("Empfaenger").Value = Left$(empfaenger, 255)
    rs.Fields("Brutto").Value = brutto
    rs.Fields("Status").Value = status
    rs.Fields("ErzeugtAm").Value = Now
    rs.Fields("ErzeugtVon").Value = Benutzer()
    rs.Fields("Rechner").Value = Rechner()
    rs.Update
    rs.Close
    db.Close
    Exit Sub
Fehler:
    If Not db Is Nothing Then db.Close
    Protokoll "Warnung", "Verarbeitungsprotokoll nicht geschrieben: " & Err.Description
End Sub

' Nach dem E-Mail-Entwurf
Public Sub VerarbeitungMail(ByVal rechNr As String)
    Dim db As Object, qd As Object
    If Not DatenAktiv() Or Len(rechNr) = 0 Then Exit Sub
    On Error GoTo Fehler
    Set db = DbOeffnen()
    Set qd = db.CreateQueryDef("", "UPDATE Verarbeitung SET MailAm = [pAm], MailVon = [pVon] WHERE RechNr = [pNr]")
    qd.Parameters("pAm").Value = Now
    qd.Parameters("pVon").Value = Benutzer()
    qd.Parameters("pNr").Value = rechNr
    qd.Execute DB_FAIL_ON_ERROR
    db.Close
    Exit Sub
Fehler:
    If Not db Is Nothing Then db.Close
    Protokoll "Warnung", "Verarbeitungsprotokoll (Mail) nicht geschrieben: " & Err.Description
End Sub

'------------------------------------------------------------------------------
' Import von Stammdaten aus einer Arbeitsmappe
'------------------------------------------------------------------------------
' Makro: Stammdaten einer älteren PDF2E-Mappe übernehmen. Nur fehlende
' Schlüssel werden angelegt, vorhandene Datensätze bleiben unverändert.
Public Sub StammdatenAusMappeUebernehmen()
    Dim pfad As Variant, wb As Workbook, ereignisse As Boolean
    If Not DatenAktiv() Then
        MsgBox "Kein Backend eingestellt (Einstellung 'Backend-Pfad').", vbExclamation, TOOL_NAME
        Exit Sub
    End If
    pfad = Application.GetOpenFilename("Excel-Arbeitsmappen (*.xlsm;*.xlsx),*.xlsm;*.xlsx", , "Mappe mit den Stammdaten auswählen")
    If VarType(pfad) = vbBoolean Then Exit Sub
    If StrComp(CStr(pfad), ThisWorkbook.FullName, vbTextCompare) = 0 Then
        MsgBox "Bitte eine andere Mappe als diese auswählen.", vbExclamation, TOOL_NAME
        Exit Sub
    End If
    On Error GoTo Fehler
    ereignisse = Application.EnableEvents
    Application.EnableEvents = False
    Set wb = Workbooks.Open(CStr(pfad), UpdateLinks:=0, ReadOnly:=True)
    Application.EnableEvents = ereignisse
    ImportAusMappe wb, False, Array("Brauerei", "Bankkonten", "Zahlungskonditionen", "Empfaenger", "Pfandartikel", BLATT_LIEFERORTE)
    wb.Close False
    StammdatenLaden
    MsgBox "Übernahme abgeschlossen. Einzelheiten im Blatt 'Protokoll'.", vbInformation, TOOL_NAME
    Exit Sub
Fehler:
    Application.EnableEvents = ereignisse
    If Not wb Is Nothing Then wb.Close False
    MsgBox "Übernahme fehlgeschlagen:" & vbLf & Err.Description, vbExclamation, TOOL_NAME
End Sub

' Blätter einer Mappe ins Backend übernehmen (nur fehlende Schlüssel).
' erstbefuellung = True: aus dieser Mappe beim Anlegen des Backends.
Private Sub ImportAusMappe(ByVal wb As Workbook, ByVal erstbefuellung As Boolean, ByVal liste As Variant)
    Dim t As Variant, ws As Worksheet, f As Variant, n As Long, z As Long, s As Long
    Dim db As Object, rs As Object, schluessel As String, neu As Long, da As Long, krit As String, nQ As Long
    Set db = DbOeffnen()
    For Each t In liste
        Set ws = Nothing
        On Error Resume Next
        Set ws = wb.Worksheets(CStr(t))
        On Error GoTo 0
        If Not ws Is Nothing Then
            f = Felder(CStr(t))
            n = UBound(f) + 1
            neu = 0: da = 0
            ' Ältere Mappe mit weniger Datenspalten: nur bis vor die Verwaltungsspalten lesen
            nQ = n
            For s = 2 To n
                If CStr(ws.Cells(1, s).Value) = "Geändert am" Then nQ = s - 1: Exit For
            Next s
            Set rs = db.OpenRecordset(CStr(t), DB_OPEN_DYNASET)
            For z = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
                schluessel = Trim$(CStr(ws.Cells(z, 1).Value))
                If IstPlatzEinstellung(CStr(t), schluessel) Then schluessel = ""     ' gehört nicht ins Backend
                If Len(schluessel) > 0 Then
                    krit = "[" & FeldName(f(0)) & "] = '" & Replace(schluessel, "'", "''") & "'"
                    If CStr(t) = "Einstellungen" Then krit = krit & " AND [Arbeitsplatz] = '" & _
                        Replace(CStr(DbWert(ws.Cells(z, 4).Value, "A")), "'", "''") & "'"
                    rs.FindFirst krit
                    If rs.NoMatch Then
                        rs.AddNew
                        For s = 1 To nQ
                            rs.Fields(FeldName(f(s - 1))).Value = DbWert(ws.Cells(z, s).Value, FeldTyp(f(s - 1)))
                        Next s
                        rs.Fields("Rev").Value = 1
                        rs.Fields("GeaendertAm").Value = Now
                        rs.Fields("GeaendertVon").Value = Benutzer()
                        rs.Update
                        neu = neu + 1
                    Else
                        da = da + 1
                    End If
                End If
            Next z
            rs.Close
            If Not erstbefuellung And neu > 0 Then
                Protokoll "Info", "Übernahme " & t & ": " & neu & " neu, " & da & " bereits vorhanden (unverändert)."
            End If
        End If
    Next t
    db.Close
End Sub

'------------------------------------------------------------------------------
' Hilfen
'------------------------------------------------------------------------------
' Platzbezogene Einstellungen (Blatt Arbeitsplatz, vom Starter gesetzt) nie ins Backend
Private Function IstPlatzEinstellung(ByVal tabelle As String, ByVal schluessel As String) As Boolean
    If tabelle <> "Einstellungen" Then Exit Function
    IstPlatzEinstellung = (schluessel = "Backend-Pfad") Or (PlatzWert(schluessel) <> vbNullChar)
End Function

Private Function DatenBlatt(ByVal name As String) As Worksheet
    On Error Resume Next
    Set DatenBlatt = ThisWorkbook.Worksheets(name)
    If DatenBlatt Is Nothing And name = BLATT_PFANDARTIKEL Then
        On Error GoTo 0
        PfandartikelSicherstellen
        Set DatenBlatt = ThisWorkbook.Worksheets(name)
    ElseIf DatenBlatt Is Nothing And name = BLATT_LIEFERORTE Then
        On Error GoTo 0
        LieferorteSicherstellen
        Set DatenBlatt = ThisWorkbook.Worksheets(name)
    End If
End Function

Private Function FeldVorhanden(ByVal rs As Object, ByVal name As String) As Boolean
    Dim fld As Object
    For Each fld In rs.Fields
        If StrComp(fld.Name, name, vbTextCompare) = 0 Then FeldVorhanden = True: Exit Function
    Next fld
End Function

Private Function OrdnerExistiert(ByVal ordner As String) As Boolean
    If Len(ordner) = 0 Then Exit Function
    OrdnerExistiert = CreateObject("Scripting.FileSystemObject").FolderExists(ordner)
End Function

' Wert aus der Zelle für die Datenbank: leer = Null
Private Function DbWert(ByVal v As Variant, ByVal typ As String) As Variant
    Dim s As String
    If IsError(v) Then DbWert = Null: Exit Function
    s = Trim$(CStr(v))
    If Len(s) = 0 Then
        If typ = "A" Then DbWert = "*" Else DbWert = Null
        Exit Function
    End If
    Select Case typ
        Case "N": If IsNumeric(v) Then DbWert = CDbl(v) Else DbWert = Null
        Case "T": DbWert = Left$(s, 255)
        Case "A": DbWert = Left$(s, 100)
        Case Else: DbWert = CStr(v)
    End Select
End Function

' Wert aus der Datenbank für die Zelle
Private Function ZellWert(ByVal v As Variant, ByVal typ As String) As Variant
    If IsNull(v) Then ZellWert = Empty: Exit Function
    If typ = "A" And v = "*" Then ZellWert = Empty: Exit Function
    ZellWert = v
End Function

Private Function Nz0(ByVal v As Variant) As Variant
    If IsNull(v) Then Nz0 = "" Else Nz0 = v
End Function

Private Function OrdnerVon(ByVal pfad As String) As String
    If InStrRev(pfad, "\") > 0 Then OrdnerVon = Left$(pfad, InStrRev(pfad, "\") - 1)
End Function

Private Function DateiStamm(ByVal pfad As String) As String
    Dim d As String
    d = Mid$(pfad, InStrRev(pfad, "\") + 1)
    If InStrRev(d, ".") > 0 Then d = Left$(d, InStrRev(d, ".") - 1)
    DateiStamm = d
End Function
