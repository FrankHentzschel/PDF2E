Attribute VB_Name = "modAdapterSmartTools"
'==============================================================================
' Adapter für die SmartTools E-Rechnung VBA-Bibliothek
'
' Die Befehle sind nur so weit ausgeschrieben, wie sie auf der Produktseite
' belegt sind. Die übrigen Zuweisungen stehen als Feldzuordnung bereit und
' werden nach Lieferung der Bibliothek anhand der Musterlösung ergänzt.
'==============================================================================
Option Explicit

Public Sub SmartToolsEinbetten(ByVal zielPdf As String)
    '--------------------------------------------------------------------------
    ' Feldzuordnung (Quelle -> E-Rechnung, BT-Nummern nach EN 16931)
    '
    ' Kopf
    '   BT-1   Rechnungsnummer        Feld("f_RechNr")
    '   BT-2   Rechnungsdatum         Feld("f_RechDatum")
    '   BT-3   Rechnungsart           Feld("f_RechArt")          (380)
    '   BT-5   Währung                Feld("f_Waehrung")         (EUR)
    '   BT-16  Lieferschein           Feld("f_LSNr")
    '   BT-72  Leistungsdatum         Feld("f_LSDatum")
    '   BT-9   Fälligkeitsdatum       Feld("f_Faellig")
    '   BT-20  Zahlungsbedingungen    Spalte B in "Zahlungskonditionen" zu Feld("f_Kondition")
    ' Verkäufer (Blatt "Brauerei")
    '   BT-27  Name                   Stamm("Name")
    '   BT-35  Straße                 Stamm("Straße")
    '   BT-38  PLZ                    Stamm("PLZ")
    '   BT-37  Ort                    Stamm("Ort")
    '   BT-40  Land                   Stamm("Land")
    '   BT-31  USt-IdNr               Stamm("USt-IdNr")
    '   BT-32  Steuernummer           Stamm("Steuernummer")
    '   BT-30  Registernummer         Stamm("Registernummer")
    '   BT-41..43 Kontakt             Stamm("Kontakt Name"), Stamm("Telefon"), Stamm("E-Mail")
    ' Käufer
    '   BT-44  Name                   Feld("f_KName")
    '   BT-46  Käufer-Kennung         Feld("f_KGLN") mit Schema 0088, falls vorhanden (z. B. REWE),
    '                                 sonst Feld("f_KundenNr")
    '   BT-48  USt-IdNr               Feld("f_KUstId")           (optional)
    '   BT-49  E-Mail                 Feld("f_KEmail")
    '   BT-50  Straße                 Feld("f_KStrasse")
    '   BT-53  PLZ / BT-52 Ort        Feld("f_KPLZ"), Feld("f_KOrt")
    '   BT-55  Land                   Feld("f_KLand")
    ' Lieferort
    '   BT-70  Name                   Feld("f_LName")
    '   BT-71  GLN (Schema 0088)      Feld("f_LGLN")
    '   BT-75..80 Anschrift           Feld("f_LZeile2"), Feld("f_LPLZ"), Feld("f_LOrt"), Feld("f_LLand")
    ' Zahlung (Blatt "Bankkonten", Zeile zu Feld("f_Konto"))
    '   BT-81  Zahlungsart            58 (SEPA-Überweisung)
    '   BT-84  IBAN                   Spalte C
    '   BT-85  Kontoinhaber           Spalte B
    '   BT-86  BIC                    Spalte D
    ' Positionen (Blatt "Positionen", Typ "Ware" und "Pfand")
    '   BT-126 Positionsnummer        laufend
    '   BT-155 Artikelnummer          Spalte B
    '   BT-153 Bezeichnung            Spalte C
    '   BT-129 Menge                  Spalte D
    '   BT-130 Einheit                Spalte E
    '   BT-146 Nettopreis             Spalte F
    '   BT-131 Nettobetrag            Spalte G
    '   BT-151 USt-Kategorie          Spalte I  (S)
    '   BT-152 USt-Satz               Spalte H
    ' Nachlässe (Typ "Nachlass") als Abschlag auf Belegebene (BG-20)
    '   BT-92  Betrag                 Abs(Spalte G)
    '   BT-97  Grund                  Spalte C
    '   BT-95  USt-Kategorie          Spalte I
    '   BT-96  USt-Satz               Spalte H
    '--------------------------------------------------------------------------

    ' Laut Produktseite belegte Aufrufe (Verweis auf die Smart-Tools-Bibliothek nötig):
    '   Set objInvoice = New SmartTools_Invoicing.XRechnungInvoice
    '   objInvoice.Seller.Name = Stamm("Name")
    '   objInvoice.Seller.PostalAddress.City = Stamm("Ort")
    '   arrInvoiceViolations = objInvoice.Validate()
    '   objZFRechnung.Save zielPdf          ' ZUGFeRD: XML einbetten, PDF/A-3
    '
    ' TODO: Klasse für ZUGFeRD, Profil EN16931 und die übrigen Eigenschaften
    '       aus der Smart-Tools-Musterlösung übernehmen.

    Err.Raise vbObjectError + 100, , "Smart-Tools-Anbindung ist noch nicht eingerichtet (modAdapterSmartTools). Bis dahin in den Einstellungen STUB verwenden.."
End Sub
