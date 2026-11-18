import sys
# Erzeugt die Starter-Mappe Start_PDF2E.xlsx bzw. Start_XL2E.xlsx (Blatt "Starter").
# Aufruf: python build_starter.py <Zieldatei.xlsx> [pdf2e|xl2e]   (benötigt openpyxl)
# Danach in Excel: Module aus src/starter importieren, als .xlsm speichern,
# einmalig das Makro StarterEinrichten ausführen (Schaltfläche "Jetzt starten").
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

out = sys.argv[1]
PROD = (sys.argv[2] if len(sys.argv) > 2 else "pdf2e").upper()
H = Font(bold=True, color="FFFFFF"); HF = PatternFill("solid", fgColor="7A4A12")
INP = PatternFill("solid", fgColor="FFF8E1")
HINW = Font(italic=True, color="777777")

wb = Workbook()
ws = wb.active
ws.title = "Starter"
ws.column_dimensions["A"].width = 34
ws.column_dimensions["B"].width = 55
ws.column_dimensions["C"].width = 80
ws["A1"] = f"Starter für {PROD}"
ws["A1"].font = Font(bold=True, size=18, color="7A4A12")
ws["A2"] = ("Startet das Programm, holt neue Versionen aus dem Update-Ordner und überträgt die Werte unten in die "
            "Programmmappe. Zum Bearbeiten dieser Mappe beim Öffnen Strg oder Umschalt gedrückt halten.")
ws["A2"].font = HINW


def kopf(z, titel):
    for j, t in enumerate([titel, "Wert", "Hinweis"], 1):
        c = ws.cell(z, j, t); c.font = H; c.fill = HF


def zeilen(z0, daten):
    for i, (a, b, c) in enumerate(daten):
        ws.cell(z0 + i, 1, a)
        x = ws.cell(z0 + i, 2, b); x.fill = INP; x.number_format = "@"
        ws.cell(z0 + i, 3, c).font = HINW
        ws.cell(z0 + i, 3).alignment = Alignment(wrap_text=True)


kopf(4, "Starter-Einstellung")
zeilen(5, [
    ("Produkt", PROD, "PDF2E oder XL2E"),
    ("Update-Ordner", rf"\\SERVER\{PROD}\Programm",
     f"Hier liegt die aktuelle {PROD}.xlsm (Mastermappe). Einzelplatz: z. B. C:\\{PROD}\\Update. Leer = keine Update-Prüfung"),
    ("Programmordner lokal", rf"C:\{PROD}",
     "Lokale Kopie der Programmmappe. Im Trust Center als vertrauenswürdigen Speicherort eintragen"),
])
# Bereich 2: Text in A9 muss genau so lauten (modStarter, BEREICH2); Liste endet an der ersten leeren Zeile
kopf(9, "Einstellung für die Programmmappe")
zeilen(10, [
    ("Backend-Pfad", rf"\\SERVER\{PROD}\Daten\{PROD}_Daten.accdb",
     "Access-Datei mit Einstellungen, Stammdaten und Verarbeitungsprotokoll; wird beim ersten Start angelegt. Leer = ohne Backend"),
] + ([("pdftotext-Pfad", r"C:\Tools\xpdf\pdftotext.exe", "Pfad zu pdftotext.exe auf diesem Platz (leer = Wert aus den Einstellungen)")] if PROD == "PDF2E" else []))
ws["A13"] = ("Diese Werte schreibt der Starter bei jedem Start in das Blatt 'Arbeitsplatz' der Programmmappe. "
             "Alle anderen Einstellungen (Ordner, Sicherung, Mailvorlage ...) pflegst du in der Programmmappe; sie liegen im Backend "
             "und gelten für alle Plätze.")
ws["A13"].font = HINW
ws.freeze_panes = "A4"
wb.save(out)
