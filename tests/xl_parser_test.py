"""Regressionstest für die Einleselogik von XL2E (src/xl2e/modParserXL.bas).

Python-Nachbau der Abläufe aus modParserXL. Er liest die Excel-Rechnungen in
tests/xlsx/ (Blatt PA400_00 aus der BSI-Software) und prüft, ob Warenwert,
Pfand, Nettosumme, MwSt und Rechnungsbetrag mit den Werten der Rechnung
übereinstimmen und ob jede Position Menge x Preis = Betrag erfüllt.

Aufruf:  python tests/xl_parser_test.py [datei.xlsx ...]   (benötigt openpyxl)
Bei Änderungen an modParserXL diese Datei mitziehen.
"""
import glob
import os
import re
import sys

import openpyxl

RE_LS = re.compile(r"Lieferschein:\s+(\d+)\s+vom\s+(\d{1,2}\.\d{1,2}\.\d{4})")
RE_INKL = re.compile(r"\(incl\.\s+([\d.,]+-?)\s+Pfand\)")
RE_SATZ = re.compile(r"([\d,]+)\s*%\s*MwSt")


def zahl(s):
    s = str(s).strip()
    if not s:
        return 0.0
    neg = s.endswith("-") or s.startswith("-")
    v = float(s.strip("-").replace(".", "").replace(",", "."))
    return -v if neg else v


def txt(ws, r, c):
    v = ws.cell(r, c).value
    return "" if v is None else str(v).strip()


def parse(pfad):
    ws = openpyxl.load_workbook(pfad).worksheets[0]
    r = {"pos": [], "pfand": [], "info": [], "mwst": 0.0, "kopf": {}}
    modus = "kopf"
    letzte = ws.max_row
    zeile = 1
    while zeile <= letzte:
        A, B, D, G, I, K, L = (txt(ws, zeile, c) for c in (1, 2, 4, 7, 9, 11, 12))
        if "Kun-Nr" in I:
            nr = txt(ws, zeile + 1, 9).split()
            r["kopf"].update(kundennr=nr[0], rechnr=nr[1], datum=txt(ws, zeile + 1, 11))
            zeile += 2
            continue
        if B == "Art-Nr":
            modus = "pos"
        elif B.startswith("LEERGUT-KONTO"):
            modus = "lg"
            zeile += 2          # zweite Überschriftenzeile überspringen
            continue
        elif B.startswith("─"):
            if modus == "lg":
                modus = "summen"
            elif modus == "lginfo":
                modus = "pos"
        elif D == "WARENWERT:":
            r["warenwert"] = zahl(L)
            modus = "nach"
        elif G == "Zwischensumme":
            r["pdf_pfand"] = zahl(RE_INKL.search(I)[1])
            r["netto"] = zahl(L)
        elif G.startswith("+") and "MwSt" in G:
            r["satz"] = zahl(RE_SATZ.search(G)[1])
            r["mwst"] += zahl(L)
        elif G == "RECHNUNGSBETRAG":
            r["brutto"] = zahl(L)
        elif modus in ("pos", "lginfo"):
            m = RE_LS.search(D)
            if m:
                r["kopf"].update(ls=m[1], lsdatum=m[2])
            elif D.startswith("Lieferort:"):
                r["kopf"]["lieferort"] = D[10:].strip()
            elif D == "Leergutrückgabe:":
                modus = "lginfo"
            elif modus == "pos" and K and L:
                ka, fl, liter = zahl(txt(ws, zeile, 8)), zahl(I), zahl(txt(ws, zeile, 10))
                preis, betrag = zahl(K), zahl(L)
                if betrag < 0 and not B:
                    r["pos"].append(("Nachlass", 1, preis, betrag))
                elif ka and abs(round(ka * preis, 2) - betrag) < 0.005:
                    r["pos"].append(("Ware", ka, preis, betrag))
                elif liter and abs(round(liter / 100 * preis, 2) - betrag) < 0.005:
                    r["pos"].append(("Ware hl", liter / 100, preis, betrag))
                else:
                    r["pos"].append(("Ware ?", ka, preis, betrag))
            elif B or D:
                r["info"].append((modus, B, D))
        elif modus == "lg":
            satz = zahl(K)
            if satz:
                lief, zur = zahl(G), zahl(txt(ws, zeile, 8))
                betrag = zahl(L)
                if abs(round((lief - zur) * satz, 2) - betrag) > 0.005:
                    r["info"].append(("FEHLER Pfand", B, betrag))
                for mg in (lief, -zur):
                    if mg:
                        r["pfand"].append((A, B, mg, satz, round(mg * satz, 2)))
        zeile += 1
    return r


def pruefe(pfad):
    r = parse(pfad)
    ware = sum(p[3] for p in r["pos"] if p[0] != "Nachlass")
    nachlass = sum(p[3] for p in r["pos"] if p[0] == "Nachlass")
    pfand = sum(p[4] for p in r["pfand"])
    netto = round(ware + nachlass + pfand, 2)
    mwst = round(netto * r["satz"] / 100 + 1e-9, 2)
    checks = [("Warenwert", ware + nachlass, r.get("warenwert")), ("Pfand", pfand, r.get("pdf_pfand")),
              ("Netto", netto, r.get("netto")), ("MwSt", mwst, r["mwst"]), ("Brutto", netto + mwst, r.get("brutto"))]
    fehler = [f"{n}: berechnet {a:.2f}, Rechnung {b}" for n, a, b in checks if b is None or abs(a - b) > 0.005]
    fehler += [f"Position ohne passende Menge: {p}" for p in r["pos"] if p[0].endswith("?")]
    fehler += [f"{i}" for i in r["info"] if i[0].startswith("FEHLER")]
    print(("OK    " if not fehler else "FEHLER") + "  " + os.path.basename(pfad), r["kopf"])
    for p in r["pos"]:
        print(f"        {p[0]:8} {p[1]:g} x {p[2]:.2f} = {p[3]:.2f}")
    for p in r["pfand"]:
        print(f"        Pfand {p[0]} {p[1]}: {p[2]:+g} x {p[3]:.2f} = {p[4]:.2f}")
    for f in fehler:
        print("        " + f)
    return not fehler


if __name__ == "__main__":
    dateien = sys.argv[1:] or sorted(glob.glob(os.path.join(os.path.dirname(__file__), "xlsx", "*.xlsx")))
    sys.exit(0 if all([pruefe(d) for d in dateien]) else 1)
