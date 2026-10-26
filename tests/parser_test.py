"""Regressionstest für die Einleselogik von modParserBSI.

Python-Nachbau der Suchmuster und Abläufe aus src/modParserBSI.bas. Er liest
die pdftotext-Ausgaben in tests/pdftotext/ und prüft, ob Warenwert, Pfand,
Nettosumme, MwSt und Rechnungsbetrag mit den Werten im PDF übereinstimmen.

Aufruf:  python tests/parser_test.py [datei.txt ...]
Ohne Argumente werden alle Dateien in tests/pdftotext/ geprüft.
Bei Änderungen an den Suchmustern in VBA diese Datei mitziehen.
"""
import glob
import os
import re
import sys

RE_LS = re.compile(r"^\s*\d+\s+Lieferschein:\s+(\d+)\s+vom\s+(\d{1,2}\.\d{1,2}\.\d{4})")
RE_POS = re.compile(r"^\s*(\d+)\s+(.+?)\s+(\d+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+-?)\s*$")
RE_AKT = re.compile(r"^\s*Aktion\s+(.+?)\s+(\d+)\s+Ka/F\s+([\d.,]+)\s+([\d.,]+)-\s*$")
RE_LG = re.compile(r"^\s*(\d{5})\s+(\S+(?: \S+)*)\s{2,}(?:.*\s)?(-?\d+-?)\s+(\d+,\d{2})\s+([\d.]*\d,\d{2}-?)\s*$")
RE_WW = re.compile(r"WARENWERT:\s+([\d.,]+-?)")
RE_ZS = re.compile(r"Zwischensumme\s+\(incl\.\s+([\d.,]+-?)\s+Pfand\)\s+([\d.,]+-?)")
RE_ST = re.compile(r"([\d,]+)\s*%\s*MwSt\s+\(auf\s+([\d.,]+-?)\s*\)\s+([\d.,]+-?)")
RE_RB = re.compile(r"RECHNUNGSBETRAG\s+EUR\s+([\d.,]+-?)")
RE_ARTNR = re.compile(r"^\d{5}\s")


def zahl(s):
    s = s.strip()
    neg = s.endswith("-") or s.startswith("-")
    v = float(s.strip("-").replace(".", "").replace(",", "."))
    return -v if neg else v


def parse(text):
    r = {"ware": 0.0, "pfand": 0.0, "mwst": 0.0, "satz": None, "zeilenfehler": []}
    modus = "pos"
    lg_rest = ""

    def pfand(m):
        menge, satz, betrag = zahl(m[3]), zahl(m[4]), zahl(m[5])
        r["pfand"] += betrag
        if abs(round(menge * satz, 2) - betrag) > 0.005:
            r["zeilenfehler"].append(m[0].strip())

    for seite in text.split("\f"):
        zeilen = seite.split("\n")
        start = next((i + 1 for i, z in enumerate(zeilen) if "ARTNR" in z and "BEZEICHNUNG" in z), None)
        if start is None:
            continue
        for l in zeilen[start:]:
            s = l.strip()
            if not s:
                continue
            if s[0] in "─═":
                lg_rest = ""
                continue
            if s.startswith("ÜBERTRAG"):
                continue
            if s.startswith("LEERGUT-KONTO"):
                modus = "lg"
                continue
            if modus == "lg" and s.startswith("SALDO"):
                continue
            if RE_LS.match(l):
                continue
            m = RE_AKT.match(l)
            if m:
                r["ware"] -= zahl(m[4])
                continue
            m = RE_WW.search(l)
            if m:
                r["pdf_warenwert"] = zahl(m[1])
                modus = "summen"
                continue
            m = RE_ZS.search(l)
            if m:
                r["pdf_pfand"], r["pdf_netto"] = zahl(m[1]), zahl(m[2])
                continue
            m = RE_ST.search(l)
            if m:
                r["satz"] = zahl(m[1])
                r["mwst"] += zahl(m[3])
                continue
            m = RE_RB.search(l)
            if m:
                r["pdf_brutto"] = zahl(m[1])
                continue
            if modus == "pos":
                m = RE_POS.match(l)
                if m:
                    menge, preis, betrag = zahl(m[3]), zahl(m[5]), zahl(m[6])
                    r["ware"] += betrag
                    if abs(round(menge * preis, 2) - betrag) > 0.005:
                        r["zeilenfehler"].append(s)
            elif modus == "lg":
                if lg_rest:
                    if not RE_ARTNR.match(s):
                        m = RE_LG.match(lg_rest + "   " + s)
                        lg_rest = ""
                        if m:
                            pfand(m)
                            continue
                    lg_rest = ""
                m = RE_LG.match(l)
                if m:
                    pfand(m)
                    continue
                if RE_ARTNR.match(s):
                    lg_rest = s
    return r


def pruefe(pfad):
    r = parse(open(pfad, encoding="utf-8").read())
    netto = round(r["ware"] + r["pfand"], 2)
    mwst = round(netto * r["satz"] / 100 + 1e-9, 2)
    checks = [
        ("Warenwert", r["ware"], r.get("pdf_warenwert")),
        ("Pfand", r["pfand"], r.get("pdf_pfand")),
        ("Netto", netto, r.get("pdf_netto")),
        ("MwSt", mwst, r["mwst"]),
        ("Brutto", netto + mwst, r.get("pdf_brutto")),
    ]
    fehler = [f"{n}: berechnet {a:.2f}, PDF {b}" for n, a, b in checks if b is None or abs(a - b) > 0.005]
    fehler += [f"Zeile Menge x Preis: {z}" for z in r["zeilenfehler"]]
    print(("OK    " if not fehler else "FEHLER") + "  " + os.path.basename(pfad))
    for f in fehler:
        print("        " + f)
    return not fehler


if __name__ == "__main__":
    dateien = sys.argv[1:] or sorted(glob.glob(os.path.join(os.path.dirname(__file__), "pdftotext", "*.txt")))
    ergebnisse = [pruefe(d) for d in dateien]
    sys.exit(0 if all(ergebnisse) else 1)
