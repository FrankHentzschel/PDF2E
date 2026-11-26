"""Regressionstest für die Einleselogik von modParserBSI.

Python-Nachbau der Suchmuster und Abläufe aus src/pdf2e/modParserBSI.bas. Er liest
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
RE_POS = re.compile(r"^\s*(\d+)\s+(.+?)\s+(\d+)\s+(?:([\d.,]+)\s+)?([\d.,]+)\s+([\d.,]+-?)(?:\s+\*)?\s*$")
RE_AKT = re.compile(r"^\s*Aktion\s+(.+?)\s+(\d+)\s+Ka/F\s+([\d.,]+)\s+([\d.,]+)-\s*$")
RE_LG = re.compile(r"^\s*(\d{5})\s+(\S+(?: \S+)*)\s{2,}(?:(.*\S)\s+)?(-?\d+-?)\s+(\d+,\d{2})\s+([\d.]*\d,\d{2}-?)\s*$")
RE_WW = re.compile(r"WARENWERT:\s+([\d.,]+-?)")
RE_ZS = re.compile(r"Zwischensumme\s+\(incl\.\s+([\d.,]+-?)\s+Pfand\)\s+([\d.,]+-?)")
RE_ST = re.compile(r"([\d,]+)\s*%\s*MwSt\s+\(auf\s+([\d.,]+-?)\s*\)\s+([\d.,]+-?)")
RE_RB = re.compile(r"RECHNUNGSBETRAG\s+EUR\s+([\d.,]+-?)")
# Nachlass auf den Warenwert, z. B. Skonto: Art, Text, Basis, Prozent, Betrag
RE_NACHLASS = re.compile(r"^\s*(\S+)\s+(?:(.*?)\s+)?([\d.]*\d,\d{2})\s+EUR?\s+(\d+,\d{2})\s+([\d.]*\d,\d{2})-\s*$")
RE_WW_LEER = re.compile(r"WARENWERT:\s*$")
RE_IG = re.compile(r"Steuerfreie innergemeinschaftliche Lieferung")
RE_USTID = re.compile(r"^\s*USt-?ID(?:-?Nr\.?)?:?\s+([A-Za-z]{2}[\w .-]*\w)\s*$")
RE_NW = re.compile(r"NETTOWERT:\s+([\d.,]+-?)")
RE_SEPA_BETRAG = re.compile(r"Forderung von\s+([\d.,]+)\s+ziehen wir mit einer SEPA-Lastschrift")
RE_SEPA_MANDAT = re.compile(r"^\s*(\S+)\s+und mit der Gl(?:ae|ä)ubiger-ID\s+(\S+)\s+am\s+(\d{1,2}\.\d{1,2}\.\d{4})\s+ein")
RE_FILIALE = re.compile(r"Lieferfiliale:\s*(\d{5})\s+([^,]+?)\s*,\s*(.+?)(?:\s*\*)?\s*$")   # PLZ, Ort, Straße
RE_ARTNR = re.compile(r"^\d{5}\s")
RE_LGREST = re.compile(r"\d+,\d{2}\s+[\d.]*\d,\d{2}-?\s*$")
RE_LG0 = re.compile(r"^\s*(\d{5})\s+(\S+(?: \S+)*)\s{2,}(.*\S)\s+(\d+,\d{2})\s*$")   # ohne Pfandbetrag


def ueberlagern(a, b):
    """Legt zwei Zeilen spaltengenau übereinander; None, wenn sich Zeichen überschneiden."""
    n = max(len(a), len(b))
    a, b = a.ljust(n), b.ljust(n)
    if any(x != " " and y != " " for x, y in zip(a, b)):
        return None
    return "".join(x if x != " " else y for x, y in zip(a, b))


def zahl(s):
    s = s.strip()
    neg = s.endswith("-") or s.startswith("-")
    v = float(s.strip("-").replace(".", "").replace(",", "."))
    return -v if neg else v


def leergut_aufteilen(mitte, diff):
    """Lieferung und Rücknahme aus den Mengen vor MENGE DIFF. bestimmen.

    mitte: Zahlen zwischen Bezeichnung und MENGE DIFF., also eine Auswahl aus
    ALTER SALDO, LIEFERUNG, ZURÜCK (leere Spalten fehlen) und NEUER SALDO.
    Es gilt LIEFERUNG - ZURÜCK = DIFF und ALT + DIFF = NEU. Gesucht ist die
    eindeutige Zuordnung, die beides erfüllt. Ergebnis (lief, zur) oder None.
    """
    werte = [zahl(t) for t in mitte.split()] if mitte else []
    if not werte:
        return None
    neu, vorne = werte[-1], werte[:-1]
    if len(vorne) > 3:
        return None
    loesungen = set()
    for maske in range(8):                      # Bit 0 = ALT, 1 = LIEF, 2 = ZUR
        if bin(maske).count("1") != len(vorne):
            continue
        it = iter(vorne)
        alt, lief, zur = [next(it) if maske & (1 << k) else 0.0 for k in range(3)]
        if abs(lief - zur - diff) < 1e-9 and abs(alt + diff - neu) < 1e-9:
            loesungen.add((lief, zur))
    return loesungen.pop() if len(loesungen) == 1 else None


def parse(text):
    r = {"ware": 0.0, "pfand": 0.0, "mwst": 0.0, "satz": None, "zeilenfehler": [], "warnungen": [], "pfandzeilen": [],
         "nachlaesse": []}
    modus = "pos"
    lg_offen = ""   # unvollständige Leergutzeile (Rohzeile mit Einrückung)

    def lg_abschliessen():
        nonlocal lg_offen
        if lg_offen and not RE_ARTNR.match(lg_offen.strip()):
            r["warnungen"].append("Leergutzeile nicht zugeordnet: " + lg_offen.strip())
        elif lg_offen:
            m = RE_LG0.match(lg_offen)
            if m:
                pfand_ohne_differenz(m)
        lg_offen = ""

    def pfand_ohne_differenz(m):
        """Zeile ohne Pfandbetrag: Lieferung = Rückgabe ergibt trotzdem zwei Positionen."""
        teile = leergut_aufteilen(m[3], 0.0)
        if teile is None:
            r["warnungen"].append("Leergut ohne Pfandbetrag nicht auswertbar: " + m[0].strip())
            return
        lief, zur = teile
        satz = zahl(m[4])
        for mg in (lief, -zur):
            if mg:
                r["pfand"] += round(mg * satz, 2)
                r["pfandzeilen"].append((m[1], mg, satz))

    def pfand(m):
        menge, satz, betrag = zahl(m[4]), zahl(m[5]), zahl(m[6])
        if abs(round(menge * satz, 2) - betrag) > 0.005:
            r["zeilenfehler"].append(m[0].strip())
        teile = leergut_aufteilen(m[3], menge)
        if teile is None:
            r["warnungen"].append("Pfand nicht aufteilbar: " + m[0].strip())
            r["pfand"] += betrag
            return
        lief, zur = teile
        for mg in (lief, -zur):
            if mg:
                r["pfand"] += round(mg * satz, 2)
                r["pfandzeilen"].append((m[1], mg, satz))

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
                lg_abschliessen()
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
            if RE_WW_LEER.search(l):
                r["pdf_warenwert"] = 0.0
                modus = "summen"
                continue
            if RE_IG.search(l):
                r["ig"] = True
                continue
            m = RE_USTID.match(l)
            if m:
                r["ustid"] = re.sub(r"[ .\-/]", "", m[1]).upper()
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
            if modus == "summen":
                m = RE_NACHLASS.match(l)
                if m:
                    basis, proz, betrag = zahl(m[3]), zahl(m[4]), zahl(m[5])
                    r["ware"] -= betrag
                    r["nachlaesse"].append((m[1], basis, proz, betrag))
                    if abs(round(basis * proz / 100 + 1e-9, 2) - betrag) > 0.005:
                        r["warnungen"].append("Nachlass Prozent x Basis: " + s)
                    if abs(basis - r.get("pdf_warenwert", 0)) > 0.005:
                        r["warnungen"].append("Nachlass Basis <> Warenwert: " + s)
                    continue
                m = RE_NW.search(l)
                if m:
                    r["pdf_warenwert"] = zahl(m[1])
                    continue
            m = RE_SEPA_BETRAG.search(l)
            if m:
                r["sepa_betrag"] = zahl(m[1])
                continue
            m = RE_SEPA_MANDAT.match(l)
            if m:
                r["sepa"] = (m[1], m[2], m[3])
                continue
            if modus == "pos":
                m = RE_POS.match(l)
                if m:
                    menge, preis, betrag = zahl(m[3]), zahl(m[5]), zahl(m[6])
                    r["ware"] += betrag
                    if abs(round(menge * preis, 2) - betrag) > 0.005:
                        r["zeilenfehler"].append(s)
            elif modus == "lg":
                m = RE_LG.match(l)
                if m:
                    lg_abschliessen()
                    pfand(m)
                    continue
                # pdftotext (Xpdf, Windows) verteilt eine Leergutzeile teils auf zwei
                # Zeilen, der Teil mit dem Pfandbetrag steht dabei VOR der Artikelnummer.
                # Beide Teile spaltengenau übereinanderlegen.
                if lg_offen:
                    z = ueberlagern(lg_offen, l)
                    if z is not None:
                        m = RE_LG.match(z)
                        if m:
                            lg_offen = ""
                            pfand(m)
                            continue
                if RE_ARTNR.match(s) or RE_LGREST.search(s):
                    lg_abschliessen()
                    lg_offen = l
                    continue
    return r


def pruefe(pfad):
    r = parse(open(pfad, encoding="utf-8").read())
    netto = round(r["ware"] + r["pfand"], 2)
    if r.get("ig"):
        mwst = 0.0                       # innergemeinschaftliche Lieferung, Kategorie K
    elif r["satz"] is None:
        mwst = None
    else:
        mwst = round(netto * r["satz"] / 100 + 1e-9, 2)
    checks = [
        ("Warenwert", r["ware"], r.get("pdf_warenwert")),
        ("Pfand", r["pfand"], r.get("pdf_pfand")),
        ("Netto", netto, r.get("pdf_netto")),
        ("MwSt", mwst or 0.0, r["mwst"]),
        ("Brutto", netto + (mwst or 0), r.get("pdf_brutto")),
    ]
    if mwst is None:
        checks.append(("MwSt-Satz", 0, None))
    if r.get("ig") and not r.get("ustid"):
        checks.append(("USt-IdNr. bei innergemeinschaftlicher Lieferung", 0, None))
    fehler = [f"{n}: berechnet {a:.2f}, PDF {b}" for n, a, b in checks if b is None or abs(a - b) > 0.005]
    fehler += [f"Zeile Menge x Preis: {z}" for z in r["zeilenfehler"]]
    fehler += r["warnungen"]
    if "sepa_betrag" in r and abs(r["sepa_betrag"] - (r.get("pdf_brutto") or 0)) > 0.005:
        fehler.append(f"SEPA-Betrag {r['sepa_betrag']} <> Rechnungsbetrag")
    print(("OK    " if not fehler else "FEHLER") + "  " + os.path.basename(pfad))
    for art, mg, satz in r["pfandzeilen"]:
        print(f"        Pfand {art}: {mg:+g} x {satz:.2f}")
    for art, basis, proz, betrag in r["nachlaesse"]:
        print(f"        Nachlass {art} {proz:.2f} % auf {basis:.2f}: -{betrag:.2f}")
    if r.get("ig"):
        print(f"        Innergemeinschaftliche Lieferung (K, 0 %), USt-IdNr. {r.get('ustid')}")
    for z in open(pfad, encoding="utf-8").read().split("\n"):
        m = RE_FILIALE.search(z)
        if m:
            print(f"        Lieferfiliale: PLZ {m[1]}, Ort {m[2]}, Straße {m[3]}")
    if os.path.basename(pfad) == "Rechnung_2604302.txt" and not any(
            (m := RE_FILIALE.search(z)) and m.groups() == ("95233", "Helmbrechts", "Frankenstr. 13")
            for z in open(pfad, encoding="utf-8").read().split("\n")):
        fehler.append("Lieferfiliale nicht erkannt")
    if "sepa" in r:
        print("        SEPA-Lastschrift: Mandat %s, Gläubiger-ID %s, Einzug %s" % r["sepa"])
    for f in fehler:
        print("        " + f)
    return not fehler


def test_differenz_null():
    """Lieferung = Rückgabe (Differenz 0, kein Pfandbetrag im PDF) muss zwei Positionen ergeben.

    Grundlage ist das echte Muster 2603587 (Windows). Die Zeile "Kasten leer" ohne
    Bewegung wird spaltengenau zu 5 geliefert / 5 zurück umgeschrieben.
    """
    pfad = os.path.join(os.path.dirname(__file__), "pdftotext", "Rechnung_2603587_windows.txt")
    text = open(pfad, encoding="utf-8").read()
    alt = next(z for z in text.split("\n") if z.startswith("99009 Kasten leer"))
    neu = alt[:38] + "5".ljust(10) + "5" + alt[49:]
    assert alt[38:49].strip() == "", "Spalten LIEFERUNG/ZURÜCK in der Musterzeile nicht leer"
    r = parse(text.replace(alt, neu))
    zeilen = [z for z in r["pfandzeilen"] if z[0] == "99009"]
    ok = zeilen == [("99009", 5.0, 1.5), ("99009", -5.0, 1.5)] and not r["warnungen"]
    print(("OK    " if ok else "FEHLER") + "  Differenz 0 (synthetisch aus 2603587): " + repr(zeilen))
    return ok


if __name__ == "__main__":
    dateien = sys.argv[1:] or sorted(glob.glob(os.path.join(os.path.dirname(__file__), "pdftotext", "*.txt")))
    ergebnisse = [pruefe(d) for d in dateien]
    if not sys.argv[1:]:
        ergebnisse.append(test_differenz_null())
    sys.exit(0 if all(ergebnisse) else 1)
