"""Gemeinsame VBA-Module zwischen den Produktordnern abgleichen.

Jedes Produkt (PDF2E, XL2E, ...) hat unter src/<produkt>/ alle Module seiner
Arbeitsmappe, so wie VbaSync sie exportiert. Die in src/gemeinsam.txt
genannten Dateien müssen in allen Produktordnern identisch sein.

  python tools/abgleich.py               prüft; Rückgabe 1, wenn Kopien abweichen
  python tools/abgleich.py --von pdf2e   kopiert die gemeinsamen Dateien aus
                                         src/pdf2e in alle anderen Produktordner
Produktordner = Unterordner von src/ mit einer modProdukt.bas.
"""
import os
import shutil
import sys

SRC = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "src")


def gemeinsame_dateien():
    with open(os.path.join(SRC, "gemeinsam.txt"), encoding="utf-8") as f:
        return [z.strip() for z in f if z.strip() and not z.lstrip().startswith("#")]


def produkte():
    return sorted(d for d in os.listdir(SRC) if os.path.isfile(os.path.join(SRC, d, "modProdukt.bas")))


def inhalt(pfad):
    """Dateiinhalt ohne Unterschied CRLF/LF, damit Checkout-Einstellungen keine Rolle spielen."""
    if not os.path.exists(pfad):
        return None
    return open(pfad, "rb").read().replace(b"\r\n", b"\n")


def pruefen():
    prods, fehler = produkte(), []
    for datei in gemeinsame_dateien():
        staende = {p: inhalt(os.path.join(SRC, p, datei)) for p in prods}
        if None in staende.values():
            fehlt = [p for p, s in staende.items() if s is None]
            fehler.append(f"{datei}: fehlt in {', '.join(fehlt)}")
        elif len(set(staende.values())) > 1:
            fehler.append(f"{datei}: unterschiedlich in {', '.join(prods)}")
    for f in fehler:
        print("ABWEICHUNG  " + f)
    print(("OK" if not fehler else "FEHLER") + f"  {len(gemeinsame_dateien())} gemeinsame Dateien in {', '.join(prods)}")
    return not fehler


def kopieren(von):
    prods = produkte()
    if von not in prods:
        sys.exit(f"Unbekannter Produktordner '{von}'. Vorhanden: {', '.join(prods)}")
    for datei in gemeinsame_dateien():
        quelle = os.path.join(SRC, von, datei)
        for p in prods:
            ziel = os.path.join(SRC, p, datei)
            if p != von and inhalt(quelle) != inhalt(ziel):
                shutil.copyfile(quelle, ziel)
                print(f"kopiert  {von}/{datei} -> {p}/")


if __name__ == "__main__":
    if len(sys.argv) == 3 and sys.argv[1] == "--von":
        kopieren(sys.argv[2])
    sys.exit(0 if pruefen() else 1)
