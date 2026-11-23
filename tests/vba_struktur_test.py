"""Statische Prüfung der VBA-Module (src/**/*.bas, *.frm, *.doccls).

Prüft, dass Deklarationen auf Modulebene (Const, Dim, Private/Public/Global
Variablen, Declare, Type, Enum, Option, Implements) im Deklarationsteil oben
stehen, also vor der ersten Sub/Function/Property. Steht so eine Zeile zwischen
Prozeduren, meldet der VBA-Editor beim Kompilieren einen Fehler.

Aufruf:  python tests/vba_struktur_test.py
"""
import glob
import os
import re
import sys

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
PROC_START = re.compile(r"^\s*(?:(?:Public|Private|Friend|Static)\s+)*(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+\w", re.I)
PROC_END = re.compile(r"^\s*End\s+(?:Sub|Function|Property)\b", re.I)
DEKL = re.compile(r"^\s*(?:(?:Public|Private|Global)\s+(?:Const|Declare|Type|Enum|WithEvents\s+)?|Const\s|Dim\s|Declare\s|Type\s|Enum\s|Option\s|Implements\s)", re.I)


def lesen(pfad):
    roh = open(pfad, "rb").read()
    try:
        return roh.decode("utf-8")
    except UnicodeDecodeError:
        return roh.decode("cp1252")


def logische_zeilen(text):
    """Fortsetzungszeilen ( _) zusammenfassen; Nummer der ersten Zeile behalten."""
    erg, puffer, start = [], "", None
    for nr, z in enumerate(text.replace("\r\n", "\n").split("\n"), 1):
        if start is None:
            start = nr
        puffer += z
        if re.search(r"\s_\s*$", z):
            puffer = puffer.rstrip()[:-1] + " "
            continue
        erg.append((start, puffer))
        puffer, start = "", None
    return erg


def pruefe(pfad):
    fehler = []
    in_proc = False
    erste_proc = None
    in_typ = False
    for nr, z in logische_zeilen(lesen(pfad)):
        code = z.split("'", 1)[0] if not z.lstrip().startswith("'") else ""
        if not code.strip() or code.lstrip().startswith("Attribute "):
            continue
        if in_proc:
            if PROC_END.match(code):
                in_proc = False
            continue
        if in_typ:
            if re.match(r"^\s*End\s+(?:Type|Enum)\b", code, re.I):
                in_typ = False
            continue
        if PROC_START.match(code):
            in_proc = True
            erste_proc = erste_proc or nr
            continue
        if DEKL.match(code):
            if re.match(r"^\s*(?:(?:Public|Private)\s+)?(?:Type|Enum)\s", code, re.I):
                in_typ = True
            if erste_proc:
                fehler.append(f"{os.path.relpath(pfad, ROOT)}:{nr}: Deklaration nach der ersten Prozedur (Zeile {erste_proc}): {code.strip()[:80]}")
    return fehler


if __name__ == "__main__":
    dateien = sorted(glob.glob(os.path.join(ROOT, "src", "**", "*.bas"), recursive=True)
                     + glob.glob(os.path.join(ROOT, "src", "**", "*.frm"), recursive=True)
                     + glob.glob(os.path.join(ROOT, "src", "**", "*.doccls"), recursive=True))
    alle = [f for d in dateien for f in pruefe(d)]
    for f in alle:
        print("FEHLER  " + f)
    print(("OK    " if not alle else "FEHLER") + f"  {len(dateien)} VBA-Dateien geprüft")
    sys.exit(1 if alle else 0)
