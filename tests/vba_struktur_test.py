"""Statische Prüfung der VBA-Module (src/**/*.bas, *.frm, *.doccls).

Prüft, dass Deklarationen auf Modulebene (Const, Dim, Private/Public/Global
Variablen, Declare, Type, Enum, Option, Implements) im Deklarationsteil oben
stehen, also vor der ersten Sub/Function/Property. Steht so eine Zeile zwischen
Prozeduren, meldet der VBA-Editor beim Kompilieren einen Fehler.
Außerdem: keine Variable/Prozedur mit dem Namen einer VBA-Funktion (z. B. mId = Mid).

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


# VBA-Funktionen und -Anweisungen. Eine Variable oder Prozedur gleichen Namens
# (VBA unterscheidet Groß-/Kleinschreibung nicht, z. B. mId = Mid) verdeckt sie,
# der Editor meldet dann beim Kompilieren z. B. "Erwartet: Datenfeld".
VBA_NAMEN = {n.lower() for n in """Abs Array Asc AscB AscW Atn CBool CByte CCur CDate CDbl CDec Choose Chr ChrB ChrW
CInt CLng CSng CStr CurDir CVar CVErr Date DateAdd DateDiff DatePart DateSerial DateValue Day Dir DoEvents
Environ EOF Error Exp FileAttr FileDateTime FileLen Filter Fix Format FreeFile GetAttr GetObject Hex Hour IIf
Input InputB InputBox InStr InStrB InStrRev Int IsArray IsDate IsEmpty IsError IsMissing IsNull IsNumeric
IsObject Join LBound LCase Left LeftB Len LenB Loc LOF Log LTrim Mid MidB Minute Month MonthName MsgBox Now
Oct Replace RGB Right RightB Rnd Round RTrim Second Seek Sgn Shell Sin Space Split Sqr Str StrComp StrConv
String StrReverse Switch Tab Tan Time Timer TimeSerial TimeValue Trim TypeName UBound UCase Val VarType
Weekday Year Kill FileCopy MkDir RmDir ChDir Randomize CreateObject""".split()}
DEKL_NAME = re.compile(r"\b(?:Dim|Private|Public|Global|Static|Const|Sub|Function|Get|Let)\s+(?:Const\s+|WithEvents\s+)?(\w+)", re.I)
WEITERE_NAMEN = re.compile(r",\s*(?:Optional\s+)?(?:ByVal\s+|ByRef\s+)?(\w+)\s*(?:\(\))?\s+As\b", re.I)
PARAM_NAME = re.compile(r"\(\s*(?:Optional\s+)?(?:ByVal\s+|ByRef\s+)?(\w+)\s*(?:\(\))?\s+As\b", re.I)


def namenskonflikte(pfad):
    fehler = []
    for nr, z in logische_zeilen(lesen(pfad)):
        code = z.split("'", 1)[0] if not z.lstrip().startswith("'") else ""
        if not code.strip() or code.lstrip().startswith("Attribute "):
            continue
        namen = DEKL_NAME.findall(code) + WEITERE_NAMEN.findall(code) + PARAM_NAME.findall(code)
        for n in namen:
            if n.lower() in VBA_NAMEN:
                fehler.append(f"{os.path.relpath(pfad, ROOT)}:{nr}: Name '{n}' verdeckt die VBA-Funktion {n}: {code.strip()[:80]}")
    return fehler


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
    alle = [f for d in dateien for f in pruefe(d) + namenskonflikte(d)]
    for f in alle:
        print("FEHLER  " + f)
    print(("OK    " if not alle else "FEHLER") + f"  {len(dateien)} VBA-Dateien geprüft")
    sys.exit(1 if alle else 0)
