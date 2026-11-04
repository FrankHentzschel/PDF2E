"""Erzeugt die Auslieferungs-ZIPs für PDF2E und XL2E.

Aufruf:  python vorlage/pakete.py <Zielordner> [<Logo-Ordner>]

Je Produkt entsteht <NAME>_<VERSION>.zip mit
  <NAME>.xlsx          Arbeitsmappe (vorlage/build_xlsx.py)
  vba/                 Module aus src/kern und src/<produkt>
  logo/                Logos aus dem Logo-Ordner (optional, siehe LOGOS)
  ANLEITUNG.md         PDF2E: ANLEITUNG.md, XL2E: ANLEITUNG_XL2E.md
Die VBA-Dateien werden in Windows-1252 mit CRLF gepackt (wie der VBA-Editor sie braucht).
"""
import os, re, subprocess, sys, tempfile, zipfile

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
PRODUKTE = {
    "pdf2e": {"name": "PDF2E", "anleitung": "ANLEITUNG.md"},
    "xl2e": {"name": "XL2E", "anleitung": "ANLEITUNG_XL2E.md"},
}
# Logo-Dateien je Produkt (relativ zum Logo-Ordner); fehlende werden übersprungen
LOGOS = {
    "pdf2e": ["pdf2e/logo/pdf2e_logo.jpg", "pdf2e/logo/pdf2e_logo.bmp", "pdf2e/logo/pdf2e.ico",
              "pdf2e/logo/officedev_logo.jpg", "pdf2e/logo/officedev_logo.bmp"],
    "xl2e": ["pdf2e/logo/officedev_logo.jpg", "pdf2e/logo/officedev_logo.bmp",
             "xl2e/logo/fiedler_logo_hoch.png"],
}


def vba_bytes(pfad):
    """Datei als cp1252 mit CRLF, unabhängig davon, wie sie ausgecheckt ist."""
    roh = open(pfad, "rb").read()
    try:
        text = roh.decode("utf-8")
    except UnicodeDecodeError:
        text = roh.decode("cp1252")
    text = text.replace("\r\n", "\n").replace("\n", "\r\n")
    return text.encode("cp1252")


def version(prod):
    s = open(os.path.join(ROOT, "src", prod, "modProdukt.bas"), "rb").read().decode("cp1252", "replace")
    return re.search(r'TOOL_VERSION As String = "([^"]+)"', s)[1]


def paket(prod, ziel, logo_ordner):
    info = PRODUKTE[prod]
    name, ver = info["name"], version(prod)
    zip_pfad = os.path.join(ziel, f"{name}_{ver}.zip")
    with tempfile.TemporaryDirectory() as tmp:
        mappe = os.path.join(tmp, f"{name}.xlsx")
        subprocess.run([sys.executable, os.path.join(ROOT, "vorlage", "build_xlsx.py"), mappe, prod], check=True)
        with zipfile.ZipFile(zip_pfad, "w", zipfile.ZIP_DEFLATED) as z:
            z.write(mappe, f"{name}.xlsx")
            for ordner in ("kern", prod):
                for f in sorted(os.listdir(os.path.join(ROOT, "src", ordner))):
                    z.writestr(f"vba/{f}", vba_bytes(os.path.join(ROOT, "src", ordner, f)))
            if logo_ordner:
                for rel in LOGOS[prod]:
                    p = os.path.join(logo_ordner, rel)
                    if os.path.exists(p):
                        z.write(p, "logo/" + os.path.basename(p))
            z.write(os.path.join(ROOT, info["anleitung"]), "ANLEITUNG.md")
        # Mappe zusätzlich einzeln ablegen
        os.replace(mappe, os.path.join(ziel, f"{name}.xlsx"))
    print(zip_pfad)


if __name__ == "__main__":
    ziel = sys.argv[1]
    logos = sys.argv[2] if len(sys.argv) > 2 else None
    os.makedirs(ziel, exist_ok=True)
    for p in PRODUKTE:
        paket(p, ziel, logos)
