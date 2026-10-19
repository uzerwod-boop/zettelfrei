"""Prüft XRechnung-Dateien (UBL) gegen die offiziellen Regeln EN 16931 und XRechnung (Schematron).

Benötigt:  pip install saxonche
Regelwerke (einmalig klonen, Pfade per Umgebungsvariable oder Standard ./_regeln):
  git clone --depth 1 https://github.com/ConnectingEurope/eInvoicing-EN16931 _regeln/en16931
  git clone --depth 1 https://github.com/itplr-kosit/xrechnung-schematron     _regeln/xrechnung
  git clone --depth 1 https://github.com/Schematron/schematron                _regeln/iso

Aufruf:  python scripts/validate_xrechnung.py rechnung.xml [weitere.xml ...]
Ausgabe: je Datei OK oder die verletzten Regeln. Exit-Code 1 bei Fehlern.
"""
import os, sys, re
from saxonche import PySaxonProcessor

BASE = os.environ.get("ZF_REGELN", os.path.join(os.path.dirname(__file__), "..", "_regeln"))
EN = os.path.join(BASE, "en16931", "ubl", "xslt", "EN16931-UBL-validation.xslt")
XR_SCH = os.path.join(BASE, "xrechnung", "src", "validation", "schematron", "ubl", "XRechnung-UBL-validation.sch")
ISO = os.path.join(BASE, "iso", "trunk", "schematron", "code")


def compile_sch(proc, sch, out):
    if os.path.exists(out):
        return out
    x = proc.new_xslt30_processor()
    step = sch
    for i, name in enumerate(["iso_dsdl_include.xsl", "iso_abstract_expand.xsl", "iso_svrl_for_xslt2.xsl"]):
        target = out if i == 2 else out + f".step{i}.xml"
        ex = x.compile_stylesheet(stylesheet_file=os.path.join(ISO, name))
        ex.transform_to_file(source_file=step, output_file=target)
        step = target
    return out


def failures(proc, xslt, xml):
    ex = proc.new_xslt30_processor().compile_stylesheet(stylesheet_file=xslt)
    svrl = ex.transform_to_string(source_file=xml)
    out = []
    for m in re.finditer(r'<svrl:failed-assert[^>]*?(?:id="([^"]*)")?[^>]*?flag="([^"]*)"[^>]*>.*?<svrl:text>(.*?)</svrl:text>', svrl, re.S):
        rule, flag, text = m.group(1) or "", m.group(2), re.sub(r"\s+", " ", m.group(3)).strip()
        out.append((flag, rule, text))
    # Reihenfolge der Attribute variiert: zweiter Durchgang ohne flag-Bedingung
    if not out:
        for m in re.finditer(r'<svrl:failed-assert(.*?)>.*?<svrl:text>(.*?)</svrl:text>', svrl, re.S):
            attrs = m.group(1)
            flag = (re.search(r'flag="([^"]*)"', attrs) or [None, "fatal"])[1]
            rule = (re.search(r'id="([^"]*)"', attrs) or [None, ""])[1]
            out.append((flag, rule, re.sub(r"\s+", " ", m.group(2)).strip()))
    return out


def main(files):
    ok_all = True
    with PySaxonProcessor(license=False) as proc:
        xr = compile_sch(proc, XR_SCH, os.path.join(BASE, "XRechnung-UBL-validation.xsl"))
        for f in files:
            probs = [("EN16931",) + p for p in failures(proc, EN, f)] + [("XRechnung",) + p for p in failures(proc, xr, f)]
            errors = [p for p in probs if p[1] in ("fatal", "error")]
            warnings = [p for p in probs if p not in errors]
            print(("OK      " if not errors else "FEHLER  ") + os.path.basename(f) + (f"  ({len(warnings)} Hinweise)" if warnings else ""))
            for src, flag, rule, text in probs:
                print(f"   [{src} {flag} {rule}] {text}")
            ok_all = ok_all and not errors
    return 0 if ok_all else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
