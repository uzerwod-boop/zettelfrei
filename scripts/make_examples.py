"""Erzeugt drei typische Excel-Dateien eines kleinen Handwerksbetriebs (Beispieldaten)."""
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

import os
OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "beispiele") + os.sep
BOLD = Font(bold=True)
HEAD = PatternFill("solid", fgColor="DDE7E3")
DATE = "DD.MM.YYYY"
EUR = '#,##0.00 "€"'


def header(ws, row, names, widths):
    for i, n in enumerate(names, 1):
        c = ws.cell(row=row, column=i, value=n)
        c.font = BOLD
        c.fill = HEAD
        ws.column_dimensions[get_column_letter(i)].width = widths[i - 1]


# 1) Stundenzettel -----------------------------------------------------------
wb = Workbook()
ws = wb.active
ws.title = "Oktober 2026"
ws["A1"] = "Fliesen Krause – Stundenzettel Oktober 2026"
ws["A1"].font = Font(bold=True, size=14)
header(ws, 3, ["Datum", "Mitarbeiter", "Projekt", "Tätigkeit", "Stunden", "Stundensatz €", "Betrag €", "Bemerkung"],
       [12, 13, 18, 30, 9, 13, 12, 24])
rows = [
    (date(2026, 10, 12), "Kevin", "Bad Schneider", "Altfliesen entfernt", 8, 48, ""),
    (date(2026, 10, 12), "Mehmet", "Bad Schneider", "Untergrund grundiert", 7.5, 52, ""),
    (date(2026, 10, 13), "Kevin", "Bad Schneider", "Abdichtung Dusche", 6, 48, ""),
    (date(2026, 10, 13), "Jonas", "Küche Yilmaz", "Aufmaß, Material bestellt", 3, 42, ""),
    (date(2026, 10, 14), "Mehmet", "Bad Schneider", "Wand gefliest 30x60", 11, 52, "Kunde wollte heute fertig"),
    (date(2026, 10, 14), "Jonas", "Praxis Dr. Hahn", "Fugen ausgebessert", 4, 42, ""),
    (date(2026, 10, 15), "Kevin", "Küche Yilmaz", "Boden verlegt", 8, 48, ""),
    (date(2026, 10, 15), "Mehmet", None, "Fahrt Baustoffhandel", 1.5, 52, ""),
    (date(2026, 10, 16), "Jonas", "Küche Yilmaz", "Verfugt, Silikon", 6.5, 42, ""),
]
r0 = 4
for i, (d, m, p, t, h, s, b) in enumerate(rows):
    r = r0 + i
    ws.cell(r, 1, d).number_format = DATE
    ws.cell(r, 2, m); ws.cell(r, 3, p); ws.cell(r, 4, t)
    ws.cell(r, 5, h); ws.cell(r, 6, s).number_format = EUR
    ws.cell(r, 7, f"=E{r}*F{r}").number_format = EUR
    ws.cell(r, 8, b)
last = r0 + len(rows) - 1
tr = last + 1
ws.cell(tr, 1, "Summe").font = BOLD
ws.cell(tr, 5, f"=SUM(E{r0}:E{last})").font = BOLD
ws.cell(tr, 7, f"=SUM(G{r0}:G{last})").font = BOLD
ws.cell(tr, 7).number_format = EUR

ws2 = wb.create_sheet("Mitarbeiter")
header(ws2, 1, ["Name", "Stundensatz €", "Telefon", "Seit"], [14, 14, 18, 12])
for i, (n, s, t, d) in enumerate([("Kevin", 48, "0176 2231904", date(2021, 3, 1)),
                                  ("Mehmet", 52, "0152 8840217", date(2018, 8, 15)),
                                  ("Jonas", 42, "0160 5519023", date(2025, 9, 1))], 2):
    ws2.cell(i, 1, n); ws2.cell(i, 2, s).number_format = EUR; ws2.cell(i, 3, t)
    ws2.cell(i, 4, d).number_format = DATE
wb.save(OUT + "Stundenzettel_Oktober_2026.xlsx")

# 2) Material / Lager --------------------------------------------------------
wb = Workbook()
ws = wb.active
ws.title = "Lager"
header(ws, 1, ["Artikel-Nr", "Artikel", "Kategorie", "Bestand", "Mindestbestand", "Einheit", "EK-Preis €", "Lagerwert €", "Lieferant"],
       [11, 32, 13, 9, 15, 10, 11, 12, 22])
items = [
    ("A-1001", "Fliesenkleber C2TE 25 kg", "Kleber", 14, 10, "Sack", 14.90, "Baustoff Kaiser GmbH"),
    ("A-1002", "Flexkleber S1 25 kg", "Kleber", 4, 8, "Sack", 17.50, "Fliesen-Großhandel Weber KG"),
    ("A-1010", "Fugenmörtel grau 5 kg", "Fuge", 9, 6, "Eimer", 11.40, "Baustoff Kaiser GmbH"),
    ("A-1011", "Fugenmörtel anthrazit 5 kg", "Fuge", 2, 4, "Eimer", 12.30, "Fliesen-Großhandel Weber KG"),
    ("A-1020", "Silikon sanitär weiß 310 ml", "Silikon", 18, 12, "Kartusche", 6.20, "Baustoff Kaiser GmbH"),
    ("A-1021", "Silikon sanitär grau 310 ml", "Silikon", 5, 12, "Kartusche", 6.20, "Baustoff Kaiser GmbH"),
    ("A-1030", "Dichtschlämme 2K 15 kg", "Abdichtung", 3, 2, "Eimer", 64.00, "Baustoff Kaiser GmbH"),
    ("A-1031", "Dichtband 10 m", "Abdichtung", 7, 5, "Rolle", 18.90, "Baustoff Kaiser GmbH"),
    ("A-1040", "Tiefengrund 10 l", "Grundierung", 1, 3, "Kanister", 29.50, "Baustoff Kaiser GmbH"),
    ("A-1050", "Abschlussprofil Alu 2,5 m", "Profile", 22, 10, "Stk", 9.80, "Fliesen-Großhandel Weber KG"),
    ("A-1051", "Eckschutzprofil Edelstahl 2,5 m", "Profile", 6, 4, "Stk", 21.40, "Fliesen-Großhandel Weber KG"),
    ("A-1060", "Feinsteinzeug 60x60 anthrazit", "Fliesen", 12.5, 0, "m²", 32.90, "Fliesen-Großhandel Weber KG"),
    ("A-1061", "Wandfliese 30x60 weiß matt", "Fliesen", 30, 0, "m²", 24.50, "Fliesen-Großhandel Weber KG"),
    ("A-1070", "Fliesenkreuze 3 mm (250 Stk)", "Zubehör", 9, 3, "Beutel", 3.90, "Baustoff Kaiser GmbH"),
]
for i, it in enumerate(items, 2):
    for j, v in enumerate(it[:7], 1):
        ws.cell(i, j, v)
    ws.cell(i, 7).number_format = EUR
    ws.cell(i, 8, f"=D{i}*G{i}").number_format = EUR
    ws.cell(i, 9, it[7])
last = 1 + len(items)
ws.cell(last + 1, 1, "Gesamt").font = BOLD
ws.cell(last + 1, 8, f"=SUM(H2:H{last})").number_format = EUR
wb.save(OUT + "Material_Lager.xlsx")

# 3) Aufträge ----------------------------------------------------------------
wb = Workbook()
ws = wb.active
ws.title = "Aufträge 2026"
ws["A1"] = "Auftragsliste 2026"
ws["A1"].font = Font(bold=True, size=14)
ws["A2"] = "Stand: 16.10.2026"
header(ws, 4, ["Auftrags-Nr", "Kunde", "Adresse", "Leistung", "Status", "Angebot netto €", "MwSt €", "Brutto €", "Rechnung am", "Bezahlt am"],
       [12, 20, 28, 30, 12, 15, 11, 12, 13, 13])
orders = [
    ("2026-031", "Schneider, Petra", "Lindenstr. 8, 10969 Berlin", "Badsanierung komplett", "In Arbeit", 14800, None, None),
    ("2026-032", "Yilmaz, Ahmet", "Sonnenallee 112, 12045 Berlin", "Küchenboden + Spiegel", "In Arbeit", 6350, None, None),
    ("2026-033", "Praxis Dr. Hahn", "Kantstr. 41, 10625 Berlin", "Fugen- und Silikonsanierung", "Fertig", 1980, date(2026, 10, 17), None),
    ("2026-028", "Wohnbau Spree eG", "Am Treptower Park 3, 12435 Berlin", "Treppenhaus 4 Etagen", "Bezahlt", 9200, date(2026, 9, 12), date(2026, 9, 30)),
    ("2026-029", "Becker, Thomas", "Dorfstr. 5, 14532 Kleinmachnow", "Terrasse Feinsteinzeug", "Fertig", 7450, date(2026, 9, 26), None),
    ("2026-034", "Neumann, Lisa", "Boxhagener Str. 20, 10245 Berlin", "Gäste-WC", "Angebot", 3100, None, None),
    ("2026-035", None, "Weserstr. 9, 12047 Berlin", "Duschwand Reparatur", "Angebot", 890, None, None),
    ("2026-033", "Hoffmann GmbH", "Friedrichstr. 77, 10117 Berlin", "Büroküche Wandfliesen", "Beauftragt", 2650, None, None),
]
r0 = 5
for i, o in enumerate(orders):
    r = r0 + i
    for j, v in enumerate(o[:5], 1):
        ws.cell(r, j, v)
    ws.cell(r, 6, o[5]).number_format = EUR
    ws.cell(r, 7, f"=ROUND(F{r}*0.19,2)").number_format = EUR
    ws.cell(r, 8, f"=F{r}+G{r}").number_format = EUR
    if o[6]: ws.cell(r, 9, o[6]).number_format = DATE
    if o[7]: ws.cell(r, 10, o[7]).number_format = DATE
last = r0 + len(orders) - 1
ws.cell(last + 1, 1, "Summe").font = BOLD
for col in "FGH":
    c = ws[f"{col}{last + 1}"]
    c.value = f"=SUM({col}{r0}:{col}{last})"
    c.number_format = EUR
    c.font = BOLD
wb.save(OUT + "Auftraege_2026.xlsx")
print("ok")

# 4) Arbeitszeit nach ArbZG --------------------------------------------------
from datetime import time
TIME = "HH:MM"
wb = Workbook()
ws = wb.active
ws.title = "KW 42"
ws["A1"] = "Arbeitszeitnachweis KW 42 / 2026"
ws["A1"].font = Font(bold=True, size=14)
header(ws, 3, ["Datum", "Mitarbeiter", "Projekt", "Beginn", "Ende", "Pause (Min.)", "Stunden", "Tätigkeit"],
       [12, 13, 18, 9, 9, 12, 9, 30])
shifts = [
    (date(2026, 10, 12), "Kevin", "Bad Schneider", time(7, 0), time(15, 30), 30, "Altfliesen entfernt"),
    (date(2026, 10, 12), "Mehmet", "Bad Schneider", time(7, 0), time(15, 0), 30, "Untergrund grundiert"),
    (date(2026, 10, 13), "Kevin", "Bad Schneider", time(7, 30), time(14, 0), 15, "Abdichtung Dusche"),
    (date(2026, 10, 13), "Jonas", "Küche Yilmaz", time(13, 0), time(22, 0), 45, "Aufmaß und Notdienst"),
    (date(2026, 10, 14), "Jonas", "Praxis Dr. Hahn", time(6, 30), time(11, 0), 0, "Fugen ausgebessert"),
    (date(2026, 10, 14), "Mehmet", "Bad Schneider", time(6, 0), time(18, 0), 45, "Wand gefliest 30x60"),
    (date(2026, 10, 15), "Kevin", "Küche Yilmaz", time(7, 0), time(15, 30), 30, "Boden verlegt"),
    (date(2026, 10, 16), "Jonas", "Küche Yilmaz", time(7, 0), time(14, 0), 30, "Verfugt, Silikon"),
]
for i, (d, m, p, b, e, pause, t) in enumerate(shifts):
    r = 4 + i
    ws.cell(r, 1, d).number_format = DATE
    ws.cell(r, 2, m); ws.cell(r, 3, p)
    ws.cell(r, 4, b).number_format = TIME
    ws.cell(r, 5, e).number_format = TIME
    ws.cell(r, 6, pause)
    ws.cell(r, 7, f"=ROUND((E{r}-D{r})*24-F{r}/60,2)").number_format = "0.00"
    ws.cell(r, 8, t)
last = 3 + len(shifts)
ws.cell(last + 1, 1, "Summe").font = BOLD
ws.cell(last + 1, 7, f"=SUM(G4:G{last})").number_format = "0.00"
wb.save(OUT + "Arbeitszeit_KW42.xlsx")

# 5) Eingangsrechnungen ------------------------------------------------------
wb = Workbook()
ws = wb.active
ws.title = "Eingangsrechnungen"
header(ws, 1, ["Rechnung am", "Lieferant", "Rechnungs-Nr", "Netto €", "MwSt-Satz", "MwSt €", "Brutto €", "Fällig am", "Bezahlt am", "Projekt"],
       [12, 28, 14, 12, 10, 11, 12, 12, 12, 18])
bills = [
    (date(2026, 9, 18), "Baustoff Kaiser GmbH", "RE-2026-4398", 412.60, 0.19, date(2026, 10, 2), date(2026, 9, 30), "Bad Schneider"),
    (date(2026, 10, 1), "Fliesen-Großhandel Weber KG", "87911", 1960.00, 0.19, date(2026, 10, 15), None, "Bad Schneider"),
    (date(2026, 10, 13), "Baustoff Kaiser GmbH", "RE-2026-4471", 311.20, 0.19, date(2026, 10, 27), None, "Bad Schneider"),
    (date(2026, 10, 14), "Werkzeug Lehmann", "WL-2209", 189.00, 0.19, date(2026, 10, 28), None, None),
]
for i, (d, l, nr, net, rate, due, paid, proj) in enumerate(bills, 2):
    ws.cell(i, 1, d).number_format = DATE
    ws.cell(i, 2, l); ws.cell(i, 3, nr)
    ws.cell(i, 4, net).number_format = EUR
    ws.cell(i, 5, rate).number_format = "0%"
    ws.cell(i, 6, f"=ROUND(D{i}*E{i},2)").number_format = EUR
    ws.cell(i, 7, f"=D{i}+F{i}").number_format = EUR
    ws.cell(i, 8, due).number_format = DATE
    if paid: ws.cell(i, 9, paid).number_format = DATE
    if proj: ws.cell(i, 10, proj)
last = 1 + len(bills)
ws.cell(last + 1, 1, "Summe").font = BOLD
for c in "DFG":
    ws[f"{c}{last+1}"] = f"=SUM({c}2:{c}{last})"; ws[f"{c}{last+1}"].number_format = EUR
wb.save(OUT + "Eingangsrechnungen_2026.xlsx")
print("ok 2")
