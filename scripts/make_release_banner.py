"""Banner für die Release-Seite: python scripts/make_release_banner.py 0.2.1"""
import asyncio, os, sys
from playwright.async_api import async_playwright
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__))); D = os.path.join(ROOT, "docs")
VER = sys.argv[1] if len(sys.argv) > 1 else "0.2.1"
F = "file://" + os.path.join(ROOT, "fonts")
logo = open(os.path.join(D, "brand", "zettelfrei-logo.svg")).read()
ITEMS_021 = [("Sicherheit", "Formeln führen keinen Code aus, keine Formel-Einschleusung, Server nur für die eigene Oberfläche"),
         ("Neue Leser", "Excel über ExcelJS, CSV im deutschen Format, PDF.js und jsPDF abgesichert"),
         ("Gutschriften", "werden mit Minus gebucht und senken den Lagerbestand"),
         ("Neuer Auftritt", "Logo, README, Lizenz nach Wahl MIT oder Apache-2.0")]
ITEMS_050 = [("Programm für Mac und Windows", "Installieren, im Dock oder Startmenü öffnen – die Datei bleibt auf dem Rechner"),
         ("Neue Oberfläche", "Wie Excel: anklicken und tippen, Assistent rechts, alles in einem Fenster"),
         ("Hineinziehen", "Excel-Dateien und Rechnungen aus dem Ordner direkt ins Fenster ziehen"),
         ("Kostenlos", "Open Source, MIT oder Apache-2.0 – auch weiter im Browser nutzbar")]
ITEMS = ITEMS_050 if VER.startswith("0.5") else ITEMS_021
HEAD = "Jetzt als <em>Programm</em>." if VER.startswith("0.5") else "Geprüft und <em>gehärtet</em>."
cells = "".join(f'<div class="it"><b>{t}</b><span>{d}</span></div>' for t, d in ITEMS)
html = f"""<html><head><style>
@font-face{{font-family:A;src:url('{F}/Archivo.ttf');font-weight:100 900}}
@font-face{{font-family:P;src:url('{F}/IBMPlexSans.ttf');font-weight:100 700}}
@font-face{{font-family:M;src:url('{F}/IBMPlexMono-Regular.ttf')}}
body{{margin:0;width:1280px;height:640px;background:#f4f5f1;color:#1d2528;font-family:P;position:relative;overflow:hidden}}
.grid{{position:absolute;inset:0;background-image:linear-gradient(#e3e7e0 1px,transparent 1px),linear-gradient(90deg,#e3e7e0 1px,transparent 1px);background-size:40px 40px;opacity:.7}}
.card{{position:absolute;left:64px;top:56px;right:64px;bottom:56px;background:#fff;border:1px solid #d9ddd6;border-radius:22px;padding:52px 60px;box-shadow:0 18px 50px rgba(15,60,50,.10)}}
.top{{display:flex;align-items:center;justify-content:space-between}}
.top svg{{height:64px;width:auto}}
.ver{{font:500 22px M;color:#0f6b5c;border:2px solid #0f6b5c;border-radius:999px;padding:8px 20px}}
h2{{font:800 58px/1.05 A;letter-spacing:-.02em;margin:44px 0 34px}} h2 em{{font-style:normal;background:linear-gradient(transparent 60%,#f6dd6a 60%)}}
.items{{display:grid;grid-template-columns:1fr 1fr;gap:22px 48px}}
.it{{display:flex;flex-direction:column;gap:4px;padding-left:18px;border-left:4px solid #0f6b5c}}
.it b{{font:700 22px P}} .it span{{font:400 18px/1.35 P;color:#5d686b}}
</style></head><body><div class="grid"></div><div class="card">
<div class="top">{logo}<span class="ver">Version {VER}</span></div>
<h2>{HEAD}</h2><div class="items">{cells}</div></div></body></html>"""
async def main():
    p = os.path.join(D, "_banner.html"); open(p, "w").write(html)
    async with async_playwright() as pw:
        b = await pw.chromium.launch(); pg = await b.new_page(viewport={"width": 1280, "height": 640}, device_scale_factor=2)
        await pg.goto("file://" + p); await pg.wait_for_timeout(500)
        await pg.screenshot(path=os.path.join(D, "release", f"v{VER}.png")); await b.close()
    os.remove(p)
asyncio.run(main())
