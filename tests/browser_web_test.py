"""Ende-zu-Ende-Test für die neue Oberfläche 0.5 (React, web/dist):
    cd web && npm ci && npm run build && cd .. && python tests/browser_web_test.py
Prüft: Beispiel öffnen, Zelle bearbeiten, Assistent (Auswertung + Eintrag mit Vorschau), Prüfen,
Verlauf und Protokoll, Belege (XRechnung, ZUGFeRD-PDF), Dokumente (E-Rechnung schreiben),
Speichern als Excel. KI-Antworten werden nachgebaut; gerechnet wird echt.
"""
import asyncio, json, os, re, socket, subprocess, sys, tempfile, time
from playwright.async_api import async_playwright
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__))); OUT = tempfile.mkdtemp(prefix="zf-web-")
ok = True
def check(c, m):
    global ok; print(("ok   " if c else "FAIL ") + m); ok = ok and bool(c)
def free_port():
    s = socket.socket(); s.bind(("127.0.0.1", 0)); p = s.getsockname()[1]; s.close(); return p
SET = {"provider": "openrouter", "keys": {"openrouter": "sk-or-test", "claude": "", "local": ""}, "models": {"openrouter": "", "claude": "", "local": ""}, "url": "http://localhost:11434/v1"}
FIRMA = {"name": "Fliesen Krause", "contactName": "Petra Krause", "phone": "030 1234567", "street": "Gneisenaustr. 22", "zip": "10961", "city": "Berlin", "email": "buero@fliesen-krause.example",
         "taxNo": "12/345/67890", "iban": "DE02120300000000202051", "zahlungsziel": "14", "bearbeiter": "Test"}
def answer(text):
    if "je Baustelle" in text:
        return {"abfragen": [{"id": "q1", "blatt": "Stunden Oktober", "spalte": "Stunden", "art": "summe", "gruppieren": "Baustelle"}], "antwort": "Stunden je Baustelle: {q1}"}
    if "Jonas" in text:
        return {"aenderungen": [{"typ": "neu", "blatt": "Stunden Oktober", "werte": {"Datum": "20.10.2026", "Mitarbeiter": "Jonas", "Baustelle": "Neubau Weber",
                "Beginn": "07:00", "Ende": "16:30", "Pause (Min.)": 30}}], "antwort": "Ich trage Jonas ein."}
    return {"antwort": "?"}
async def main():
    if not os.path.exists(os.path.join(ROOT, "web", "dist", "index.html")):
        print("web/dist fehlt – zuerst: cd web && npm ci && npm run build"); sys.exit(1)
    port = free_port(); env = dict(os.environ, PORT=str(port)); env.pop("ANTHROPIC_API_KEY", None); env.pop("OPENROUTER_API_KEY", None); env.pop("ZF_KLASSISCH", None)
    srv = subprocess.Popen(["node", "server.js"], cwd=ROOT, env=env, stdout=subprocess.DEVNULL); time.sleep(0.8)
    errs = []
    try:
        async with async_playwright() as p:
            b = await p.chromium.launch(); ctx = await b.new_context(viewport={"width": 1440, "height": 900}, accept_downloads=True)
            await ctx.add_init_script("try{if(!sessionStorage.getItem('t')){sessionStorage.setItem('t',1);localStorage.clear();}localStorage.setItem('zf-ai-settings'," + json.dumps(json.dumps(SET)) +
                                      ");localStorage.setItem('zf-firma'," + json.dumps(json.dumps(FIRMA)) + ")}catch(e){}")
            async def route(r):
                body = json.loads(r.request.post_data); msg = body["messages"][-1]["content"]
                if isinstance(msg, list): msg = " ".join(x.get("text", "") for x in msg)
                m = re.search(r'Nutzers: """(.*?)"""', msg, re.S)
                await r.fulfill(status=200, content_type="application/json", body=json.dumps({"choices": [{"message": {"content": json.dumps(answer(m.group(1) if m else ""))}}]}))
            await ctx.route("https://openrouter.ai/**", route)
            pg = await ctx.new_page(); pg.on("pageerror", lambda e: errs.append(str(e)))
            await pg.goto(f"http://127.0.0.1:{port}/"); await pg.wait_for_timeout(700)
            check(await pg.locator(".empty-state").count() == 1, "Startseite ist die neue Oberfläche")
            await pg.select_option('select[aria-label="Beispiel öffnen"]', "Krause_Betrieb_Oktober_2026.xlsx"); await pg.wait_for_timeout(800)
            await pg.click('.sheettabs button:text-is("Stunden Oktober")'); await pg.wait_for_timeout(300)
            # Zelle bearbeiten
            await pg.click('td[data-r="15"][data-k="5"]'); await pg.keyboard.type("45"); await pg.keyboard.press("Enter"); await pg.wait_for_timeout(300)
            check((await pg.locator('td[data-r="15"][data-k="5"]').inner_text()).strip() == "45", "Zelle tippen und mit Enter speichern")
            # Assistent
            await pg.fill("#cmdText", "Wie viele Stunden je Baustelle?"); await pg.press("#cmdText", "Enter"); await pg.wait_for_timeout(900)
            check(await pg.locator("#cmdOut .gcard").count() >= 1, "Auswertung je Baustelle mit Diagramm")
            await pg.fill("#cmdText", "Jonas heute 7 bis 16:30 bei Weber"); await pg.press("#cmdText", "Enter"); await pg.wait_for_timeout(900)
            check(await pg.locator("#cmdOut .plan").count() == 1, "Änderung erst als Vorschau")
            await pg.click("#cmdOut .plan [data-ok]"); await pg.wait_for_timeout(700)
            check("9,00" in await pg.locator('td[data-r="31"]').nth(6).inner_text() or "9,00" in await pg.locator('tr:has(td[data-r="31"])').inner_text(), "neue Zeile, Stunden selbst gerechnet (9,00)")
            # Prüfen
            await pg.click("#t-check"); await pg.wait_for_timeout(300)
            n = await pg.locator("#issues li button").count(); check(n >= 5, f"Prüfen listet Auffälligkeiten ({n})")
            # Verlauf
            await pg.click("#t-hist"); await pg.wait_for_timeout(300)
            check(await pg.locator("#hist li").count() == 3, "Verlauf: Öffnen, Zelle, Assistent (kein Doppel-Eintrag)")
            check("„30“ → „45“" in await pg.locator("#protList").inner_text(), "Protokoll zeigt alt → neu")
            # Belege
            await pg.click("#t-pdf"); await pg.wait_for_timeout(300)
            await pg.select_option("#pdfTarget", label="Lager"); await pg.click('[data-sampledoc="Rechnung_Weber_88213_XRechnung.xml"]'); await pg.wait_for_timeout(700)
            check(await pg.locator("#pdfOut .einv").count() == 1 and await pg.locator("#pdfOut .plan").count() == 1, "XRechnung exakt gelesen, Vorschau fürs Lager")
            await pg.click("#pdfOut .plan [data-ok]"); await pg.wait_for_timeout(500)
            check(await pg.locator("#pdfOut .plan.done").count() == 1, "Beleg übernommen")
            await pg.click('[data-sampledoc="Rechnung_Kaiser_RE-2026-4502_ZUGFeRD.pdf"]'); await pg.wait_for_timeout(1500)
            check("ZUGFeRD" in await pg.locator("#pdfOut .einv-h").inner_text(), "ZUGFeRD-PDF: Rechnung aus dem PDF gelesen")
            check(await pg.locator("#cmdOut .bot").count() >= 2, "Chat bleibt beim Wechsel der Bereiche erhalten")
            # Dokumente
            await pg.select_option('select[aria-label="Beispiel öffnen"]', "Auftraege_2026.xlsx"); await pg.wait_for_timeout(800)
            await pg.click("#t-docs"); await pg.wait_for_timeout(500)
            await pg.fill("#invMail", "kunde@example.com")
            names = []; pg.on("download", lambda d: names.append(d.suggested_filename))
            await pg.click('#docs [data-act="xr"]'); await pg.wait_for_timeout(2000)
            check(any(n.endswith(".xml") for n in names) and any(n.endswith(".pdf") for n in names), "E-Rechnung: XML und PDF erstellt " + str(names))
            im = await pg.locator("#invMsg").inner_text(); check("erstellt" in im, "Bestätigung bleibt stehen " + ("" if "erstellt" in im else repr(im)))
            # Speichern
            async with pg.expect_download() as d:
                await pg.click('button:has-text("Als Excel speichern")')
            dl = await d.value; path = os.path.join(OUT, dl.suggested_filename); await dl.save_as(path)
            check(os.path.getsize(path) > 5000, "Als Excel gespeichert")
            # Persistenz
            await pg.reload(); await pg.wait_for_timeout(700)
            check("Auftraege" in await pg.locator(".topbar, header").first.inner_text(), "nach Neuladen wieder da")
            check(not errs, "keine Skriptfehler " + "; ".join(errs[:3]))
            await b.close()
    finally:
        srv.terminate()
    print("\nALLES OK" if ok else "\nFEHLER"); sys.exit(0 if ok else 1)
asyncio.run(main())
