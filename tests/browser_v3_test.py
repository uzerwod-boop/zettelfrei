"""Ende-zu-Ende-Test für Version 0.3/0.4 (Playwright, openpyxl):
    python tests/browser_v3_test.py
Prüft: neuer Arbeitsplatz, Assistent als Chat (Enter sendet), Auswertung „je …“ mit Diagramm,
Rechnen mit Ergebnissen, Verknüpfung zweier Blätter, Auswertung als Excel-Blatt speichern.
KI-Antworten werden nachgebaut (kein Schlüssel nötig); gerechnet wird echt.
"""
import asyncio, json, os, re, socket, subprocess, sys, tempfile, time
from playwright.async_api import async_playwright
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__))); OUT = tempfile.mkdtemp(prefix="zf-v3-")
ok = True
def check(c, m):
    global ok; print(("ok   " if c else "FAIL ") + m); ok = ok and bool(c)
def free_port():
    s = socket.socket(); s.bind(("127.0.0.1", 0)); p = s.getsockname()[1]; s.close(); return p
SET = {"provider": "openrouter", "keys": {"openrouter": "sk-or-test", "claude": "", "local": ""}, "models": {"openrouter": "", "claude": "", "local": ""}, "url": "http://localhost:11434/v1"}
def answer(text):
    if "je Baustelle" in text:
        return {"abfragen": [{"id": "q1", "blatt": "Stunden Oktober", "spalte": "Stunden", "art": "summe", "gruppieren": "Baustelle"}], "antwort": "Stunden je Baustelle: {q1}"}
    if "Anteil" in text:
        return {"abfragen": [{"id": "q1", "blatt": "Stunden Oktober", "spalte": "Stunden", "art": "summe", "filter": [{"spalte": "Baustelle", "op": "=", "wert": "Neubau Weber"}]},
                             {"id": "q2", "blatt": "Stunden Oktober", "spalte": "Stunden", "art": "summe"}, {"id": "q3", "ausdruck": "q1/q2*100", "prozent": True}],
                "antwort": "Neubau Weber: {q1} von {q2} Stunden, also {q3}."}
    if "offener Rechnung" in text:
        return {"abfragen": [{"id": "q1", "blatt": "Aufträge", "spalte": "Baustelle", "art": "liste", "filter": [{"spalte": "Status", "op": "=", "wert": "offen"}]},
                             {"id": "q2", "blatt": "Stunden Oktober", "spalte": "Stunden", "art": "summe", "filter": [{"spalte": "Baustelle", "op": "in", "wert": "{q1}"}]}],
                "antwort": "Auf Baustellen mit offener Rechnung: {q2} Stunden."}
    return {"antwort": "?"}
async def main():
    port = free_port(); env = dict(os.environ, PORT=str(port)); env.pop("ANTHROPIC_API_KEY", None); env.pop("OPENROUTER_API_KEY", None)
    srv = subprocess.Popen(["node", "server.js"], cwd=ROOT, env=env, stdout=subprocess.DEVNULL); time.sleep(0.8)
    errs = []
    async with async_playwright() as p:
        b = await p.chromium.launch(); ctx = await b.new_context(viewport={"width": 1440, "height": 900}, accept_downloads=True)
        await ctx.add_init_script("try{localStorage.clear();localStorage.setItem('zf-ai-settings'," + json.dumps(json.dumps(SET)) + ")}catch(e){}")
        async def route(r):
            body = json.loads(r.request.post_data); msg = body["messages"][-1]["content"]
            if isinstance(msg, list): msg = " ".join(x.get("text", "") for x in msg)
            m = re.search(r'Nutzers: """(.*?)"""', msg, re.S)
            await r.fulfill(status=200, content_type="application/json", body=json.dumps({"choices": [{"message": {"content": json.dumps(answer(m.group(1) if m else ""))}}]}))
        await ctx.route("https://openrouter.ai/**", route)
        pg = await ctx.new_page(); pg.on("pageerror", lambda e: errs.append(str(e)))
        await pg.goto(f"http://127.0.0.1:{port}/"); await pg.wait_for_timeout(600)
        await pg.select_option("#exampleSel", "Krause_Betrieb_Oktober_2026.xlsx"); await pg.wait_for_timeout(600)
        check(await pg.locator(".rail #t-cmd[aria-selected=true]").count() == 1, "Leiste links, Assistent aktiv")
        check(await pg.locator(".sheet #sheetTabs button").count() == 5, "Blätter unten in der Tabelle")
        check("Zum Prüfen" in await pg.locator("#sumbar").inner_text(), "Kennzahlen über der Tabelle")
        await pg.fill("#cmdText", "Stunden je Baustelle?"); await pg.press("#cmdText", "Enter"); await pg.wait_for_timeout(900)
        check(await pg.locator("#cmdOut .me").count() == 1 and (await pg.locator("#cmdText").input_value()) == "", "Enter sendet, Nachricht steht im Chat")
        check(await pg.locator("#cmdOut .gcard .gbar").count() >= 4, "Balkendiagramm je Baustelle")
        check("Neubau Weber" in await pg.locator("#cmdOut .gcard .gbar").first.inner_text(), "größte Baustelle oben")
        await pg.fill("#cmdText", "Anteil Neubau Weber an allen Stunden?"); await pg.press("#cmdText", "Enter"); await pg.wait_for_timeout(900)
        check(re.search(r"\d+,\d %", await pg.locator("#cmdOut .msg").nth(1).inner_text()) is not None, "Prozent aus zwei Ergebnissen berechnet")
        await pg.fill("#cmdText", "Wie viele Stunden auf Baustellen mit offener Rechnung?"); await pg.press("#cmdText", "Enter"); await pg.wait_for_timeout(900)
        t3 = await pg.locator("#cmdOut .msg").nth(2).inner_text(); check("26,25 Stunden" in t3, "zwei Blätter verknüpft (26,25 Std.) " + ("" if "26,25" in t3 else repr(t3)))
        check(await pg.locator("#cmdOut .msg").count() == 3, "Chatverlauf bleibt stehen")
        await pg.locator("#cmdOut .gcard [data-sheet]").first.click(); await pg.wait_for_timeout(400)
        check(await pg.locator("#sheetTabs button", has_text="Auswertung").count() == 1, "Auswertung als neues Blatt")
        async with pg.expect_download() as d:
            await pg.click("#saveBtn")
        dl = await d.value; path = os.path.join(OUT, dl.suggested_filename); await dl.save_as(path)
        await b.close()
    srv.terminate()
    import openpyxl
    wbk = openpyxl.load_workbook(path); names = wbk.sheetnames
    aus = [n for n in names if n.startswith("Auswertung")]
    check(bool(aus) and wbk[aus[0]]["A1"].value == "Baustelle" and wbk[aus[0]]["A2"].value == "Neubau Weber", "Auswertungsblatt in der gespeicherten Excel-Datei")
    check("Kunden" in names and wbk["Kunden"].sheet_state == "hidden", "übrige Blätter unverändert")
    check(not errs, "keine Skriptfehler " + (str(errs) if errs else ""))
    print("\nALLES OK" if ok else "\nFEHLER"); sys.exit(0 if ok else 1)
asyncio.run(main())
