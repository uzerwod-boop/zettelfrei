"""Ende-zu-Ende: Liste als PDF (z. B. Odoo-Pipeline) ohne geöffnete Datei hineinziehen -> neues Blatt.
    cd web && npm run build && cd .. && python tests/browser_pdflist_test.py [eigene.pdf]
KI-Antworten werden nachgebaut (je Bildteil ein Stück der Liste, mit einer Doppelzeile an der Schnittkante).
"""
import asyncio, json, os, re, socket, subprocess, sys, tempfile, time
from playwright.async_api import async_playwright
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__))); OUT = tempfile.mkdtemp(prefix="zf-pdfl-")
ok = True
def check(c, m):
    global ok; print(("ok   " if c else "FAIL ") + m); ok = ok and bool(c)
def free_port():
    s = socket.socket(); s.bind(("127.0.0.1", 0)); p = s.getsockname()[1]; s.close(); return p
SET = {"provider": "openrouter", "keys": {"openrouter": "sk-or-test", "claude": "", "local": ""}, "models": {"openrouter": "", "claude": "", "local": ""}, "url": "http://localhost:11434/v1"}
SP = ["Verkaufschance", "Phase", "Betrag", "Zuständig", "Anruf", "To-do"]
ROWS = [[f"Firma {i:02d} GmbH", "2. Qualifizierung" if i % 3 else "1. Erstkontakt", "0,00 €", "O" if i % 2 else "A",
         f"{20 + i % 6}.11.2026" if i % 2 else "", "" if i % 2 else f"{25 + i % 5}.11.2026"] for i in range(1, 31)]
def chunk(part, parts):
    per = -(-len(ROWS) // parts); a = (part - 1) * per; rows = ROWS[a:a + per]
    if part > 1: rows = [ROWS[a - 1]] + rows          # Schnittkante doppelt gelesen
    return {"titel": "Pipeline", "spalten": SP, "zeilen": rows, "gesamt": 30 if part == 1 else None, "unsicher": ["Zeile 4, Anruf: Datum verschmiert"] if part == 2 else []}

async def main():
    port = free_port(); env = dict(os.environ, PORT=str(port)); [env.pop(k, None) for k in ("ANTHROPIC_API_KEY", "OPENROUTER_API_KEY", "ZF_KLASSISCH")]
    srv = subprocess.Popen(["node", "server.js"], cwd=ROOT, env=env, stdout=subprocess.DEVNULL); time.sleep(0.8)
    errs, calls, maxt = [], [], []
    try:
        async with async_playwright() as p:
            b = await p.chromium.launch(); ctx = await b.new_context(viewport={"width": 1440, "height": 900})
            # Test-PDF: 3 Seiten Hochformat mit einer Liste
            pg0 = await ctx.new_page()
            await pg0.set_content("<style>td{border:1px solid #999;padding:6px;font:11px sans-serif}</style><table>" + "".join(f"<tr>{''.join(f'<td>{v}</td>' for v in r)}</tr>" for r in ROWS * 3) + "</table>")
            pdf = sys.argv[1] if len(sys.argv) > 1 else os.path.join(OUT, "Pipeline.pdf")
            if len(sys.argv) <= 1: await pg0.pdf(path=pdf, format="A4")
            await ctx.add_init_script("try{if(!sessionStorage.getItem('t')){sessionStorage.setItem('t',1);localStorage.clear();}localStorage.setItem('zf-ai-settings'," + json.dumps(json.dumps(SET)) + ")}catch(e){}")
            async def route(r):
                body = json.loads(r.request.post_data); msg = body["messages"][-1]["content"]
                text = " ".join(x.get("text", "") for x in msg) if isinstance(msg, list) else msg
                imgs = sum(1 for x in msg if isinstance(x, dict) and x.get("type") == "image_url") if isinstance(msg, list) else 0
                m = re.search(r"Teil (\d+) von (\d+)", text); calls.append((m and m.groups(), imgs, "GENAU diese Spalten" in text))
                await asyncio.sleep(0.2)
                res = chunk(int(m.group(1)), int(m.group(2))) if m else {"antwort": "?"}
                maxt.append(body.get("max_tokens"))
                if m and m.group(1) == "2" and len(sys.argv) <= 1:   # Teil 2: zweimal abgeschnittene Antwort, dann in zwei Hälften
                    n2 = sum(1 for c in calls if c[0] and c[0][0] == "2")
                    if n2 <= 2: await r.fulfill(status=200, content_type="application/json", body=json.dumps({"choices": [{"message": {"content": json.dumps(res)[:120]}}]})); return
                    z = res["zeilen"]; h = len(z) // 2; res = dict(res, zeilen=z[:h + 1] if n2 == 3 else z[h:])
                await r.fulfill(status=200, content_type="application/json", body=json.dumps({"choices": [{"message": {"content": json.dumps(res)}}]}))
            await ctx.route("https://openrouter.ai/**", route)
            pg = await ctx.new_page(); pg.on("pageerror", lambda e: errs.append(str(e)))
            await pg.goto(f"http://127.0.0.1:{port}/"); await pg.wait_for_selector(".empty")
            check(await pg.locator(".empty", has_text="PDF").count() == 1, "Startseite erwähnt PDF-Listen")
            # PDF auf die leere Startseite ziehen
            data = open(pdf, "rb").read()
            dt = await pg.evaluate_handle("(b)=>{const d=new DataTransfer();d.items.add(new File([new Uint8Array(b)],'Pipeline.pdf',{type:'application/pdf'}));return d;}", list(data))
            await pg.dispatch_event(".sheetwrap", "drop", {"dataTransfer": dt})
            await pg.wait_for_selector(".listcard", timeout=30000)
            n = len(calls); check(n >= 2 and all(c[1] == 1 for c in calls), f"{n} Anfragen, je ein Bildteil")
            check(all(x == 8000 for x in maxt), "Antwortlimit 8000 Tokens für Listen")
            if len(sys.argv) <= 1: check(sum(1 for c in calls if c[0] and c[0][0] == "2") == 4, "abgeschnittene Antwort: neu versucht, dann in zwei Hälften gelesen")
            check(not calls[0][2] and all(c[2] for c in calls[1:]), "ab Teil 2 mit festen Spalten")
            if len(sys.argv) <= 1:
                head = await pg.inner_text(".listcard .einv-h"); check("30 Einträge" in head, "30 Einträge (Kantendoppel entfernt): " + head)
                check(await pg.locator(".listcard", has_text="Vollständig").count() == 1, "Abgleich mit Gesamtzahl laut Dokument")
                check(await pg.locator(".listcard .planwarn", has_text="verschmiert").count() == 1, "unsichere Stellen werden gezeigt")
            await pg.screenshot(path=os.path.join(OUT, "liste.png"))
            await pg.fill(".listcard input", "Aktivitäten"); await pg.click("#listCreate")
            await pg.wait_for_selector(".sheettabs >> text=Aktivitäten", timeout=5000)
            check(True, "Blatt „Aktivitäten“ angelegt, ohne vorher eine Excel-Datei zu öffnen")
            cells = await pg.locator("#grid tbody tr").count()
            check(cells >= 30, f"Tabelle zeigt {cells} Zeilen")
            await pg.screenshot(path=os.path.join(OUT, "blatt.png"))
            check(not errs, "keine JS-Fehler " + "; ".join(errs[:3]))
            await b.close()
    finally:
        srv.terminate()
    print("Bilder:", OUT); sys.exit(0 if ok else 1)
asyncio.run(main())
