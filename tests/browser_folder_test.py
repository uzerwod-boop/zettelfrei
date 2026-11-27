"""Ende-zu-Ende: Ordner mit Excel-Dateien und einer PDF-Liste öffnen -> eine Arbeitsmappe mit allen Blättern.
    cd web && npm run build && cd .. && python tests/browser_folder_test.py
KI-Antworten für die PDF-Liste werden nachgebaut; Zusammenführen und Speichern laufen echt.
"""
import asyncio, json, os, re, shutil, socket, subprocess, sys, tempfile, time
from playwright.async_api import async_playwright
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__))); OUT = tempfile.mkdtemp(prefix="zf-dir-")
ok = True
def check(c, m):
    global ok; print(("ok   " if c else "FAIL ") + m); ok = ok and bool(c)
def free_port():
    s = socket.socket(); s.bind(("127.0.0.1", 0)); p = s.getsockname()[1]; s.close(); return p
SET = {"provider": "openrouter", "keys": {"openrouter": "sk-or-test", "claude": "", "local": ""}, "models": {"openrouter": "", "claude": "", "local": ""}, "url": "http://localhost:11434/v1"}
ROWS = [[f"Firma {i:02d} GmbH", "2. Qualifizierung", "0,00 €", "O", f"{20 + i % 6}.11.2026"] for i in range(1, 13)]
def chunk(part, parts):
    per = -(-len(ROWS) // parts); return {"titel": "Aktivitäten", "spalten": ["Verkaufschance", "Phase", "Betrag", "Zuständig", "Anruf"], "zeilen": ROWS[(part - 1) * per:part * per], "gesamt": 12 if part == 1 else None, "unsicher": []}

async def main():
    port = free_port(); env = dict(os.environ, PORT=str(port)); [env.pop(k, None) for k in ("ANTHROPIC_API_KEY", "OPENROUTER_API_KEY", "ZF_KLASSISCH")]
    srv = subprocess.Popen(["node", "server.js"], cwd=ROOT, env=env, stdout=subprocess.DEVNULL); time.sleep(0.8)
    errs = []
    folder = os.path.join(OUT, "Test-Ordner"); os.makedirs(folder)
    shutil.copy(os.path.join(ROOT, "beispiele", "Auftraege_2026.xlsx"), os.path.join(folder, "CRM.xlsx"))
    subprocess.run(["node", "-e", """
global.window=global;const XLSX=require('./vendor/xlsx.full.min.js');
const ws=XLSX.utils.aoa_to_sheet([['Kampagne','Mailings','Klicks','Status'],['Erstkontakt',4,0,'Neu'],['Qualifizierung',4,0,'Neu']]);
const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'Sheet1');require('fs').writeFileSync(process.argv[1],Buffer.from(XLSX.write(wb,{type:'array',bookType:'xlsx'})));""", os.path.join(folder, "Marketing-Automation.xlsx")], cwd=ROOT, check=True)
    open(os.path.join(folder, ".DS_Store"), "w").write("x")
    try:
        async with async_playwright() as p:
            b = await p.chromium.launch(); ctx = await b.new_context(viewport={"width": 1440, "height": 900}, accept_downloads=True)
            pg0 = await ctx.new_page()
            await pg0.set_content("<table>" + "".join(f"<tr>{''.join(f'<td>{v}</td>' for v in r)}</tr>" for r in ROWS) + "</table>")
            await pg0.pdf(path=os.path.join(folder, "Pipeline.pdf"), format="A4")
            await ctx.add_init_script("try{if(!sessionStorage.getItem('t')){sessionStorage.setItem('t',1);localStorage.clear();}localStorage.setItem('zf-ai-settings'," + json.dumps(json.dumps(SET)) + ")}catch(e){}")
            async def route(r):
                body = json.loads(r.request.post_data); msg = body["messages"][-1]["content"]
                text = " ".join(x.get("text", "") for x in msg) if isinstance(msg, list) else msg
                m = re.search(r"Teil (\d+) von (\d+)", text)
                await r.fulfill(status=200, content_type="application/json", body=json.dumps({"choices": [{"message": {"content": json.dumps(chunk(int(m.group(1)), int(m.group(2))) if m else {"antwort": "?"})}}]}))
            await ctx.route("https://openrouter.ai/**", route)
            pg = await ctx.new_page(); pg.on("pageerror", lambda e: errs.append(str(e)))
            await pg.goto(f"http://127.0.0.1:{port}/"); await pg.wait_for_selector("#emptyFolder")
            # Ordner über „Ordner öffnen“ wählen
            await pg.locator("#emptyFolder ~ input, .empty input[webkitdirectory]").last.set_input_files(folder)
            await pg.wait_for_selector(".sheettabs", timeout=10000)
            tabs = await pg.locator(".sheettabs [role=tab]").all_inner_texts()
            check("Marketing-Automation" in tabs, "Excel mit „Sheet1“ heißt nach der Datei: " + ", ".join(tabs))
            check(len(tabs) >= 2, f"{len(tabs)} Blätter aus zwei Excel-Dateien")
            title = await pg.inner_text(".filebar, header"); check("Test-Ordner" in title, "Arbeitsmappe heißt wie der Ordner")
            await pg.wait_for_selector(".listcard", timeout=20000)
            check(await pg.locator(".listcard", has_text="12 Einträge").count() == 1, "PDF aus dem Ordner wird als Liste gelesen")
            await pg.click("#listCreate"); await pg.wait_for_selector(".sheettabs >> text=Aktivitäten", timeout=5000)
            tabs = await pg.locator(".sheettabs [role=tab]").all_inner_texts(); check("Aktivitäten" in tabs, "drittes Blatt aus der PDF: " + ", ".join(tabs))
            await pg.screenshot(path=os.path.join(OUT, "ordner.png"))
            # Speichern: alle Blätter in einer Datei
            async with pg.expect_download() as dl: await pg.click("text=Als Excel speichern")
            path = os.path.join(OUT, "out.xlsx"); await (await dl.value).save_as(path)
            names = subprocess.run(["node", "-e", "global.window=global;const X=require('./vendor/xlsx.full.min.js');console.log(X.read(require('fs').readFileSync(process.argv[1])).SheetNames.join('|'))", path], cwd=ROOT, capture_output=True, text=True).stdout.strip()
            check("Marketing-Automation" in names and "Aktivitäten" in names, "gespeicherte Datei enthält alle Blätter: " + names)
            # Weitere Datei per Rechtsklick auf ein Blatt hinzufügen
            await pg.locator(".sheettabs [role=tab]").first.click(button="right"); await pg.click("text=Datei als Blatt hinzufügen")
            await pg.locator(".sheettabs input[type=file]").set_input_files(os.path.join(ROOT, "beispiele", "Material_Lager.xlsx"))
            await pg.wait_for_selector(".sheettabs >> text=Lager", timeout=5000); check(True, "Datei per Rechtsklick als Blatt hinzugefügt")
            # PDF auf die Tabelle gezogen (nicht als Liste gewählt): Knopf „Als neues Blatt übernehmen“
            data = open(os.path.join(folder, "Pipeline.pdf"), "rb").read()
            dt = await pg.evaluate_handle("(b)=>{const d=new DataTransfer();d.items.add(new File([new Uint8Array(b)],'Pipeline 2.pdf',{type:'application/pdf'}));return d;}", list(data))
            await pg.dispatch_event(".sheetwrap", "drop", {"dataTransfer": dt})
            try: await pg.wait_for_selector("#asList", timeout=20000)
            except Exception: await pg.screenshot(path=os.path.join(OUT, "fail.png")); print(await pg.inner_text("#pdfOut")); raise
            await pg.click("#asList")
            await pg.wait_for_selector(".listcard", timeout=20000); await pg.click("#listCreate")
            await pg.wait_for_selector(".sheettabs >> text=Aktivitäten 2", timeout=5000); check(True, "Beleg-Ergebnis lässt sich als neues Blatt übernehmen")
            check(not errs, "keine JS-Fehler " + "; ".join(errs[:3]))
            await b.close()
    finally:
        srv.terminate()
    print("Bilder:", OUT); sys.exit(0 if ok else 1)
asyncio.run(main())
