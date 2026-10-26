"""Ende-zu-Ende-Test im echten Browser (optional, braucht Python + Playwright):
    pip install playwright && python -m playwright install chromium
    python tests/browser_test.py
Testet: ohne KI, OpenRouter, Claude API mit Schlüssel im Browser, lokales Modell, eigener Server.
Die KI-Antworten werden nachgebaut, es wird kein echter Schlüssel gebraucht.
"""
import asyncio, json, os, subprocess, sys, time, socket, threading, http.server
from playwright.async_api import async_playwright

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ANSWER = {"aenderungen": [{"typ": "aendern", "blatt": "Oktober 2026", "zeile": 11, "werte": {"Projekt": "Küche Yilmaz"}}],
          "abfragen": [{"id": "q1", "blatt": "Oktober 2026", "spalte": "Stunden", "art": "summe", "filter": [{"spalte": "Mitarbeiter", "op": "=", "wert": "Kevin"}]}],
          "antwort": "Kevin hatte {q1} Stunden.", "rueckfrage": None}
ok = True
def check(cond, msg):
    global ok
    print(("ok   " if cond else "FAIL ") + msg)
    ok = ok and cond

def free_port():
    s = socket.socket(); s.bind(("127.0.0.1", 0)); p = s.getsockname()[1]; s.close(); return p

class Fake(http.server.BaseHTTPRequestHandler):
    def do_POST(self):
        self.rfile.read(int(self.headers["content-length"]))
        body = json.dumps({"content": [{"type": "text", "text": json.dumps(ANSWER)}]}).encode()
        self.send_response(200); self.send_header("content-type", "application/json"); self.end_headers(); self.wfile.write(body)
    def log_message(self, *a): pass

async def scenario(b, url, name, setup=None, route=None):
    ctx = await b.new_context(accept_downloads=True)
    pg = await ctx.new_page()
    errs = []
    pg.on("pageerror", lambda e: errs.append(str(e)))
    seen = {}
    if route:
        await ctx.route(route["pattern"], lambda r: route["handler"](r, seen))
    await pg.goto(url); await pg.wait_for_timeout(800)
    if setup: await setup(pg); await pg.wait_for_timeout(600)
    chip = await pg.locator("#aiText").inner_text()
    print(f"\n[{name}] KI-Status: {chip}")
    await pg.fill("#cmdText", "Wie viele Stunden hatte Kevin? Und die Fahrt am 15.10. gehört zu Küche Yilmaz.")
    await pg.click("#cmdBtn"); await pg.wait_for_timeout(1200)
    if name == "ohne KI":
        check("KI einrichten" in await pg.locator("#cmdOut").inner_text(), "ohne KI: Hinweis mit Link zu den Einstellungen")
        td = pg.locator('td[data-r="4"][data-k="4"]'); await td.click(); await pg.keyboard.type("9"); await pg.keyboard.press("Enter")
        check((await pg.locator("tr.foot td").nth(4).inner_text()) == "56,5", "ohne KI: Formel/Summe rechnet nach Bearbeiten")
        async with pg.expect_download() as d:
            await pg.click("#saveBtn")
        dl = await d.value
        check(dl.suggested_filename == "Stundenzettel_Oktober_2026.xlsx", "ohne KI: Speichern lädt .xlsx herunter")
    else:
        ans = await pg.locator(".answer").first.inner_text()
        check("Kevin hatte 22 Stunden" in ans, f"{name}: Antwort mit exakt gerechneter Summe")
        await pg.click(".plan [data-ok]"); await pg.wait_for_timeout(300)
        check((await pg.locator('td[data-r="11"][data-k="2"]').inner_text()) == "Küche Yilmaz", f"{name}: Änderung übernommen")
        for k, v in seen.items(): check(v, f"{name}: {k}")
    check(not errs, f"{name}: keine Skriptfehler {errs if errs else ''}")
    await ctx.close()

async def main():
    port = free_port(); fake_port = free_port()
    fake = http.server.HTTPServer(("127.0.0.1", fake_port), Fake); threading.Thread(target=fake.serve_forever, daemon=True).start()
    # 1) ohne Schlüssel gestartet: KI im Browser einstellbar
    env = dict(os.environ, PORT=str(port)); env.pop("ANTHROPIC_API_KEY", None)
    srv = subprocess.Popen(["node", "server.js"], cwd=ROOT, env=env, stdout=subprocess.DEVNULL)
    time.sleep(0.8)
    async with async_playwright() as p:
        b = await p.chromium.launch()
        base = f"http://127.0.0.1:{port}/klassisch/"
        await scenario(b, base, "ohne KI")

        async def set_claude(pg):
            await pg.click("#aiChip"); await pg.select_option("#setProvider", "claude"); await pg.fill("#setKey", "sk-ant-test"); await pg.click("#setSave")
        async def claude_route(r, seen):
            h = r.request.headers; body = json.loads(r.request.post_data)
            seen["Schlüssel im Header"] = h.get("x-api-key") == "sk-ant-test"
            seen["Browser-Header gesetzt"] = h.get("anthropic-dangerous-direct-browser-access") == "true"
            seen["Tabelle im Prompt"] = "Blatt „Oktober 2026“" in body["messages"][0]["content"][-1]["text"]
            await r.fulfill(status=200, content_type="application/json", headers={"access-control-allow-origin": "*"},
                            body=json.dumps({"content": [{"type": "text", "text": "```json\n" + json.dumps(ANSWER) + "\n```"}]}))
        await scenario(b, base, "Claude API im Browser", set_claude, {"pattern": "https://api.anthropic.com/**", "handler": claude_route})

        async def set_or(pg):
            await pg.click("#aiChip"); await pg.select_option("#setProvider", "openrouter"); await pg.fill("#setKey", "sk-or-test"); await pg.click("#setSave")
        async def or_route(r, seen):
            h = r.request.headers; body = json.loads(r.request.post_data)
            seen["Bearer-Schlüssel"] = h.get("authorization") == "Bearer sk-or-test"
            seen["Standardmodell anthropic/claude-sonnet-5.5"] = body.get("model") == "anthropic/claude-sonnet-5.5"
            seen["App-Name an OpenRouter"] = h.get("x-openrouter-title") == "Zettelfrei"
            await r.fulfill(status=200, content_type="application/json", headers={"access-control-allow-origin": "*"},
                            body=json.dumps({"choices": [{"message": {"content": json.dumps(ANSWER)}}]}))
        await scenario(b, base, "OpenRouter", set_or, {"pattern": "https://openrouter.ai/**", "handler": or_route})

        async def set_local(pg):
            await pg.click("#aiChip"); await pg.select_option("#setProvider", "local"); await pg.fill("#setUrl", "http://localhost:11434/v1"); await pg.fill("#setModel", "qwen2.5"); await pg.click("#setSave")
        async def local_route(r, seen):
            body = json.loads(r.request.post_data)
            seen["OpenAI-Format und Modell"] = body.get("model") == "qwen2.5" and body["messages"][0]["role"] == "user"
            await r.fulfill(status=200, content_type="application/json", headers={"access-control-allow-origin": "*"},
                            body=json.dumps({"choices": [{"message": {"content": json.dumps(ANSWER)}}]}))
        await scenario(b, base, "lokales Modell", set_local, {"pattern": "http://localhost:11434/**", "handler": local_route})
        srv.terminate()

        # 2) mit Schlüssel auf dem Server: Automatisch nutzt den Server
        port2 = free_port()
        env2 = dict(os.environ, PORT=str(port2), ANTHROPIC_API_KEY="server-key", ZF_UPSTREAM=f"http://127.0.0.1:{fake_port}/v1/messages")
        srv2 = subprocess.Popen(["node", "server.js"], cwd=ROOT, env=env2, stdout=subprocess.DEVNULL); time.sleep(0.8)
        await scenario(b, f"http://127.0.0.1:{port2}/klassisch/", "eigener Server")
        srv2.terminate()
        await b.close()
    fake.shutdown()
    print("\nALLES OK" if ok else "\nFEHLER")
    sys.exit(0 if ok else 1)

asyncio.run(main())
