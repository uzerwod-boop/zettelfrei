"""Screenshots für README (hell/dunkel) und das Vorschaubild für GitHub/LinkedIn.
    python scripts/make_screens.py
"""
import asyncio, os, subprocess, time, socket
from playwright.async_api import async_playwright
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__))); D = os.path.join(ROOT, "docs")

def free_port():
    s = socket.socket(); s.bind(("127.0.0.1", 0)); p = s.getsockname()[1]; s.close(); return p

async def app_shot(b, port, scheme, out):
    ctx = await b.new_context(viewport={"width": 1360, "height": 820}, device_scale_factor=2, color_scheme=scheme)
    pg = await ctx.new_page(); await pg.goto(f"http://127.0.0.1:{port}/"); await pg.wait_for_timeout(600)
    await pg.select_option("#exampleSel", "Material_Lager.xlsx"); await pg.wait_for_timeout(400)
    await pg.click("#t-pdf"); await pg.click('[data-sampledoc="Rechnung_Kaiser_RE-2026-4502_ZUGFeRD.pdf"]'); await pg.wait_for_timeout(1500)
    await pg.add_style_tag(content=".toast,#toast{display:none!important}")
    await pg.screenshot(path=out); await ctx.close()

async def social(b, out):
    mark = open(os.path.join(D, "brand", "zettelfrei-mark.svg")).read()
    font = "file://" + os.path.join(ROOT, "fonts")
    html = f"""<html><head><style>
@font-face{{font-family:A;src:url('{font}/Archivo.ttf');font-weight:100 900}}
@font-face{{font-family:P;src:url('{font}/IBMPlexSans.ttf');font-weight:100 700}}
body{{margin:0;width:1280px;height:640px;background:#f4f5f1;font-family:P;color:#1d2528;position:relative;overflow:hidden}}
.l{{position:absolute;left:84px;top:96px;width:560px}}
.mk svg{{width:96px;height:96px}}
h1{{font:800 96px/1 A;letter-spacing:-.02em;margin:26px 0 18px}}
.t{{font:600 34px/1.25 P;margin:0 0 30px}} .t b{{background:linear-gradient(transparent 58%,#f6dd6a 58%);font-weight:700}}
.chips{{display:flex;flex-wrap:wrap;gap:10px}} .chips span{{font:500 19px P;padding:6px 14px;border:2px solid #0f6b5c;color:#0f6b5c;border-radius:999px}}
.by{{position:absolute;left:84px;bottom:48px;font:500 20px P;color:#5d686b}}
.s{{position:absolute;left:690px;top:70px;width:760px;border-radius:14px;box-shadow:0 20px 60px rgba(15,60,50,.25);border:1px solid #d9ddd6}}
</style></head><body><div class="l"><div class="mk">{mark}</div><h1>Zettelfrei</h1>
<p class="t">Excel bleibt.<br><b>Das Tippen fällt weg.</b></p>
<div class="chips"><span>E-Rechnung</span><span>Arbeitszeit</span><span>Mahnwesen</span><span>Abnahme</span></div></div>
<img class="s" src="file://{D}/screenshot-hell.png"><div class="by">Open Source · Bohdan Sychevskyi</div></body></html>"""
    p = os.path.join(D, "_social.html"); open(p, "w").write(html)
    ctx = await b.new_context(viewport={"width": 1280, "height": 640}, device_scale_factor=1)
    pg = await ctx.new_page(); await pg.goto("file://" + p); await pg.wait_for_timeout(500)
    await pg.screenshot(path=out); await ctx.close(); os.remove(p)

async def main():
    port = free_port(); env = dict(os.environ, PORT=str(port)); env.pop("ANTHROPIC_API_KEY", None); env.pop("OPENROUTER_API_KEY", None)
    srv = subprocess.Popen(["node", "server.js"], cwd=ROOT, env=env, stdout=subprocess.DEVNULL); time.sleep(0.8)
    try:
        async with async_playwright() as p:
            b = await p.chromium.launch()
            await app_shot(b, port, "light", os.path.join(D, "screenshot-hell.png"))
            await app_shot(b, port, "dark", os.path.join(D, "screenshot-dunkel.png"))
            await social(b, os.path.join(D, "brand", "social-preview.png"))
            await b.close()
    finally: srv.terminate()
asyncio.run(main())
