"""Demo-GIF für das README: ZUGFeRD-Rechnung einlesen, Lagerbestand steigt.
    python scripts/make_gif.py   (braucht Playwright und ffmpeg)
"""
import asyncio, os, subprocess, time, socket, tempfile, glob
from playwright.async_api import async_playwright
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__))); OUT = os.path.join(ROOT, "docs", "demo.gif")
START = 0
W, H = 1200, 980
CURSOR = """(()=>{const c=document.createElement('div');c.id='fakecur';
c.style.cssText='position:fixed;z-index:99999;left:0;top:0;width:22px;height:22px;pointer-events:none;transition:transform .45s cubic-bezier(.3,.7,.3,1)';
c.innerHTML='<svg viewBox="0 0 24 24" width="22" height="22"><path d="M3 2l7 19 2.6-7.6L20 11z" fill="#1d2528" stroke="#fff" stroke-width="1.6" stroke-linejoin="round"/></svg>';
document.body.appendChild(c);window.__cur=(x,y)=>{c.style.transform=`translate(${x}px,${y}px)`};
document.addEventListener('mousedown',()=>{c.animate([{scale:1},{scale:.8},{scale:1}],{duration:250})});})()"""

def free_port():
    s = socket.socket(); s.bind(("127.0.0.1", 0)); p = s.getsockname()[1]; s.close(); return p

async def move_click(pg, sel, wait=650):
    box = await pg.locator(sel).first.bounding_box()
    x, y = box["x"] + box["width"] / 2, box["y"] + box["height"] / 2
    await pg.evaluate(f"window.__cur({x-3},{y-3})"); await pg.wait_for_timeout(wait)
    await pg.locator(sel).first.click()

async def main():
    port = free_port(); env = dict(os.environ, PORT=str(port)); env.pop("ANTHROPIC_API_KEY", None); env.pop("OPENROUTER_API_KEY", None)
    srv = subprocess.Popen(["node", "server.js"], cwd=ROOT, env=env, stdout=subprocess.DEVNULL); time.sleep(0.8)
    tmp = tempfile.mkdtemp(prefix="zf-gif-")
    try:
        async with async_playwright() as p:
            b = await p.chromium.launch()
            ctx = await b.new_context(viewport={"width": W, "height": H}, record_video_dir=tmp, record_video_size={"width": W, "height": H}, color_scheme="light")
            t0 = time.time(); pg = await ctx.new_page(); await pg.goto(f"http://127.0.0.1:{port}/"); await pg.wait_for_timeout(500)
            await pg.add_style_tag(content=".toast{display:none!important}")
            await pg.evaluate(CURSOR); await pg.evaluate("window.__cur(600,400)")
            await pg.select_option("#exampleSel", "Material_Lager.xlsx"); await pg.wait_for_timeout(400)
            global START; START = time.time() - t0; await pg.wait_for_timeout(700)
            await move_click(pg, "#t-pdf"); await pg.wait_for_timeout(500)
            await move_click(pg, '[data-sampledoc="Rechnung_Kaiser_RE-2026-4502_ZUGFeRD.pdf"]'); await pg.wait_for_timeout(1800)
            await pg.wait_for_timeout(900)
            bb = await pg.locator("#pdfOut .plan [data-ok]").bounding_box(); print("Knopf y", bb["y"], "von", H)
            await move_click(pg, "#pdfOut .plan [data-ok]", 900); await pg.wait_for_timeout(600)
            box = await pg.locator('td[data-r="2"][data-k="3"]').bounding_box()
            await pg.evaluate(f"window.__cur({box['x']+box['width']/2},{box['y']+box['height']/2+4})"); await pg.wait_for_timeout(2600)
            await ctx.close(); await b.close()
        vid = glob.glob(os.path.join(tmp, "*.webm"))[0]
        pal = os.path.join(tmp, "pal.png"); vf = "fps=12,scale=960:-1:flags=lanczos"
        subprocess.run(["ffmpeg", "-y", "-loglevel", "error", "-ss", f"{START:.2f}", "-i", vid, "-vf", vf + ",palettegen=max_colors=128:stats_mode=diff", pal], check=True)
        subprocess.run(["ffmpeg", "-y", "-loglevel", "error", "-ss", f"{START:.2f}", "-i", vid, "-i", pal, "-lavfi", vf + "[x];[x][1:v]paletteuse=dither=bayer:bayer_scale=4:diff_mode=rectangle", OUT], check=True)
        print("ok", OUT, os.path.getsize(OUT) // 1024, "KB")
    finally: srv.terminate()
asyncio.run(main())
