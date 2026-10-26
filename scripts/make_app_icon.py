"""App-Symbol für die Desktop-Version (1024 × 1024, macOS-Raster) aus dem Logo."""
import asyncio, pathlib
from playwright.async_api import async_playwright
ROOT = pathlib.Path(__file__).resolve().parent.parent
mark = (ROOT / "docs/brand/zettelfrei-mark-dark.svg").read_text().replace('stroke="#3a4446"', 'stroke="#56656a"')
HTML = f"""<html><body style="margin:0;background:transparent">
<div style="position:absolute;left:100px;top:100px;width:824px;height:824px;border-radius:186px;
 background:linear-gradient(160deg,#26322f 0%,#18201f 100%);box-shadow:inset 0 2px 0 rgba(255,255,255,.08);
 display:flex;align-items:center;justify-content:center">
 <div style="width:500px;height:500px">{mark.replace('width="64" height="64"', 'width="500" height="500"')}</div></div></body></html>"""
async def main():
    async with async_playwright() as p:
        b = await p.chromium.launch(); pg = await b.new_page(viewport={"width": 1024, "height": 1024})
        await pg.set_content(HTML)
        await pg.screenshot(path=str(ROOT / "web/src-tauri/app-icon.png"), omit_background=True); await b.close()
asyncio.run(main())
