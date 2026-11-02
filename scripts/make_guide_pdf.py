"""Erste-Schritte-Anleitung als PDF (für E-Mails an Pilotbetriebe): python scripts/make_guide_pdf.py"""
import asyncio, os, markdown
from playwright.async_api import async_playwright
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__))); D = os.path.join(ROOT, "docs"); F = "file://" + os.path.join(ROOT, "fonts")
md = open(os.path.join(D, "Erste-Schritte.md"), encoding="utf-8").read()
body = markdown.markdown(md, extensions=["tables"]).replace("<blockquote>", '<div class="tip">').replace("</blockquote>", "</div>")
mark = open(os.path.join(D, "brand", "zettelfrei-logo.svg"), encoding="utf-8").read()
html = f"""<html><head><meta charset="utf-8"><style>
@font-face{{font-family:A;src:url('{F}/Archivo.ttf');font-weight:100 900}}
@font-face{{font-family:P;src:url('{F}/IBMPlexSans.ttf');font-weight:100 700}}
@font-face{{font-family:M;src:url('{F}/IBMPlexMono-Regular.ttf')}}
@page{{size:A4;margin:16mm 16mm 18mm}}
body{{font:10.5pt/1.5 P;color:#1d2528}}
.logo svg{{height:34px;width:auto}} h1{{font:800 22pt/1.1 A;margin:14px 0 6px;letter-spacing:-.01em}}
h2{{font:800 14pt A;margin:20px 0 6px;color:#0f6b5c;break-after:avoid}} p{{margin:6px 0}}
table{{border-collapse:collapse;width:100%;margin:8px 0;font-size:9.5pt}} th,td{{border:1px solid #d9ddd6;padding:6px 8px;vertical-align:top;text-align:left}} th{{background:#eef1ec}}
code{{font:9pt M;background:#eef1ec;padding:1px 4px;border-radius:4px}} a{{color:#0f6b5c}}
.tip{{border-left:4px solid #f6dd6a;background:#fbf7e3;padding:6px 12px;margin:10px 0;border-radius:0 8px 8px 0}}
hr{{border:0;border-top:1px solid #d9ddd6;margin:14px 0}} li{{margin:3px 0}} ul,ol{{padding-left:20px}}
</style></head><body><div class="logo">{mark}</div>{body.replace('<h1>Zettelfrei – Erste Schritte</h1>','<h1>Erste Schritte</h1>')}</body></html>"""
async def main():
    p = os.path.join(D, "_guide.html"); open(p, "w", encoding="utf-8").write(html)
    async with async_playwright() as pw:
        b = await pw.chromium.launch(); pg = await b.new_page(); await pg.goto("file://" + p); await pg.wait_for_timeout(400)
        await pg.pdf(path=os.path.join(D, "Zettelfrei-Erste-Schritte.pdf"), format="A4", print_background=True,
                     display_header_footer=True, header_template="<span></span>",
                     footer_template='<div style="font:8px sans-serif;color:#5d686b;width:100%;text-align:center">Zettelfrei · Erste Schritte · Seite <span class="pageNumber"></span>/<span class="totalPages"></span></div>',
                     margin={"top": "16mm", "bottom": "18mm", "left": "16mm", "right": "16mm"})
        await b.close()
    os.remove(p)
asyncio.run(main())
