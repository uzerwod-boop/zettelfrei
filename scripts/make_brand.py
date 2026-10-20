"""Erzeugt Logo und Markenzeichen von Zettelfrei (docs/brand/*.svg).
Text wird in Pfade umgewandelt (Archivo, OFL), damit das Logo überall gleich aussieht – auch auf GitHub.
    pip install fonttools && python scripts/make_brand.py
"""
import os
from fontTools.ttLib import TTFont
from fontTools.varLib.instancer import instantiateVariableFont
from fontTools.pens.svgPathPen import SVGPathPen
from fontTools.pens.transformPen import TransformPen

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "docs", "brand"); os.makedirs(OUT, exist_ok=True)
ACCENT, ACCENT_D, MARKER, INK, INK_D, PAPER = "#0f6b5c", "#4fbfa6", "#f6dd6a", "#1d2528", "#e7ebe6", "#f4f5f1"

font = instantiateVariableFont(TTFont(os.path.join(ROOT, "fonts", "Archivo.ttf")), {"wght": 800, "wdth": 100})
gs, cmap, hmtx = font.getGlyphSet(), font.getBestCmap(), font["hmtx"]
UPM = font["head"].unitsPerEm

def text_path(s, size, x0, base, track=0.0):
    pen = SVGPathPen(gs); x = 0; sc = size / UPM
    for ch in s:
        g = cmap[ord(ch)]
        gs[g].draw(TransformPen(pen, (sc, 0, 0, -sc, x0 + x, base)))
        x += hmtx[g][0] * sc + track * size
    return pen.getCommands(), x - track * size

# Zeichen: 4×4 Tabellenzellen, die gefüllten bilden ein Z; die gelbe Zelle ist die markierte Zelle.
Z = ["####", "..y.", ".#..", "####"]
THEMES = {"light": (ACCENT, "#cdd2ca"), "dark": (ACCENT_D, "#3a4446")}

def mark(x=0, y=0, s=64, theme="light", outline=True):
    fill, line = THEMES[theme]; c = s * 13 / 64; g = s * 4 / 64; r = s * 3.5 / 64; sw = s * 2.2 / 64
    out = [f'<g transform="translate({x} {y})">']
    for ri, row in enumerate(Z):
        for ci, ch in enumerate(row):
            px, py = round(ci * (c + g), 2), round(ri * (c + g), 2)
            if ch == "#": out.append(f'<rect x="{px}" y="{py}" width="{round(c,2)}" height="{round(c,2)}" rx="{round(r,2)}" fill="{fill}"/>')
            elif ch == "y": out.append(f'<rect x="{px}" y="{py}" width="{round(c,2)}" height="{round(c,2)}" rx="{round(r,2)}" fill="{MARKER}"/>')
            elif outline: out.append(f'<rect x="{round(px+sw/2,2)}" y="{round(py+sw/2,2)}" width="{round(c-sw,2)}" height="{round(c-sw,2)}" rx="{round(r-sw/2,2)}" fill="none" stroke="{line}" stroke-width="{round(sw,2)}"/>')
    out.append("</g>"); return "\n".join(out)

def svg(w, h, body, title="Zettelfrei"):
    return f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {w} {h}" width="{w}" height="{h}" role="img" aria-label="{title}"><title>{title}</title>\n{body}\n</svg>\n'

# Zeichen allein: groß mit leeren Zellen, klein (Favicon, ≤ 32 px) nur die gefüllten
open(os.path.join(OUT, "zettelfrei-mark.svg"), "w").write(svg(64, 64, mark()))
open(os.path.join(OUT, "zettelfrei-mark-dark.svg"), "w").write(svg(64, 64, mark(theme="dark")))
open(os.path.join(OUT, "zettelfrei-mark-small.svg"), "w").write(svg(64, 64, mark(outline=False)))
# Logo: Zeichen + „Zettel“ + „frei“ in Akzentfarbe
for name, theme, ink, acc in (("zettelfrei-logo.svg", "light", INK, ACCENT), ("zettelfrei-logo-dark.svg", "dark", INK_D, ACCENT_D)):
    d1, w1 = text_path("Zettel", 76, 92, 62.5, track=-0.012)
    d2, w2 = text_path("frei", 76, 92 + w1 - 0.012 * 76, 62.5, track=-0.012)
    W = int(92 + w1 + w2 + 6)
    open(os.path.join(OUT, name), "w").write(svg(W, 72, mark(0, 4, theme=theme) + f'\n<path d="{d1}" fill="{ink}"/>\n<path d="{d2}" fill="{acc}"/>'))
print("ok", sorted(os.listdir(OUT)))
