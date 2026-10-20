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

def mark(x=0, y=0, s=64):
    """Zeichen: Blatt mit Eselsohr (Zettel), darin Tabellenlinien und ein Z als Textmarker-Strich."""
    k = s / 64
    T = lambda v: round(v * k, 2)
    return f'''<g transform="translate({x} {y})">
  <path d="M{T(14)} 0H{T(44)}L{T(64)} {T(20)}V{T(50)}A{T(14)} {T(14)} 0 0 1 {T(50)} {T(64)}H{T(14)}A{T(14)} {T(14)} 0 0 1 0 {T(50)}V{T(14)}A{T(14)} {T(14)} 0 0 1 {T(14)} 0Z" fill="{ACCENT}"/>
  <path d="M{T(44)} 0V{T(14)}A{T(6)} {T(6)} 0 0 0 {T(50)} {T(20)}H{T(64)}Z" fill="#8fd0c0"/>
  <g stroke="#ffffff" stroke-opacity=".22" stroke-width="{T(2)}"><path d="M{T(10)} {T(25)}H{T(54)}M{T(10)} {T(39)}H{T(54)}M{T(25)} {T(12)}V{T(54)}M{T(39)} {T(22)}V{T(54)}"/></g>
  <path d="M{T(17)} {T(19)}H{T(46)}L{T(18)} {T(46)}H{T(47)}" fill="none" stroke="{MARKER}" stroke-width="{T(7.5)}" stroke-linecap="round" stroke-linejoin="round"/>
</g>'''

def svg(w, h, body, title="Zettelfrei"):
    return f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {w} {h}" width="{w}" height="{h}" role="img" aria-label="{title}"><title>{title}</title>\n{body}\n</svg>\n'

# Zeichen allein
open(os.path.join(OUT, "zettelfrei-mark.svg"), "w").write(svg(64, 64, mark()))
# Logo: Zeichen + Wortmarke, für hellen und dunklen Hintergrund
for name, ink in (("zettelfrei-logo.svg", INK), ("zettelfrei-logo-dark.svg", INK_D)):
    d, w = text_path("Zettelfrei", 58, 92, 60, track=-0.012)
    open(os.path.join(OUT, name), "w").write(svg(int(92 + w + 4), 76, mark(0, 6) + f'\n<path d="{d}" fill="{ink}"/>'))
print("ok", os.listdir(OUT))
