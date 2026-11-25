"""Measuring copy of Arial Nova Cond for LibreOffice renders (EN layout).
Takes the real ArialNovaCond.ttf (fonts/), drops kerning (PowerPoint does not kern this text,
LibreOffice would) and widens every advance by FACTOR (default 0.25%), which makes LibreOffice
break lines like PowerPoint: tested on 43 published EN commentary boxes (12 days), 43 exact.
(Unchanged real font: 40 exact, 3 one line short. +0.5%: 38 exact, 5 one line long.)
Usage: python3 make_measure_font.py fonts/ArialNovaCond.ttf ~/.fonts/ArialNovaCond-measure.ttf [factor]"""
import sys
from fontTools.ttLib import TTFont
from fontTools.pens.transformPen import TransformPen
from fontTools.pens.ttGlyphPen import TTGlyphPen
src, out = sys.argv[1], sys.argv[2]
sc = float(sys.argv[3]) if len(sys.argv) > 3 else 1.0025
f = TTFont(src)
for tag in ("GPOS", "kern"):
    if tag in f:
        del f[tag]
gs = f.getGlyphSet(); glyf = f["glyf"]
for n in f.getGlyphOrder():
    g = glyf[n]
    if g.isComposite():
        for c in g.components:
            c.x = int(round(c.x * sc))
    else:
        pen = TTGlyphPen(gs); gs[n].draw(TransformPen(pen, (sc, 0, 0, 1, 0, 0))); glyf[n] = pen.glyph()
    adv, lsb = f["hmtx"].metrics[n]
    f["hmtx"].metrics[n] = (int(round(adv * sc)), int(round(lsb * sc)))
f.save(out)
