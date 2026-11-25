"""Compare a built DMW deck against the team's published deck, cell by cell.

Checks slide 1 to 3 tables (text and +/- colour on % and bps cells), the slide 1 headline boxes,
and whether slides 4 and 5 (Technicals) match. Prints a count, then one line per difference.

Usage:  python compare_ar.py <my_deck.pptx> <published_deck.pptx>
"""
import sys, re
from pptx import Presentation
from lxml import etree
A='{http://schemas.openxmlformats.org/drawingml/2006/main}'
def n(s): return re.sub(r'\s+',' ',s.replace('\xa0',' ')).strip()
def color(cell_or_tf):
    tf=cell_or_tf.text_frame if hasattr(cell_or_tf,'text_frame') else cell_or_tf
    for p in tf.paragraphs:
        for r in p.runs:
            if not r.text.strip(): continue
            rp=r._r.find(A+'rPr')
            if rp is None: return None
            sf=rp.find(A+'solidFill')
            if sf is None: return None
            v=sf[0].get('val')
            return None if v in ('tx1','000000','241F20','262626','404040') or sf[0].tag.endswith('schemeClr') else v
    return None
def title(t):
    return n(' '.join(t.cell(r,c).text for r in range(min(2,len(t.rows))) for c in range(len(t.columns))))[:40]
if len(sys.argv) != 3:
    sys.exit("Usage: python compare_ar.py <my_deck.pptx> <published_deck.pptx>")
mine,pub=Presentation(sys.argv[1]),Presentation(sys.argv[2])
tot=0; bad=[]
for si in range(3):
    M,P=mine.slides[si],pub.slides[si]
    mt={title(s.table):s.table for s in M.shapes if s.has_table}
    pt={title(s.table):s.table for s in P.shapes if s.has_table}
    for k,pt_ in pt.items():
        if k not in mt: bad.append(f"S{si+1} table only in published: {k}"); continue
        m_=mt[k]
        if len(m_.rows)!=len(pt_.rows): bad.append(f"S{si+1} '{k}': rows mine {len(m_.rows)} vs pub {len(pt_.rows)}")
        for r in range(min(len(m_.rows),len(pt_.rows))):
            for c in range(min(len(m_.columns),len(pt_.columns))):
                a,b=m_.cell(r,c),pt_.cell(r,c)
                if not n(b.text) and not n(a.text): continue
                tot+=1
                if n(a.text)!=n(b.text): bad.append(f"S{si+1} '{k}' r{r}c{c}: mine '{n(a.text)}' | pub '{n(b.text)}'")
                elif re.search(r'[%]|bps',b.text) and color(a)!=color(b): bad.append(f"S{si+1} '{k}' r{r}c{c} '{n(b.text)}': colour mine {color(a)} vs pub {color(b)}")
    for k in mt:
        if k not in pt: bad.append(f"S{si+1} table only in mine: {k}")
    if si==0:
        ms={s.shape_id:s for s in M.shapes};
        for s in P.shapes:
            if s.has_text_frame and 600_000<s.top<2_150_000 and s.shape_id in ms and n(s.text_frame.text):
                tot+=1; a=ms[s.shape_id]
                if n(a.text_frame.text)!=n(s.text_frame.text): bad.append(f"S1 shape {s.shape_id}: mine '{n(a.text_frame.text)}' | pub '{n(s.text_frame.text)}'")
                elif '%' in s.text_frame.text and color(a.text_frame)!=color(s.text_frame): bad.append(f"S1 shape {s.shape_id} colour mine {color(a.text_frame)} vs pub {color(s.text_frame)}")
for si in (3,4):
    mtxt=sorted(n(s.text_frame.text) for s in mine.slides[si].shapes if s.has_text_frame and n(s.text_frame.text))
    ptxt=sorted(n(s.text_frame.text) for s in pub.slides[si].shapes if s.has_text_frame and n(s.text_frame.text))
    mtab=[[n(c.text) for r in s.table.rows for c in r.cells] for s in mine.slides[si].shapes if s.has_table]
    ptab=[[n(c.text) for r in s.table.rows for c in r.cells] for s in pub.slides[si].shapes if s.has_table]
    tot+=1
    if mtxt!=ptxt or sorted(mtab)!=sorted(ptab): bad.append(f"S{si+1}: text/tables differ from published")
print(f"compared {tot} items, {len(bad)} differences")
print("\n".join(bad))
