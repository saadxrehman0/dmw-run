"""Insert Arabic commentary into the AR DMW deck, matching the previous deck's run structure.

Latin-script spans get their own lang="en-US" run, exactly as the published decks do:
a single mixed-language run makes renderers lose the bidi order of the whole slide.

Every bullet is space-normalised before insertion (proofread Arabic can carry double spaces),
and Latin-script names left in the Arabic are printed as warnings (see check_ar_text.py;
tickers and acronym indices are allowed).
"""
import copy, json, re, sys
from pptx import Presentation
from lxml import etree
from check_ar_text import normalise_spaces, latin_flags

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
LATIN = re.compile(r"[A-Za-z][A-Za-z0-9&.\-’ ]*")


def split_runs(text):
    """[(is_latin, segment)] - Latin letter spans separated from the Arabic around them."""
    out, i = [], 0
    for m in LATIN.finditer(text):
        seg = m.group().rstrip()
        if not seg:
            continue
        s, e = m.start(), m.start() + len(seg)
        if s > i:
            out.append((False, text[i:s]))
        out.append((True, seg))
        i = e
    if i < len(text):
        out.append((False, text[i:]))
    return out


def protos(src_shape):
    """Prototype paragraph plus an Arabic and a Latin rPr, taken from the previous deck."""
    para = next(p for p in src_shape.text_frame._txBody.findall(f"{{{A}}}p") if p.findall(f"{{{A}}}r"))
    ar = lat = None
    for r in para.findall(f"{{{A}}}r"):
        rpr = r.find(f"{{{A}}}rPr")
        if rpr is None:
            continue
        if rpr.get("lang", "").startswith("ar") and ar is None:
            ar = rpr
        if rpr.get("lang", "").startswith("en") and lat is None:
            lat = rpr
    if ar is None:
        ar = para.find(f"{{{A}}}r/{{{A}}}rPr")
    if lat is None:
        lat = copy.deepcopy(ar)
        lat.set("lang", "en-US")
    return para, ar, lat


def fill(dst_shape, src_shape, bullets):
    proto, ar_rpr, lat_rpr = protos(src_shape)
    ppr = proto.find(f"{{{A}}}pPr")
    tx = dst_shape.text_frame._txBody
    for p_ in tx.findall(f"{{{A}}}p"):
        tx.remove(p_)
    for b in bullets:
        b = normalise_spaces(b)
        for span in latin_flags(b):
            print(f"  WARNING Latin text in Arabic bullet: '{span}'")
        p_ = etree.SubElement(tx, f"{{{A}}}p")
        if ppr is not None:
            p_.append(copy.deepcopy(ppr))
        for is_latin, seg in split_runs(b):
            r = etree.SubElement(p_, f"{{{A}}}r")
            rpr = copy.deepcopy(lat_rpr if is_latin else ar_rpr)
            for hl in rpr.findall(f"{{{A}}}highlight"):
                rpr.remove(hl)
            r.append(rpr)
            t = etree.SubElement(r, f"{{{A}}}t")
            t.text = seg
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


def shape(prs, slide_idx, sid):
    return next(s for s in prs.slides[slide_idx].shapes if s.shape_id == sid)


def main(cur_path, prev_path, draft_path, out_path):
    d = json.load(open(draft_path, encoding="utf-8"))
    cur, prev = Presentation(cur_path), Presentation(prev_path)
    for si, sid, key in [(0, 22, "s1_markets"), (0, 76, "s1_econ"), (0, 87, "s1_corp"), (1, 8, "s2")]:
        fill(shape(cur, si, sid), shape(prev, si, sid), d[key])
        print(f"slide {si+1} shape {sid}: {len(d[key])} bullets")
    cur.save(out_path)
    print("saved", out_path)


if __name__ == "__main__":
    main(*sys.argv[1:5])
