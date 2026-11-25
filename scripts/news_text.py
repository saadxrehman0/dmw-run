"""Shared text helpers for the v2 news cells: HarfBuzz line breaking, the cell height model, the
Arabic run split and the EN/AR parity checks.

Height model (fitted 25 Nov 2026 on 16 news cells of the CEO finals A and B, measured in
PowerPoint by export_and_measure.ps1): a news cell is 11.0 pt per line plus 4.0 pt between
paragraphs (EN), and 0.5 pt less in AR. With HarfBuzz line breaks the model hit all 16 cells within
0.5 pt (EN exact). The cells use exact 10.8 pt line spacing and 3.6 pt after each paragraph.
"""
import copy, re
from lxml import etree
from build_dmw import text_width_in

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
E = 914400
PT_LINE, PT_PARA, AR_ADJ = 11.0, 4.0, -0.5
FONT = {"en": "ArialNovaCond.ttf", "ar": "anbApercu-Regular.otf"}
SIZE = {"en": 9.5, "ar": 9.0}
PLACEHOLDER = re.compile(r"\[[^\]]*(paste|لصق)[^\]]*\]")
# A Latin span in Arabic text. Leading digits belong to the span when a letter follows (1Q27):
# splitting "1" from "Q27" made PowerPoint show "Q27.1" (13 Oct 2026, v1).
LATIN = re.compile(r"(?:\d+(?=[A-Za-z]))?[A-Za-z][A-Za-z0-9&.\-’ ]*")


def n_lines(text, lang, avail_in):
    words, n, cur = text.split(" "), 1, ""
    for w in words:
        t = f"{cur} {w}" if cur else w
        wid = text_width_in(t, SIZE[lang], FONT[lang], rtl=(lang == "ar"))
        if wid is None:
            raise SystemExit(f"Font {FONT[lang]} or uharfbuzz missing: cannot measure the news text.")
        if wid <= avail_in or not cur:
            cur = t
        else:
            n, cur = n + 1, w
    return n


def text_avail_in(width_emu, tc):
    """Line room in a news cell: cell width minus cell margins minus the bullet indent."""
    pr = tc.find(f"{{{A}}}tcPr")
    ml = int(pr.get("marL", 91440)) if pr is not None else 91440
    mr = int(pr.get("marR", 91440)) if pr is not None else 91440
    lvl = tc.find(f".//{{{A}}}lvl1pPr")
    ind = int(lvl.get("marL", 0)) if lvl is not None else 0
    return (width_emu - ml - mr - ind) / E


def cell_height_pt(paras, lang, avail_in):
    if not paras:
        return 0.0
    lines = sum(n_lines(p, lang, avail_in) for p in paras)
    return lines * PT_LINE + (len(paras) - 1) * PT_PARA + (AR_ADJ if lang == "ar" else 0)


def _clean_rpr(rpr):
    rp = copy.deepcopy(rpr)
    for x in rp.findall(f"{{{A}}}highlight"):
        rp.remove(x)
    rp.attrib.pop("err", None)
    return rp


def fill_cell(tc, paras, lang):
    """Replace a news cell's paragraphs, keeping the first placeholder paragraph's pPr and rPr."""
    tb = tc.find(f"{{{A}}}txBody")
    ps = tb.findall(f"{{{A}}}p")
    p0 = ps[0]
    ppr = p0.find(f"{{{A}}}pPr")
    rpr = p0.find(f"{{{A}}}r/{{{A}}}rPr")
    if rpr is None:
        rpr = etree.Element(f"{{{A}}}rPr")
    end = p0.find(f"{{{A}}}endParaRPr")
    for p in ps:
        tb.remove(p)
    for t in paras:
        p = etree.SubElement(tb, f"{{{A}}}p")
        if ppr is not None:
            p.append(copy.deepcopy(ppr))
        if lang == "en":
            _run(p, t, rpr, None)
        else:
            pos = 0
            for m in LATIN.finditer(t):
                seg = m.group().rstrip(" .")
                if not seg:
                    continue
                if m.start() > pos:
                    _run(p, t[pos:m.start()], rpr, "ar-SA")
                _run(p, seg, rpr, "en-US")
                pos = m.start() + len(seg)
            if pos < len(t):
                _run(p, t[pos:], rpr, "ar-SA")
        if end is not None:
            p.append(_clean_rpr(end))
    if not paras:   # a txBody needs one paragraph or PowerPoint asks to repair the file
        p = etree.SubElement(tb, f"{{{A}}}p")
        if ppr is not None:
            p.append(copy.deepcopy(ppr))


def _run(p, text, rpr, lang):
    r = etree.SubElement(p, f"{{{A}}}r")
    rp = _clean_rpr(rpr)
    if lang:
        rp.set("lang", lang)
    r.append(rp)
    t = etree.SubElement(r, f"{{{A}}}t")
    t.text = text


# ---------------------------------------------------------------- parity checks (from v1)
def numbers(s, lang):
    s = re.sub(r"(?<![A-Za-z0-9])(\d)[QHM](\d\d)(?![0-9])", lambda m: "20" + m.group(2), s)
    # period labels (Q3, 3Q, H1, 9M) carry no figure: AR usually writes them as words, so drop them
    # in both languages (false CONFLICT on the 4 Nov replay, 26 Nov 2026)
    s = re.sub(r"(?<![A-Za-z0-9])(?:Q[1-4]|[1-4]Q|H[12]|[12]H|(?:3|6|9|12)M)(?![A-Za-z0-9])", " ", s)
    if lang == "en":
        s = re.sub(r"(?:\$|USD\s?)(\d)\b(?![.,]\d)", r"\1", s)
    else:
        s = s.translate(str.maketrans("٠١٢٣٤٥٦٧٨٩٫٬", "0123456789.,"))
        s = re.sub(r"(دولارين|مليوني|ملياري)", " 2 ", s)
        s = re.sub(r"(\d+(?:\.\d+)?)\s*(?:ألف|آلاف)", lambda m: f"{round(float(m.group(1)) * 1000):,}", s)
    xs = [x.rstrip(".,") for x in re.findall(r"\d[\d,]*\.?\d*", s)]
    years = sorted(set(x for x in xs if re.fullmatch(r"20\d\d", x)))
    return sorted(x for x in xs if not re.fullmatch(r"20\d\d", x)) + years


def tickers(s):
    return sorted(re.findall(r"\(([A-Z0-9]+ AB)\)", s))
