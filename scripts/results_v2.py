"""Results Announcements table in the Template v2 style (Saad, 25 Nov 2026).

On results days the table goes under "Major Corporate News" on slide 1, in place of that day's
Corporate bullet, and the rows that do not fit continue at the top of slide 2, under the same
header, before the Corporate news; the calendars move down. fill_commentary.py --results drives it.

The v2 master has no Results frame, so the table is made from the master's own Economic Calendar
rows (title, header, data and source rows of s2.flow): same fonts, sizes, fills and borders.
Rows, sorting and figures come from place_results.py (v1 logic, unchanged): DMW - Results.xlsx,
recomputed change, YoY % and variance, ANB-covered first, other periods last and starred.
Company names follow the v2 rule: Mapping long name if it fits its cell on one line, else short.
"""
import copy, re
from lxml import etree

import build_dmw as B
import place_results as PR

A = B.A
E = B.E
NCOL = 9
# logical (EN) widths in inches; sum = 7.70, the s2.flow width
WIDTHS = (1.50, 0.60, 0.95, 0.72, 0.72, 0.72, 0.62, 0.92, 0.95)
ROW_IN = 0.15
TITLE = {"en": "Results Announcements", "ar": "إعلانات النتائج"}
# Saad, 25 Nov 2026: a second navy bar right under "Major Corporate News" reads as a new section.
# The Results title is a light sub-header instead: the column-header style (light fill, bold,
# dark text) at 8 pt, so the table sits inside the Corporate section.
TITLE_STYLE = "rule"         # "sub" (light sub-header) or "bar" (navy section bar)
TITLE_PT = 8.0
SOURCE = {"en": "Source: Tadawul, Bloomberg", "ar": "المصدر: تداول، بلومبرغ"}
AR_Q = {"1": "الأول", "2": "الثاني", "3": "الثالث", "4": "الرابع"}


def prev_period(q):
    m = re.fullmatch(r"([1-4])Q(\d\d)", q.upper())
    return f"{m.group(1)}Q{int(m.group(2)) - 1:02d}" if m else ""


def ar_period(q):
    m = re.fullmatch(r"([1-4])Q(\d\d)", q.upper())
    return f"الربع {AR_Q[m.group(1)]} من عام 20{m.group(2)}" if m else q


def heads(lang, quarter):
    """Two header rows as [(logical col, span, text)]."""
    if lang == "en":
        h1 = [(3, 2, "Net Profit (SAR mn)"), (5, 2, "Change YoY"), (7, 1, "Consensus"), (8, 1, "Variance")]
        h2 = [(0, 1, "Company"), (1, 1, "TASI Code"), (2, 1, "Bloomberg Code"), (3, 1, quarter),
              (4, 1, prev_period(quarter)), (5, 1, "SAR mn"), (6, 1, "%"), (7, 1, "SAR mn"), (8, 1, "%")]
    else:
        h1 = [(3, 2, "صافي الربح (مليون ريال)"), (5, 2, "التغير السنوي"), (7, 1, "الإجماع"), (8, 1, "الفرق")]
        # period labels stay in Latin (3Q26), as in the AR news text: the Arabic words wrap the
        # 0.72 in column and grow the header row (4 Nov replay)
        h2 = [(0, 1, "اسم الشركة"), (1, 1, "رمز تاسي"), (2, 1, "رمز بلومبرغ"), (3, 1, quarter),
              (4, 1, prev_period(quarter)), (5, 1, "مليون ريال"), (6, 1, "%"), (7, 1, "مليون ريال"),
              (8, 1, "%")]
    for h in (h1, h2):
        used = {c + k for c, s, _ in h for k in range(s)}
        h += [(c, 1, "") for c in range(NCOL) if c not in used]
        h.sort()
    return h1, h2


class Protos:
    """Prototype rows and cells taken from the deck's own s2.flow (Economic Calendar block)."""
    def __init__(self, flow, lang):
        t = flow.table
        fr = B.flow_rows(t, lang)
        h, first, src = fr["ec"]
        trs = t._tbl.findall(f"{{{A}}}tr")
        self.lang = lang
        self.title, self.head, self.data, self.src = trs[h], trs[h + 1], trs[first], trs[src]
        self.gap = trs[h - 1]
        M = B.Mirror(lang)
        n = len(t.columns)
        tc_at = lambda tr, logical, span=1: tr.findall(f"{{{A}}}tc")[M.col(t, logical, span) if lang == "ar" else logical]
        # text cells: logical col 1 (Event, span 2); numbers: logical col 3 (Prior)
        self.cell = {}
        for key, tr in (("head", self.head), ("data", self.data)):
            self.cell[(key, "text")] = tc_at(tr, 1, 2)
            self.cell[(key, "num")] = tc_at(tr, 3)
        self.cell[("title", "text")] = tr0 = self.title.findall(f"{{{A}}}tc")[0]
        self.cell[("src", "text")] = self.src.findall(f"{{{A}}}tc")[0]
        self.n_src = n


def _set_text(tc, text):
    tb = tc.find(f"{{{A}}}txBody")
    ps = tb.findall(f"{{{A}}}p")
    for p in ps[1:]:
        tb.remove(p)
    p = ps[0]
    rs = p.findall(f"{{{A}}}r")
    if not text:
        # an empty cell keeps no run: an empty run takes the default size and grows the row
        # (4 Nov replay: the 0.05 in gap row rendered 0.23 in)
        end = p.find(f"{{{A}}}endParaRPr")
        if end is None and rs:
            end = copy.deepcopy(rs[0].find(f"{{{A}}}rPr"))
            end.tag = f"{{{A}}}endParaRPr"
            p.append(end)
        for r in rs:
            p.remove(r)
        for rpr in tc.iter(f"{{{A}}}endParaRPr"):
            for x in rpr.findall(f"{{{A}}}highlight"):
                rpr.remove(x)
        return
    if not rs:
        r = etree.Element(f"{{{A}}}r")
        rpr = p.find(f"{{{A}}}endParaRPr")
        r.append(copy.deepcopy(rpr) if rpr is not None else etree.Element(f"{{{A}}}rPr"))
        if rpr is not None:
            r[0].tag = f"{{{A}}}rPr"
            rpr.addprevious(r)
        else:
            p.append(r)
        etree.SubElement(r, f"{{{A}}}t")
        rs = [r]
    for r in rs[1:]:
        p.remove(r)
    rs[0].find(f"{{{A}}}t").text = text
    if re.fullmatch(r"[A-Za-z0-9 .,*%\-]+", text) and re.search(r"[A-Za-z]", text):
        # Latin text in a right-to-left cell: "3Q26" showed as "26Q3" in the AR PDF (4 Nov replay)
        ppr = p.find(f"{{{A}}}pPr")
        if ppr is not None and ppr.get("rtl") == "1":
            ppr.set("rtl", "0")
        rs[0].find(f"{{{A}}}rPr").set("lang", "en-US")
    for rpr in tc.iter(f"{{{A}}}rPr", f"{{{A}}}endParaRPr"):
        for x in rpr.findall(f"{{{A}}}highlight"):
            rpr.remove(x)


def _row(proto_tr, cells, lang, kinds, protos, key, h_in=ROW_IN):
    """cells: [(logical col, span, text)] covering all 9 columns; kinds: logical col -> text|num."""
    tr = etree.Element(f"{{{A}}}tr")
    tr.set("h", str(int(round(h_in * E))))
    phys = [None] * NCOL
    for c, span, text in cells:
        p0 = NCOL - c - span if lang == "ar" else c
        tc = copy.deepcopy(protos.cell[(key, kinds.get(c, "text"))])
        for att in ("gridSpan", "hMerge", "rowSpan", "vMerge"):
            tc.attrib.pop(att, None)
        _set_text(tc, text)
        if span > 1:
            tc.set("gridSpan", str(span))
        phys[p0] = tc
        for k in range(1, span):
            m = copy.deepcopy(tc)
            m.attrib.pop("gridSpan", None)
            m.set("hMerge", "1")
            _set_text(m, "")
            phys[p0 + k] = m
    for tc in phys:
        tr.append(tc)
    return tr


def band_row(proto_tr, lang, h_emu):
    """A full-width row (section header, gap) rebuilt for 9 columns from a s2.flow row."""
    class P: cell = {("x", "text"): proto_tr.findall(f"{{{A}}}tc")[0]}
    text = "".join(t.text or "" for t in proto_tr.findall(f"{{{A}}}tc")[0].iter(f"{{{A}}}t"))
    tr = _row(None, [(0, NCOL, text)], lang, {}, P, "x")
    tr.set("h", str(h_emu))
    return tr


def make_frame(flow, lang, rows, quarter, with_title, with_source, notes, names_for, lead=(), tail=()):
    """A new graphicFrame (copy of s2.flow's) holding the results table, with optional full-width
    rows (copies of s2.flow rows) before and after it. Returns (frame, height EMU)."""
    pr = Protos(flow, lang)
    lead = [band_row(tr, lang, int(tr.get("h"))) for tr in lead]
    tail = [band_row(tr, lang, int(tr.get("h"))) for tr in tail]
    fr = copy.deepcopy(flow._element)
    tbl = fr.find(f".//{{{A}}}tbl")
    for tr in tbl.findall(f"{{{A}}}tr"):
        tbl.remove(tr)
    grid = tbl.find(f"{{{A}}}tblGrid")
    for gc in grid.findall(f"{{{A}}}gridCol"):
        grid.remove(gc)
    ws = WIDTHS[::-1] if lang == "ar" else WIDTHS
    for w in ws:
        etree.SubElement(grid, f"{{{A}}}gridCol").set("w", str(int(round(w * E))))
    numk = {c: "num" for c in range(3, NCOL)}
    out = list(lead)
    if with_title:
        if TITLE_STYLE == "bar":
            out.append(_row(pr.title, [(0, NCOL, TITLE[lang])], lang, {}, pr, "title", 0.20))
        elif TITLE_STYLE == "rule":
            # white row, bold navy text, a navy rule under it (the v1 Results look)
            tr = _row(pr.data, [(0, NCOL, TITLE[lang])], lang, {}, pr, "data", 0.20)
            navy = pr.title.find(f".//{{{A}}}tcPr/{{{A}}}solidFill/{{{A}}}srgbClr")
            navy = navy.get("val") if navy is not None else "002060"
            for rpr in tr.iter(f"{{{A}}}rPr"):
                rpr.set("sz", str(int(TITLE_PT * 100))); rpr.set("b", "1")
                for f in rpr.findall(f"{{{A}}}solidFill"):
                    rpr.remove(f)
                sf = etree.Element(f"{{{A}}}solidFill"); etree.SubElement(sf, f"{{{A}}}srgbClr").set("val", navy)
                rpr.insert(0, sf)
            for tcpr in tr.iter(f"{{{A}}}tcPr"):
                lnb = tcpr.find(f"{{{A}}}lnB")
                if lnb is not None:
                    lnb.set("w", "12700")
                    for f in lnb.findall(f"{{{A}}}solidFill"):
                        f.find(f"{{{A}}}srgbClr").set("val", navy)
            out.append(tr)
        else:
            tr = _row(pr.head, [(0, NCOL, TITLE[lang])], lang, {}, pr, "head", 0.20)
            for rpr in tr.iter(f"{{{A}}}rPr"):
                rpr.set("sz", str(int(TITLE_PT * 100))); rpr.set("b", "1")
            out.append(tr)
    h1, h2 = heads(lang, quarter)
    out.append(_row(pr.head, h1, lang, numk, pr, "head"))
    out.append(_row(pr.head, h2, lang, numk, pr, "head"))
    for x in rows:
        name = names_for(x)
        v = [name, x["code"], x["bbg"], PR.f1(x["cur"]), PR.f1(x["pri"]), PR.f1(x["chg"]), PR.f1(x["pct"]),
             PR.f1(x["cons"]), PR.f1(x["var"])]
        out.append(_row(pr.data, [(c, 1, v[c]) for c in range(NCOL)], lang, numk, pr, "data"))
    if with_source:
        tr = _row(pr.src, [(0, NCOL, SOURCE[lang])], lang, {}, pr, "src", ROW_IN * (1 + len(notes)))
        if notes:
            tc = [x for x in tr.findall(f"{{{A}}}tc") if x.get("gridSpan")][0]
            txb = tc.find(f"{{{A}}}txBody")
            p0 = txb.findall(f"{{{A}}}p")[0]
            for n in notes:
                np_ = copy.deepcopy(p0)
                np_.findall(f"{{{A}}}r")[0].find(f"{{{A}}}t").text = n
                txb.append(np_)
        out.append(tr)
    out += tail
    for tr in out:
        tbl.append(tr)
    h = sum(int(tr.get("h")) for tr in out)
    xfrm = fr.find(".//{http://schemas.openxmlformats.org/presentationml/2006/main}xfrm")
    xfrm.find(f"{{{A}}}ext").set("cy", str(h))
    nv = fr.find(".//{http://schemas.openxmlformats.org/presentationml/2006/main}cNvPr")
    return fr, h


def height_in(n_rows, with_title, with_source, n_notes):
    return (0.20 if with_title else 0) + ROW_IN * (2 + n_rows) + (ROW_IN * (1 + n_notes) if with_source else 0)
