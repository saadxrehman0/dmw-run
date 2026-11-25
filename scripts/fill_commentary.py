#!/usr/bin/env python3
"""Fill the DMW EN and AR commentary boxes from the ANB Rough doc in one pass.

Usage:
  python3 fill_commentary.py --doc READ_DOC.json --en "ANB Daily Market Watch - EN.pptx" \
      --ar "ANB Daily Market Watch - AR.pptx" --outdir OUT [--force] [--render] [--en-skip ec,co]

--doc     the saved JSON of a Google Docs read_doc of ANB Rough (both tabs).
--en/--ar today's decks as the team left them (patched, never rebuilt).
--force   also overwrite boxes that no longer hold the build's placeholder text
          (default: a box the team already filled is left alone and reported).
--render  also write PNGs of slides 1-2 of both outputs for a visual check.
--en-skip English boxes a teammate pastes by hand, left untouched: mk (Markets), ec (Economic),
          co (Corporate, slide 1 and slide 2). The Arabic deck is always filled in full: Arabic
          pasted by hand scrambles brackets, tickers and quarter labels (13 Oct 2026). The English
          boxes are still sized and placed for the planned text, and the report says where the
          Corporate items break between slide 1 and slide 2 so the teammate pastes to match.

What it does:
  1. Markets box: Global bullet then TASI bullet from the Draft tab write-ups.
     Economic box: the finalized Economic items, in order.
     Corporate: the finalized Corporate items, in order, slide 1 box then slide 2 box.
     Source: the Final - Sequenced tab if the doc has one, else the Draft tab's
     "Economic News" / "Corporate News" sections (single-tab layout since 6 Oct 2026).
  2. Measures every paragraph's real height with one LibreOffice render of a scratch
     deck (each paragraph in its own colour), so the layout needs no trial renders.
  3. Layout, never shrinking fonts, one house gap (SECTION_GAP, 0.10 in, ink to ink) between
     every pair of blocks: Markets text, Economic header, Economic text, Corporate header,
     Corporate text on slide 1. Slide 2: Corporate text, Event calendar, Economic calendar, with
     a wider CAL_GAP (0.20 in) above each calendar (Saad, 25 Nov 2026).
     Blocks flow top-down and spare space stays at the bottom of the page.
     Corporate items fill slide 1 down to 10.60 in; EN and AR break at the SAME item (the
     smaller of the two fits). The rest go to slide 2, where the calendars follow the news.
     If items still do not fit, the lowest-ranked Corporate items are left out of BOTH
     decks (the team trims from the bottom) and listed in the report.
     EN is measured with a copy of the real Arial Nova Cond made for LibreOffice
     (make_measure_font.py: kerning off, +0.25% width). On 43 published EN boxes it broke
     lines exactly where PowerPoint did in all 43. AR uses the real Apercu Arabic font.
  4. Checks: no placeholders left, every bullet ends with a source tag, EN vs AR numbers
     and tickers match (quarter notation normalised).
Requires: python-pptx, lxml, numpy, Pillow, LibreOffice (soffice) and pdftoppm, and the
anb Apercu / Apercu Arabic fonts plus an "Arial Nova Cond" font (or the 0.82-scaled
Liberation Sans proxy) installed for the renderer.
"""
import argparse, copy, json, os, re, shutil, subprocess, sys, tempfile
import numpy as np
from lxml import etree
from PIL import Image
from pptx import Presentation

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
EMU = 914400
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from build_dmw import SECTION_GAP, PAGE_BOTTOM, clean_hidden_titles   # one source for the house spacing

GAP = SECTION_GAP       # visible gap, ink to ink, between consecutive blocks (0.10 in)
CAL_GAP = 0.20          # slide 2: visible gap above each calendar table, EN and AR (Saad, 25 Nov 2026,
                        # from his 24 Nov hand layout: 0.21-0.31 in, evened to 0.20)
S1_LIMIT = PAGE_BOTTOM  # slide 1 corporate text must end by this (10.60 in, page number band starts 10.66)
TABLE_BOTTOM = PAGE_BOTTOM   # slide 2: nothing may end below this
EN_MARGIN = 0.0         # EN reserve when deciding what fits; 0 since the measuring font matched PowerPoint on 43/43 boxes
TAB_TOP = 1.0           # where tables sit on the scratch measuring slides
MEASURE_FONT = "ArialNovaCond-measure.ttf"
HDR = {
    "en": {"mk": "MARKETS PERFORMANCE & NEWS", "ec": "MAJOR ECONOMIC/SECTOR NEWS", "co": "MAJOR CORPORATE NEWS"},
    "ar": {"mk": "أداء الأسواق/الأخبار", "ec": "أخبار اقتصادية/قطاعية رئيسية", "co": "أهم أخبار الشركات"},
}
PLACEHOLDER = re.compile(r"\[.*(paste|لصق).*\]")
# A Latin span in Arabic text. Leading digits belong to the span when a letter follows (1Q27,
# 2Q26, 1H26): splitting "1" into the Arabic run and "Q27" into the English run made PowerPoint
# show "Q27.1" (13 Oct 2026). Trailing spaces and full stops stay in the Arabic run.
LATIN = re.compile(r"(?:\d+(?=[A-Za-z]))?[A-Za-z][A-Za-z0-9&.\-’ ]*")
log = []


def note(level, msg):
    log.append((level, msg))


# ---------------------------------------------------------------- doc parsing
def read_doc(path):
    d = json.load(open(path, encoding="utf-8"))
    d = d.get("content", d)
    tabs = {}
    for t in d["tabs"]:
        paras = []
        for el in t["documentTab"]["body"]["content"]:
            if "paragraph" not in el:
                continue
            p = el["paragraph"]
            txt = "".join(r.get("textRun", {}).get("content", "") for r in p["elements"]).rstrip("\n")
            paras.append((txt, p["paragraphStyle"].get("namedStyleType", "")))
        tabs[t["tabProperties"]["title"].strip()] = paras
    return tabs


def news_items(paras):
    """{'Economic': [{'en','ar','link'}], 'Corporate': [...]} from a tab's paragraphs."""
    out, sec, cur = {"Economic": [], "Corporate": []}, None, None
    for txt, style in paras:
        if style.startswith("HEADING"):
            # "Economic" / "Corporate" (old Final - Sequenced) or "Economic News" / "Corporate News"
            # (single Draft tab, 6 Oct 2026); any other heading (Global, TASI, calendars) ends the
            # section; an empty heading paragraph changes nothing
            h = txt.strip()
            if h:
                sec = next((k for k in out if h.startswith(k)), None)
            continue
        if sec is None or not txt.strip():
            continue
        if txt.startswith("http"):
            if cur: cur["link"] = txt
        elif re.match(r"[A-Za-z0-9\"'(]", txt.strip()):
            cur = {"en": txt.strip()}; out[sec].append(cur)
        elif cur is not None:
            cur["ar"] = txt.strip()
    for s in out:
        for it in out[s]:
            if "ar" not in it:
                note("CONFLICT", f"{s} item has no Arabic paragraph: {it['en'][:60]}")
    return out


def writeups(paras):
    """Global and TASI write-ups: the paragraphs under a heading starting "Global" / "TASI"
    (6 Oct 2026 layout; the old "TASI Write up:" / "Global Write up:" titles also match),
    up to the next heading."""
    def block(title):
        i = next((k for k, (t, st) in enumerate(paras) if t.strip().startswith(title)
                  and (st.startswith("HEADING") or "Write up" in t)), None)
        if i is None:
            note("FLAG", f"No '{title}' heading found in the Draft tab.")
            return []
        out = []
        for t, st in paras[i + 1:]:
            if st.startswith("HEADING") and t.strip():
                break
            if re.match(r"(TASI|Global) Write up", t.strip()):
                break
            if t.strip():
                out.append(t.strip())
        return out
    tasi = block("TASI")
    glob = block("Global")
    for name, L in (("TASI", tasi), ("Global", glob)):
        if not L:
            note("FLAG", f"The {name} write-up under its Draft tab heading is empty; the Markets box will "
                         f"miss it. Check whether the approved text was posted to the Draft tab.")
    en = lambda L: next((x for x in L if re.match(r"[A-Za-z]", x)), None)
    ar = lambda L: next((x for x in L if re.match(r"[؀-ۿ]", x)), None)
    return {"en": [en(glob), en(tasi)], "ar": [ar(glob), ar(tasi)]}


# ---------------------------------------------------------------- deck shapes
def _flat(shapes):
    for s in shapes:
        yield s
        if s.shape_type == 6:  # group: the team sometimes groups a header with its box (AR, Oct 2026)
            yield from _flat(s.shapes)


def _t(s):
    return re.sub(r"\s+", " ", s.text_frame.text.replace("​", "")).strip()


def find_shapes(prs, lang, strict=True):
    """Slide 1 Markets, Economic and Corporate boxes and the slide 2 Corporate box.
    strict=False (check_final.py on the team's saved decks) returns None for a box that is not
    there instead of stopping: on results days the slide 1 Corporate box is replaced by the results
    table, and some AR finals have no slide 2 Corporate header (back-test of 70 finals, 21 Oct 2026)."""
    s1, s2 = prs.slides[0], prs.slides[1]
    h = HDR[lang]
    hv = set(h.values())
    def hdr(slide, text):
        c = [s for s in _flat(slide.shapes) if s.has_text_frame and _t(s) == text and s.width > 5 * EMU]
        if not c:
            if strict:
                sys.exit(f"{lang}: header '{text}' not found")
            return None
        return min(c, key=lambda s: s.top)
    sh1 = list(_flat(s1.shapes))
    bodies = [s for s in sh1 if s.has_text_frame and s.width > 6.5 * EMU and s.height > 0.3 * EMU
              and _t(s) not in hv]
    out = {"mk_h": hdr(s1, h["mk"]), "ec_h": hdr(s1, h["ec"]), "co_h": hdr(s1, h["co"])}
    heads = sorted([out[k + "_h"] for k in ("mk", "ec", "co") if out[k + "_h"] is not None], key=lambda x: x.top)
    for k in ("mk", "ec", "co"):
        hd = out[k + "_h"]
        if hd is None:
            out[k] = None; continue
        nxt = next((x.top for x in heads if x.top > hd.top), None)
        below = sorted([b for b in bodies if b.top >= hd.top - int(0.05 * EMU)
                        and (nxt is None or b.top < nxt)], key=lambda b: b.top)
        if not below:
            # a box left empty and collapsed (no item fitted on slide 1) is shorter than 0.3 in;
            # take the first wide text box under the header instead of stopping (13 Oct 2026 test)
            below = sorted([s for s in sh1 if s.has_text_frame and s.width > 6.5 * EMU
                            and s.height > 0.01 * EMU and s.top >= hd.top + int(0.05 * EMU)
                            and (nxt is None or s.top < nxt) and _t(s) not in hv], key=lambda b: b.top)
        if not below:
            if strict:
                sys.exit(f"{lang}: no text box under '{h[k]}'")
            out[k] = None; continue
        out[k] = below[0]
    # slide 2 Corporate box: the first text box under the slide 2 Corporate header. Not the tallest
    # box: once the box is filled with nothing and collapsed, the tallest is a stray empty box or a
    # page header (12 Oct 2026). Without a header, the first wide text box with text.
    h2 = hdr(s2, h["co"]) if strict else (hdr(s2, h["co"]) or None)
    sh2 = list(_flat(s2.shapes))
    b2 = sorted([s for s in sh2 if s.has_text_frame and s.width > 6.5 * EMU
                 and (h2 is None or (s.shape_id != h2.shape_id and s.top >= h2.top - int(0.05 * EMU)))
                 and _t(s) not in hv and (h2 is not None or len(_t(s)) > 40)], key=lambda s: s.top)
    if not strict:   # a saved deck: skip an empty stray box above the real one (21 Oct 2026 EN final)
        b2 = [b for b in b2 if _t(b)] or b2
    out["s2"] = b2[0] if b2 else (None if not strict else
                                  max((s for s in sh2 if s.has_text_frame), key=lambda s: s.height))
    out["tables"] = sorted([s for s in s2.shapes if s.has_table], key=lambda s: s.top)
    return out


# ---------------------------------------------------------------- filling
def fill_en(shape, paras, colors=None):
    tb = shape.text_frame._txBody
    p0 = tb.find(f"{{{A}}}p")
    ppr = p0.find(f"{{{A}}}pPr"); rpr = p0.find(f"{{{A}}}r/{{{A}}}rPr")
    if rpr is None:
        rpr = etree.Element(f"{{{A}}}rPr"); rpr.set("lang", "en-US"); rpr.set("sz", "950")
    for p in tb.findall(f"{{{A}}}p"):
        tb.remove(p)
    for i, t in enumerate(paras):
        p = etree.SubElement(tb, f"{{{A}}}p")
        if ppr is not None:
            p.append(copy.deepcopy(ppr))
        r = etree.SubElement(p, f"{{{A}}}r")
        rp = clean_rpr(rpr, colors[i] if colors else None)
        r.append(rp)
        etree.SubElement(r, f"{{{A}}}t").text = t


def fill_ar(shape, paras, colors=None):
    tb = shape.text_frame._txBody
    for p in tb.findall(f"{{{A}}}p"):
        tb.remove(p)
    for i, t in enumerate(paras):
        p = etree.SubElement(tb, f"{{{A}}}p")
        ppr = etree.SubElement(p, f"{{{A}}}pPr"); ppr.set("marR", "0")
        # justified right-to-left, like the published AR decks (their boxes inherit algn="just"
        # rtl="1"); "r" left the Arabic ragged (13 Oct 2026)
        ppr.set("algn", "just"); ppr.set("rtl", "1")
        pos = 0
        for m in LATIN.finditer(t):
            seg = m.group().rstrip(" .")
            if not seg:
                continue
            if m.start() > pos:
                ar_run(p, t[pos:m.start()], "ar-SA", colors[i] if colors else None)
            ar_run(p, seg, "en-US", colors[i] if colors else None)
            pos = m.start() + len(seg)
        if pos < len(t):
            ar_run(p, t[pos:], "ar-SA", colors[i] if colors else None)


def ar_run(p, text, lang, color):
    r = etree.SubElement(p, f"{{{A}}}r")
    rp = etree.Element(f"{{{A}}}rPr"); rp.set("lang", lang); rp.set("dirty", "0")
    r.append(clean_rpr(rp, color))
    t = etree.SubElement(r, f"{{{A}}}t"); t.text = text
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


def clean_rpr(rpr, color=None):
    rp = copy.deepcopy(rpr)
    for tag in ("highlight", "solidFill"):
        for x in rp.findall(f"{{{A}}}{tag}"):
            rp.remove(x)
    rp.attrib.pop("err", None)
    if color:
        sf = etree.Element(f"{{{A}}}solidFill")
        etree.SubElement(sf, f"{{{A}}}srgbClr").set("val", "%02X%02X%02X" % color)
        rp.insert(0, sf)
    return rp


def fill(lang, shape, paras, colors=None):
    (fill_en if lang == "en" else fill_ar)(shape, paras, colors)
    # A txBody needs at least one paragraph, or PowerPoint asks to repair the file. This happens
    # when no Corporate item is left for slide 2 (12 Oct 2026): leave one empty paragraph.
    tb = shape.text_frame._txBody
    if tb.find(f"{{{A}}}p") is None:
        etree.SubElement(tb, f"{{{A}}}p")


# ---------------------------------------------------------------- measuring
PAL6 = [(220, 0, 0), (0, 0, 220), (0, 160, 0), (200, 0, 200), (230, 120, 0), (0, 170, 170)]


def palette(n):
    """Six well-separated colours in rotation; same-colour paragraphs are told apart by position."""
    return [PAL6[i % len(PAL6)] for i in range(n)]


def build_measure_deck(path, lang, boxes, out):
    """Scratch deck: one new slide per box, holding a clone of that box at the top of the page,
    each paragraph in its own colour; then one slide per slide-2 calendar table, cloned at
    TAB_TOP, to measure its real rendered height. Original slides are dropped."""
    prs = Presentation(path)
    sh = find_shapes(prs, lang)
    layout = prs.slide_layouts[len(prs.slide_layouts) - 1]
    jobs = []
    def new_slide():
        s = prs.slides.add_slide(layout)
        for ph in list(s.placeholders):
            ph._element.getparent().remove(ph._element)
        return s
    for key, paras in boxes:
        src = sh[key]
        s = new_slide()
        el = copy.deepcopy(src._element)
        s.shapes._spTree.append(el)
        box = s.shapes[-1]
        box.top = int(0.3 * EMU); box.height = int(11.0 * EMU)
        cols = palette(len(paras))
        fill(lang, box, paras, cols)
        jobs.append((key, cols))
    for i, t in enumerate(sh["tables"]):
        s = new_slide()
        s.shapes._spTree.append(copy.deepcopy(t._element))
        s.shapes[-1].top = int(TAB_TOP * EMU)
        jobs.append((f"tab{i}", (t.left / EMU, (t.left + t.width) / EMU)))
    lst = prs.slides._sldIdLst
    n_orig = len(lst) - len(jobs)
    for sid in list(lst)[:n_orig]:
        prs.part.drop_rel(sid.rId); lst.remove(sid)
    prs.save(out)
    return jobs


def ink_rows(png, x0, x1, dpi=100):
    a = np.asarray(Image.open(png).convert("L")).astype(int)[:, int(x0 * dpi):int(x1 * dpi)]
    return (a < 200).sum(axis=1) >= 2


def measure_table(png, xs, dpi=100):
    """(top offset, bottom offset) in inches of a table's ink relative to its frame top."""
    ink = ink_rows(png, xs[0] + 0.02, xs[1] - 0.02, dpi)
    ys = [y / dpi for y in np.where(ink)[0] if TAB_TOP - 0.05 <= y / dpi <= 10.5]
    if not ys:
        return None
    return (ys[0] - TAB_TOP, ys[-1] + 1 / dpi - TAB_TOP)


def render(pptx, outdir, prefix, dpi=100):
    subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", outdir, pptx],
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    pdf = os.path.join(outdir, os.path.splitext(os.path.basename(pptx))[0] + ".pdf")
    subprocess.run(["pdftoppm", "-r", str(dpi), "-png", pdf, os.path.join(outdir, prefix)], check=True)
    return sorted(f for f in os.listdir(outdir) if f.startswith(prefix) and f.endswith(".png"))


def measure(png, cols, top_in=0.3, dpi=100):
    """Relative (top, bottom) in inches of each coloured paragraph, from the box top.
    Rows of one colour are grouped into blocks split by gaps over 0.1 in; the k-th paragraph
    drawn in a colour takes that colour's k-th block."""
    a = np.asarray(Image.open(png).convert("RGB")).astype(int)
    blocks = {}
    for c in set(cols):
        m = (np.abs(a - np.array(c)).sum(axis=2) < 60).sum(axis=1) >= 2
        ys = np.where(m)[0]
        bl = []
        for y in ys:
            if bl and y - bl[-1][1] <= 0.1 * dpi:
                bl[-1][1] = y
            else:
                bl.append([y, y])
        blocks[c] = bl
    # the colour match misses the faint anti-aliased rows at the edge of the text (descenders);
    # extend each block up and down through rows that still carry ink, up to 0.04 in (Arabic
    # descenders reach further than Latin ones).
    ink = ((a.sum(axis=2) / 3) < 200).sum(axis=1) >= 2
    reach = int(0.04 * dpi)
    for c in blocks:
        for bl in blocks[c]:
            for _ in range(reach):
                if bl[1] + 1 < len(ink) and ink[bl[1] + 1]:
                    bl[1] += 1
                else:
                    break
            for _ in range(reach):
                if bl[0] - 1 >= 0 and ink[bl[0] - 1]:
                    bl[0] -= 1
                else:
                    break
    seen, res = {}, []
    for c in cols:
        k = seen.get(c, 0); seen[c] = k + 1
        b = blocks[c][k] if k < len(blocks[c]) else None
        res.append((b[0] / dpi - top_in, (b[1] + 1) / dpi - top_in) if b else None)
    return res


# ---------------------------------------------------------------- layout
def margin(lang):
    return EN_MARGIN if lang == "en" else 0.0


def plan_s1(lang, sh, rel, n_corp):
    """Slide 1 positions in inches and k1max, how many Corporate items fit on slide 1.
    rel: key -> list of (top, bottom) of each paragraph's ink relative to its box top."""
    inch = lambda e: e / EMU
    L = {}
    mk_top = inch(sh["mk"].top)
    mk_end = mk_top + rel["mk"][-1][1]
    ec_hoff = inch(sh["ec"].top - sh["ec_h"].top)
    co_hoff = inch(sh["co"].top - sh["co_h"].top)
    ec_h = mk_end + GAP; ec_top = ec_h + ec_hoff
    ec_end = ec_top + rel["ec"][-1][1] if rel["ec"] else ec_top
    co_h = ec_end + GAP; co_top = co_h + co_hoff
    L.update(mk=(mk_top, ec_h - mk_top - 0.02), ec_h=ec_h, ec=(ec_top, co_h - ec_top - 0.02),
             co_h=co_h, co=(co_top, S1_LIMIT - co_top))
    r1 = rel["c1"]
    k1 = 0
    while k1 < min(n_corp, len(r1)) and co_top + r1[k1][1] <= S1_LIMIT - margin(lang):
        k1 += 1
    L["k1max"] = k1
    L["c1_ends"] = [co_top + r[1] for r in r1]
    return L


def plan_s2(lang, sh, rel, k1, n):
    """Slide 2 with n Corporate items starting at item k1: the news, then the Event calendar,
    then the Economic calendar, each CAL_GAP under the ink of the block above.
    Returns positions and the bottom of the last block's ink."""
    inch = lambda e: e / EMU
    s2_top = inch(sh["s2"].top)
    r2 = rel["c2"]
    text_end = s2_top + (r2[k1 + n - 1][1] - r2[k1][0] + r2[0][0]) if n else s2_top
    y = text_end
    tops, heights = [], []
    for i, t in enumerate(sh["tables"]):
        off_top, off_bot = rel[f"tab{i}"]
        vis_top = y + CAL_GAP
        tops.append(vis_top - off_top)
        heights.append(off_bot)
        y = vis_top - off_top + off_bot
    return {"s2": (s2_top, max(text_end - s2_top + 0.04, 0.3)), "tables": tops, "table_h": heights,
            "s2_text_end": text_end, "s2_end": y, "n2": n}


def apply_layout(sh, L):
    E = lambda x: int(round(max(x, 0.05) * EMU))   # never a negative size, even on a broken plan
    sh["mk"].height = E(L["mk"][1])
    sh["ec_h"].top = E(L["ec_h"]); sh["ec"].top = E(L["ec"][0]); sh["ec"].height = E(L["ec"][1])
    sh["co_h"].top = E(L["co_h"]); sh["co"].top = E(L["co"][0]); sh["co"].height = E(L["co"][1])
    sh["s2"].height = E(L["s2"][1])
    for t, top, h in zip(sh["tables"], L["tables"], L["table_h"]):
        t.top = E(top); t.height = E(h)


def text_rows(png, dpi=100):
    """Rows holding ink that is not a header/table bar. Bars span most of the width; LibreOffice
    draws auto-fit header bars up to 0.04 in taller than PowerPoint does, so gaps are taken from
    the frame top (where PowerPoint draws the bar) to the last ink that is not bar."""
    a = np.asarray(Image.open(png).convert("L")).astype(int)[:, int(0.45 * dpi):int(8.05 * dpi)]
    n = (a < 200).sum(axis=1)
    return (n >= 2) & (n < 0.6 * a.shape[1])


def gap_above(ink, y, dpi=100):
    i = int(y * dpi) - 1
    while i > 0 and not ink[i]:
        i -= 1
    return y - (i + 1) / dpi


def settle(path, sh, L, lang, outdir):
    """One render of the saved deck, then nudge every header and table so the visible gap above
    it is exactly SECTION_GAP. Corrects the few hundredths by which the scratch measurement can
    miss (Arabic first-line ink in particular). Moves of more than 0.08 in are not applied but
    flagged, as they mean something else is wrong."""
    d = tempfile.mkdtemp(dir=outdir)
    pngs = render(path, d, "s")
    ink1, ink2 = text_rows(os.path.join(d, pngs[0])), text_rows(os.path.join(d, pngs[1]))
    shutil.rmtree(d, ignore_errors=True)
    E = lambda x: int(round(x * EMU))
    moves = []
    # slide 1: Economic block moves by d1; Corporate block by d1 plus its own correction
    d1 = GAP - gap_above(ink1, L["ec_h"])
    d2 = d1 + GAP - gap_above(ink1, L["co_h"])
    # slide 2: each table moves with the one above it plus its own correction
    tops = [t + o for t, o in zip(L["tables"], L["tab_off"])]
    dt, acc = [], 0.0
    for i, y in enumerate(tops):
        acc = (acc if i else 0.0) + CAL_GAP - gap_above(ink2, y)
        dt.append(acc)
    for name, dv in [("Economic header", d1), ("Corporate header", d2)] + [(f"table {i + 1}", v) for i, v in enumerate(dt)]:
        if abs(dv) > 0.08:
            note("FLAG", f"{lang} settle: {name} would need a {dv:+.2f} in move; not applied, check the layout.")
            return False
    if abs(d1) >= 0.01:
        L["ec_h"] += d1; L["ec"] = (L["ec"][0] + d1, L["ec"][1] + d2 - d1)
        L["mk"] = (L["mk"][0], L["mk"][1] + d1)
    if abs(d2) >= 0.01:
        L["co_h"] += d2; L["co"] = (L["co"][0] + d2, max(S1_LIMIT - (L["co"][0] + d2), 0.3))
        end = L.get("s1_end", 0) + d2
        if L.get("k1") and end > S1_LIMIT + 0.01:
            note("FLAG", f"{lang} settle: slide 1 Corporate text now ends at {end:.2f} in, past {S1_LIMIT} in.")
    L["tables"] = [t + v for t, v in zip(L["tables"], dt)]
    if L["tables"]:
        L["s2"] = (L["s2"][0], L["s2"][1] + dt[0])
        if L["tables"][-1] + L["table_h"][-1] > TABLE_BOTTOM + 0.01:
            note("FLAG", f"{lang} settle: slide 2 now ends at {L['tables'][-1] + L['table_h'][-1]:.2f} in, past {TABLE_BOTTOM} in.")
    apply_layout(sh, L)
    moves = [f"{n} {v:+.2f}" for n, v in [("Econ", d1), ("Corp", d2)] + [(f"T{i + 1}", v) for i, v in enumerate(dt)] if abs(v) >= 0.01]
    if moves:
        note("OK", f"{lang} settle: nudged " + ", ".join(moves) + " in to hit the house gap.")
    return True


def spacing_report(png1, png2, L, lang, dpi=100):
    """Measure the finished render: every visible gap between consecutive blocks, and the page
    bottom. Gaps within 0.02 in of SECTION_GAP are OK."""
    out = []
    rows = text_rows
    ink1 = rows(png1); ink2 = rows(png2)
    checks = [("slide 1 Markets text to Economic header", ink1, L["ec_h"], GAP),
              ("slide 1 Economic text to Corporate header", ink1, L["co_h"], GAP)]
    names = ["Event calendar", "Economic calendar"]
    above = "Corporate news"
    for i, top in enumerate(L["tables"]):
        # the visible top of a table is its frame top plus the measured offset
        checks.append((f"slide 2 {above} to {names[i] if i < 2 else 'table'}", ink2, top + L["tab_off"][i], CAL_GAP))
        above = names[i] if i < 2 else "table"
    for label, ink, y, want in checks:
        g = gap_above(ink, y)
        lvl = "OK" if abs(g - want) <= 0.02 else "FLAG"
        out.append((lvl, f"{lang} spacing: {label} {g:.2f} in (house gap {want:.2f})"))
    for label, ink, lim in (("slide 1", ink1, S1_LIMIT + 0.02), ("slide 2", ink2, TABLE_BOTTOM + 0.02)):
        # scan to 10.70 in: below the limit plus tolerance (a 10.62 in scan could never show an
        # overrun, fixed 13 Oct 2026) and above the page number digits (about 10.74 in)
        ys = np.where(ink[: int(round(10.70 * dpi))])[0]
        last = (ys[-1] + 1) / dpi if len(ys) else 0
        if last > lim:
            out.append(("FLAG", f"{lang} {label}: content runs to {last:.2f} in, past the {lim - 0.02:.2f} in limit."))
    return out


def proxy_font_check():
    """EN measurement is only right with the measuring copy of Arial Nova Cond installed
    (make_measure_font.py: the real font, kerning off, +0.25% width)."""
    try:
        f = subprocess.run(["fc-match", "-f", "%{file}", "Arial Nova Cond"], capture_output=True, text=True).stdout
    except FileNotFoundError:
        return
    if MEASURE_FONT not in f:
        sys.exit(f"Arial Nova Cond resolves to {f!r}. Run: python3 make_measure_font.py fonts/ArialNovaCond.ttf "
                 f"~/.fonts/{MEASURE_FONT} && fc-cache -f   (and keep no other Arial Nova Cond in ~/.fonts), "
                 "or the EN layout will be wrong.")


# ---------------------------------------------------------------- checks
def numbers(s, lang):
    # period labels (1Q27, 1H26) count as their year in both languages: the Arabic keeps them in
    # Latin (13 Oct 2026) or writes الربع الأول من عام 2027
    s = re.sub(r"(?<![A-Za-z0-9])(\d)[QHM](\d\d)(?![0-9])", lambda m: "20" + m.group(2), s)
    if lang == "en":   # "$2 a barrel" / "USD 2" against Arabic "دولارين" (13 Oct 2026 test)
        s = re.sub(r"(?:\$|USD\s?)(\d)\b(?![.,]\d)", r"\1", s)
    else:              # dual forms the Arabic uses for 2
        s = re.sub(r"(دولارين|مليوني|ملياري)", " 2 ", s)
        # "22 ألف" is 22,000 in the English (20 and 21 Oct 2026 finals flagged by mistake)
        s = re.sub(r"(\d+(?:\.\d+)?)\s*(?:ألف|آلاف)", lambda m: f"{round(float(m.group(1)) * 1000):,}", s)
    xs = [x.rstrip(".,") for x in re.findall(r"\d[\d,]*\.?\d*", s)]
    years = sorted(set(x for x in xs if re.fullmatch(r"20\d\d", x)))   # "1Q27 and 2Q27" is one Arabic phrase
    return sorted(x for x in xs if not re.fullmatch(r"20\d\d", x)) + years


def tickers(s):
    return sorted(re.findall(r"\(([A-Z0-9]+ AB)\)", s))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--doc", required=True); ap.add_argument("--en", required=True); ap.add_argument("--ar", required=True)
    ap.add_argument("--outdir", required=True); ap.add_argument("--force", action="store_true"); ap.add_argument("--render", action="store_true")
    ap.add_argument("--results", action="store_true",
                    help="results day: no Corporate item on slide 1 (the Results table goes there)")
    ap.add_argument("--en-skip", default="", help="English boxes a teammate pastes by hand, comma separated: mk, ec, co")
    ap.add_argument("--economic", help="which Draft-tab Economic items go in, 1-based, e.g. 1,4 "
                    "(needed when the doc holds more or fewer than two; the team runs two)")
    a = ap.parse_args()
    en_skip = {x.strip() for x in a.en_skip.split(",") if x.strip()}
    if en_skip - {"mk", "ec", "co"}:
        sys.exit("--en-skip takes mk, ec and co only (comma separated).")
    if "co" in en_skip:
        en_skip.add("s2")   # Corporate continues on slide 2
    os.makedirs(a.outdir, exist_ok=True)
    tabs = read_doc(a.doc)
    # news source: Final - Sequenced while that tab existed; since 6 Oct 2026 the doc has one
    # tab, Draft, which the team finalizes in the morning (Economic News / Corporate News)
    src = "Final - Sequenced" if "Final - Sequenced" in tabs else "Draft"
    final = news_items(tabs[src])
    if src == "Final - Sequenced" and not (final["Economic"] or final["Corporate"]):
        src = "Draft"; final = news_items(tabs["Draft"])
    note("OK", f"News taken from the '{src}' tab, in its order.")
    wu = writeups(tabs["Draft"])
    ec, co = final["Economic"], final["Corporate"]
    if a.economic:
        pick = [int(x) - 1 for x in a.economic.split(",")]
        ec = [ec[i] for i in pick]
        note("OK", f"Economic items {a.economic} used, as asked.")
    elif len(ec) != 2:
        # 13 Oct 2026 test: four Economic items in the doc pushed every Corporate item off slide 1.
        # The team trims Economic to two in the deck, not always in the doc: ask, never guess.
        sys.exit(f"{src} Economic has {len(ec)} items; the team runs two. Ask which go in and rerun with "
                 "--economic, e.g. --economic 1,4:\n" + "\n".join(f"  {i + 1}. {x['en'][:90]}" for i, x in enumerate(ec)))
    content = {lang: {"mk": [x for x in wu[lang] if x], "ec": [i[lang] for i in ec], "co": [i[lang] for i in co]}
               for lang in ("en", "ar")}
    for lang in ("en", "ar"):
        if not content[lang]["mk"]:
            # with no Markets text the layout cannot be planned (this crashed with an IndexError
            # until 13 Oct 2026); the Arabic deck always needs both write-ups
            sys.exit(f"The Draft tab has no {'English' if lang == 'en' else 'Arabic'} Global or TASI write-up under its "
                     "Global / TASI headings. Post the approved write-ups to the Draft tab, then rerun. Nothing was filled.")
    for lang in ("en", "ar"):
        for k, v in content[lang].items():
            for t in v:
                if not re.search(r"\)\s*$", t):
                    note("FLAG", f"{lang} {k}: bullet without a closing source tag: {t[:50]}")

    # 1. measure (one render per language)
    proxy_font_check()
    tmp = tempfile.mkdtemp(dir=a.outdir)
    rels = {}
    for lang, path in (("en", a.en), ("ar", a.ar)):
        c = content[lang]
        mpath = os.path.join(tmp, f"measure_{lang}.pptx")
        jobs = build_measure_deck(path, lang, [("mk", c["mk"]), ("ec", c["ec"]), ("co", c["co"]), ("s2", c["co"])], mpath)
        pngs = render(mpath, tmp, f"m_{lang}")
        rel = {}
        for (key, cols), png in zip(jobs, pngs):
            if key.startswith("tab"):
                rel[key] = measure_table(os.path.join(tmp, png), cols)
                if rel[key] is None:
                    sys.exit(f"{lang} {key}: calendar table could not be measured.")
                continue
            rel[{"co": "c1", "s2": "c2"}.get(key, key)] = measure(os.path.join(tmp, png), cols)
        for k, v in list(rel.items()):
            if k.startswith("tab") or not any(x is None for x in v):
                continue
            if k in ("c1", "c2"):
                # long news days: items past the bottom of the scratch page cannot be on this
                # slide anyway; keep the measured ones, the rest count as not fitting
                rel[k] = v[:v.index(None)]
            else:
                sys.exit(f"{lang} {k}: a paragraph could not be measured (text ran off the scratch page?)")
        rels[lang] = rel

    # 2. plan: one slide 1 break for both decks, then as many slide 2 items as fit in both
    plans, prs, shs = {}, {}, {}
    for lang, path in (("en", a.en), ("ar", a.ar)):
        prs[lang] = Presentation(path)
        clean_hidden_titles(prs[lang], note)
        shs[lang] = find_shapes(prs[lang], lang)
        plans[lang] = plan_s1(lang, shs[lang], rels[lang], len(co))
    k1 = min(p["k1max"] for p in plans.values())
    if a.results:
        # results day: the Results table takes the slide 1 Corporate space and every Corporate item
        # goes to slide 2 (all 47 results days in the 15 Jul to 21 Oct 2026 finals)
        k1 = 0
        note("OK", "Results day: all Corporate items placed on slide 2; Task 10 puts the Results table under "
                   "the slide 1 Corporate header.")
    elif len({p["k1max"] for p in plans.values()}) > 1:
        note("OK", "Slide 1 Corporate break set to " + str(k1) + " items in both decks (fits: "
             + ", ".join(f"{l.upper()} {p['k1max']}" for l, p in plans.items()) + ").")
    n2 = 0
    while k1 + n2 < min([len(co)] + [len(rels[l]["c2"]) for l in plans]) and all(
            plan_s2(l, shs[l], rels[l], k1, n2 + 1)["s2_end"] <= TABLE_BOTTOM - margin(l) for l in plans):
        n2 += 1
    n_fit = k1 + n2
    for lang in plans:
        L2 = plan_s2(lang, shs[lang], rels[lang], k1, n2)
        L2["tab_off"] = [rels[lang][f"tab{i}"][0] for i in range(len(shs[lang]["tables"]))]
        plans[lang].update(L2); plans[lang]["k1"] = k1
        plans[lang]["s1_end"] = plans[lang]["c1_ends"][k1 - 1] if k1 else plans[lang]["co"][0]
        if k1 == 0 and not a.results:
            note("FLAG", f"{lang}: no Corporate item fits on slide 1; the slide 1 Corporate header has nothing under it. "
                         "Move the header to slide 2 only, or shorten the Markets/Economic text.")
        if L2["s2_end"] > TABLE_BOTTOM:
            note("FLAG", f"{lang}: slide 2 calendars end at {L2['s2_end']:.2f} in, past {TABLE_BOTTOM} in, even with no news on slide 2.")
    if n_fit < len(co):
        note("FLAG", f"Only {n_fit} of {len(co)} corporate items fit; left out of both decks (lowest ranked): "
             + "; ".join(i["en"][:60] for i in co[n_fit:]))

    # 3. fill and save
    for lang in ("en", "ar"):
        p, L = prs[lang], plans[lang]
        sh = shs[lang]
        c = content[lang]
        targets = [("mk", sh["mk"], c["mk"]), ("ec", sh["ec"], c["ec"]),
                   ("co", sh["co"], c["co"][:k1]), ("s2", sh["s2"], c["co"][k1:n_fit])]
        for key, shape, paras in targets:
            if lang == "en" and key in en_skip:
                continue   # the teammate pastes this box
            if not PLACEHOLDER.search(shape.text_frame.text) and shape.text_frame.text.strip() and not a.force:
                note("FLAG", f"{lang} {key}: box already has team text, left unchanged (use --force to overwrite).")
                continue
            fill(lang, shape, paras)
        apply_layout(sh, L)
        out = os.path.join(a.outdir, os.path.basename(a.en if lang == "en" else a.ar))
        p.save(out)
        if lang == "en" and en_skip:
            # boxes still hold placeholders, so a render cannot be used to settle the gaps
            names = {"mk": "Markets", "ec": "Economic", "co": "Corporate"}
            note("FLAG", "en: left for the teammate to paste: " + ", ".join(names[k] for k in ("mk", "ec", "co") if k in en_skip)
                 + f". Corporate breaks after item {k1}: items 1 to {k1} on slide 1, {k1 + 1} to {n_fit} on slide 2."
                 + " English gaps not settled and EN/AR figures not compared here; run check_final.py after the paste.")
        elif settle(out, sh, L, lang, a.outdir):
            p.save(out)
        note("OK", f"{lang}: markets {len(c['mk'])}, economic {len(c['ec'])}, corporate slide 1 {k1}, slide 2 {n_fit - k1}.")

    # 4. checks on the outputs
    outs = {lang: Presentation(os.path.join(a.outdir, os.path.basename(a.en if lang == 'en' else a.ar))) for lang in ("en", "ar")}
    paras = {}
    for lang, p in outs.items():
        sh = find_shapes(p, lang)
        paras[lang] = [x.text for k in ("mk", "ec", "co", "s2") for x in sh[k].text_frame.paragraphs if x.text.strip()]
        if any(PLACEHOLDER.search(t) for t in paras[lang]) and not (lang == "en" and en_skip):
            note("FLAG", f"{lang}: a placeholder is still in a commentary box.")
    if len(paras["en"]) != len(paras["ar"]) and not en_skip:
        note("CONFLICT", f"EN has {len(paras['en'])} paragraphs, AR {len(paras['ar'])}.")
    for i, (e, r) in enumerate(zip(paras["en"], paras["ar"]) if not en_skip else []):
        if numbers(e, "en") != numbers(r, "ar") or tickers(e) != tickers(r):
            note("FLAG", f"EN/AR mismatch in paragraph {i + 1}: EN {numbers(e, 'en')} {tickers(e)} vs AR {numbers(r, 'ar')} {tickers(r)}")
    if a.render:
        for lang in ("en", "ar"):
            pngs = render(os.path.join(a.outdir, os.path.basename(a.en if lang == "en" else a.ar)), a.outdir, f"check_{lang}")
            if lang == "en" and en_skip:
                continue   # rendered for a look only; its gaps mean nothing until the teammate pastes
            for lvl, m in spacing_report(os.path.join(a.outdir, pngs[0]), os.path.join(a.outdir, pngs[1]), plans[lang], lang):
                note(lvl, m)
    shutil.rmtree(tmp, ignore_errors=True)
    json.dump({k: {kk: vv for kk, vv in v.items()} for k, v in plans.items()}, open(os.path.join(a.outdir, "layout.json"), "w"), indent=1)
    for lvl, m in log:
        print(lvl, m)


if __name__ == "__main__":
    main()
