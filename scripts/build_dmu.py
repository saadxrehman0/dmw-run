"""Build the Daily Market Update (DMU) deck (v2, 21 Oct 2026).

Rebuilt after the review of the team's 80 DMUs (1 Jul to 21 Oct 2026); see
references/dmu-structure.md and references/dmu-writing.md for the rules it follows.

Usage:
  python build_dmu.py --template "<previous or day-folder DMU>.pptx" --data "DMW - Data.xlsx" \
      --date YYYY-MM-DD --out "Daily Market Update - YYYYMMDD.pptx" \
      [--technicals "Technicals -ENG-YYYYMMDD.pptx"] [--content content.json] \
      [--forecasts "DMU Commodity Data - YYYYMMDD.xlsx"] [--names name_overrides_dmu.json] \
      [--recap recap.png] [--futures futures.png] [--results-template "<a DMU with a Results slide>.pptx"] \
      [--heatmap heatmap.png] [--team-technical "<team's DMU in the day folder>.pptx"]

--heatmap: Saad's IMAP snip, padded black to the Saudi slide heatmap frame (replaces heat.py, 10 Nov 2026).
--team-technical: the TASI TECHNICAL slide copied whole from the team's DMU (all shapes, pictures
re-linked); --technicals is then ignored (replaces post.py, 10 Nov 2026).

content.json (all keys optional):
  global_headlines   [5 str]  slide 2; [0] = Global write-up exactly as in the final DMW deck
  events             [str]    "DD-Mon-YY: <event>: x (P), y (F)"
  commodity_headlines[2-3 str] oil first, then gold (last settlement story)
  commodity_captions [2 str]  [oil caption, gold caption]: the move on the chart's last date
  saudi_headlines    [3-4 str] DMW Economic item first, then Corporate items, copied from the final DMW
  results            {"quarter": "3Q26", "prior": "3Q25",
                      "rows": [[Company, TASI code, BBG code, NP, NP prior, chg SAR mn, chg %,
                                consensus, var %, anbc est, var %], ...],
                      "commentary": [str, ...]}           # omit or [] rows = no Results slide
  target_prices      {name: value}  normally omitted: Consensus TP is left blank for the team

Slides are found by title, never by position (decks run 6 to 8 slides).
Writes <out>_checks.md: FLAG = needs a human, INFO = for the record.
"""
import argparse, copy, datetime, hashlib, json, os, re
from decimal import Decimal, ROUND_HALF_UP
import openpyxl
from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor

GREEN, RED = RGBColor(0x00, 0xB0, 0x50), RGBColor(0xFF, 0x00, 0x00)
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
C = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
R_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
EMU = 914400
MONTH_FOLDER = {1: "01. Jan", 2: "02. Feb", 3: "03. Mar", 4: "04. April", 5: "05. May", 6: "06. Jun",
                7: "07. Jul", 8: "08. Aug", 9: "09. Sep", 10: "10. Oct", 11: "11. Nov", 12: "12. Dec"}
checks = []
LEVEL = re.compile(r"\d{1,2},\d{3}(?:\.\d+)?(?:/\d{1,3})?(?:\s*[–-]\s*\d{1,2},\d{3}(?:\.\d+)?)?")


def note(level, msg): checks.append(f"{level}: {msg}")


# ---------------------------------------------------------------- number formats (house rules)
def hu(v, nd=1):
    return Decimal(repr(float(v))).quantize(Decimal(1).scaleb(-nd), rounding=ROUND_HALF_UP)


def fnum(v, nd=1):
    if v is None or (isinstance(v, str) and not re.match(r"^-?[\d.]+$", v.strip())):
        return str(v).strip() if v not in (None, "") else "-"
    return f"{hu(float(v), nd):,.{nd}f}"


def fpct(v):
    if v is None or (isinstance(v, str) and not re.match(r"^-?[\d.]+$", v.strip())):
        return "-", None
    x = float(v) * 100
    t = f"{hu(x):.1f}%"
    if x < 0 and not t.startswith("-"): t = "-" + t
    if t == "0.0%": return t, "default"
    return t, (RED if x < 0 else GREEN)


# ---------------------------------------------------------------- slide and shape lookup
def flat(shapes):
    for s in shapes:
        yield s
        if s.shape_type == 6: yield from flat(s.shapes)


def slide_title(slide):
    for s in flat(slide.shapes):
        if getattr(s, "has_text_frame", False) and s.has_text_frame:
            for p in s.text_frame.paragraphs:
                t = p.text.strip().upper()
                if t in ("GLOBAL MARKETS", "COMMODITY MARKETS", "SAUDI MARKET") or \
                        t.startswith(("RESULT ANNOUNCEMENTS", "SUKUK", "TASI TECHNICAL")):
                    return t
    return ""


def find_slides(prs, prefix):
    return [s for s in prs.slides if slide_title(s).startswith(prefix)]


def one_slide(prs, prefix):
    found = find_slides(prs, prefix)
    if not found: raise SystemExit(f"no '{prefix}' slide in the template")
    return found[0]


def text_shapes(slide):
    return [s for s in flat(slide.shapes) if getattr(s, "has_text_frame", False) and s.has_text_frame]


def body_box(slide, header_text):
    """The body box under a header: nearest wide box below it."""
    hdr = next((s for s in text_shapes(slide) if s.text_frame.text.strip() == header_text), None)
    if hdr is None: return None
    cands = [s for s in text_shapes(slide) if s is not hdr and s.top >= hdr.top - EMU * 0.05
             and s.text_frame.text.strip() != header_text and s.width > 3 * EMU and s.height > 0.5 * EMU]
    return min(cands, key=lambda s: s.top - hdr.top) if cands else None


def charts(slide):
    return sorted([s for s in flat(slide.shapes) if getattr(s, "has_chart", False) and s.has_chart],
                  key=lambda s: s.top)


def pictures(slide):
    return [s for s in flat(slide.shapes) if s.shape_type == 13]


# ---------------------------------------------------------------- text, keeping run formatting
def set_cell(cell_or_tf, text, color=None):
    tf = cell_or_tf.text_frame if hasattr(cell_or_tf, "text_frame") else cell_or_tf
    p = tf.paragraphs[0]
    if p.runs:
        p.runs[0].text = text
        for r in p.runs[1:]: r._r.getparent().remove(r._r)
    else:
        p.add_run().text = text
    for extra in tf.paragraphs[1:]: extra._p.getparent().remove(extra._p)
    if text == "":
        # an empty run takes its line height from endParaRPr; without one the cell grows to the 18 pt
        # default (1 Nov 2026: blank Consensus TP cells stretched the leaders/laggards rows)
        rpr = p.runs[0]._r.find(A + "rPr")
        end = p._p.find(A + "endParaRPr")
        if end is not None: p._p.remove(end)
        end = copy.deepcopy(rpr) if rpr is not None else etree.Element(A + "rPr")
        end.tag = A + "endParaRPr"
        p._p.append(end)
        p._p.remove(p.runs[0]._r)   # no empty run: the paragraph keeps only the sized end mark
    if not p.runs:
        return
    if color == "default":
        rpr = p.runs[0]._r.find(A + "rPr")
        if rpr is not None:
            for fill in rpr.findall(A + "solidFill"): rpr.remove(fill)
    elif color is not None:
        p.runs[0].font.color.rgb = color


def _proto_paragraph(tf):
    for p in tf.paragraphs:
        if p.runs: return copy.deepcopy(p._p)
    return copy.deepcopy(tf.paragraphs[0]._p)


def set_lines(tf, lines):
    """One paragraph per line, every line in the format of the template's first run."""
    proto = _proto_paragraph(tf)
    body = tf._txBody
    for p in list(tf.paragraphs): body.remove(p._p)
    for line in lines:
        np_ = copy.deepcopy(proto)
        runs = np_.findall(A + "r")
        for r in runs[1:]: np_.remove(r)
        for br in np_.findall(A + "br"): np_.remove(br)   # template line breaks (source row, 29 Oct 2026)
        if runs:
            runs[0].find(A + "t").text = line
        else:
            r = etree.SubElement(np_, A + "r"); etree.SubElement(r, A + "t").text = line
        # keep endParaRPr last
        end = np_.find(A + "endParaRPr")
        if end is not None: np_.remove(end); np_.append(end)
        body.append(np_)


def set_event_lines(tf, lines):
    """Events: 'DD-Mon-YY: ' in the template's date run (bold), the rest in its text run (regular)."""
    proto = _proto_paragraph(tf)
    runs = proto.findall(A + "r")
    if not runs:
        return set_lines(tf, lines)
    if len(runs) < 2:   # template had no events (or one-run lines): build the date/text pair
        date_r, text_r = copy.deepcopy(runs[0]), copy.deepcopy(runs[0])
    else:
        date_r, text_r = runs[0], runs[1]
    for r, b in ((date_r, "1"), (text_r, None)):
        rpr = r.find(A + "rPr")
        if rpr is None: rpr = etree.Element(A + "rPr"); r.insert(0, rpr)
        if b: rpr.set("b", b)
        elif rpr.get("b") in ("1", "true"): rpr.set("b", "0")
    body = tf._txBody
    for p in list(tf.paragraphs): body.remove(p._p)
    for line in lines:
        m = re.match(r"^(\d{2}-[A-Z][a-z]{2}-\d{2,4}:\s*)(.*)$", line.strip())
        head, rest = (m.group(1), m.group(2)) if m else ("", line)
        np_ = copy.deepcopy(proto)
        for r in np_.findall(A + "r"): np_.remove(r)
        end = np_.find(A + "endParaRPr")
        dr, tr = copy.deepcopy(date_r), copy.deepcopy(text_r)
        dr.find(A + "t").text, tr.find(A + "t").text = head, rest
        for r in (dr, tr):
            if end is not None: end.addprevious(r)
            else: np_.append(r)
        body.append(np_)


def set_rich_paragraphs(tf, paras):
    """paras = [[(text, bold), ...], ...]; base format = template's first run, bold per source run."""
    proto = _proto_paragraph(tf)
    base_r = proto.find(A + "r")
    body = tf._txBody
    for p in list(tf.paragraphs): body.remove(p._p)
    for runs in paras:
        np_ = copy.deepcopy(proto)
        for r in np_.findall(A + "r"): np_.remove(r)
        end = np_.find(A + "endParaRPr")
        for text, bold in runs:
            r = copy.deepcopy(base_r) if base_r is not None else etree.Element(A + "r")
            if r.find(A + "t") is None: etree.SubElement(r, A + "t")
            r.find(A + "t").text = text
            rpr = r.find(A + "rPr")
            if rpr is None:
                rpr = etree.Element(A + "rPr"); r.insert(0, rpr)
            if bold: rpr.set("b", "1")
            elif rpr.get("b") in ("1", "true"): rpr.set("b", "0")
            elif "b" in rpr.attrib: del rpr.attrib["b"]
            if end is not None: end.addprevious(r)
            else: np_.append(r)
        body.append(np_)


# ---------------------------------------------------------------- charts
def excel_serial(d):
    if isinstance(d, datetime.datetime): d = d.date()
    return (d - datetime.date(1899, 12, 30)).days


def parse_ref(f):
    m = re.match(r"^'?([^'!]+)'?!\$([A-Z]+)\$(\d+)(?::\$([A-Z]+)\$(\d+))?$", f.strip())
    return (m.group(1), m.group(2), int(m.group(3)), int(m.group(5)) if m.group(5) else None) if m else None


def col_idx(letters):
    n = 0
    for ch in letters: n = n * 26 + ord(ch) - 64
    return n


def read_col(ws, col, start, end):
    c = col_idx(col)
    return [r[0] for r in ws.iter_rows(min_row=start, max_row=end, min_col=c, max_col=c, values_only=True)]


def last_row(ws, col, start, limit=6000):
    vals = read_col(ws, col, start, limit)
    last = None
    for i, v in enumerate(vals):
        if v not in (None, ""): last = start + i
    return last


def fill_num_cache(cache, vals):
    for pt in cache.findall(C + "pt"): cache.remove(pt)
    cache.find(C + "ptCount").set("val", str(len(vals)))
    for i, v in enumerate(vals):
        if v is None or v == "": continue
        pt = etree.SubElement(cache, C + "pt", idx=str(i))
        etree.SubElement(pt, C + "v").text = repr(float(v)) if not isinstance(v, int) else str(v)


def fill_str_cache(cache, vals):
    for pt in cache.findall(C + "pt"): cache.remove(pt)
    cache.find(C + "ptCount").set("val", str(len(vals)))
    for i, v in enumerate(vals):
        pt = etree.SubElement(cache, C + "pt", idx=str(i))
        etree.SubElement(pt, C + "v").text = str(v)


def pe_label(v):
    return v.strftime("Dec-%y") if isinstance(v, (datetime.datetime, datetime.date)) else str(v)


def refresh_chart(chart_shape, wb, data_path, day, label):
    """Rebuild each series' cache from the data file, extend the ranges to the last filled row,
    and point the external link at today's data file."""
    cs = chart_shape.chart._chartSpace
    sers = cs.findall(".//" + C + "ser")
    if not sers: note("FLAG", f"{label}: no series found"); return
    cat_f = sers[0].find(C + "cat").find(".//" + C + "f")
    ref = parse_ref(cat_f.text) if cat_f is not None else None
    if ref is None: note("FLAG", f"{label}: category range unreadable; chart left as it was"); return
    sheet, col, start, _ = ref
    if sheet not in wb.sheetnames: note("FLAG", f"{label}: sheet '{sheet}' not in data file"); return
    ws = wb[sheet]
    end = last_row(ws, col, start)
    if end is None: note("FLAG", f"{label}: no data in {sheet}!{col}{start}:"); return
    # dated series: stop at the last weekday before the edition (trailing weekend carry-forward rows
    # and rows dated on/after the edition are left out, as on the team's Monday decks)
    while end > start:
        v = read_col(ws, col, end, end)[0]
        if isinstance(v, datetime.datetime) and (v.date() >= day or v.weekday() >= 5):
            end -= 1; continue
        break
    last_cat = None
    for ser in sers:
        for tag in ("cat", "val"):
            el = ser.find(C + tag)
            if el is None: continue
            f = el.find(".//" + C + "f")
            r = parse_ref(f.text) if f is not None else None
            if r is None: continue
            sh, c, s0, _ = r
            f.text = re.sub(r"(:\$[A-Z]+\$)\d+$", lambda m: m.group(1) + str(end), f.text)
            vals = read_col(wb[sh], c, s0, end)
            nc, sc = el.find(".//" + C + "numCache"), el.find(".//" + C + "strCache")
            if tag == "cat":
                if nc is not None:
                    vals = [excel_serial(v) if isinstance(v, (datetime.datetime, datetime.date)) else v for v in vals]
                    fill_num_cache(nc, vals)
                elif sc is not None:
                    vals = [pe_label(v) for v in vals]
                    fill_str_cache(sc, vals)
                last_cat = vals[-1]
            else:
                if nc is not None: fill_num_cache(nc, vals)
    # external link -> today's data file
    for rel in chart_shape.chart.part.rels.values():
        if rel.is_external and "DMW" in rel.target_ref:
            old = rel.target_ref
            new = re.sub(r"\\\d{2}\.%20[A-Za-z]+\\\d{8}\\",
                         lambda m: "\\" + MONTH_FOLDER[day.month].replace(" ", "%20") + "\\" + day.strftime("%Y%m%d") + "\\",
                         old)
            if new != old:
                rel._target = new
                rel.__dict__["target_ref"] = new   # target_ref is a cached property
            if day.strftime("%Y%m%d") not in rel.target_ref:
                note("FLAG", f"{label}: external link not re-pointed ({old[-80:]})")
    if isinstance(last_cat, (int, float)) and last_cat > 40000:
        last_cat = (datetime.date(1899, 12, 30) + datetime.timedelta(days=int(last_cat))).isoformat()
    note("INFO", f"{label}: rebuilt from {sheet}!{col}{start}:{col}{end}, last category {last_cat}")
    return last_cat


# ---------------------------------------------------------------- pictures
# Fixed picture frames on the Global slide, inches (x, y, w, h) (Saad, 29 Oct 2026). Shaped like the
# Bloomberg snips (Oct 2026 medians: recap 3.39, futures 3.04) so nothing is stretched; recap at full
# left-column width, futures at the valuations table's width. The team's hand-set frames drifted
# (recap 6.75-7.07 x 1.95-2.20, futures 5.16-5.82 x 1.72-2.12) and stretched images by up to 8-16%.
FRAMES = {"recap": (0.25, 1.17, 6.92, 2.04), "futures": (7.50, 3.85, 5.43, 1.79)}
RECAP_LABEL_TOP = 0.88   # MARKET RECAP label moved up 0.09 in so the taller recap keeps its bottom at 3.21
PAD_FLAG = 0.10          # more padding than this usually means a bad crop


def pad_to_ratio(image_path, ratio, label):
    """Pad the snip with its own edge colour (recap: Bloomberg dark grey; futures: near-white) to the
    frame's width/height ratio, centred. Returns the path of the padded copy."""
    from PIL import Image
    from statistics import median
    im = Image.open(image_path).convert("RGB")
    w, h = im.size
    edge = [im.getpixel((x, y)) for x in range(0, w, max(1, w // 40)) for y in (0, h - 1)] + \
           [im.getpixel((x, y)) for y in range(0, h, max(1, h // 20)) for x in (0, w - 1)]
    fill = tuple(int(median(c[i] for c in edge)) for i in range(3))
    if w / h < ratio: nw, nh = round(h * ratio), h
    else: nw, nh = w, round(w / ratio)
    pad = max(nw / w, nh / h) - 1
    if pad > PAD_FLAG:
        note("FLAG", f"{label}: snip {w}x{h}px needs {pad * 100:.0f}% padding to fit the frame: check the crop")
    out = Image.new("RGB", (nw, nh), fill)
    out.paste(im, ((nw - w) // 2, (nh - h) // 2))
    p = os.path.splitext(image_path)[0] + "_framed.png"
    out.save(p)
    note("INFO", f"{label}: {w}x{h}px padded {pad * 100:.1f}% with {fill} to {nw}x{nh}px")
    return p


def replace_picture(slide, pic, image_path, label, frame=None):
    """Swap the picture for a new image part. With `frame` (x, y, w, h in inches) the picture is set to
    that fixed frame and the snip is padded to its shape first, so nothing is stretched. Without it the
    template frame is kept as it is."""
    from PIL import Image
    if frame:
        image_path = pad_to_ratio(image_path, frame[2] / frame[3], label)
        pic.left, pic.top, pic.width, pic.height = (int(v * EMU) for v in frame)
    img_part, rid = slide.part.get_or_add_image_part(image_path)
    blip = pic._element.blipFill.find(A + "blip")
    blip.set(R_NS + "embed", rid)
    w, h = Image.open(image_path).size
    old_ratio = pic.height / pic.width
    new_ratio = h / w
    if abs(new_ratio / old_ratio - 1) > 0.08:
        note("FLAG", f"{label}: new image shape {w}x{h}px differs from the frame by "
                     f"{(new_ratio / old_ratio - 1) * 100:+.0f}% in height: check the crop")
    note("INFO", f"{label}: replaced with {os.path.basename(image_path)} ({w}x{h}px)")


# Fixed right column on the Commodities slide, inches (Saad, 29 Oct 2026). The team placed it by hand
# and it drifted every day (84 decks: charts x 7.67 to 8.58, widths 4.6 to 5.1). Medians of those decks.
COMMODITY_LAYOUT = {"oil chart": (7.83, 1.15, 5.07, 3.10), "gold chart": (7.83, 4.40, 5.07, 3.10),
                    "oil caption": (7.83, 0.87, 5.50, None), "gold caption": (7.83, 4.25, 5.50, None),
                    "USD/bbl": (7.92, 1.24, None, None)}


def fix_commodity_layout(slide, chs, caps):
    def put(s, spec):
        x, y, w, h = spec
        s.left, s.top = int(x * EMU), int(y * EMU)
        if w: s.width = int(w * EMU)
        if h: s.height = int(h * EMU)
    if len(chs) >= 2:
        put(chs[0], COMMODITY_LAYOUT["oil chart"]); put(chs[1], COMMODITY_LAYOUT["gold chart"])
    if len(caps) == 2:
        put(caps[0], COMMODITY_LAYOUT["oil caption"]); put(caps[1], COMMODITY_LAYOUT["gold caption"])
    usd = [s for s in text_shapes(slide) if s.text_frame.text.strip() == "USD/bbl" and s.left > 0]
    for s in usd: put(s, COMMODITY_LAYOUT["USD/bbl"])
    # A caption that wraps to two lines ran into its chart's own title (2 Nov 2026: the 15-word gold
    # headline, 5.48 in). Captions stay word for word, so the chart under a wrapping caption moves down
    # one line (0.20 in) and loses the same height, keeping its bottom edge (USD/bbl label moves with oil).
    if len(chs) >= 2 and len(caps) == 2:
        for k, (cap, ch) in enumerate(zip(caps, chs[:2])):
            w_in = caption_width_in(cap.text_frame.text.strip())
            if w_in > 5.50 - 0.20:
                ch.top += int(0.20 * EMU); ch.height -= int(0.20 * EMU)
                if k == 0:
                    for s in usd: s.top += int(0.20 * EMU)
                note("INFO", f"commodities: {'oil' if k == 0 else 'gold'} caption wraps ({w_in:.2f} in): "
                             "chart below moved down 0.20 in")
    note("INFO", "commodities: right column set to the fixed layout (charts x 7.83, 5.07 x 3.10 in)")


_BOLD_W = None


def caption_width_in(text, pt=12):
    """Width of a caption line in inches, Arial Nova Cond Bold (font_widths.json 'bold', 1/1000 em)."""
    global _BOLD_W
    if _BOLD_W is None:
        p = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "data", "font_widths.json")
        _BOLD_W = json.load(open(p, encoding="utf-8"))["bold"]
    return sum(_BOLD_W.get(str(ord(c)), 600.0) for c in text) / 1000 * pt / 72


def pic_at(slide, x, y, tol=0.6):
    c = [p for p in pictures(slide) if abs(p.left / EMU - x) < tol and abs(p.top / EMU - y) < tol]
    return c[0] if c else None


def md5(blob): return hashlib.md5(blob).hexdigest()


# ---------------------------------------------------------------- slide copy and delete
def delete_slide(prs, slide):
    sldIdLst = prs.slides._sldIdLst
    for sid in list(sldIdLst):
        if prs.part.related_part(sid.rId) is slide.part:
            prs.part.drop_rel(sid.rId); sldIdLst.remove(sid); return


def copy_slide_in(prs, src_slide, before_slide):
    """Copy a slide (shapes, pictures, table) from another deck; place it before before_slide."""
    layout = next((l for l in prs.slide_layouts if l.name == src_slide.slide_layout.name), before_slide.slide_layout)
    new = prs.slides.add_slide(layout)
    for ph in list(new.placeholders): ph._element.getparent().remove(ph._element)
    tree = new.shapes._spTree
    for el in src_slide.shapes._spTree.iterchildren():
        if el.tag.endswith("}nvGrpSpPr") or el.tag.endswith("}grpSpPr"): continue
        el2 = copy.deepcopy(el)
        for node in el2.iter():
            for attr in (R_NS + "embed", R_NS + "link"):
                rid = node.get(attr)
                if rid and rid in src_slide.part.rels:
                    rel = src_slide.part.rels[rid]
                    if rel.is_external:
                        node.set(attr, new.part.relate_to(rel.target_ref, rel.reltype, is_external=True))
                    else:
                        tgt = rel.target_part
                        if "image" in rel.reltype:
                            import io
                            _, nrid = new.part.get_or_add_image_part(io.BytesIO(tgt.blob))
                            node.set(attr, nrid)
                        else:
                            node.set(attr, new.part.relate_to(tgt, rel.reltype))
        tree.append(el2)
    # move into place
    sldIdLst = prs.slides._sldIdLst
    ids = list(sldIdLst)
    new_id = ids[-1]
    target_idx = next(i for i, sid in enumerate(ids) if prs.part.related_part(sid.rId) is before_slide.part)
    sldIdLst.remove(new_id); sldIdLst.insert(target_idx, new_id)
    return new


# ---------------------------------------------------------------- data blocks
def block(ws, top, left_col, right_col, n):
    out = {}
    for r in range(top, top + n):
        label = ws.cell(r, left_col).value
        if label: out[str(label).strip()] = [ws.cell(r, c).value for c in range(left_col + 1, right_col + 1)]
    return out


# ---------------------------------------------------------------- results slide
def fill_results(slide, res):
    tbl_shape = next((s for s in flat(slide.shapes) if getattr(s, "has_table", False) and s.has_table), None)
    if tbl_shape is None: note("FLAG", "results slide: no table"); return
    tbl = tbl_shape.table
    rows = list(tbl.rows)
    src_row = next((i for i, r in enumerate(rows) if r.cells[0].text.strip().lower().startswith("source")), len(rows))
    hdr = 2
    data_rows = rows[hdr:src_row]
    want = res.get("rows", [])
    tbl_el = tbl._tbl
    proto = copy.deepcopy(data_rows[0]._tr) if data_rows else None
    for r in data_rows[len(want):]: tbl_el.remove(r._tr)
    for _ in range(len(want) - len(data_rows)):
        new_tr = copy.deepcopy(proto)
        anchor = list(tbl.rows)[hdr + len(data_rows) - 1]._tr if data_rows else rows[hdr - 1]._tr
        anchor.addnext(new_tr); data_rows = list(tbl.rows)[hdr:hdr + len(data_rows) + 1]
    rows = list(tbl.rows)
    for r, vals in zip(rows[hdr:hdr + len(want)], want):
        for c, v in zip(r.cells, vals):
            if isinstance(v, (int, float)): v = fnum(v)
            set_cell(c, "-" if v is None else str(v))
    if res.get("quarter"):
        set_cell(rows[1].cells[3], res["quarter"])
        if res.get("prior"): set_cell(rows[1].cells[4], res["prior"])
    # source row: "Source: Tadawul, Bloomberg", plus a "*FY26"-style note when a row is starred
    # (29 Oct 2026 team final: PURITY* with "*FY26" on a second line, 8 pt)
    rows = list(tbl.rows)
    src = next((r for r in rows if r.cells[0].text.strip().lower().startswith("source")), None)
    if src is not None:
        foot = res.get("footnote") or ""
        if any(str(v[0]).endswith("*") for v in want) and not foot:
            note("FLAG", "results: a starred row but no results.footnote (e.g. '*FY26')")
        set_lines(src.cells[0].text_frame, ["Source: Tadawul, Bloomberg"] + ([foot] if foot else []))
    # table frame = sum of its rows; commentary box sits just under the table (team final, 29 Oct 2026:
    # box top 0.07 in above the last row's bottom, about 0.48 in per paragraph)
    rows_h = sum(r.height for r in tbl.rows)
    tbl_shape.height = rows_h
    com = [s for s in text_shapes(slide) if s.top > tbl_shape.top + EMU and s.width > 10 * EMU]
    paras = res.get("commentary") or []
    if com:
        box = com[0]
        # with a "*FY26" note on the source row's second line, start below the table so the note shows
        box.top = int(tbl_shape.top + rows_h + (0.02 if res.get("footnote") else -0.07) * EMU)
        box.height = int(max(0.5, 0.48 * len(paras)) * EMU)
        set_lines(box.text_frame, paras if paras else [""])
        if box.top + box.height > 7.0 * EMU:
            note("FLAG", "results: commentary runs past 7.0 in: consider results.split_at")
    covered = [v[0] for v in want if len(v) > 9 and v[9] not in ("-", None, "")]
    if covered and not paras:
        note("FLAG", f"results: anbc-covered names {covered} have no commentary (analyst input)")
    note("INFO", f"results slide: {len(want)} rows, {len(res.get('commentary', []))} commentary paragraphs")


# ---------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--template", required=True); ap.add_argument("--data", required=True)
    ap.add_argument("--date", required=True); ap.add_argument("--out", required=True)
    ap.add_argument("--technicals"); ap.add_argument("--content"); ap.add_argument("--forecasts")
    ap.add_argument("--names"); ap.add_argument("--recap"); ap.add_argument("--futures")
    ap.add_argument("--results-template")
    ap.add_argument("--heatmap"); ap.add_argument("--team-technical")
    a = ap.parse_args()
    day = datetime.date.fromisoformat(a.date)
    prs = Presentation(a.template)
    tmpl_hash = {(slide_title(s)[:12], round(p.left / EMU, 1), round(p.top / EMU, 1)): md5(p.image.blob)
                 for s in prs.slides for p in pictures(s)}
    wb = openpyxl.load_workbook(a.data, read_only=True, data_only=True)
    ws = wb["PPT-ENG"]
    content = json.load(open(a.content, encoding="utf-8")) if a.content else {}
    names = {}
    if a.names and os.path.exists(a.names):
        names = json.load(open(a.names, encoding="utf-8"))
        names = names.get("en", names) if isinstance(names, dict) else {}

    # ---- data file date
    try:
        face = wb["Face"]
        dates = [c for r in face.iter_rows(max_row=30, max_col=15, values_only=True) for c in r
                 if isinstance(c, datetime.datetime)]
        if dates and day not in [d.date() for d in dates]:
            note("FLAG", f"data file Face dates {sorted(set(d.date() for d in dates))[:3]} do not include {day}")
    except KeyError:
        pass

    # ---- cover
    cover = prs.slides[0]
    dtxt = f"{day.day} {day.strftime('%B %Y')}"
    hit = False
    for s in text_shapes(cover):
        for p in s.text_frame.paragraphs:
            if re.search(r"\d{1,2} [A-Z][a-z]+ \d{4}", p.text) and p.runs:
                p.runs[0].text = re.sub(r"\d{1,2} [A-Z][a-z]+ \d{4}", dtxt, p.runs[0].text) \
                    if re.search(r"\d{1,2} [A-Z][a-z]+ \d{4}", p.runs[0].text) else dtxt
                for r in p.runs[1:]:
                    if re.search(r"\d", r.text): r.text = ""
                hit = True
    note("INFO" if hit else "FLAG", f"cover date {'set to ' + dtxt if hit else 'not found'}")

    # ---- global
    g = one_slide(prs, "GLOBAL MARKETS")
    intl = block(ws, 109, 4, 8, 10)
    vt = next(s for s in flat(g.shapes) if getattr(s, "has_table", False) and s.has_table)
    for r in list(vt.table.rows)[1:]:
        label = r.cells[0].text.strip()
        if label not in intl:
            note("FLAG", f"valuations: '{label}' not in PPT-ENG D109:D118; row left as it was"); continue
        close, chg, pe, ytd = intl[label]
        set_cell(r.cells[1], fnum(close)); set_cell(r.cells[2], *fpct(chg))
        set_cell(r.cells[3], fnum(pe)); set_cell(r.cells[4], *fpct(ytd))
    note("INFO", "global: comparative valuations refreshed from PPT-ENG D109:H118")
    for label in ("JAP (NIKKEI 225)", "CHN (SHCOMP)", "Hong Kong (HSI)"):
        if label in intl and intl[label][1] is not None and abs(float(intl[label][1])) < 0.0002:
            note("FLAG", f"{label} moved {float(intl[label][1]) * 100:.3f}%: check for a market holiday "
                         "before the recap tiles or headlines report a move")
    box = body_box(g, "HEADLINE UPDATES")
    if content.get("global_headlines") and box is not None:
        set_lines(box.text_frame, content["global_headlines"]); note("INFO", "global: headlines filled")
    else:
        note("FLAG", "global: HEADLINE UPDATES not supplied, still the template's text")
    ev = next((s for s in text_shapes(g) if re.match(r"\d{2}-[A-Z][a-z]{2}-\d{2,4}:", s.text_frame.text.strip())), None)
    if ev is None:
        ev_hdr = next((s for s in text_shapes(g) if s.text_frame.text.strip() == "EVENTS CALENDAR"), None)
        ev = min((s for s in text_shapes(g) if ev_hdr is not None and s.top > ev_hdr.top and s is not ev_hdr),
                 key=lambda s: s.top - ev_hdr.top, default=None)
    if content.get("events") is not None and ev is not None:
        if content["events"]:
            set_event_lines(ev.text_frame, content["events"])
        else:
            set_lines(ev.text_frame, [""])
        note("INFO", f"global: {len(content['events'])} events lines")
    elif ev is not None:
        note("FLAG", f"global: EVENTS CALENDAR not supplied, still '{ev.text_frame.text.strip()[:60]}'")
    rp = pic_at(g, 0.25, 1.26)
    fp = pic_at(g, 7.5, 3.85)
    lbl = next((s for s in text_shapes(g) if s.text_frame.text.strip() == "MARKET RECAP"), None)
    if lbl is not None: lbl.top = int(RECAP_LABEL_TOP * EMU)
    if a.recap and rp is not None: replace_picture(g, rp, a.recap, "global: market recap picture", FRAMES["recap"])
    else: note("FLAG", "global: market recap picture not supplied (capture Markets > Top Securities on bloomberg.com)")
    if a.futures and fp is not None: replace_picture(g, fp, a.futures, "global: stock futures picture", FRAMES["futures"])
    else: note("FLAG", "global: stock futures picture not supplied (capture bloomberg.com Markets > Futures, Today in the Markets, Popular Indices tab; not the Stocks page)")

    # ---- commodities
    cm = one_slide(prs, "COMMODITY MARKETS")
    com = block(ws, 122, 4, 8, 12)
    fc = {}
    if a.forecasts:
        fws = openpyxl.load_workbook(a.forecasts, data_only=True).worksheets[0]
        for row in fws.iter_rows(values_only=True):
            if row and row[0] and len(row) > 6: fc[str(row[0]).strip()] = (row[5], row[6])
    ct = next(s for s in flat(cm.shapes) if getattr(s, "has_table", False) and s.has_table)
    for r in list(ct.table.rows)[1:]:
        label = r.cells[0].text.strip()
        if not label or label.startswith("*"): continue
        if label not in com:
            note("FLAG", f"commodities: '{label}' not in PPT-ENG D122:D133"); continue
        close, chg, mtd, ytd = com[label]
        set_cell(r.cells[1], fnum(close)); set_cell(r.cells[2], *fpct(chg))
        set_cell(r.cells[3], *fpct(mtd)); set_cell(r.cells[4], *fpct(ytd))
        if label in fc:
            for ci, v in zip((5, 6), fc[label]):
                set_cell(r.cells[ci], "NA" if v is None or str(v).strip().upper() == "NA" else fnum(v))
    note("INFO", "commodities: table refreshed from PPT-ENG D122:H133" +
         (f"; forecasts from {os.path.basename(a.forecasts)}" if a.forecasts else
          "; forecast columns kept (weekly, Sundays: pass --forecasts)"))
    if day.weekday() == 6 and not a.forecasts:
        note("FLAG", "Sunday: the weekly commodity forecast file is due (DMU Commodity Data - <date>.xlsx)")
    chs = charts(cm)
    if len(chs) >= 2:
        refresh_chart(chs[0], wb, a.data, day, "commodities: oil chart")
        refresh_chart(chs[1], wb, a.data, day, "commodities: gold chart")
    else:
        note("FLAG", f"commodities: expected 2 charts, found {len(chs)}")
    box = body_box(cm, "HEADLINE UPDATES")
    if content.get("commodity_headlines") and box is not None:
        set_lines(box.text_frame, content["commodity_headlines"]); note("INFO", "commodities: headlines filled")
    else:
        note("FLAG", "commodities: HEADLINE UPDATES not supplied, still the template's text")
    caps = sorted([s for s in text_shapes(cm) if s.left > 7 * EMU and s.height < 0.5 * EMU
                   and len(s.text_frame.text.strip()) > 20 and "USD" not in s.text_frame.text], key=lambda s: s.top)
    if content.get("commodity_captions") and len(caps) == 2:
        for s, txt in zip(caps, content["commodity_captions"]): set_cell(s.text_frame, txt)
        note("INFO", "commodities: captions filled (top = oil, bottom = gold)")
    else:
        for s in caps: note("FLAG", f"commodities: caption not supplied, still '{s.text_frame.text.strip()[:70]}'")
    fix_commodity_layout(cm, chs, caps)

    # ---- saudi
    sa = one_slide(prs, "SAUDI MARKET")
    tabs = [s for s in flat(sa.shapes) if getattr(s, "has_table", False) and s.has_table]
    strip = {s.table.rows[0].cells[0].text.strip().upper(): s for s in tabs}
    top = [c.value for c in next(ws.iter_rows(min_row=6, max_row=6, min_col=3, max_col=20))]
    perf = [c.value for c in next(ws.iter_rows(min_row=8, max_row=8, min_col=3, max_col=20))]
    for key, off in (("TASI", 0), ("NOMU", 9)):
        if key not in strip: note("FLAG", f"saudi: {key} strip not found"); continue
        t = strip[key].table
        set_cell(t.rows[2].cells[0], fnum(top[off], 2)); set_cell(t.rows[2].cells[3], fnum(top[off + 3]))
        set_cell(t.rows[2].cells[6], f"{fpct(top[off + 6])[0]} YTD")
        for i in range(9): set_cell(t.rows[4].cells[i], *fpct(perf[off + i]))
    note("INFO", "saudi: TASI / NOMU strips refreshed from PPT-ENG C6:T8")
    tps = content.get("target_prices", {})
    for key, col in (("LEADERS", 4), ("LAGGARDS", 11)):
        sh = next((s for k, s in strip.items() if k.startswith(key)), None)
        if sh is None: note("FLAG", f"saudi: {key} table not found"); continue
        t = sh.table
        src = block(ws, 85, col, col + 4, 5)
        for r, (full, vals) in zip(list(t.rows)[2:], src.items()):
            short = names.get(full, full)
            if short == full and full not in names.values():
                note("FLAG", f"{key}: no short name for '{full}', full name used")
            close, chg, vol, val = vals
            set_cell(r.cells[0], short); set_cell(r.cells[1], fnum(close)); set_cell(r.cells[2], *fpct(chg))
            set_cell(r.cells[3], fnum(vol)); set_cell(r.cells[4], fnum(val))
            tp = tps.get(full, tps.get(short))
            set_cell(r.cells[5], fnum(tp) if tp not in (None, "") else "")
    note("INFO", "saudi: leaders / laggards refreshed from PPT-ENG D85:O89; Consensus TP left blank for the team")
    pe = charts(sa)
    if pe: refresh_chart(pe[0], wb, a.data, day, "saudi: P/E chart")
    else: note("FLAG", "saudi: P/E chart not found")
    hp = max((p for p in pictures(sa) if p.width > 5 * EMU and p.top < 4.5 * EMU),
             key=lambda p: p.width * p.height, default=None)
    if a.heatmap and hp is not None:
        fr = (hp.left / EMU, hp.top / EMU, hp.width / EMU, hp.height / EMU)
        replace_picture(sa, hp, a.heatmap, "saudi: heatmap picture", fr)
    elif a.heatmap:
        note("FLAG", "saudi: heatmap picture frame not found; heatmap not placed")
    else:
        note("FLAG", "saudi: heatmap not supplied (Saad: SASEIDX Index, IMAP, previous session)")
    box = body_box(sa, "HEADLINE UPDATES")
    if content.get("saudi_headlines") and box is not None:
        set_lines(box.text_frame, content["saudi_headlines"]); note("INFO", "saudi: headlines filled")
    else:
        note("FLAG", "saudi: HEADLINE UPDATES not supplied, still the template's text")

    # ---- results slides
    tech = one_slide(prs, "TASI TECHNICAL")
    res = content.get("results")
    rs = find_slides(prs, "RESULT ANNOUNCEMENTS")
    if res and res.get("rows"):
        if not rs:
            src = None
            if a.results_template:
                src = next(iter(find_slides(Presentation(a.results_template), "RESULT ANNOUNCEMENTS")), None)
            if src is None:
                note("FLAG", "results supplied but the template has no Results slide: pass --results-template")
            else:
                rs = [copy_slide_in(prs, src, tech)]
                note("INFO", "results slide copied in from the results template")
        if rs:
            for extra in rs[1:]: delete_slide(prs, extra)
            split = res.get("split_at")
            if split and 0 < split < len(res["rows"]):
                second = copy_slide_in(prs, rs[0], tech)
                fill_results(rs[0], dict(res, rows=res["rows"][:split]))
                fill_results(second, dict(res, rows=res["rows"][split:], commentary=[]))
                for sh in text_shapes(second):
                    if len(sh.text_frame.text) > 100 or (sh.text_frame.text.strip() == "" and sh.height > EMU):
                        sh._element.getparent().remove(sh._element)
                note("INFO", f"results split over two slides at row {split} (table continues, no commentary on slide 2)")
            else:
                fill_results(rs[0], res)
                if len(res["rows"]) > 12 and res.get("commentary"):
                    note("FLAG", f"{len(res['rows'])} result rows plus commentary: in Aug 2026 the team split such "
                                 "days over two slides (set results.split_at)")
    elif rs:
        for s in rs: delete_slide(prs, s)
        note("INFO", f"no results today: {len(rs)} Results slide(s) from the template removed")
    for s in find_slides(prs, "SUKUK"):
        note("FLAG", "Sukuk Issuances slide in the template: the team decides whether it stays; check it")

    # ---- technicals
    team_tech = None
    if a.team_technical:
        team_tech = next(iter(find_slides(Presentation(a.team_technical), "TASI TECHNICAL")), None)
        if team_tech is None:
            note("FLAG", f"technical: no TASI TECHNICAL slide in {os.path.basename(a.team_technical)}")
    if team_tech is not None:
        copy_slide_in(prs, team_tech, tech); delete_slide(prs, tech)
        note("INFO", f"technical: slide copied whole from the team's {os.path.basename(a.team_technical)} "
                     f"({len(team_tech.shapes)} shapes, pictures re-linked)")
        if a.technicals: note("INFO", "technical: --technicals ignored (the team's slide is used whole)")
    elif a.technicals:
        tp = Presentation(a.technicals).slides[0]
        src_pic = max((s for s in tp.shapes if s.shape_type == 13), key=lambda s: s.width * s.height)
        dst_pic = max(pictures(tech), key=lambda s: s.width * s.height)
        import io
        _, rid = tech.part.get_or_add_image_part(io.BytesIO(src_pic.image.blob))
        dst_pic._element.blipFill.find(A + "blip").set(R_NS + "embed", rid)
        t5 = next((s for s in tp.shapes if s.has_text_frame and s.name == "Text 5"), None)
        dst_title = next((s for s in text_shapes(tech) if 0.8 * EMU < s.top < 1.7 * EMU and s.text_frame.text.strip()
                          and s.text_frame.text.strip().upper() != "MARKET OUTLOOK"), None)
        if t5 is not None and dst_title is not None:
            set_cell(dst_title.text_frame, t5.text_frame.text.strip())
        brief = next((s for s in tp.shapes if s.has_text_frame and s.text_frame.text.strip().startswith("Market Brief")), None)
        out_box = max((s for s in text_shapes(tech) if len(s.text_frame.text) > 200),
                      key=lambda s: len(s.text_frame.text), default=None)
        if brief is not None and out_box is not None:
            paras = [p for p in brief.text_frame.paragraphs if p.text.strip()]
            heads = [i for i, p in enumerate(paras) if p.text.strip().lower() in ("market brief", "outlook")]
            out_i = next((i for i in heads if paras[i].text.strip().lower() == "outlook"), len(paras))
            brief_ps = [p for i, p in enumerate(paras) if i not in heads and i < out_i]
            trend_ps = [p for i, p in enumerate(paras) if i not in heads and i > out_i]

            def runs_of(ps):
                """Join the paragraphs; bold the index levels (9,999 / 99,999 and ranges), as the team does."""
                text = " ".join(re.sub(r"\s+", " ", p.text).strip() for p in ps)
                out, pos = [], 0
                for m in LEVEL.finditer(text):
                    if m.start() > pos: out.append((text[pos:m.start()], False))
                    out.append((m.group(0), True)); pos = m.end()
                if pos < len(text): out.append((text[pos:], False))
                return out
            rich = [runs_of(brief_ps)] + ([runs_of(trend_ps)] if trend_ps else [])
            set_rich_paragraphs(out_box.text_frame, rich)
            nb = sum(1 for ps in rich for t, b in ps if b)
            note("INFO", f"technical: chart, title and outlook from the Technicals deck (brief joined into one "
                         f"paragraph, trend line second, {nb} bold runs kept)")
        note("FLAG", "technical: the three right-hand images come from the technical team: confirm they are today's")
    else:
        note("FLAG", "technical: no Technicals deck given, slide still the template's")

    prs.save(a.out)

    # ---- pictures still identical to the template
    new = Presentation(a.out)
    for s in new.slides:
        for p in pictures(s):
            k = (slide_title(s)[:12], round(p.left / EMU, 1), round(p.top / EMU, 1))
            if not k[0] or p.width < 1.5 * EMU: continue
            if tmpl_hash.get(k) == md5(p.image.blob):
                note("FLAG", f"{k[0].title()} picture at ({k[1]}, {k[2]}) in is unchanged from the template")
    rep = os.path.splitext(a.out)[0] + "_checks.md"
    with open(rep, "w", encoding="utf-8") as f:
        f.write(f"# DMU build checks - {a.date}\n\n" + "\n".join(f"- {c}" for c in checks) + "\n")
    print("\n".join(checks)); print("wrote", a.out, "and", rep)


if __name__ == "__main__":
    main()
