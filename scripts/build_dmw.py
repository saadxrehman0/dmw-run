#!/usr/bin/env python3
"""Build the ANB Daily Market Watch deck on Template v2 (EN 5.14 / AR 5.26, CEO approved 25 Nov 2026).

One engine for both languages (--lang en|ar; build_dmw_ar.py is the AR entry point). The v2
template is all native tables with fixed names (s1.title, s1.cards, s1.index, s1.charts, s1.news,
s2.flow, s3.tasi, s3.markets, s3.global, ...) and the AR tables are exact column mirrors of the EN
ones, so each fill is written once in EN logical columns and mirrored for AR.

What the build fills (numbers, names, calendars, charts; commentary is fill_commentary.py's job):
  slide 1  date, the two headline cards, the Index Performance row, the two charts
  slide 2  Event and Economic calendars inside s2.flow (rows cloned or removed); the corporate news
           row above them is left for the commentary fill
  slide 3  every s3 table from PPT-ENG / PPT-AR; company names long or short by measured width
           (Mapping.xlsx, v2 rule of 25 Nov 2026); the sector YTD chart
  4 and 5  appended from the Technicals deck when given

Conflicts are reported, never resolved. Writes <out>.pptx and <out>_checks.md.
"""
import argparse, copy, json, re, sys
from datetime import datetime
from pathlib import Path

import openpyxl
from pptx import Presentation
from lxml import etree
from pptx.oxml.ns import qn

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import dmw_common as C                     # noqa: E402
from dmw_common import A, R, NS, RED, GREEN, rnd, num, pct, log, checks   # noqa: E402

E = 914400
PAGE_BOTTOM = C.PAGE_BOTTOM                  # 10.60 in: nothing may end below it (v2 rule)
SECTION_GAP = C.SECTION_GAP
clean_hidden_titles = C.clean_hidden_titles  # still imported by fill_commentary.py
REPO = HERE.parent
TEMPLATES = {"en": REPO / "templates" / "DMW Template v2 - EN.pptx",
             "ar": REPO / "templates" / "DMW Template v2 - AR.pptx"}
FONT_DIRS = [REPO / "fonts", HERE / "fonts", Path.home() / ".fonts",
             Path("C:/Windows/Fonts"), Path.home() / "AppData/Local/Microsoft/Windows/Fonts"]
INDEX_MAX_CHARS = 6        # "446.2%" fits an Index Performance cell; 7 characters do not (v2 rule)
NAME_SAFETY_IN = 0.01      # slack kept inside a cell before a name counts as wrapping


# ---------------------------------------------------------------- template access
def shape(slide, name):
    for sh in slide.shapes:
        if sh.name == name:
            return sh
    raise SystemExit(f"Template v2: shape '{name}' not found. Is this the v2 template?")


class Mirror:
    """EN logical column -> physical column. AR tables are left-right mirrors of the EN ones."""
    def __init__(self, lang):
        self.ar = lang == "ar"

    def col(self, table, c, span=1):
        return len(table.columns) - c - span if self.ar else c


def strip_highlight(cell):
    for rpr in cell._tc.iter(f"{{{A}}}rPr", f"{{{A}}}endParaRPr"):
        for h in rpr.findall(f"{{{A}}}highlight"):
            rpr.remove(h)


def put(cell, text, color="keep"):
    """Write a cell keeping its first run's format; the template's placeholder highlight goes."""
    C.set_cell(cell, text, color)
    strip_highlight(cell)


def color_of(v, text):
    """House colours: negative red (even -0.0%), positive green unless it shows as 0.0%."""
    if v is None or isinstance(v, str):
        return None
    return RED if v < 0 else (GREEN if re.search(r"[1-9]", text) else None)


def put_value(cell, kind, v):
    if kind == "n1":
        put(cell, num(v, 1))
    elif kind == "n2":
        put(cell, num(v, 2))
    elif kind in ("pct", "bps"):
        txt = pct(v) if kind == "pct" else C.bps(v)
        put(cell, txt, color_of(v, txt))
    else:
        raise ValueError(kind)


# ---------------------------------------------------------------- text width (real font, shaped)
_FONTS = {}


def font_path(fname):
    for d in FONT_DIRS:
        p = d / fname
        if p.exists():
            return p
    return None


def text_width_in(text, pt, fname="anbApercu-Regular.otf", rtl=False):
    """Advance width of a shaped line in inches (HarfBuzz, the deck's own font), or None if the
    font or HarfBuzz is missing. anb Apercu carries the Arabic glyphs too (v2: one font)."""
    try:
        import uharfbuzz as hb
    except ImportError:
        return None
    if fname not in _FONTS:
        p = font_path(fname)
        if p is None:
            _FONTS[fname] = None
        else:
            face = hb.Face(hb.Blob.from_file_path(str(p)))
            _FONTS[fname] = (hb.Font(face), face.upem)
    f = _FONTS[fname]
    if f is None:
        return None
    font, upem = f
    buf = hb.Buffer()
    buf.add_str(text)
    buf.guess_segment_properties()
    if rtl:
        buf.direction = "rtl"
    hb.shape(font, buf, {"kern": True, "liga": True})
    return sum(p.x_advance for p in buf.glyph_positions) / upem * pt / 72


def cell_room_in(table, r, c, span=1):
    tc = table.cell(r, c)._tc
    w = sum(table.columns[c + k].width for k in range(span))
    pr = tc.find(f"{{{A}}}tcPr")
    ml = int(pr.get("marL", 91440)) if pr is not None else 91440
    mr = int(pr.get("marR", 91440)) if pr is not None else 91440
    return (w - ml - mr) / E - NAME_SAFETY_IN


def run_pt(cell, default=7.0):
    runs = cell.text_frame.paragraphs[0].runs
    return runs[0].font.size.pt if runs and runs[0].font.size else default


def wraps(text, cell, room, lang):
    w = text_width_in(text, run_pt(cell), rtl=(lang == "ar"))
    return None if w is None else w > room


def norm_en(s):
    s = re.sub(r"[^a-z0-9 ]", " ", str(s).lower().replace("&", " and "))
    s = re.sub(r"\b(co|company|ltd|the|cos)\b", " ", s)
    return re.sub(r"\s+", " ", s).strip()


# ---------------------------------------------------------------- company names (Mapping.xlsx)
class Names:
    """Long or short company name (v2 rule, 25 Nov 2026): the long name when it fits its cell on one
    line, else the short name, measured in the deck font for each language. Slide 3 rows are matched
    to Mapping.xlsx by Tadawul code, from the data file's own lookup in 'TASI All Com'."""

    COLS = (("long_en", "Long Name - English"), ("short_en", "Short Name - English"),
            ("long_ar", "Long Name - Arabic"), ("short_ar", "Short Name - Arabic"))

    def __init__(self, mapping_xlsx, data_xlsx, overrides_en=None, overrides_ar=None):
        self.by_code, self.by_en = {}, {}
        ws = openpyxl.load_workbook(mapping_xlsx, read_only=True)["mapping"]
        hdr = [str(c).strip() if c is not None else "" for c in next(ws.iter_rows(max_row=1, values_only=True))]
        ix = {h: i for i, h in enumerate(hdr)}
        miss = [h for _, h in self.COLS if h not in ix] + ([] if "Code" in ix else ["Code"])
        if miss:
            raise SystemExit(f"Mapping.xlsx: headers not found {miss}.")
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[ix["Code"]] is None:
                continue
            rec = {k: (str(row[ix[h]]).strip() if row[ix[h]] else "") for k, h in self.COLS}
            self.by_code[str(row[ix["Code"]]).strip()] = rec
            for k in ("long_en", "short_en"):
                if rec[k]:
                    self.by_en.setdefault(norm_en(rec[k]), rec)
        self.code_of = {}
        dwb = openpyxl.load_workbook(data_xlsx, read_only=True, data_only=True)
        if "TASI All Com" in dwb.sheetnames:
            sym2code, disp2sym = {}, {}
            for row in dwb["TASI All Com"].iter_rows(min_row=3, max_col=27, values_only=True):
                if row[0] is not None and row[1] is not None:            # A code, B symbol
                    sym2code[str(row[1]).strip()] = str(row[0]).strip()
                if len(row) > 24 and row[22] is not None and row[24] is not None:   # W symbol, Y name
                    disp2sym[str(row[24]).strip()] = str(row[22]).strip()
            self.code_of = {d: sym2code[s] for d, s in disp2sym.items() if s in sym2code}
        else:
            log("FLAG", "Data file has no 'TASI All Com' sheet: slide 3 names cannot be matched to Mapping.xlsx by code.")
        load = lambda p: json.loads(Path(p).read_text(encoding="utf-8")) if p and Path(p).exists() else {}
        self.ov_en, self.ov_ar = load(overrides_en), load(overrides_ar)

    def choose(self, lang, rec, fallback, cell, room, where, short_only=False):
        long_, short = (rec["long_en"], rec["short_en"]) if lang == "en" else (rec["long_ar"], rec["short_ar"])
        cands = [x for x in ((short,) if short_only and short else (long_, short)) if x] or [fallback]
        for name in cands:
            w = wraps(name, cell, room, lang)
            if w is None:
                log("FLAG", f"{where}: font or HarfBuzz missing, name width not measured; long name used.")
                return cands[0]
            if not w:
                return name
        log("FLAG", f"{where}: short name '{cands[-1]}' still wraps ({room:.2f} in). Needs a shorter name "
                    "(this is the one name case that pings).")
        return cands[-1]


# ---------------------------------------------------------------- dates
DAYS_AR = ["الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت", "الأحد"]
MONTHS_AR = ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر"]
IND = str.maketrans("0123456789", "٠١٢٣٤٥٦٧٨٩")


def ar_long_date(d):
    """House form since 15 Oct 2026, Arabic-Indic digits (v2 rule): 'الأحد، ١٨ أكتوبر ٢٠٢٦'."""
    return f"{DAYS_AR[d.weekday()]}، {str(d.day).translate(IND)} {MONTHS_AR[d.month-1]} {str(d.year).translate(IND)}"


def ar_cal_date(d):
    return f"{str(d.day).translate(IND)}- {MONTHS_AR[d.month-1]}-{str(d.year).translate(IND)}"


def en_long_date(d):
    return f"{d:%A, %B} {d.day}, {d.year}"


# ---------------------------------------------------------------- slide 1
def fill_date(slide, lang, d):
    t = shape(slide, "s1.title").table
    if lang == "en":
        # The v2 master holds a datetime2 field. PowerPoint refreshes a field to the PC's date on
        # open or export, so a deck built for another day (replays, a run past midnight) would show
        # the wrong date. Turn the field into a plain run with the same formatting and the report date.
        cell = t.cell(1, 3)
        for fld in cell._tc.findall(".//a:fld", NS):
            r = etree.SubElement(fld.getparent(), qn("a:r"))
            fld.addprevious(r)
            rpr = fld.find("a:rPr", NS)
            if rpr is not None:
                r.append(rpr)
            etree.SubElement(r, qn("a:t"))
            fld.getparent().remove(fld)
        ts = cell._tc.findall(".//a:t", NS)
        ts[0].text = en_long_date(d)
        for x in ts[1:]:
            x.text = ""
        log("OK", f"Slide 1 date '{en_long_date(d)}' (static text, not a date field).")
    else:
        put(t.cell(0, 0), ar_long_date(d))
        log("OK", f"Slide 1 date '{ar_long_date(d)}'.")


def digit_runs(cell):
    return [r for r in cell.text_frame.paragraphs[0].runs if re.search(r"\d", r.text)]


def fill_cards(slide, lang, g):
    """Each card's close and DoD share one cell (two runs); value traded and YTD are the digit runs
    of their own cells."""
    t = shape(slide, "s1.cards").table
    M = Mirror(lang)
    cards = {"TASI": dict(c=1, close=g["C6"], vt=g["F6"], ytd=g["E8"], dod=g["C8"], unit="bn"),
             "NOMU": dict(c=5, close=g["L6"], vt=g["O6"], ytd=g["N8"], dod=g["L8"], unit="mn")}
    for key, d in cards.items():
        c = M.col(t, d["c"])
        cd, vt, ytd = digit_runs(t.cell(2, c)), digit_runs(t.cell(3, c)), digit_runs(t.cell(4, c))
        if len(cd) < 2 or not vt or not ytd:
            raise SystemExit(f"s1.cards {key}: close/DoD, value or YTD runs not found.")
        cd[0].text = f"{rnd(d['close'], 2):,.2f}"
        for run, v in ((cd[-1], d["dod"]), (ytd[-1], d["ytd"])):
            run.text = pct(v)
            C.set_run_color(run._r.get_or_add_rPr(), color_of(v, run.text))
        if lang == "en":
            vt[-1].text = f"SAR {num(d['vt'], 1)} {d['unit']}"
        else:
            vt[-1].text = f"{num(d['vt'], 1)} {'مليار' if d['unit'] == 'bn' else 'مليون'} ريال"
        for r in (2, 3, 4):
            strip_highlight(t.cell(r, c))
    log("OK", "Slide 1 cards (TASI, NOMU) filled from PPT-ENG.")


def fill_index(slide, lang, g, g_ar=None):
    t = shape(slide, "s1.index").table
    M = Mirror(lang)
    vals = [g[f"{c}8"] for c in "CDEFGHIJKLMNOPQRST"]
    if g_ar is not None:
        ar_row = [g_ar.get(f"{c}7") for c in "GHJKLMNOPQRTUVWXY"]
        try:
            same = [f"{x:.10f}" for x in ar_row] == [f"{x:.10f}" for x in list(reversed(vals))[:17]]
        except (TypeError, ValueError):
            same = False
        if not same:
            log("CONFLICT", "Index Performance: PPT-AR row 7 differs from PPT-ENG row 8 (reversed). EN values used.")
    for i, v in enumerate(vals):
        txt = pct(v)
        put(t.cell(5, M.col(t, i)), txt, color_of(v, txt))
        if len(txt) > INDEX_MAX_CHARS:
            log("FLAG", f"Index Performance: '{txt}' is {len(txt)} characters; the cell holds {INDEX_MAX_CHARS}.")
    log("OK", "Slide 1 Index Performance row filled (18 values).")


def swap_picture(slide, pic, chart, label):
    """A Graphs-deck chart into a v2 chart frame: fitted inside the frame with its aspect kept,
    centred across and top-aligned, so the frame never grows and nothing below moves."""
    rid = C.add_image_part(slide, chart["blob"], chart["ext"])
    blip = pic._element.find(".//a:blip", NS)
    blip.set(f"{{{R}}}embed", rid)
    for ext in blip.findall(f"{{{A}}}extLst"):
        blip.remove(ext)
    src = pic._element.find(".//a:srcRect", NS)
    if src is not None:
        src.getparent().remove(src)
    fx, fy, fw, fh = pic.left, pic.top, pic.width, pic.height
    ar = chart["width"] / chart["height"]
    w, h = (fw, int(fw / ar)) if fw / fh < ar else (int(fh * ar), fh)
    pic.left, pic.top, pic.width, pic.height = fx + (fw - w) // 2, fy, w, h
    log("OK", f"{label}: {w/E:.2f} x {h/E:.2f} in inside the {fw/E:.2f} x {fh/E:.2f} in frame.")


def fill_charts(slide, names, charts, label):
    pics = sorted((shape(slide, n) for n in names), key=lambda s: s.left)
    if len(charts) != len(pics):
        log("FLAG", f"{label}: Graphs deck has {len(charts)} chart(s) for {len(pics)} frame(s). Charts not updated.")
        return
    for pic, ch in zip(pics, sorted(charts, key=lambda c: c["left"])):
        swap_picture(slide, pic, ch, f"{label} ({pic.name})")


# ---------------------------------------------------------------- slide 2 calendars (inside s2.flow)
HEADS = {"en": ("Event Calendar", "Economic Calendar"), "ar": ("تقويم الأحداث", "التقويم الاقتصادي")}
SRC_WORD = {"en": "Source", "ar": "المصدر"}
SOURCES = {"en": ("Source: Argaam, Bloomberg", "Source: Argaam, Bloomberg, Investing.com"),
           "ar": ("المصدر: أرقام، بلومبرغ", "المصدر: أرقام، بلومبرغ، وإنفستينغ.كوم")}


def flow_rows(t, lang):
    """{'ev'|'ec': (header row, first data row, source row)} in s2.flow."""
    text = lambda r: "".join(t.cell(r, c).text for c in range(len(t.columns))).strip()
    out = {}
    for key, title in zip(("ev", "ec"), HEADS[lang]):
        h = next((r for r in range(len(t.rows)) if text(r) == title), None)
        if h is None:
            raise SystemExit(f"s2.flow: '{title}' row not found.")
        s = next((r for r in range(h + 2, len(t.rows)) if text(r).startswith(SRC_WORD[lang])), None)
        if s is None:
            raise SystemExit(f"s2.flow: source row under '{title}' not found.")
        out[key] = (h, h + 2, s)
    return out


def set_rows(t, first, src, n):
    """Make rows first..src-1 exactly n rows by cloning the first data row (n=0 keeps none)."""
    tbl = t._tbl
    data = tbl.findall(f"{{{A}}}tr")[first:src]
    proto = data[0]
    for tr in data[1:]:
        tbl.remove(tr)
    anchor = proto
    for _ in range(max(n - 1, 0)):
        new = copy.deepcopy(proto)
        anchor.addnext(new)
        anchor = new
    if n == 0:
        tbl.remove(proto)


def flag_wrap(t, r, c, span, text, lang, label):
    room = cell_room_in(t, r, c, span)
    if wraps(text, t.cell(r, c), room, lang):
        log("FLAG", f"{label} '{text[:60]}' wraps to two lines ({room:.2f} in): its row grows.")


def calendar_rows(lang, events, econ, names, tr=None, ar_cal=(None, None)):
    """Display rows. EN from the calendar file; AR from its Arabic columns, glossary as fallback."""
    if lang == "en":
        ev = [(f"{d:%d-%b-%y}", comp, e, names.by_en.get(norm_en(comp))) for d, comp, e in events]
        return ev, [(f"{d:%d-%b-%y}", e, p, s) for d, e, p, s in econ]
    AR = dmw_ar()
    ev_ar, ec_ar = ar_cal[0] or [], ar_cal[1] or []
    if ev_ar and len(ev_ar) != len(events):
        log("FLAG", f"Event calendar: {len(ev_ar)} Arabic rows vs {len(events)} English rows in the file; Arabic columns not used.")
        ev_ar = []
    if ec_ar and len(ec_ar) != len(econ):
        log("FLAG", f"Economic calendar: {len(ec_ar)} Arabic rows vs {len(econ)} English rows in the file; Arabic columns not used.")
        ec_ar = []
    ev = []
    for i, (d, comp, event) in enumerate(events):
        e_ar, c_ar = ev_ar[i] if i < len(ev_ar) else ("", "")
        if not e_ar:
            e_ar, _ = tr.ev(event)
            if not e_ar:
                log("FLAG", f"Event calendar: no Arabic wording for '{event}'. Placeholder left.")
        if not c_ar:
            c_ar, _ = tr.comp(comp)
            if not c_ar:
                log("FLAG", f"Event calendar: no Arabic name for '{comp}'. Placeholder left.")
        rec = names.by_en.get(norm_en(comp))
        if rec and c_ar and c_ar.strip() not in (rec["long_ar"], rec["short_ar"]):
            log("FLAG", f"Event calendar: Arabic name '{c_ar}' for '{comp}' is not the Mapping long or short name; "
                        "calendar wording kept.")
            rec = None
        ev.append((ar_cal_date(d), c_ar or AR.PH.format(comp), e_ar or AR.PH.format(event), rec))
    ec = []
    for i, (d, e, p, s) in enumerate(econ):
        e_ar = ec_ar[i] if i < len(ec_ar) else ""
        if not e_ar:
            e_ar, _ = tr.eco(e)
            if not e_ar:
                log("FLAG", f"Economic calendar: no Arabic wording for '{e}'. Placeholder left.")
        if e == "US EIA Crude Oil Inventories" and e_ar and "EIA" not in e_ar and "معلومات الطاقة" not in e_ar:
            e_ar = "مخزونات النفط الخام الأمريكية (EIA)"      # client form, 25 Nov 2026
            log("OK", "Economic calendar: EIA row set to the client form with (EIA).")
        ec.append((ar_cal_date(d), e_ar or AR.PH.format(e), p, s))
    return ev, ec


def put_label(cell, text, lang):
    if lang == "ar":
        dmw_ar().set_ar_label(cell, text)     # Latin spans in their own en-US runs
        strip_highlight(cell)
    else:
        put(cell, text)


def fill_calendars(slide, lang, events, econ, names, tr=None, ar_cal=(None, None)):
    t = shape(slide, "s2.flow").table
    M = Mirror(lang)
    ev_rows, ec_rows = calendar_rows(lang, events, econ, names, tr, ar_cal)
    _, f, s = flow_rows(t, lang)["ec"]          # economic first: its rows sit below the event block
    set_rows(t, f, s, len(ec_rows))
    for i, (d, ev, prior, est) in enumerate(ec_rows):
        r = f + i
        put(t.cell(r, M.col(t, 0)), d)
        put_label(t.cell(r, M.col(t, 1, 2)), ev, lang)
        put(t.cell(r, M.col(t, 3)), prior)
        put(t.cell(r, M.col(t, 4)), est)
        flag_wrap(t, r, M.col(t, 1, 2), 2, ev, lang, "Economic calendar event")
    _, f, s = flow_rows(t, lang)["ev"]
    set_rows(t, f, s, len(ev_rows))
    for i, (d, comp, ev, rec) in enumerate(ev_rows):
        r = f + i
        cc = M.col(t, 1)
        room = cell_room_in(t, r, cc)
        if rec:
            name = names.choose(lang, rec, comp, t.cell(r, cc), room, f"Event calendar '{comp}'")
        else:
            name = comp
            if wraps(comp, t.cell(r, cc), room, lang):
                log("FLAG", f"Event calendar: company '{comp}' wraps and has no Mapping.xlsx match for a short name.")
        put(t.cell(r, M.col(t, 0)), d)
        put_label(t.cell(r, cc), name, lang)
        put_label(t.cell(r, M.col(t, 2, 3)), ev, lang)
        flag_wrap(t, r, M.col(t, 2, 3), 3, ev, lang, "Event calendar event")
    rows = flow_rows(t, lang)
    for key, src in zip(("ev", "ec"), SOURCES[lang]):
        put(t.cell(rows[key][2], 0), src)
    for label, n in (("Event", len(ev_rows)), ("Economic", len(ec_rows))):
        if n == 0:
            log("FLAG", f"{label} calendar: no rows; the table shows its header and source only. Confirm there are none today.")
    log("OK", f"Slide 2 calendars: {len(ev_rows)} events, {len(ec_rows)} economic releases.")


# ---------------------------------------------------------------- slide 3
# s3.tasi blocks: header row and side (L = EN columns 0-4, R = 6-10); data rows header+2 .. +6
TASI_BLOCKS = {"top_sec": (2, 0), "top_co": (2, 6), "lag_sec": (10, 0), "lag_co": (10, 6),
               "val_co": (18, 6), "vol_co": (26, 6)}
TITLES = {"en": {"top_sec": "Top Performers – Sectors", "top_co": "Top Performers – Companies",
                 "lag_sec": "Top Laggards – Sectors", "lag_co": "Top Laggards – Companies",
                 "val_co": "Value Leaders – Companies", "vol_co": "Volume Leaders – Companies",
                 "intl": "International Markets", "reg": "Regional Markets",
                 "comm": "Commodity Prices", "rates": "Interest Rates"},
          "ar": {"top_sec": "الأكثر ارتفاعاً – القطاعات", "top_co": "الأكثر ارتفاعاً – الشركات",
                 "lag_sec": "الأكثر انخفاضاً – القطاعات", "lag_co": "الأكثر انخفاضاً – الشركات",
                 "val_co": "الأكثر نشاطاً بالقيمة – الشركات", "vol_co": "الأكثر نشاطاً بالكمية – الشركات",
                 "intl": "الأسواق العالمية", "reg": "الأسواق الإقليمية",
                 "comm": "أسعار السلع الأساسية", "rates": "أسعار الفائدة"}}
# source blocks (columns, first row, last row). PPT-AR runs right to left: values ..., label last.
EN_SRC = {"top_sec": ("DEFGH", 77, 81), "lag_sec": ("KLMNO", 77, 81), "top_co": ("DEFGH", 85, 89),
          "lag_co": ("KLMNO", 85, 89), "val_co": ("DEFGH", 93, 97), "vol_co": ("KLMNO", 93, 97),
          "intl": ("DEFGH", 109, 118), "reg": ("JKLMN", 109, 118), "comm": ("DEFGH", 122, 133),
          "rates": ("JKLMN", 122, 125)}
AR_SRC = {"top_sec": ("RSTUV", 76, 80), "lag_sec": ("JKLMN", 76, 80), "top_co": ("RSTUV", 84, 88),
          "lag_co": ("JKLMN", 84, 88), "val_co": ("RSTUV", 92, 96), "vol_co": ("JKLMN", 92, 96),
          "intl": ("RSTUV", 102, 111), "reg": ("JKLMN", 102, 111), "comm": ("JKLMN", 115, 126),
          "rates": ("RSTUV", 115, 118)}
SEC, CO = ["n1", "pct", "n1", "pct"], ["n1", "pct", "n1", "n1"]
SPECS = {"top_sec": SEC, "lag_sec": SEC, "top_co": CO, "lag_co": CO, "val_co": CO, "vol_co": CO,
         "intl": SEC, "reg": SEC, "comm": ["n1", "pct", "pct", "pct"], "rates": ["n2", "bps", "bps", "bps"]}
SIDE_BLOCKS = {"intl": ("s3.markets", 0, 10), "reg": ("s3.markets", 13, 10),
               "comm": ("s3.global", 0, 12), "rates": ("s3.global", 15, 4)}


def src_rows(g, lang, key):
    cols, r0, r1 = (EN_SRC if lang == "en" else AR_SRC)[key]
    rows = [[g.get(f"{c}{r}") for c in cols] for r in range(r0, r1 + 1)]
    if lang == "ar":
        rows = [[row[-1]] + list(reversed(row[:-1])) for row in rows]
    return rows                 # [label, v1, v2, v3, v4] in EN logical order


def plain_label(lang, key, label, cell):
    """Sector, market, commodity and rate labels. EN: the data label; a standing label that differs
    from the template is a CONFLICT. AR: the data label with the team's fixed spellings, or the
    template's wording when it is the same characters reordered."""
    cur = cell.text.strip()
    if lang == "en":
        if key not in ("top_sec", "lag_sec") and cur and not cur.startswith("[") and cur != label:
            log("CONFLICT", f"{TITLES['en'][key]}: template row '{cur}' vs data row '{label}'.")
        return label
    label = dmw_ar().ar_label(label)
    same = lambda x: sorted(re.sub(r"[\s()]", "", x))
    if key not in ("top_sec", "lag_sec") and cur and not cur.startswith("[") and (same(cur) == same(label) or key == "comm"):
        return cur
    return label


def company_label(lang, key, i, label, code, names, t, r, c):
    room = cell_room_in(t, r, c)
    where = f"Slide 3 {TITLES['en'][key]} row {i+1}"
    rec = names.by_code.get(code) if code else None
    if rec:
        # slide 3 company rows: the Tadawul trading name (Mapping "Short Name"), as in Tadawul's
        # Detailed Daily Report, so slide 3 and the Results table read the same (Saad, 25 Nov 2026)
        return names.choose(lang, rec, label, t.cell(r, c), room, where, short_only=True)
    name = names.ov_en.get(label, label) if lang == "en" else names.ov_ar.get(dmw_ar().ar_label(label), dmw_ar().ar_label(label))
    log("FLAG", f"{where}: '{label}' not matched to Mapping.xlsx by code; the data file's name used.")
    if wraps(name, t.cell(r, c), room, lang):
        log("FLAG", f"{where}: '{name}' wraps ({room:.2f} in). Needs a short name.")
    return name


def fill_block(t, M, lang, key, hdr, base, rows, names=None, codes=None):
    title = t.cell(hdr, M.col(t, base, 5)).text.strip()
    if title != TITLES[lang][key]:
        raise SystemExit(f"Slide 3 row {hdr}: expected '{TITLES[lang][key]}', template has '{title}'.")
    for i, row in enumerate(rows):
        r = hdr + 2 + i
        lc = M.col(t, base)
        label = str(row[0]).strip() if row[0] is not None else ""
        if names is not None:
            label = company_label(lang, key, i, label, codes[i], names, t, r, lc)
        else:
            label = plain_label(lang, key, label, t.cell(r, lc))
        put_label(t.cell(r, lc), label, lang)
        for j, (kind, v) in enumerate(zip(SPECS[key], row[1:]), start=1):
            put_value(t.cell(r, M.col(t, base + j)), kind, v)


def fill_slide3(slide, lang, g, g_en, names, charts):
    M = Mirror(lang)
    t = shape(slide, "s3.tasi").table
    for key, (hdr, base) in TASI_BLOCKS.items():
        rows = src_rows(g, lang, key)
        codes = None
        if key.endswith("_co"):
            cols, r0, r1 = EN_SRC[key]
            codes = [names.code_of.get(str(g_en.get(f"{cols[0]}{r}")).strip()) for r in range(r0, r1 + 1)]
            if lang == "ar":            # the AR rows must be the same companies, in the same order
                for i, (a, e) in enumerate(zip(rows, src_rows(g_en, "en", key))):
                    if a[1] != e[1] or a[2] != e[2]:
                        log("CONFLICT", f"Slide 3 {TITLES['en'][key]} row {i+1}: PPT-AR {a[1]} / {a[2]} vs PPT-ENG {e[1]} / {e[2]}.")
                        codes[i] = None
        fill_block(t, M, lang, key, hdr, base, rows, names if codes is not None else None, codes)
    for key, (tname, hdr, n) in SIDE_BLOCKS.items():
        rows = src_rows(g, lang, key)
        fill_block(shape(slide, tname).table, M, lang, key, hdr, 0, rows)
    log("OK", f"Slide 3 tables filled from {'PPT-ENG' if lang == 'en' else 'PPT-AR'}.")
    fill_charts(slide, ["s3.chart1"], charts.get(3, []), "Slide 3 sector YTD chart")


# ---------------------------------------------------------------- Technicals (slides 4-5)
def append_technicals(prs, tech, lang, report_date):
    blank = next((l for l in prs.slide_layouts if l.name.lower() == "blank"), prs.slide_layouts[-1])
    new = []
    for src in list(tech.slides)[:2]:
        s = prs.slides.add_slide(blank)
        C.replace_with_technicals(s, src)
        new.append(s)
    log("OK", f"Slides 4-5 appended from the Technicals {lang.upper()} deck.")
    if lang == "ar":
        AR = dmw_ar()
        AR.tidy_technicals_ar(new)
        AR.ensure_s5_rule(new[1], log)
        td = [sh for sh in new[0].shapes if sh.has_text_frame and re.match(AR.DAY_RE, sh.text_frame.text.strip())]
        want, bare = ar_long_date(report_date), (lambda x: re.sub(r"[\s,،]+", " ", x).strip())
        if not td:
            log("FLAG", "Technicals AR: date not found on slide 4.")
        elif td[0].text_frame.text.strip() != want:
            if bare(td[0].text_frame.text) == bare(want):
                runs = td[0].text_frame.paragraphs[0].runs
                runs[0].text = want
                for x in runs[1:]:
                    x._r.getparent().remove(x._r)
                log("OK", f"Technicals AR date rewritten to the house form '{want}'.")
            else:
                log("CONFLICT", f"Technicals AR date '{td[0].text_frame.text.strip()}' vs report date '{want}'.")


# ---------------------------------------------------------------- main
_AR = None


def dmw_ar():
    global _AR
    if _AR is None:
        import dmw_common_ar
        _AR = dmw_common_ar
    return _AR


def s2_bottom_in(slide):
    """s2.flow bottom from its row heights, before any text growth (the commentary fill and
    check_final measure the real one); catches calendars that alone overflow."""
    sh = shape(slide, "s2.flow")
    return (sh.top + sum(r.height for r in sh.table.rows)) / E


def build(a):
    lang = a.lang
    rd = datetime.strptime(a.date, "%Y-%m-%d").date()
    g_en, face = C.load_data(a.data)
    g = g_en
    if lang == "ar":
        wb = openpyxl.load_workbook(a.data, data_only=True, read_only=True)
        g = {c.coordinate: c.value for row in wb["PPT-AR"].iter_rows(min_row=1, max_row=140, max_col=30)
             for c in row if c.value is not None}
    events, econ = C.load_calendar(a.calendar, rd)
    charts, graphs_date = C.ole_previews(a.graphs)
    tech = Presentation(a.technicals) if a.technicals else None
    names = Names(a.mapping, a.data, a.names, a.ar_names)
    prs = Presentation(a.template or TEMPLATES[lang])
    if len(prs.slides) != 3:
        raise SystemExit(f"Template v2 has 3 slides; this one has {len(prs.slides)}.")
    if tech is not None and len(tech.slides) != 2:
        raise SystemExit("Expected a 2-slide Technicals deck.")
    if lang == "en":
        C.cross_checks(g_en, face, rd, graphs_date, tech, a.writeup, events, econ)
        if a.chart_state:
            earlier = [d.date() for d in face if d.date() < rd]
            C.chart_staleness(charts, g_en, rd, max(earlier) if earlier else rd, a.chart_state)
    elif rd not in [d.date() for d in face]:
        log("CONFLICT", "Data file Face sheet does not carry the report date.")

    s1, s2, s3 = prs.slides[0], prs.slides[1], prs.slides[2]
    fill_date(s1, lang, rd)
    fill_cards(s1, lang, g_en)
    fill_index(s1, lang, g_en, g if lang == "ar" else None)
    fill_charts(s1, ["s1.chart1", "s1.chart2"], charts.get(1, []), "Slide 1 charts")
    tr = ar_cal = None
    if lang == "ar":
        tr = dmw_ar().Translator(a.glossary, a.mapping)
        ar_cal = dmw_ar().load_calendar_ar(a.calendar)
    fill_calendars(s2, lang, events, econ, names, tr, ar_cal or (None, None))
    fill_slide3(s3, lang, g, g_en, names, charts)
    for sl in (prs.slides[1], prs.slides[2]):
        drop_header_rule(sl)
    if tech is not None:
        append_technicals(prs, tech, lang, rd)
    else:
        log("FLAG", "No Technicals deck: slides 1-3 only. Patch slides 4-5 in when it lands.")
    b = s2_bottom_in(s2)
    log("OK" if b <= PAGE_BOTTOM else "FLAG",
        f"Slide 2 flow ends at {b:.2f} in on row heights, before news text (limit {PAGE_BOTTOM:.2f} in).")
    out = Path(a.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    prs.save(out)
    lines = [f"# DMW {lang.upper()} build checks (Template v2) - {rd:%d %b %Y}", ""]
    for lvl in ("STALE", "CONFLICT", "FLAG", "OK"):
        items = list(dict.fromkeys(m for l, m in checks if l == lvl))
        if items:
            lines += [f"## {lvl} ({len(items)})", ""] + [f"- {m}" for m in items] + [""]
    rep = out.with_name(out.stem + "_checks.md")
    rep.write_text("\n".join(lines), encoding="utf-8")
    return out, rep, lines


def parser(default_lang="en"):
    ap = argparse.ArgumentParser(description="Build the DMW deck on Template v2.")
    ap.add_argument("--lang", choices=["en", "ar"], default=default_lang)
    ap.add_argument("--template", help="default: templates/DMW Template v2 - EN|AR.pptx in this repo")
    ap.add_argument("--data", required=True)
    ap.add_argument("--graphs", required=True)
    ap.add_argument("--calendar", required=True)
    ap.add_argument("--technicals")
    ap.add_argument("--writeup", help="EN only: the Task 4 write-up docx, cross-checked against the data")
    ap.add_argument("--date", required=True, help="report date YYYY-MM-DD")
    ap.add_argument("--mapping", default=str(REPO / "data" / "Mapping.xlsx"),
                    help="names sheet: pass the state copy (state/Mapping.xlsx)")
    ap.add_argument("--names", default=str(C.data_file("name_overrides.json")), help="EN name overrides (fallback only)")
    ap.add_argument("--ar-names", default=str(C.data_file("name_overrides_ar.json")), help="AR name overrides (fallback only)")
    ap.add_argument("--glossary", default=str(C.data_file("ar_glossary.json")))
    ap.add_argument("--chart-state", help="chart fingerprints from the last build (EN stale-chart check)")
    ap.add_argument("--out", required=True)
    return ap


def drop_header_rule(slide):
    """Slides 2-3: no navy rule under the logo header. The slide's first navy band sits right under
    it, and the rule showed as a solid line over part of the band's top edge (Saad, 25 Nov 2026)."""
    for sh in slide.shapes:
        if not (sh.has_table and sh.name.endswith(".header")):
            continue
        for tcpr in sh.table._tbl.iter(qn("a:tcPr")):
            for ln in tcpr:
                if ln.tag in (qn("a:lnB"), qn("a:lnT")) and ln.find(qn("a:noFill")) is None:
                    for x in list(ln):
                        ln.remove(x)
                    ln.set("w", "0")
                    etree.SubElement(ln, qn("a:noFill"))


def main(default_lang="en"):
    out, rep, lines = build(parser(default_lang).parse_args())
    print("\n".join(lines))
    print(f"Saved {out}\nSaved {rep}")


if __name__ == "__main__":
    main()
