#!/usr/bin/env python3
"""Build the ANB Daily Market Watch (AR) deck. Mirrors build_dmw.py (EN).

Slides 1-3 from DMW - Data.xlsx sheet PPT-AR, DMW Graphs - AR.pptx and the calendar file
(Arabic calendar wording comes from the manual calendar file itself; ar_glossary.json
+ Mapping.xlsx are only a fallback). Slides 4-5 from Technicals -AR-.
Untranslated calendar rows are left as a yellow placeholder and FLAGGED, never guessed.
"""
import argparse, copy, json, re, sys
from datetime import datetime
from pathlib import Path
import openpyxl
sys.path.insert(0, str(Path(__file__).resolve().parent))
import mapcols  # noqa: E402
from lxml import etree
from pptx import Presentation
from pptx.util import Emu
import legacy_build_dmw_v1 as en
from legacy_build_dmw_v1 import (A, NS, RED, GREEN, rnd, num, pct, bps, set_cell, colored, set_run_color,
                       add_picture, remove_shape, replace_with_technicals, ole_previews, log, checks)

DAYS = ["الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت", "الأحد"]
MONTHS = ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر"]
IND = str.maketrans("0123456789", "٠١٢٣٤٥٦٧٨٩")


AR_DATE_STYLE = {"arabic_comma": False}   # set from the template's own date box


def ar_long_date(d):
    # Two published styles: older "الأحد,١٨ أكتوبر, ٢٠٢٦" and, from 15 Oct 2026,
    # "الأحد، ١٨ أكتوبر ٢٠٢٦" (Arabic comma, no comma before the year). Follow the template.
    if AR_DATE_STYLE["arabic_comma"]:
        return f"{DAYS[d.weekday()]}، {str(d.day).translate(IND)} {MONTHS[d.month-1]} {str(d.year).translate(IND)}"
    return f"{DAYS[d.weekday()]},{str(d.day).translate(IND)} {MONTHS[d.month-1]}, {str(d.year).translate(IND)}"


DAY_RE = r"^(%s)\s*[,،]" % "|".join(DAYS)


def ar_cal_date(d):
    return f"{str(d.day).translate(IND)}- {MONTHS[d.month-1]}-{str(d.year).translate(IND)}"


def norm_en(s):
    s = re.sub(r"[^a-z0-9 ]", " ", s.lower().replace("&", " and "))
    s = re.sub(r"\b(co|company|ltd|the|cos)\b", " ", s)
    return re.sub(r"\s+", " ", s).strip()


class Translator:
    def __init__(self, glossary, names_xlsx):
        g = json.loads(Path(glossary).read_text(encoding="utf-8"))
        self.company = {norm_en(k): v for k, v in g["company"].items()}
        self.event = g["event"]; self.econ = {norm_en(k): v for k, v in g["economic"].items()}
        self.codes = {}
        if names_xlsx and Path(names_xlsx).exists():
            wb_ = openpyxl.load_workbook(names_xlsx, read_only=True)
            # the team's Mapping.xlsx (sheet "mapping") since 29 Oct 2026; old "TASI CODE" files still read
            ws = wb_["mapping"] if "mapping" in wb_.sheetnames else wb_["TASI CODE"]
            ix = mapcols.idx(ws)   # columns by header name (1 Nov 2026: columns reordered)
            for row in ws.iter_rows(min_row=2, values_only=True):
                en, ar, short = (mapcols.get(row, ix, k) for k in ("en", "ar", "short"))
                if en and ar:
                    self.codes[norm_en(str(en))] = str(ar).strip()
                    if short:   # calendar file sometimes uses the short English name
                        self.codes.setdefault(norm_en(str(short)), str(ar).strip())

    def comp(self, en_name):
        k = norm_en(en_name)
        if k in self.company:
            return self.company[k].strip(), "glossary"
        if k in self.codes:
            return self.codes[k], "Mapping"
        return None, None

    def ev(self, text):
        t = re.sub(r"\d+(?:[.,]\d+)*", "{n}", text.strip()); nums = re.findall(r"\d+(?:[.,]\d+)*", text)
        for k, v in self.event.items():
            if k.lower() == t.lower() and v.count("{n}") == len(nums):
                for n in nums:
                    v = v.replace("{n}", n, 1)
                return v, "glossary"
        return None, None

    def eco(self, text):
        v = self.econ.get(norm_en(text))
        return (v, "glossary") if v else (None, None)


PH = "[ترجمة مطلوبة: {}]"


def find_tbl(slide, title):
    for sh in slide.shapes:
        if sh.has_table:
            t = sh.table
            head = " ".join(t.cell(r, c).text for r in range(min(2, len(t.rows))) for c in range(len(t.columns)))
            if title in head:
                return sh
    raise KeyError(title)


def build_slide1(slide, g, ge, charts, report_date):
    dt = [sh for sh in slide.shapes if sh.has_text_frame and re.match(DAY_RE, sh.text_frame.text.strip())][0]
    AR_DATE_STYLE["arabic_comma"] = "،" in dt.text_frame.text
    ts = dt.text_frame._txBody.findall(".//a:t", NS); ts[0].text = ar_long_date(report_date)
    for t in ts[1:]:
        t.text = ""
    cards = {True: dict(close=ge["C6"], vt=ge["F6"], ytd=ge["E8"], dod=ge["C8"]),      # right side = TASI
             False: dict(close=ge["L6"], vt=ge["O6"], ytd=ge["N8"], dod=ge["L8"])}
    pairs = en.card_pairs(slide)   # AR: TASI card on the right, NOMU on the left
    nomu_gap = en.card_gap(*pairs["left"]) if "left" in pairs else None
    for sh in slide.shapes:
        if not sh.has_text_frame or not (1_100_000 < sh.top < 2_150_000):
            continue
        d = cards[sh.left > 3_800_000]; t = sh.text_frame.text; runs = sh.text_frame.paragraphs[0].runs
        if not runs:
            continue
        if re.fullmatch(r"\s*[\d,]+\.\d{2}\s*", t):
            runs[0].text = f"{rnd(d['close'], 2):,.2f}"
        elif t.startswith("القيمة المتداولة"):
            numrun = [r for r in runs if re.search(r"\d", r.text)]
            # TASI value traded is in billions, NOMU in millions (EN: SAR bn / SAR mn). The
            # template carries no unit (every AR build to 4 Nov 2026 left it out), so the number
            # run always gets the full "<value> مليار ريال" / "<value> مليون ريال", and any unit
            # left in a later run is removed (user, 4 Nov 2026).
            is_tasi = sh.left > 3_800_000
            want = "مليار" if is_tasi else "مليون"
            if numrun:
                numrun[0].text = f"{num(d['vt'], 1)} {want} ريال"
                for r in runs[runs.index(numrun[0]) + 1:]:
                    if re.search(r"مليار|مليون|ريال", r.text):
                        r._r.getparent().remove(r._r)
                log("OK", "Slide 1 %s card value traded written with unit '%s ريال' (EN: SAR %s)." %
                    ("TASI" if is_tasi else "NOMU", want, "bn" if is_tasi else "mn"))
            else:
                log("FLAG", "Slide 1 %s card: value-traded number run not found." % ("TASI" if is_tasi else "NOMU"))
        elif re.fullmatch(r"\s*-?[\d.]+%\s*", t):
            is_ytd = sh.top > 1_850_000
            v = d["ytd"] if is_ytd else d["dod"]
            runs[0].text = pct(v)
            for r in runs[1:]:
                r._r.getparent().remove(r._r)
            # house rule: a positive change that shows as 0.0% stays default colour (18 Oct 2026)
            set_run_color(runs[0]._r.get_or_add_rPr(), RED if v < 0 else (GREEN if re.search(r"[1-9]", pct(v)) else None))
    if len(pairs) == 2:
        en.place_dod(pairs, nomu_gap, "AR slide 1", log)
    else:
        log("FLAG", "AR slide 1: headline card close/DoD boxes not found on both cards; DoD boxes not placed.")
    tbl = [sh for sh in slide.shapes if sh.has_table and "السوق الرئيسية" in " ".join(c.text for c in sh.table.rows[0].cells)][0].table
    vals = list(reversed([ge[f"{c}8"] for c in "CDEFGHIJKLMNOPQRST"]))
    ar_row = [g.get(f"{c}7") for c in ["G", "H", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "T", "U", "V", "W", "X", "Y"]]
    if [f"{x:.10f}" for x in ar_row] != [f"{x:.10f}" for x in vals[:17]]:
        log("CONFLICT", "Index Performance: PPT-AR row 7 differs from PPT-ENG row 8 (reversed).")
    for i, v in enumerate(vals):
        colored(tbl.cell(3, i), pct(v), v)
    band = [s for s in slide.shapes if s.shape_type == 13 and 3_300_000 < s.top < 5_700_000]
    ch = charts.get(1, [])
    if len(ch) != 2:
        log("FLAG", f"Graphs AR slide 1: expected 2 charts, found {len(ch)}.")
    else:
        fb = [(x, 3_700_000, 3_380_000, int(3_380_000 * c["height"] / c["width"]))
              for x, c in zip((380_000, 4_022_639), ch)]
        slots = en.inherit_slots(band, fb, "AR slide 1 charts")
        en.check_pair(slots, "AR slide 1 charts")
        for sh in band:
            remove_shape(sh)
        for c, slot, side in zip(ch, slots, ("left (P/E & DY)", "right (TASI)")):
            geom = en.fit_slot(slot, c, f"AR slide 1 {side} chart")
            add_picture(slide, c["blob"], c["ext"], *geom, "DMW chart")
        log("OK", "Slide 1 charts replaced from Graphs AR deck (P/E & DY left, TASI right).")
    for sh in list(slide.shapes):
        if sh.has_table and "إعلانات النتائج" in sh.table.cell(0, 0).text:
            remove_shape(sh); log("OK", "Removed template Results Announcements table.")


def commentary(prs):
    s1, s2 = prs.slides[0], prs.slides[1]
    HDRS = {"أداء الأسواق/الأخبار", "أخبار اقتصادية/قطاعية رئيسية", "أهم أخبار الشركات"}
    # height threshold is 300k, not 700k: a commentary box the previous day's author
    # resized down to its pasted text is still the box to overwrite, not one to clone beside.
    # The team sometimes groups each header with its body (AR deck, 15 Oct 2026), so look
    # inside groups too; a box in a group is still the box to overwrite.
    def flat(shapes):
        for sh in shapes:
            if sh.shape_type == 6:
                yield from flat(sh.shapes)
            else:
                yield sh
    s1_all = list(flat(s1.shapes))
    bodies = sorted([sh for sh in s1_all if sh.has_text_frame and not getattr(sh, "has_table", False)
                     and sh.width > 6_500_000 and sh.height > 300_000
                     and sh.text_frame.text.strip() not in HDRS], key=lambda s: s.top)
    corp_hdr = [sh for sh in s1_all if sh.has_text_frame and sh.text_frame.text.strip().startswith("أهم أخبار الشركات")]
    if len(bodies) == 2 and corp_hdr and not any(b.top > corp_hdr[0].top for b in bodies):
        # corporate body missing in template (a Results table sat there): clone the econ body
        # Append the clone to the slide itself, never to the group the source box sits in
        # (19 Oct 2026: bodies were grouped with their headers, the clone landed inside a
        # group and the lookup below on s1.shapes found nothing).
        el = copy.deepcopy(bodies[-1]._element); s1.shapes._spTree.append(el)
        el.find('.//p:cNvPr', NS).set('id', str(max(sh.shape_id for sh in s1.shapes) + 1))
        b = [sh for sh in s1.shapes if sh._element is el][0]
        b.top = Emu(corp_hdr[0].top + 175_000); b.height = Emu(max(9_150_000 - b.top, 300_000))
        bodies.append(b)
        log("FLAG", "Slide 1: no text box under 'أهم أخبار الشركات' in template; added one.")
    labels = [["[الأسواق العالمية – لصق]", "[تاسي – لصق]"], ["[أخبار اقتصادية/قطاعية – لصق]"], ["[أخبار الشركات – لصق، تكملة في الصفحة 2]"]]
    for b, lab in zip(bodies, labels):
        en.placeholder_paragraphs(b, lab)
    if len(bodies) != 3:
        log("FLAG", f"Slide 1: expected 3 commentary boxes, found {len(bodies)}.")
    big = max((sh for sh in en.flat_shapes(s2.shapes) if sh.has_text_frame and not getattr(sh, "has_table", False)),
              key=lambda s: s.height)
    en.placeholder_paragraphs(big, ["[أخبار الشركات (تكملة) – لصق]"])
    log("OK", "Commentary boxes on slides 1-2 set to yellow placeholders.")


def load_calendar_ar(path):
    """Arabic calendar wording, taken straight from the manual calendar file.

    The team writes the Arabic alongside the English in the calendar sheet (Sheet1 since
    10 Oct 2026, found by header like the EN reader), so
    that file is the source of truth, not the glossary. RTL column order:
      events   H = الحدث, I = الشركة, J = التاريخ   (anchored on the EN date in C)
      economic U = الحدث, V = التاريخ                (anchored on the EN date in N)
    Prior and estimate are the same raw values as the English columns, so they
    keep their existing formatting.
    """
    ws, hr = en.cal_sheet(path)

    def txt(cell):
        v = cell.value
        t = str(v).strip() if v is not None else ""
        return "" if t in ("_", "-", "–") else t  # the file's "_" blank shows as empty, like EN "-"

    k = en.cal_columns(ws, hr)
    ev, ec, r = [], [], hr + 1
    while isinstance(ws.cell(r, k["ev_date"]).value, datetime):
        ev.append((txt(ws.cell(r, k["ev_event_ar"])), txt(ws.cell(r, k["ev_company_ar"]))))
        r += 1
    r = hr + 1
    while isinstance(ws.cell(r, k["ec_date"]).value, datetime):
        ec.append(txt(ws.cell(r, k["ec_event_ar"])))
        r += 1
    return ev, ec

def build_slide2(slide, events, econ, tr, ar_cal=(None, None)):
    ev = find_tbl(slide, "تقويم الأحداث"); ec = find_tbl(slide, "التقويم الاقتصادي")
    ev_ar, ec_ar = ar_cal[0] or [], ar_cal[1] or []
    if ev_ar and len(ev_ar) != len(events):
        log("FLAG", f"Event calendar: {len(ev_ar)} Arabic rows vs {len(events)} English rows "
                    "in the manual file. Rows may be misaligned.")
        ev_ar = []
    if ec_ar and len(ec_ar) != len(econ):
        log("FLAG", f"Economic calendar: {len(ec_ar)} Arabic rows vs {len(econ)} English rows "
                    "in the manual file. Rows may be misaligned.")
        ec_ar = []
    ev_rows, ec_rows = [], []
    for i, (d, comp, event) in enumerate(events):
        e_ar, c_ar = ev_ar[i] if i < len(ev_ar) else ("", "")
        sheet_name, _ = tr.comp(comp)
        if not e_ar:
            e_ar, _ = tr.ev(event)
            if not e_ar:
                log("FLAG", f"Event calendar: no Arabic wording for event '{event}' in the "
                            "manual file or the glossary. Placeholder left.")
        if not c_ar:
            c_ar = sheet_name
            if not c_ar:
                log("FLAG", f"Event calendar: no Arabic name for company '{comp}' in the "
                            "manual file or the code sheet. Placeholder left.")
        elif sheet_name and c_ar != sheet_name:
            log("FLAG", f"Event calendar: manual file Arabic name '{c_ar}' for '{comp}' differs "
                        f"from the code sheet '{sheet_name}'. Manual file used.")
        ev_rows.append((e_ar or PH.format(event), c_ar or PH.format(comp), ar_cal_date(d)))
    for i, (d, e, p, s) in enumerate(econ):
        e_ar = ec_ar[i] if i < len(ec_ar) else ""
        if not e_ar:
            e_ar, _ = tr.eco(e)
            if not e_ar:
                log("FLAG", f"Economic calendar: no Arabic wording for '{e}' in the manual file "
                            "or the glossary. Placeholder left.")
        if e == "US EIA Crude Oil Inventories" and e_ar and "EIA" not in e_ar and "معلومات الطاقة" not in e_ar:
            # client form (25 Nov 2026, Muryum via Saad); Arabic written by Claude. The calendar
            # owner's own Arabic wins when it already names the EIA (e.g. "وفقًا لإدارة معلومات الطاقة")
            log("OK", f"Economic calendar: '{e_ar}' -> client form with (EIA).")
            e_ar = "مخزونات النفط الخام الأمريكية (EIA)"
        ec_rows.append((s, p, e_ar or PH.format(e), ar_cal_date(d)))
    # Header row count is read from the template, not assumed: the econ table has
    # carried both 2 and 3 header rows. Data starts after the column-label row
    # (the one whose last cell is "التاريخ"); the source row is always last.
    def first_data_row(sh):
        t = sh.table
        for r in range(min(4, len(t.rows))):
            if t.cell(r, len(t.columns) - 1).text.strip() == "التاريخ":
                return r + 1
        log("FLAG", f"{t.cell(0, 0).text.strip()}: column-label row not found; assumed 2 header rows.")
        return 2
    # House rule (user, 13 Oct 2026): the AR calendar title bars are bold, like the EN ones.
    for tbl_sh in (ev, ec):
        for para in tbl_sh.table.cell(0, 0).text_frame.paragraphs:
            for run in para.runs:
                run.font.bold = True
    ev_first, ec_first = first_data_row(ev), first_data_row(ec)
    en.resize_table(ev, len(ev_rows), ev_first, ev_rows)
    en.resize_table(ec, len(ec_rows), ec_first, ec_rows)
    # Latin inside an Arabic event name ("(EIA)", "(API)") goes in its own en-US run, as on
    # slide 3, so PowerPoint keeps the brackets in place (25 Nov 2026)
    for r in range(ec_first, ec_first + len(ec_rows)):
        cell = ec.table.cell(r, 2)
        if LATIN_SPAN.search(cell.text):
            set_ar_label(cell, cell.text)
    ROW = 162000
    ev_h = 434340 + ROW * len(ev_rows) + sum(ROW // 2 for r in ev_rows if len(r[0]) + len(r[1]) > 95)
    ec_h = (472440 if ec_first == 2 else 634440) + ROW * len(ec_rows) + sum(ROW // 2 for r in ec_rows if len(r[2]) > 70)
    en.stack_calendars(slide, ev, ec, ev_h, ec_h)
    for tbl_sh in (ev, ec):
        for r in range(len(tbl_sh.table.rows)):
            for c in range(len(tbl_sh.table.columns)):
                if tbl_sh.table.cell(r, c).text.startswith("[ترجمة"):
                    rp = tbl_sh.table.cell(r, c).text_frame.paragraphs[0].runs[0]._r.get_or_add_rPr()
                    hl = etree.Element(f"{{{A}}}highlight"); x = etree.SubElement(hl, f"{{{A}}}srgbClr"); x.set("val", "FFFF00"); rp.insert(0, hl)
    log("OK", f"Slide 2 calendars: {len(ev_rows)} events, {len(ec_rows)} economic releases.")


AR_SHORT = {}   # full Arabic name -> short name the team supplies (state/name_overrides_ar.json)


def ar_label(name):
    """Slide 3 labels from the data file, with the spellings the team corrects by hand
    (15 Oct 2026 final): Persian ک/ی, الإتصالات, ارامكو, and a detached و ('و الشخصية')."""
    name = name.replace("ک", "ك").replace("ی", "ي")
    name = name.replace("الإتصالات", "الاتصالات").replace("ارامكو", "أرامكو")
    return re.sub(r"(?<=\s)و\s+(?=[ء-ي])", "و", name)


LATIN_SPAN = re.compile(r"(?:\d+(?=[A-Za-z]))?[A-Za-z][A-Za-z0-9&.\-’ ]*")


def set_ar_label(cell, text):
    """Write a slide 3 label with each Latin span (S&P 500, NIKKEI 225, MSCI) in its own en-US
    run, as the template has them. One ar-SA run made PowerPoint treat the digits as Arabic and
    show "(500S&P)", "(225NIKKEI)", "(100FTSE)" (24 Nov 2026; LibreOffice renders it fine)."""
    set_cell(cell, text)
    if not LATIN_SPAN.search(text):
        return
    base = cell.text_frame.paragraphs[0].runs[0]._r
    pieces, pos = [], 0
    for m in LATIN_SPAN.finditer(text):
        s, e = m.start(), m.start() + len(m.group().rstrip())
        if s > pos:
            pieces.append((text[pos:s], False))
        pieces.append((text[s:e], True)); pos = e
    if pos < len(text):
        pieces.append((text[pos:], False))
    for piece, latin in pieces:
        nr = copy.deepcopy(base); nr.find(f"{{{A}}}t").text = piece
        rp = nr.find(f"{{{A}}}rPr")
        if rp is not None:
            rp.set("lang", "en-US" if latin else rp.get("lang", "ar-SA"))
        base.addprevious(nr)
    base.getparent().remove(base)


def fill(slide, title, g, cols, r0, r1, spec, label=True):
    t = find_tbl(slide, title).table
    # data rows: first row after the header rows whose last cell is not a header
    start = 2
    for i in range(r1 - r0 + 1):
        r = start + i; row = [g.get(f"{c}{r0+i}") for c in cols]
        for j, (kind, v) in enumerate(zip(spec, row)):
            if kind == "label":
                if label:
                    name = ar_label(str(v).strip())
                    name = AR_SHORT.get(name, name)
                    # the fixed template carries the team's wording of the standing labels ("العالم (MSCI)",
                    # "S&P 500", "مؤشر ناسداك"); keep it when it is the data file's label reordered or
                    # respaced (same characters), so the team no longer re-fixes it daily (29 Oct 2026)
                    cur = t.cell(r, j).text.strip()
                    same = lambda x: sorted(re.sub(r"[\s()]", "", x))
                    set_ar_label(t.cell(r, j), cur if cur and same(cur) == same(name) else name)
            elif kind == "n1":
                set_cell(t.cell(r, j), num(v, 1))
            elif kind == "n2":
                set_cell(t.cell(r, j), num(v, 2))
            elif kind == "pct":
                colored(t.cell(r, j), pct(v), v)
            elif kind == "bps":
                colored(t.cell(r, j), bps(v), v)


def build_slide3(slide, g, charts):
    sec = ["pct", "n1", "pct", "n1", "label"]; co = ["n1", "n1", "pct", "n1", "label"]
    fill(slide, "أفضل القطاعات أداءً", g, "RSTUV", 76, 80, sec)
    fill(slide, "أضعف القطاعات أداءً", g, "JKLMN", 76, 80, sec)
    fill(slide, "أفضل الشركات أداءً", g, "RSTUV", 84, 88, co)
    fill(slide, "أضعف الشركات أداءً", g, "JKLMN", 84, 88, co)
    fill(slide, "قادة القيمة", g, "RSTUV", 92, 96, co)
    fill(slide, "قادة الحجم", g, "JKLMN", 92, 96, co)
    fill(slide, "الأسواق العالمية", g, "RSTUV", 102, 111, sec)
    fill(slide, "الأسواق الإقليمية", g, "JKLMN", 102, 111, sec)
    fill(slide, "أسعار السلع الأساسية", g, "JKLMN", 115, 126, ["pct", "pct", "pct", "n1", "label"], label=False)
    fill(slide, "أسعار الفائدة", g, "RSTUV", 115, 118, ["bps", "bps", "bps", "n2", "label"])
    log("OK", "Slide 3 tables filled from PPT-AR.")
    check_ar_name_widths(slide)
    ch = charts.get(3, [])
    if len(ch) != 1:
        log("FLAG", f"Graphs AR slide 3: expected 1 chart, found {len(ch)}."); return
    old = [s for s in slide.shapes if s.shape_type == 13 and 3_400_000 < s.top < 5_800_000 and s.left > 3_900_000]
    c = ch[0]; cx = 3_370_708
    fb = [(4_005_070, 3_587_004, cx, int(cx * c["height"] / c["width"]))]
    # one frame size for EN and AR (en.CHART3_IN, 2 Nov 2026); the AR chart keeps its right edge
    s0 = en.inherit_slots(old, fb, "AR slide 3 sector chart")[0]
    slot = en.chart3_slot(s0, "right", ref_edge=s0[0] + s0[2], ref_top=s0[1], label="AR slide 3 sector chart")
    for sh in old:
        remove_shape(sh)
    geom = en.fit_slot(slot, c, "AR slide 3 sector chart", tol=en.CHART3_TOL)
    add_picture(slide, c["blob"], c["ext"], *geom, "Sector YTD chart")
    log("OK", "Slide 3 sector YTD chart replaced from Graphs AR deck.")


AR_FONT = "Arial"


def check_ar_name_widths(slide):
    """Flag company names that will wrap in the slide 3 company tables (a wrapped name makes
    the table taller and pushes it into the one below). Width measured with the real Apercu
    Arabic font and Arabic shaping (Pillow + raqm); the column is the table's last (name) column."""
    font_path = Path(__file__).with_name("fonts") / "ApercuArabic-Regular-Pro.otf"
    try:
        from PIL import ImageFont, features
        if not features.check("raqm"):
            log("FLAG", "AR name width check skipped: Pillow has no raqm shaping."); return
    except ImportError:
        return
    cache = {}
    for sh in slide.shapes:
        if not sh.has_table or "الشركات" not in sh.table.cell(0, 0).text:
            continue
        t = sh.table; c = len(t.columns) - 1
        col_in = t.columns[c].width / 914400
        for r in range(2, len(t.rows)):
            name = t.cell(r, c).text.strip()
            if not name:
                continue
            runs = t.cell(r, c).text_frame.paragraphs[0].runs
            pt = runs[0].font.size.pt if runs and runs[0].font.size else 7.0
            if pt not in cache:
                cache[pt] = ImageFont.truetype(str(font_path), int(pt * 100), layout_engine=ImageFont.Layout.RAQM)
            w = cache[pt].getlength(name, direction="rtl", language="ar") / 100 / 72
            if w > col_in:
                log("FLAG", f"AR name '{name}' is {w:.2f} in wide and the column is {col_in:.2f} in: it will wrap "
                            "and push the table down. Needs a shorter Arabic name.")


def fix_arabic_fonts(prs, n_slides=3):
    """Force an Arabic-capable complex-script font. The template's 'anb Apercu' is not installed on
    most PCs; PowerPoint then garbles the Arabic. Published decks render the Arabic in Arial."""
    n = 0
    for slide in list(prs.slides)[:n_slides]:
        for rpr in slide._element.iter(f"{{{A}}}rPr", f"{{{A}}}endParaRPr", f"{{{A}}}defRPr"):
            r = rpr.getparent()
            t = r.find(f"{{{A}}}t") if r is not None else None
            arabic = t is not None and t.text and any("؀" <= c <= "ۿ" for c in t.text)
            cs = rpr.find(f"{{{A}}}cs")
            if cs is not None and cs.get("typeface") != AR_FONT:
                cs.set("typeface", AR_FONT)
                for k in ("panose", "pitchFamily", "charset"):
                    cs.attrib.pop(k, None)
                n += 1
            elif cs is None and arabic:
                cs = etree.SubElement(rpr, f"{{{A}}}cs"); cs.set("typeface", AR_FONT)
                # schema order: ... latin, ea, cs, sym, hlinkClick ...
                rpr.remove(cs)
                idx = len(rpr)
                for i, ch in enumerate(rpr):
                    if etree.QName(ch).localname in ("sym", "hlinkClick", "hlinkMouseOver", "rtl", "extLst"):
                        idx = i; break
                rpr.insert(idx, cs)
                n += 1
    log("OK", f"Arabic text set to {AR_FONT} (complex-script font) on slides 1-{n_slides}: {n} runs.")


# Typos the team corrects by hand in the AR Technicals slides (back-test of the 70 finals, 15 Jul to
# 21 Oct 2026): thousands written with dots (10.132.38) never reach a final deck, and the footer's
# "التقرير.هذا" and "هاتف: .800" were fixed on 15 Oct. Wording edits (idafa) are left to the team.
TECH_AR_FIXES = [
    (re.compile(r"(?<![\d.,])(\d{1,3})\.(\d{3})\.(\d{1,2})(?![\d.,])"), r"\1,\2.\3", "dotted thousands"),
    (re.compile(r"التقرير\.(?=\S)"), "التقرير. ", "missing space after full stop"),
    (re.compile(r"هاتف:\s*\.(?=\d)"), "هاتف: ", "stray full stop before the phone number"),
    (re.compile(r"الإتصالات"), "الاتصالات", "hamza on الاتصالات"),
]


S5_RULE_XML = ('<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
               'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:nvSpPr>'
               '<p:cNvPr id="{id}" name="Header rule"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>'
               '<a:xfrm><a:off x="365760" y="658368"/><a:ext cx="7040880" cy="0"/></a:xfrm>'
               '<a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="15875"><a:solidFill>'
               '<a:srgbClr val="002060"/></a:solidFill><a:prstDash val="solid"/></a:ln></p:spPr>'
               '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>')


def ensure_s5_rule(slide, logf=None):
    """AR slide 5 (Stock of the Day) needs the navy header rule under the title (0.40, 0.72 in,
    7.70 in wide, 1.25 pt, 002060). The EN Technicals deck carries it; the AR one keeps arriving
    without it and the team adds it by hand (4 and 5 Nov 2026; user, 5 Nov 2026)."""
    from lxml import etree
    E = 914400
    for sh in slide.shapes:
        if sh.height < 9144 and sh.width > 7 * E and 0.6 * E < sh.top < 0.85 * E:
            return False
    ids = [int(x.get("id")) for x in slide._element.iter() if x.tag.endswith("}cNvPr") and x.get("id", "").isdigit()]
    el = etree.fromstring(S5_RULE_XML.format(id=max(ids + [1]) + 1))
    slide.shapes._spTree.append(el)
    if logf:
        logf("OK", "AR slide 5: navy header rule added under the title (missing from the Technicals AR deck).")
    return True


def tidy_technicals_ar(slides):
    n = {}
    for sl in slides:
        frames = []
        for sh in en.flat_shapes(sl.shapes):
            if getattr(sh, "has_table", False) and sh.has_table:   # the key-figures table (15 Oct 2026)
                frames += [c.text_frame for row in sh.table.rows for c in row.cells]
            elif sh.has_text_frame:
                frames.append(sh.text_frame)
        for tf in frames:
            for p in tf.paragraphs:
                for r in p.runs:
                    for rx, rep, why in TECH_AR_FIXES:
                        new, k = rx.subn(rep, r.text)
                        if k:
                            r.text = new; n[why] = n.get(why, 0) + k
    for why, k in n.items():
        log("OK", f"Technicals AR: fixed {k} x {why}.")


def main():
    ap = argparse.ArgumentParser()
    for k in ("template", "data", "graphs", "calendar", "technicals", "date", "out"):
        ap.add_argument("--" + k, required=k not in ("technicals", "template"))
    ap.add_argument("--glossary", default=str(en.data_file("ar_glossary.json")))
    ap.add_argument("--names", default=str(Path(__file__).resolve().parent.parent / "data" / "Mapping.xlsx"))
    ap.add_argument("--ar-names", default=str(en.data_file("name_overrides_ar.json")),
                    help="short Arabic names for slide 3 (full name -> short name)")
    ap.add_argument("--arial-arabic", action="store_true",
                    help="force Arial for Arabic text (only for PCs without the anb Apercu font)")
    a = ap.parse_args()
    if a.ar_names and Path(a.ar_names).exists():
        AR_SHORT.update(json.loads(Path(a.ar_names).read_text(encoding="utf-8")))
    rd = datetime.strptime(a.date, "%Y-%m-%d").date()
    wb = openpyxl.load_workbook(a.data, data_only=True, read_only=True)
    g = {c.coordinate: c.value for row in wb["PPT-AR"].iter_rows(min_row=1, max_row=140, max_col=30) for c in row if c.value is not None}
    ge, face = en.load_data(a.data)
    if rd not in [d.date() for d in face]:
        log("CONFLICT", "Data file Face sheet does not carry the report date.")
    events, econ = en.load_calendar(a.calendar, rd)
    charts, _ = ole_previews(a.graphs)
    tr = Translator(a.glossary, a.names)
    import snap_spec
    prs = Presentation(a.template or str(Path(__file__).resolve().parent.parent / "templates" / "DMW Template - AR.pptx"))
    snap_spec.snap(prs, "AR", log, fix=True)
    tech = Presentation(a.technicals) if a.technicals else None
    for sh in prs.slides[0].shapes:   # date style from the template, before any date is compared
        if sh.has_text_frame and re.match(DAY_RE, sh.text_frame.text.strip()):
            AR_DATE_STYLE["arabic_comma"] = "،" in sh.text_frame.text
    if tech is not None:
        td = [sh.text_frame.text.strip() for sh in tech.slides[0].shapes if sh.has_text_frame and re.match(DAY_RE, sh.text_frame.text.strip())]
        # The Technicals author often writes the date in the old punctuation style. Same date in another
        # style is fixed on slide 4 after the paste below (user, 20 Oct 2026); a different date is a CONFLICT.
        bare = lambda t: re.sub(r"[\s,،]+", " ", t).strip()
        tech_date_fix = bool(td) and td[0] != ar_long_date(rd) and bare(td[0]) == bare(ar_long_date(rd))
        if not td or (td[0] != ar_long_date(rd) and not tech_date_fix):
            log("CONFLICT", f"Technicals AR date {td[:1]} vs report date '{ar_long_date(rd)}'.")
    en.clean_hidden_titles(prs, log)
    build_slide1(prs.slides[0], g, ge, charts, rd)
    commentary(prs)
    build_slide2(prs.slides[1], events, econ, tr, ar_cal=load_calendar_ar(a.calendar))
    en.drop_results_s2(prs.slides[1], lambda t: "إعلانات النتائج" in t or t.strip().lower().startswith("result"),
                       "أهم أخبار الشركات")
    build_slide3(prs.slides[2], g, charts)
    if tech is not None:
        replace_with_technicals(prs.slides[3], tech.slides[0]); replace_with_technicals(prs.slides[4], tech.slides[1])
        log("OK", "Slides 4-5 replaced from Technicals AR deck.")
        tidy_technicals_ar([prs.slides[3], prs.slides[4]])
        ensure_s5_rule(prs.slides[4], log)
        if tech_date_fix:
            for sh in prs.slides[3].shapes:
                if sh.has_text_frame and re.match(DAY_RE, sh.text_frame.text.strip()):
                    runs = sh.text_frame.paragraphs[0].runs
                    runs[0].text = ar_long_date(rd)
                    for r in runs[1:]:
                        r._r.getparent().remove(r._r)
            log("OK", f"Technicals AR date '{td[0]}' rewritten to the house form '{ar_long_date(rd)}' on slide 4.")
    else:
        log("FLAG", "No Technicals AR deck supplied: slides 4-5 are the template's old Technicals pages. Delete them "
                    "(Task 7 step 4) or rebuild with --technicals before sending.")
    if a.arial_arabic:
        fix_arabic_fonts(prs)
    snap_spec.snap(prs, "AR", log, fix=False)      # self-check: the build moved nothing fixed
    out = Path(a.out); prs.save(out)
    lines = [f"# DMW AR build checks - {rd:%d %b %Y}", ""]
    for lvl in ("STALE", "CONFLICT", "FLAG", "OK"):
        items = list(dict.fromkeys(m for l, m in checks if l == lvl))
        if items:
            lines += [f"## {lvl} ({len(items)})", ""] + [f"- {m}" for m in items] + [""]
    rep = out.with_name(out.stem + "_checks.md"); rep.write_text("\n".join(lines), encoding="utf-8"); print("\n".join(lines))


if __name__ == "__main__":
    main()
