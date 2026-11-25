"""Place the Results Announcements tables in the filled EN and AR DMW decks (Task 10).

First written for the 4 Nov 2026 edition; generalised the same day (no hard-coded rows, names,
calendar drops or slide 1 row count).

What it does, in both decks:
  - reads DMW - Results.xlsx (sheet Results, data from row 6: A period, B short name, C TASI code,
    D Bloomberg code, E current, F prior, H YoY %, I consensus, J variance);
  - recomputes change, YoY % and variance from the workbook's own figures with exact decimals,
    half-up to 1 dp (a cell the workbook shows as "-" stays "-");
  - sorts: ANB-covered first (--universe, "ANB Universe" sheet only), then uncovered, each by
    current net profit high to low (losses last); rows for another period than --quarter go at
    the very end, starred (the period footnote is reported for the user to add);
  - clones yesterday's table frames (title, header rows, one data row, source row) from the
    previous edition's decks, so the table looks exactly like the team's;
  - slide 1: removes the empty Corporate box and puts the table under the Corporate header; the
    number of rows on slide 1 is the most that fits above --limit in BOTH decks (the same split
    in EN and AR), unless --n1 is given;
  - slide 2: the rest under the Corporate header, with the source line; the news box and the
    calendar tables (and groups) move down by the table's height;
  - Arabic names: Long Name - Arabic from the names sheet, matched on the Bloomberg code; when it
    would wrap in the AR name column (measured with the AR deck font when it is staged, else a
    character count), the sheet's Short Name - Arabic; --ar-short JSON (Bloomberg code -> name)
    overrides both;
  - crude oil inventory rows in the economic calendar: Prior and Estimate set to "-"
    (calendar rule, 21 Oct 2026); rows are never removed for room (user, 4 Nov 2026);
  - --skip: rows left out because they arrived after the 10:30 Results cut-off (user, 4 Nov 2026);
  - --drop-cal: economic calendar rows to remove (text fragments, ";"-separated) when the
    calendar owner dropped them after the build.

Usage:
  python place_results.py --results "DMW - Results.xlsx" --names Mapping.xlsx \
      --prev "<previous edition's day folder>" --en <filled EN deck> --ar <filled AR deck> \
      --quarter 3Q26 [--universe "ANB Coverage Universe Market Cap - <date>.xlsx"] \
      [--ar-short state/name_overrides_results_ar.json] [--drop-cal "KSA Budget Balance;..."] \
      [--skip "SAICO;AYYAN"] [--n1 N] [--limit 10.60] --outdir <dir>
Writes "<outdir>/ANB Daily Market Watch - EN.pptx" and "- AR.pptx" and prints a short report.
"""
import argparse, copy, json, os, re, sys
from decimal import Decimal, ROUND_HALF_UP
import openpyxl
from lxml import etree
from pptx import Presentation

E = 914400
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
CORP = {"EN": ("MAJOR CORPORATE NEWS",), "AR": ("أهم أخبار الشركات",)}
CRUDE = ("Crude Oil Inventories", "مخزونات النفط الخام")
ROW_H, SRC_H = int(0.16 * E), int(0.12 * E)


def d(v):
    return Decimal(repr(v)) if isinstance(v, float) else Decimal(v)


def num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def f1(v):
    if not num(v):
        return "-"
    return f"{d(v).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP):,.1f}"


def read_rows(path, quarter):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["Results"]
    rows = []
    for r in _with_formulas(path, wb, ws.iter_rows(min_row=6, values_only=True)):
        if not (r[1] and num(r[2]) and isinstance(r[3], str) and r[3].strip() not in ("", "-") and num(r[4])):
            continue
        cur, pri, cons = r[4], r[5], r[8]
        chg = float(d(cur) - d(pri)) if num(pri) else "-"
        pct = float((d(cur) / d(pri) - 1) * 100) if num(r[7]) and num(pri) and pri else "-"
        var = float((d(cur) / d(cons) - 1) * 100) if num(r[9]) and num(cons) and cons else "-"
        period = str(r[0] or "").strip()
        rows.append(dict(period=period, other=bool(quarter and period and period.upper() != quarter.upper()),
                         name=r[1].strip(), code=str(int(r[2])), bbg=r[3].strip(), cur=cur, pri=pri,
                         chg=chg, pct=pct, cons=cons, var=var))
    return rows



def _with_formulas(path, wb, it):
    """Rows of the Results sheet with the formula cells filled in where the workbook has no cached
    values (a workbook last saved by openpyxl, as the unattended S4 saves it, reads as None there;
    4 Nov replay, 26 Nov 2026). C and D are XLOOKUPs of B into TASI CODE (C -> D code, C -> I
    Bloomberg); H and J are the YoY and consensus % (only whether they are numbers matters here)."""
    rows = [list(r) for r in it]
    if all(r[2] is not None for r in rows if r[1]):
        yield from rows
        return
    code = {}
    if "TASI CODE" in wb.sheetnames:
        for c in wb["TASI CODE"].iter_rows(min_row=1, values_only=True):
            if len(c) > 8 and c[2]:
                code.setdefault(str(c[2]).strip().upper(), (c[3], c[8]))
    for r in rows:
        r += [None] * (10 - len(r))
        if r[1] and r[2] is None:
            r[2], r[3] = code.get(str(r[1]).strip().upper(), ("-", "-"))
        e, f, i = r[4], r[5], r[8]
        if r[7] is None and num(e) and num(f):
            r[7] = "-" if e * f < 0 or not f else (e - f) / f * 100
        if r[9] is None and num(e):
            r[9] = "-" if not num(i) or e * i < 0 or not i else (e - i) / i * 100
        yield r

def covered_set(path):
    if not path:
        return set()
    ws = openpyxl.load_workbook(path, data_only=True)["ANB Universe"]
    out = set()
    for r in ws.iter_rows(values_only=True):
        for v in r[:6]:
            if v is None:
                continue
            s = str(v).strip().upper()
            if s:
                out.add(s)
                out.add(s.replace(" AB", "").replace(" AB EQUITY", ""))
    return out


def is_covered(x, cov):
    keys = {x["bbg"].upper(), x["bbg"].upper().replace(" AB", ""), x["code"], x["name"].upper()}
    return bool(keys & cov)


def sort_rows(rows, cov):
    for x in rows:
        x["covered"] = is_covered(x, cov)
    rows.sort(key=lambda x: (x["other"], not x["covered"], -x["cur"]))
    for x in rows:
        if x["other"]:
            x["name"] = x["name"].rstrip("*") + "*"
    return rows


def names_sheet(path):
    ws = openpyxl.load_workbook(path, data_only=True)["mapping"]
    hdr = [str(c.value).strip() if c.value else "" for c in next(ws.iter_rows(min_row=1, max_row=1))]
    ib = next(i for i, h in enumerate(hdr) if "bloomberg" in h.lower() and "name" not in h.lower())
    il, ishort = hdr.index("Long Name - Arabic"), hdr.index("Short Name - Arabic")
    out = {}
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r[ib]:
            k = str(r[ib]).strip().upper()
            out[k if k.endswith(" AB") else k + " AB"] = (r[il], r[ishort])
    return out


def text_width_in(text, size_pt, font_path):
    try:
        from PIL import ImageFont
        f = ImageFont.truetype(font_path, int(size_pt * 20))
        return f.getlength(text) / 20 / 72
    except Exception:
        return len(text) * size_pt * 0.45 / 72


def results_frame(prs, si):
    for sh in prs.slides[si].shapes:
        if sh.has_table:
            t = sh.table.cell(0, 0).text.strip()
            if t.startswith("Result") or "النتائج" in t:
                return sh
    return None


def walk(shs):
    for s in shs:
        yield s
        if s.shape_type == 6:
            yield from walk(s.shapes)


def hdr_shape(slide, texts):
    return next(s for s in walk(slide.shapes)
                if s.has_text_frame and s.text_frame.text.strip() in texts and s.width > 5 * E)


def settext(tc, txt):
    txb = tc.find(A + "txBody")
    ps = txb.findall(A + "p")
    for p in ps[1:]:
        txb.remove(p)
    rs = ps[0].findall(A + "r")
    if not rs:
        r = etree.SubElement(ps[0], A + "r")
        etree.SubElement(r, A + "t")
        rs = [r]
    for r in rs[1:]:
        ps[0].remove(r)
    rs[0].find(A + "t").text = txt


def cells(x, lang):
    v = [x["name"], x["code"], x["bbg"], f1(x["cur"]), f1(x["pri"]), f1(x["chg"]), f1(x["pct"]),
         f1(x["cons"]), f1(x["var"])]
    if lang == "AR":
        v[0] = x.get("ar") or "TK"
        if x.get("other") and not v[0].endswith("*"):
            v[0] += "*"             # 5 Nov 2026: the AR short name (أبو معطي) came out without the star
        v = v[::-1]
    return v


def head_rows(frame_el, lang):
    """Header rows of a cloned frame: everything before the first data row (EN 3, AR 4)."""
    trs = frame_el.find(".//" + A + "tbl").findall(A + "tr")
    for i, tr in enumerate(trs):
        cs = ["".join(t.text or "" for t in tc.iter(A + "t")).strip() for tc in tr.findall(A + "tc")]
        if any(c.replace(",", "").replace(".", "").isdigit() and len(c) == 4 for c in cs):   # TASI code cell
            return i
    return 3 if lang == "EN" else 4


AR_ORD = {"1": "الأول", "2": "الثاني", "3": "الثالث", "4": "الرابع"}


def period_note(period, lang):
    """Footnote line for a different-period row, as the team writes it (1 Nov 2026 Amwaj:
    '*For 1H26' / '* للنصف الأول من عام 2026'; 5 Nov 2026 Abo Moati 2Q27)."""
    if lang == "EN":
        return f"*For {period}"
    m = re.fullmatch(r"([1-4])Q(\d\d)", period.upper())
    if m:
        return f"* للربع {AR_ORD[m.group(1)]} من عام 20{m.group(2)}"
    m = re.fullmatch(r"([12])H(\d\d)", period.upper())
    if m:
        return f"* للنصف {'الأول' if m.group(1) == '1' else 'الثاني'} من عام 20{m.group(2)}"
    m = re.fullmatch(r"9M(\d\d)", period.upper())
    if m:
        return f"* للأشهر التسعة الأولى من عام 20{m.group(1)}"
    m = re.fullmatch(r"FY(\d\d)", period.upper())
    if m:
        return f"* للسنة المالية 20{m.group(1)}"
    return f"* {period}"


def add_notes(src_tr, notes):
    """Add footnote lines under the source line in the source row's text cell."""
    for tc in src_tr.findall(A + "tc"):
        txt = "".join(t.text or "" for t in tc.iter(A + "t"))
        if "Source" in txt or "المصدر" in txt:
            txb = tc.find(A + "txBody")
            p0 = txb.findall(A + "p")[0]
            for n in notes:
                np_ = copy.deepcopy(p0)
                rs = np_.findall(A + "r")
                for r in rs[1:]:
                    np_.remove(r)
                rs[0].find(A + "t").text = n
                txb.append(np_)
            return True
    return False


def build(tmpl_sh, data, lang, with_source, notes=()):
    fr = copy.deepcopy(tmpl_sh._element)
    tbl = fr.find(".//" + A + "tbl")
    trs = tbl.findall(A + "tr")
    nh = head_rows(fr, lang)
    head, row_t = trs[:nh], trs[nh]
    src = trs[-1] if with_source else None
    for tr in trs[nh:]:
        tbl.remove(tr)
    for x in data:
        tr = copy.deepcopy(row_t)
        tr.set("h", str(ROW_H))
        for tc, t in zip(tr.findall(A + "tc"), cells(x, lang)):
            settext(tc, t)
        tbl.append(tr)
    if src is not None:
        src.set("h", str(SRC_H * (1 + len(notes))))
        if notes:
            add_notes(src, notes)
        tbl.append(src)
    h = sum(int(tr.get("h")) for tr in tbl.findall(A + "tr"))
    fr.find(".//" + P + "xfrm").find(A + "ext").set("cy", str(h))
    return fr, h


def head_height(tmpl_sh, lang):
    trs = tmpl_sh._element.find(".//" + A + "tbl").findall(A + "tr")
    return sum(int(tr.get("h")) for tr in trs[:head_rows(tmpl_sh._element, lang)])


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--results", required=True)
    ap.add_argument("--names", required=True, help="Mapping.xlsx (the state copy)")
    ap.add_argument("--prev", required=True, help="previous edition's day folder (yesterday's tables)")
    ap.add_argument("--en", required=True)
    ap.add_argument("--ar", required=True)
    ap.add_argument("--quarter", default="")
    ap.add_argument("--universe", default="")
    ap.add_argument("--ar-short", default="")
    ap.add_argument("--ar-font", default=os.path.join(os.path.dirname(os.path.abspath(__file__)), "fonts",
                                                      "ApercuArabic-Regular-Pro.otf"))
    ap.add_argument("--drop-cal", default="")
    ap.add_argument("--skip", default="", help="rows left out (after the 10:30 cut-off): names or Bloomberg codes, ';'-separated")
    ap.add_argument("--n1", type=int, default=None)
    ap.add_argument("--limit", type=float, default=10.60)
    ap.add_argument("--outdir", required=True)
    a = ap.parse_args()

    skip = {t.strip().upper() for t in a.skip.split(";") if t.strip()}
    rows = [x for x in read_rows(a.results, a.quarter) if not ({x["name"].upper(), x["bbg"].upper(), x["code"]} & skip)]
    rows = sort_rows(rows, covered_set(a.universe))
    if not rows:
        sys.exit("No result rows in the workbook.")
    nm = names_sheet(a.names)
    over = json.load(open(a.ar_short, encoding="utf-8")) if a.ar_short and os.path.exists(a.ar_short) else {}

    decks, frames = {}, {}
    for lang, path in (("EN", a.en), ("AR", a.ar)):
        decks[lang] = Presentation(path)
        y = Presentation(os.path.join(a.prev, f"ANB Daily Market Watch - {lang}.pptx"))
        t1, t2 = results_frame(y, 0), results_frame(y, 1) or results_frame(y, 0)
        if t1 is None:
            t1 = t2
        if t1 is None:
            sys.exit(f"No Results table in the previous edition's {lang} deck: pass --prev for the last results day.")
        frames[lang] = (t1, t2)

    # Arabic names: override, else long name, else short name when the long one wraps
    t2ar = frames["AR"][0]
    tbl = t2ar._element.find(".//" + A + "tbl")
    gc = tbl.find(A + "tblGrid").findall(A + "gridCol")
    name_w = int(gc[-1].get("w")) / E - 0.10   # AR name column is the last (right-most); minus cell margins
    wraps = []
    for x in rows:
        k = x["bbg"].upper()
        if k in over:
            x["ar"] = over[k]
            continue
        long_, short = nm.get(k, (None, None))
        x["ar"] = long_
        if long_ and text_width_in(long_, 7, a.ar_font) > name_w and short:
            x["ar"] = short
            wraps.append(f"{x['bbg']}: {long_} -> {short}")

    # slide 1 rows: most that fit above the limit in both decks
    if a.n1 is None:
        fit = []
        for lang in ("EN", "AR"):
            s1 = decks[lang].slides[0]
            h1 = hdr_shape(s1, CORP[lang])
            room = int(a.limit * E) - (h1.top + h1.height) - head_height(frames[lang][0], lang)
            fit.append(max(0, room // ROW_H))
        n1 = min(min(fit), len(rows))
    else:
        n1 = min(a.n1, len(rows))

    drops = [t.strip() for t in a.drop_cal.split(";") if t.strip()]
    os.makedirs(a.outdir, exist_ok=True)
    report = []
    for lang in ("EN", "AR"):
        p = decks[lang]
        t1, t2 = frames[lang]
        s1 = p.slides[0]
        h1 = hdr_shape(s1, CORP[lang])
        top1 = h1.top + h1.height
        for sh in list(s1.shapes):
            if sh.has_text_frame and not sh.text_frame.text.strip() and sh.top >= h1.top and sh.height > 0.5 * E:
                sh._element.getparent().remove(sh._element)
        last_on_1 = n1 == len(rows)
        if n1:
            notes = sorted({period_note(x["period"], lang) for x in rows if x["other"]}) if last_on_1 else ()
            fr, h = build(t1 if not last_on_1 else t2, rows[:n1], lang, last_on_1, notes)
            fr.find(".//" + P + "xfrm").find(A + "off").set("y", str(top1))
            s1.shapes._spTree.append(fr)
            report.append(f"{lang} slide 1: {n1} rows, bottom {(top1 + h) / E:.2f} in")
        s2 = p.slides[1]
        if not last_on_1:
            h2 = hdr_shape(s2, CORP[lang])
            top2 = h2.top + h2.height
            notes = sorted({period_note(x["period"], lang) for x in rows if x["other"]})
            fr2, hh = build(t2, rows[n1:], lang, True, notes)
            fr2.find(".//" + P + "xfrm").find(A + "off").set("y", str(top2))
            news = [s for s in walk(s2.shapes) if s.has_text_frame and len(s.text_frame.text) > 200]
            delta = top2 + hh - h2.top - h2.height
            if news:
                news[0].top = news[0].top + delta
            for sh in s2.shapes:
                if sh.has_table:
                    sh.top = sh.top + delta
                if sh.shape_type == 6:
                    g = sh._element.find(".//" + P + "grpSpPr").find(A + "xfrm")
                    kids = list(sh.shapes)
                    lo = min(k.top for k in kids)
                    hi = max(k.top + k.height for k in kids)
                    for t in ("off", "chOff"):
                        g.find(A + t).set("y", str(lo))
                    for t in ("ext", "chExt"):
                        g.find(A + t).set("cy", str(hi - lo))
            s2.shapes._spTree.append(fr2)
            report.append(f"{lang} slide 2: {len(rows) - n1} rows, {top2 / E:.2f} to {(top2 + hh) / E:.2f} in; "
                          f"news and calendars moved {delta / E:.2f} in")
        for sh in s2.shapes:
            if not sh.has_table:
                continue
            tb = sh.table._tbl
            cut = 0
            for tr in list(tb.findall(A + "tr")):
                txt = "".join(t.text or "" for t in tr.iter(A + "t"))
                if drops and any(k in txt for k in drops):
                    cut += int(tr.get("h"))
                    tb.remove(tr)
            if cut:
                sh.height = sh.height - cut
                report.append(f"{lang} calendar rows dropped ({cut / E:.2f} in)")
            for r in sh.table.rows:
                tx = [c.text for c in r.cells]
                if any(k in t for t in tx for k in CRUDE):
                    for c in r.cells:
                        t = c.text.strip()
                        if any(ch.isdigit() for ch in t) and (t[-1:] in "MKB%" or t.endswith(("M-", "K-", "B-", "%-"))):
                            settext(c._tc, "-")
        # the page-number box sits at 10.66 to 10.90 in: content only (5 Nov 2026 false 10.90 in warning)
        bottom = max((s.top + s.height) for s in walk(s2.shapes)
                     if s.top is not None and s.height is not None and s.top < a.limit * E)
        if bottom > a.limit * E:
            report.append(f"{lang} slide 2 ends at {bottom / E:.2f} in, past {a.limit:.2f} "
                          "(crude rows stay; move the lowest-ranked Corporate items out of both decks, or accept)")
        p.save(os.path.join(a.outdir, f"ANB Daily Market Watch - {lang}.pptx"))

    print("\n".join(report))
    print(f"{len(rows)} rows: " + ", ".join(("[C] " if x["covered"] else "") + x["name"] for x in rows))
    other = [x for x in rows if x["other"]]
    if other:
        print("Different-period rows (starred, last; footnote added under the source line): "
              + ", ".join(f"{x['name']} {x['period']}" for x in other))
    if wraps:
        print("AR short names used (long name wraps): " + "; ".join(wraps))
    miss = [x["bbg"] for x in rows if not x.get("ar")]
    if miss:
        print("Missing AR names (TK in deck): " + ", ".join(miss))


if __name__ == "__main__":
    main()
