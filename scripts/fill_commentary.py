#!/usr/bin/env python3
"""Fill the news cells of the v2 DMW decks (EN and AR) from the edition's news JSON.

Usage:
  python3 fill_commentary.py --news NEWS.json --en "<built EN deck>" --ar "<built AR deck>" \
      --outdir OUT [--results-day] [--force]

NEWS.json (written by the run from the edition's Claude Doc; see read_news_doc.py):
  {"markets":   {"en": [global, tasi], "ar": [global, tasi]},
   "economic":  [{"en": ..., "ar": ..., "short_en": ..., "short_ar": ...}, ...],   # in order
   "corporate": [{"en": ..., "ar": ...}, ...]}                                    # ranked
  short_en / short_ar are optional shorter versions of an Economic item, used only by the
  slide 1 overflow fallback.

v2 rules (Template v2, approved 25 Nov 2026):
  - s1.news: Markets (Global then TASI), Economic items, and ONE Corporate bullet (the top-ranked
    item). On a results day (--results-day) the slide 1 Corporate band becomes the "Results
    Announcements" band over the Results table and every Corporate item goes to slide 2.
  - s2.flow: at most 7 Corporate items at the top; the calendars follow in the same table, so they
    move down with the news. Items past 7 are left out of both decks and reported.
  - Nothing may end below 10.60 in on either slide (PAGE_BOTTOM).
  - Slide 1 overflow fallback: use the shorter version of the LAST Economic item; if still over,
    the shorter versions of both. Without the needed short text the script stops and says how many
    lines to cut (exit 3), so the run writes the short text and reruns. It pings only if the short
    texts still overflow.
  - Slide 2 overflow: the lowest-ranked Corporate items are left out of both decks, one at a time.
  - EN and AR always carry the same items (same split, same drops).
Heights come from news_text.cell_height_pt (PowerPoint-fitted model, HarfBuzz line breaks); the
row heights are written into the decks so the file is laid out before PowerPoint opens it.
Checks: placeholders left, a closing source tag on every bullet, EN vs AR numbers and tickers.
"""
import argparse, json, os, re, sys
from pptx import Presentation

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import news_text as N                              # noqa: E402
from dmw_common import PAGE_BOTTOM                 # noqa: E402

E = 914400
S2_MAX = 7
NEWS_ROWS = {"mk": 2, "ec": 6, "co": 10}           # s1.news text rows (0-based)
S2_ROWS = (0, 1, 2, 3)                              # s2.flow: Corporate header, gap, text, gap
log = []


def note(level, msg):
    log.append((level, msg))


def shape(slide, name):
    for sh in slide.shapes:
        if sh.name == name:
            return sh
    sys.exit(f"Shape '{name}' not found: is this a Template v2 deck?")


def news_avail(sh, r):
    t = sh.table
    span = int(t.cell(r, 0)._tc.get("gridSpan", 1))
    w = sum(t.columns[c].width for c in range(span))
    return N.text_avail_in(w, t.cell(r, 0)._tc)


def plan(prs, lang, content):
    """Heights (in) of every row of s1.news and s2.flow for this content, and the bottoms."""
    s1, s2 = shape(prs.slides[0], "s1.news"), shape(prs.slides[1], "s2.flow")
    h1 = [r.height / E for r in s1.table.rows]
    for k, r in NEWS_ROWS.items():
        h1[r] = max(N.cell_height_pt(content[k], lang, news_avail(s1, r)) / 72, 0.0)
    h2 = [r.height / E for r in s2.table.rows]
    if content["s2"]:
        h2[2] = N.cell_height_pt(content["s2"], lang, news_avail(s2, 2)) / 72
    else:
        for r in S2_ROWS:
            h2[r] = 0.0                      # the rows are removed
    return dict(h1=h1, h2=h2, b1=s1.top / E + sum(h1), b2=s2.top / E + sum(h2))


def content_for(lang, news, ec_short, s1_items, s2_items):
    ec = []
    for i, it in enumerate(news["economic"]):
        ec.append(it.get(f"short_{lang}") if i in ec_short else it[lang])
    return {"mk": news["markets"][lang], "ec": ec,
            "co": [it[lang] for it in s1_items], "s2": [it[lang] for it in s2_items]}


def apply(prs, lang, content, L):
    s1, s2 = shape(prs.slides[0], "s1.news"), shape(prs.slides[1], "s2.flow")
    for k, r in NEWS_ROWS.items():
        N.fill_cell(s1.table.cell(r, 0)._tc, content[k], lang)
    for r, h in enumerate(L["h1"]):
        s1.table.rows[r].height = int(round(h * E))
    s1.height = sum(r.height for r in s1.table.rows)
    if content["s2"]:
        N.fill_cell(s2.table.cell(2, 0)._tc, content["s2"], lang)
        s2.table.rows[2].height = int(round(L["h2"][2] * E))
    else:
        tbl = s2.table._tbl
        trs = tbl.findall(f"{{{N.A}}}tr")
        for r in S2_ROWS:
            tbl.remove(trs[r])
    s2.height = sum(r.height for r in s2.table.rows)


def leftover_placeholders(prs, lang):
    out = []
    for si, s in enumerate(list(prs.slides)[:3], 1):
        for sh in s.shapes:
            if sh.has_table:
                for row in sh.table.rows:
                    for c in row.cells:
                        if N.PLACEHOLDER.search(c.text):
                            out.append(f"slide {si} {sh.name}: {c.text.strip()[:50]}")
    return out


class Results:
    """Results day: rows from DMW - Results.xlsx (v1 rules) and the v2 table (results_v2.py)."""
    def __init__(self, a):
        import openpyxl
        import place_results as PR
        import results_v2 as RV
        from build_dmw import text_width_in
        self.RV = RV
        if not a.quarter:
            sys.exit("--quarter is needed with --results (e.g. 3Q26).")
        rows = PR.sort_rows(PR.read_rows(a.results, a.quarter), PR.covered_set(a.universe))
        skip = {x.strip().upper() for x in a.skip.split(";") if x.strip()}
        self.rows = [x for x in rows if not ({x["name"].rstrip("*").upper(), x["code"], x["bbg"].upper()} & skip)]
        others = sorted({x["period"] for x in self.rows if x["other"]})
        self.notes = {"en": [PR.period_note(p, "EN") for p in others], "ar": [PR.period_note(p, "AR") for p in others]}
        self.quarter = a.quarter
        self.names = {}
        if a.mapping:
            ws = openpyxl.load_workbook(a.mapping, read_only=True)["mapping"]
            it = ws.iter_rows(values_only=True)
            hdr = [str(c).strip() if c else "" for c in next(it)]
            ix = {h: i for i, h in enumerate(hdr)}
            for r in it:
                if r[ix["Code"]] is not None:
                    self.names[str(r[ix["Code"]]).strip()] = {
                        "en": [str(r[ix[k]]).strip() for k in ("Long Name - English", "Short Name - English") if r[ix[k]]],
                        "ar": [str(r[ix[k]]).strip() for k in ("Long Name - Arabic", "Short Name - Arabic") if r[ix[k]]]}
        room = RV.WIDTHS[0] - 2 * 0.02 - 0.01

        def chooser(lang):
            def f(x):
                star = "*" if x["other"] else ""
                cands = self.names.get(x["code"], {}).get(lang) or ([x["name"].rstrip("*")] if lang == "en" else [])
                if not cands:
                    note("FLAG", f"Results: no Arabic name in Mapping.xlsx for {x['code']} {x['name']}; the English name used.")
                    cands = [x["name"].rstrip("*")]
                for n in cands:
                    w = text_width_in(n + star, 7.0, RV.NAME_FONT[lang], rtl=(lang == "ar"))
                    if w is not None and w <= room:
                        return n + star
                note("FLAG", f"Results {lang.upper()}: '{cands[-1]}' wraps in the name column ({room:.2f} in).")
                return cands[-1] + star
            return f
        self.name_for = {"en": chooser("en"), "ar": chooser("ar")}
        note("OK", f"Results: {len(self.rows)} rows from {os.path.basename(a.results)} ({a.quarter}); "
                   f"{len(others)} other-period footnote(s).")

    def h(self, k, last, lang):
        return self.RV.height_in(k, True, last, len(self.notes[lang]) if last else 0)


_S2 = {}


def s2_protos(prs):
    """Copies of s2.flow rows 0-3 (Corporate band, gap, text, gap) as built, taken before apply()
    can remove them on a day with no slide 2 Corporate news."""
    import copy
    if id(prs) not in _S2:
        trs = shape(prs.slides[1], "s2.flow").table._tbl.findall(f"{{{N.A}}}tr")
        _S2[id(prs)] = [copy.deepcopy(tr) for tr in trs[0:4]]
    return _S2[id(prs)]


def lead_rows(prs, lang=None, res=None, cont=True):
    """Copies of the s2.flow band + gap rows (the Results band on slide 2) and the gap under the news.
    The s2.flow Corporate band stays where it is: Results is its own section (25 Nov 2026)."""
    import copy
    trs = s2_protos(prs)
    # the band only: the table's light header sits right under the navy band, as in the calendars
    # (Saad, 25 Nov 2026: no white gap between them)
    lead = [copy.deepcopy(tr) for tr in trs[0:1]]
    if res is not None:
        res.RV.set_row_text(lead[0], res.RV.TITLE[lang] + (res.RV.CONT[lang] if cont else ""))
    return lead, trs[3:4]


def with_results(prs, lang, L, res, n1, has_news):
    """Slide bottoms once the Results table is in: n1 rows on slide 1, the rest on slide 2."""
    n = len(res.rows)
    # slide 1: the gap row under the band goes (no white gap above the table's header)
    b1 = L["b1"] - L["h1"][NEWS_ROWS["co"] - 1] + res.h(n1, n1 == n, lang) if n1 else L["b1"]
    b2 = L["b2"]
    if n1 < n:
        lead, tail = lead_rows(prs)
        lead_h = sum(int(tr.get("h")) for tr in lead) / E
        tail_h = sum(int(tr.get("h")) for tr in tail) / E
        b2 = L["b2"] + lead_h + res.h(n - n1, True, lang) + tail_h
    return b1, b2


def place_results(prs, lang, res, n1):
    """Insert the Results tables: under the slide 1 Corporate header and at the top of slide 2."""
    from pptx.oxml.ns import qn
    n = len(res.rows)
    RV = res.RV
    s1, s2 = shape(prs.slides[0], "s1.news"), shape(prs.slides[1], "s2.flow")
    # slide 1: the empty Corporate cell row goes; the table sits under the header and gap rows,
    # and the header band reads "Results Announcements" (its own section, 25 Nov 2026)
    tbl = s1.table._tbl
    trs1 = tbl.findall(f"{{{N.A}}}tr")
    tbl.remove(trs1[NEWS_ROWS["co"]])
    if n1:
        RV.set_row_text(trs1[NEWS_ROWS["co"] - 2], RV.TITLE[lang])
        tbl.remove(trs1[NEWS_ROWS["co"] - 1])      # no white gap between the band and the table
    s1.height = sum(r.height for r in s1.table.rows)
    if n1:
        fr, h = RV.make_frame(s2, lang, res.rows[:n1], res.quarter, True, n1 == n,
                              res.notes[lang] if n1 == n else [], res.name_for[lang])
        _place(fr, s1, s1.left, s1.top + s1.height, "s1.results", 9001)
    if n1 < n:
        lead, tail = lead_rows(prs, lang, res, cont=n1 > 0)
        fr, h = RV.make_frame(s2, lang, res.rows[n1:], res.quarter, True, True, res.notes[lang], res.name_for[lang],
                              lead=lead, tail=tail)
        top = s2.top
        _place(fr, s2, s2.left, top, "s2.results", 9002)
        s2.top = top + h


def _place(fr, anchor, left, top, name, sid):
    P = "http://schemas.openxmlformats.org/presentationml/2006/main"
    nv = fr.find(f".//{{{P}}}cNvPr")
    nv.set("name", name); nv.set("id", str(sid))
    off = fr.find(f".//{{{P}}}xfrm/{{{N.A}}}off")
    off.set("x", str(left)); off.set("y", str(top))
    anchor._element.addnext(fr)


def checks(news):
    items = [("markets", i, news["markets"]["en"][i], news["markets"]["ar"][i]) for i in range(2)]
    items += [("economic", i, x["en"], x["ar"]) for i, x in enumerate(news["economic"])]
    items += [("corporate", i, x["en"], x["ar"]) for i, x in enumerate(news["corporate"])]
    for sec, i, en, ar in items:
        for lang, t in (("en", en), ("ar", ar)):
            if not re.search(r"\)\s*\.?$", t.strip()):
                note("FLAG", f"{lang} {sec} {i + 1}: no closing source tag: {t[:60]}")
        if sec != "markets":
            ne, na = N.numbers(en, "en"), N.numbers(ar, "ar")
            if ne != na:
                note("CONFLICT", f"{sec} {i + 1}: EN numbers {ne} vs AR {na}: {en[:50]}")
            te, ta = N.tickers(en), N.tickers(ar)
            if ta and te != ta:
                note("CONFLICT", f"{sec} {i + 1}: EN tickers {te} vs AR {ta}")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--news", required=True)
    ap.add_argument("--en", required=True)
    ap.add_argument("--ar", required=True)
    ap.add_argument("--outdir", required=True)
    ap.add_argument("--results-day", action="store_true",
                    help="leave the slide 1 Corporate cell for the Results table")
    ap.add_argument("--limit", type=float, default=PAGE_BOTTOM)
    ap.add_argument("--results", help="DMW - Results.xlsx (results day: v2 Results table, implies --results-day)")
    ap.add_argument("--quarter", help="reporting quarter, e.g. 3Q26 (with --results)")
    ap.add_argument("--universe", help="ANB coverage universe workbook (covered companies first)")
    ap.add_argument("--mapping", help="Mapping.xlsx for the Results company names (with --results)")
    ap.add_argument("--skip", default="", help="Results rows left out (';'-separated names or codes)")
    a = ap.parse_args()
    news = json.load(open(a.news, encoding="utf-8"))
    for lang in ("en", "ar"):
        if len(news["markets"][lang]) != 2:
            sys.exit(f"markets.{lang} needs the Global and the TASI bullet (2 items).")
    if len(news["economic"]) != 2:
        note("FLAG", f"{len(news['economic'])} Economic items (the house runs 2).")
    checks(news)

    decks = {"en": Presentation(a.en), "ar": Presentation(a.ar)}
    res = None
    if a.results:
        a.results_day = True
        res = Results(a)
    corp = list(news["corporate"])
    k1 = 0 if a.results_day else min(1, len(corp))
    s1_items, rest = corp[:k1], corp[k1:]
    if len(rest) > S2_MAX:
        for it in rest[S2_MAX:]:
            note("FLAG", f"Corporate item left out (slide 2 holds {S2_MAX}): {it['en'][:70]}")
        rest = rest[:S2_MAX]

    # slide 1: overflow fallback on the Economic items
    n_ec = len(news["economic"])
    tries = [set(), {n_ec - 1}, set(range(n_ec))] if n_ec else [set()]
    for ec_short in tries:
        missing = [i for i in ec_short for lang in ("en", "ar") if not news["economic"][i].get(f"short_{lang}")]
        if missing:
            L = {lang: plan(decks[lang], lang, content_for(lang, news, set(), s1_items, rest)) for lang in ("en", "ar")}
            over = {lang: L[lang]["b1"] - a.limit for lang in L if L[lang]["b1"] > a.limit}
            lines = {lang: int(-(-over[lang] * 72 // N.PT_LINE)) for lang in over}
            for k, m in log:
                print(f"{k}: {m}")
            print(f"SHORTEN: slide 1 ends {', '.join(f'{k.upper()} {L[k]['b1']:.2f} in' for k in L)} "
                  f"(limit {a.limit:.2f}). Write short_en and short_ar for Economic item {max(missing) + 1} "
                  f"(cut at least {', '.join(f'{v} line(s) {k.upper()}' for k, v in lines.items())}) and rerun.")
            sys.exit(3)
        C = {lang: content_for(lang, news, ec_short, s1_items, rest) for lang in ("en", "ar")}
        L = {lang: plan(decks[lang], lang, C[lang]) for lang in ("en", "ar")}
        if all(L[lang]["b1"] <= a.limit for lang in L):
            if ec_short:
                note("OK", f"Slide 1 overflow: short text used for Economic item(s) {sorted(i + 1 for i in ec_short)}.")
            break
    else:
        note("PING", "Slide 1 still ends below the limit with the short Economic texts: "
             + ", ".join(f"{k.upper()} {L[k]['b1']:.2f} in" for k in L))

    # results day: the most Results rows on slide 1 that fit in BOTH decks, the rest on slide 2
    n1 = 0
    if res:
        n = len(res.rows)
        n1 = max([k for k in range(n + 1)
                  if all(with_results(decks[g], g, L[g], res, k, bool(rest))[0] <= a.limit for g in L)] or [0])
        if n1 == 0 and n:
            note("FLAG", "Results: no row fits on slide 1; the whole table goes to slide 2.")
        for g in L:
            L[g]["b1"], L[g]["b2"] = with_results(decks[g], g, L[g], res, n1, bool(rest))
        note("OK", f"Results: {n1} row(s) on slide 1, {n - n1} on slide 2.")

    # slide 2: drop lowest-ranked Corporate items until both decks fit
    while rest and any(L[lang]["b2"] > a.limit for lang in L):
        gone = rest.pop()
        note("FLAG", f"Slide 2 full: Corporate item left out of both decks: {gone['en'][:70]}")
        C = {lang: content_for(lang, news, ec_short, s1_items, rest) for lang in ("en", "ar")}
        L = {lang: plan(decks[lang], lang, C[lang]) for lang in ("en", "ar")}
        if res:
            for g in L:
                L[g]["b1"], L[g]["b2"] = with_results(decks[g], g, L[g], res, n1, bool(rest))
    if any(L[lang]["b2"] > a.limit for lang in L):
        note("PING", "Slide 2 calendars alone end below the limit: "
             + ", ".join(f"{k.upper()} {L[k]['b2']:.2f} in" for k in L))

    os.makedirs(a.outdir, exist_ok=True)
    for lang in ("en", "ar"):
        s2_protos(decks[lang])
        apply(decks[lang], lang, C[lang], L[lang])
        if res:
            place_results(decks[lang], lang, res, n1)
        left = leftover_placeholders(decks[lang], lang)
        for x in left:
            if "Results" not in x:
                note("FLAG", f"{lang.upper()} placeholder left: {x}")
        out = os.path.join(a.outdir, f"ANB Daily Market Watch - {lang.upper()}.pptx")
        decks[lang].save(out)
        note("OK", f"{lang.upper()}: slide 1 news ends {L[lang]['b1']:.2f} in, slide 2 ends {L[lang]['b2']:.2f} in "
                   f"(limit {a.limit:.2f}); {len(s1_items)} Corporate on slide 1, {len(rest)} on slide 2. Saved {out}")
    for k, m in log:
        print(f"{k}: {m}")
    json.dump({"s1_corporate": len(s1_items), "s2_corporate": len(rest),
               "bottoms": {k: [round(L[k]["b1"], 3), round(L[k]["b2"], 3)] for k in L},
               "log": log}, open(os.path.join(a.outdir, "fill_commentary.json"), "w"), ensure_ascii=False, indent=1)
    return 2 if any(k == "PING" for k, _ in log) else 0


if __name__ == "__main__":
    sys.exit(main())
