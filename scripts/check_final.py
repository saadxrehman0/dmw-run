#!/usr/bin/env python3
"""Final check of the v2 DMW decks (EN and AR) before the PDFs. Read only: never edits a deck.

Usage:
  python3 check_final.py --en EN.pptx --ar AR.pptx --calendar "Manual File-Calendar.xlsx" --date YYYY-MM-DD \
      --outdir OUT [--results "DMW - Results.xlsx" --quarter 3Q26] [--layout ppt_layout.csv]

Checks (v2 port of the v1 checks, 25 Nov 2026):
  1. EN/AR parity of every news paragraph (same count per section; numbers and tickers match).
  2. Calendars as in the calendar file: same rows, dates, Prior and Estimate in both decks, and the
     EN event wording (catches hand edits after the build, e.g. 24 Nov 2026 estimate 92.1 -> 92.4).
  3. Page bottom 10.60 in on slides 1-2: from PowerPoint's own layout when --layout is given
     (export_measure.ps1), else from the row heights in the file (the fill's fitted model).
  4. Results tables: every row's figures as recomputed from DMW - Results.xlsx, EN and AR alike.
  5. Technicals slides 4-5: the same levels in EN and AR (v1 check, unchanged).
  6. No placeholder text or yellow highlight left on slides 1-3.
Prints OK / FLAG lines and writes OUT/final_check.md. Exit 1 if anything is flagged.
"""
import argparse, csv, os, re, sys
from datetime import datetime
from pptx import Presentation
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import dmw_common as C                                  # noqa: E402
import news_text as N                                    # noqa: E402
from legacy_check_final_v1 import check_technicals, note as _v1note, out as v1log   # noqa: E402

E = 914400
log = v1log          # one list, so the v1 technicals check reports here too
note = _v1note
SECS = ("markets", "economic", "corporate")
HEADS = {"markets performance & news": "markets", "major economic/sector news": "economic",
         "major corporate news": "corporate", "أداء الأسواق/الأخبار": "markets",
         "أخبار اقتصادية/قطاعية رئيسية": "economic", "أهم أخبار الشركات": "corporate"}
RESULTS_HEADS = {"results announcements", "result announcements", "إعلانات النتائج"}
CAL = {"event calendar": "ev", "economic calendar": "ec", "تقويم الأحداث": "ev", "التقويم الاقتصادي": "ec"}
AI = str.maketrans("٠١٢٣٤٥٦٧٨٩٫٬", "0123456789.,")


def tables(prs, si):
    return sorted([s for s in prs.slides[si].shapes if s.has_table], key=lambda s: s.top)


def read(path):
    """News paragraphs per section, calendar rows (logical EN column order) and results rows."""
    prs = Presentation(path)
    ar = any("ar" in (r._r.rPr.get("lang") or "").lower() for sh in prs.slides[0].shapes if sh.name == "s1.news"
             for row in sh.table.rows for c in row.cells for p in c.text_frame.paragraphs for r in p.runs
             if r._r.rPr is not None)
    news = {k: [] for k in SECS}
    cal = {"ev": [], "ec": []}
    res = []
    for si in (0, 1):
        for sh in tables(prs, si):
            rows = [[c.text.strip() for c in r.cells] for r in sh.table.rows]
            if ar:
                rows = [r[::-1] for r in rows]
            if sh.name in ("s1.news", "s2.flow"):
                sec = "corporate" if sh.name == "s2.flow" else None
                for ri, r in enumerate(sh.table.rows):
                    first = " ".join(x for x in rows[ri] if x).strip()   # a full-width cell, either deck
                    key = first.lower()
                    if key in HEADS or first in HEADS:
                        sec = HEADS.get(key) or HEADS[first]; continue
                    if key in CAL or first in CAL:
                        sec = CAL.get(key) or CAL[first]; continue
                    if key in RESULTS_HEADS or first in RESULTS_HEADS:   # results days (5 Nov replay)
                        sec = None; continue
                    if sec in news:
                        cell = r.cells[0]
                        news[sec] += [p.text.strip() for p in cell.text_frame.paragraphs if p.text.strip()]
                    elif sec in cal and first and not first.startswith(("Source", "المصدر", "Date", "التاريخ")):
                        cal[sec].append(rows[ri])
            elif sh.name in ("s1.results", "s2.results"):
                res += [r for r in rows if re.fullmatch(r"\d{4}", r[1].translate(AI))]
    return prs, news, cal, res


def check_parity(en, ar):
    for sec in SECS:
        e, a = en[sec], ar[sec]
        if len(e) != len(a):
            note("FLAG", f"{sec}: {len(e)} EN paragraphs vs {len(a)} AR.")
        for i, (x, y) in enumerate(zip(e, a), 1):
            if sec != "markets" and N.numbers(x, "en") != N.numbers(y, "ar"):
                note("FLAG", f"{sec} {i}: EN numbers {N.numbers(x, 'en')} vs AR {N.numbers(y, 'ar')}: {x[:50]}")
            te, ta = N.tickers(x), N.tickers(y)
            if ta and te != ta:
                note("FLAG", f"{sec} {i}: EN tickers {te} vs AR {ta}.")
    note("OK", "News parity checked: " + ", ".join(f"{s} {len(en[s])}" for s in SECS) + " paragraphs.")


def check_calendar(cal_en, cal_ar, path, d):
    events, econ = C.load_calendar(path, d)
    fmt = lambda v: "" if v is None else str(v).strip()
    want_ec = [(f"{x[0]:%d-%b-%y}", fmt(x[2]), fmt(x[3])) for x in econ]
    for lang, cal in (("EN", cal_en), ("AR", cal_ar)):
        if len(cal["ev"]) != len(events):
            note("FLAG", f"{lang} Event calendar: {len(cal['ev'])} rows vs {len(events)} in the calendar file.")
        if len(cal["ec"]) != len(econ):
            note("FLAG", f"{lang} Economic calendar: {len(cal['ec'])} rows vs {len(econ)} in the calendar file.")
        for r, (dt, p, s) in zip(cal["ec"], want_ec):
            got = (r[3].translate(AI), r[4].translate(AI))
            if got != (p, s):
                note("FLAG", f"{lang} Economic calendar {dt} '{r[1][:40]}': deck Prior/Estimate {got} vs file {(p, s)}.")
    for r, x in zip(cal_en["ec"], econ):
        if r[1] != str(x[1]).strip():
            note("FLAG", f"EN Economic calendar: deck '{r[1]}' vs file '{x[1]}'.")
    for r, x in zip(cal_en["ev"], events):
        if r[2] != str(x[2]).strip():
            note("FLAG", f"EN Event calendar: deck '{r[2]}' vs file '{x[2]}'.")
    note("OK", f"Calendars checked against the file: {len(events)} events, {len(econ)} economic releases.")


def bottoms_from_file(prs):
    out = {}
    for si in (0, 1):
        b = 0
        for sh in prs.slides[si].shapes:
            if sh.has_table and not sh.name.endswith(".page"):
                b = max(b, (sh.top + sum(r.height for r in sh.table.rows)) / E)
        out[si + 1] = b
    return out


def bottoms_from_layout(layout, fname):
    out = {}
    for r in csv.DictReader(open(layout, encoding="utf-8-sig")):
        if os.path.basename(r["file"]) == fname or r["file"].endswith(fname):
            s = int(r["slide"])
            if s in (1, 2) and not r["shape"].endswith(".page"):
                out[s] = max(out.get(s, 0), float(r["bottom_in"]))
    return out


def check_results(res_rows, path, quarter, lang):
    import place_results as PR
    want = {x["bbg"]: x for x in PR.read_rows(path, quarter)}
    bad = 0
    for r in res_rows:
        r = [c.translate(AI) for c in r]
        x = want.get(r[2])
        if not x:
            note("FLAG", f"{lang} Results: {r[2]} is not in the Results file."); bad += 1; continue
        exp = [PR.f1(x[k]) for k in ("cur", "pri", "chg", "pct", "cons", "var")]
        if r[3:9] != exp:
            note("FLAG", f"{lang} Results {r[2]}: deck {r[3:9]} vs file {exp}."); bad += 1
    note("OK" if not bad else "FLAG", f"{lang} Results: {len(res_rows)} rows checked against the file, {bad} off.")


def leftovers(prs, lang):
    for si in range(3):
        for sh in prs.slides[si].shapes:
            if not sh.has_table:
                continue
            for row in sh.table.rows:
                for c in row.cells:
                    if N.PLACEHOLDER.search(c.text):
                        note("FLAG", f"{lang} slide {si + 1} {sh.name}: placeholder left '{c.text.strip()[:40]}'.")
                    # a TK or a bracketed note reached the PDFs and the email (5 Nov replay)
                    if re.search(r"(?<![A-Za-z])TK(?![A-Za-z])", c.text):
                        note("FLAG", f"{lang} slide {si + 1} {sh.name}: TK left '{c.text.strip()[:40]}'.")
                    if re.search(r"\[[^\]]{12,}\]", c.text) and not N.PLACEHOLDER.search(c.text):
                        note("FLAG", f"{lang} slide {si + 1} {sh.name}: bracketed note left '{c.text.strip()[:40]}'.")
                    if c._tc.find(f".//{{{N.A}}}highlight") is not None and c.text.strip():
                        note("FLAG", f"{lang} slide {si + 1} {sh.name}: yellow highlight left on '{c.text.strip()[:40]}'.")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--en", required=True); ap.add_argument("--ar", required=True)
    ap.add_argument("--calendar"); ap.add_argument("--date")
    ap.add_argument("--results"); ap.add_argument("--quarter")
    ap.add_argument("--layout", help="ppt_layout.csv from export_measure.ps1")
    ap.add_argument("--limit", type=float, default=C.PAGE_BOTTOM)
    ap.add_argument("--outdir", required=True)
    a = ap.parse_args()
    pe, ne, ce, re_ = read(a.en)
    pa, na, ca, ra = read(a.ar)
    check_parity(ne, na)
    if a.calendar and a.date:
        check_calendar(ce, ca, a.calendar, datetime.strptime(a.date, "%Y-%m-%d").date())
    for lang, prs, path in (("EN", pe, a.en), ("AR", pa, a.ar)):
        if a.layout:
            b, src = bottoms_from_layout(a.layout, os.path.basename(path)), "PowerPoint layout"
            if not b:
                b, src = bottoms_from_file(prs), "file row heights (deck not in the layout file)"
        else:
            b, src = bottoms_from_file(prs), "file row heights"
        for s, v in sorted(b.items()):
            note("OK" if v <= a.limit + 0.005 else "FLAG",
                 f"{lang} slide {s} ends {v:.2f} in (limit {a.limit:.2f}; {src}).")
        leftovers(prs, lang)
    if a.results:
        check_results(re_, a.results, a.quarter, "EN")
        check_results(ra, a.results, a.quarter, "AR")
    check_technicals(a.en, a.ar)
    os.makedirs(a.outdir, exist_ok=True)
    lines = [f"- {k}: {m}" for k, m in log]
    open(os.path.join(a.outdir, "final_check.md"), "w", encoding="utf-8").write("# DMW final check\n\n" + "\n".join(lines) + "\n")
    print("\n".join(lines))
    return 1 if any(k == "FLAG" for k, _ in log) else 0


if __name__ == "__main__":
    sys.exit(main())
