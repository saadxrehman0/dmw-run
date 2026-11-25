#!/usr/bin/env python3
"""Pre-PDF check of the team's saved DMW decks (EN and AR), run before the PDFs are exported.

Usage:
  python3 check_final.py --en "ANB Daily Market Watch - EN.pptx" --ar "ANB Daily Market Watch - AR.pptx" \
      --calendar "Manual File-Calendar.xlsx" --outdir OUT [--no-render]

Catches the faults that reached published decks (6 Oct 2026, plus 4 added 13 Oct 2026):
  1. EN/AR figures that disagree in the commentary (Knight Frank: USD 2.1 tn in EN, 2.07 in AR).
     Paragraphs are paired in deck order (Markets, Economic, Corporate slide 1, Corporate slide 2);
     every number and Bloomberg ticker must match.
  2. Calendar rows edited by hand away from the calendar file, in either deck
     (AR GDPNow row trimmed to drop "GDPNow (3Q26)").
  3. Header gaps off the 0.10 in house gap on slide 1 (AR Economic and Corporate headers at 0.03 in),
     and content running past the 10.60 in page bottom on slide 1 or slide 2 (slide 2 matters most
     on results days, when all Corporate news moves there).
     Measured on a LibreOffice render of the saved deck, not the PDF, so it works before export.
  4. Technicals slides 4-5 whose figures differ between EN and AR (13 Oct 2026: intraday high
     10,356.47 vs 10,366.47).
Prints OK / FLAG lines and writes them to OUT/final_check.md. Read only: never edits a deck.
"""
import argparse, os, re, sys, tempfile
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import numpy as np
from PIL import Image
from pptx import Presentation
from datetime import datetime
import build_dmw as en
import build_dmw_ar as ar
from fill_commentary import find_shapes, render, numbers
from build_dmw import SECTION_GAP, PAGE_BOTTOM

EMU = 914400
# Rows are scanned down to here when looking for the last ink on a page. It must sit below
# PAGE_BOTTOM plus the tolerance, or an overrun can never be seen (until 13 Oct 2026 the scan
# stopped at 10.62 in, so the bottom check always passed), and above the page number, whose
# digits start at about 10.74 in.
SCAN_BOTTOM = 10.70
BOTTOM_TOL = 0.02
out = []


def note(level, msg):
    out.append((level, msg))


def tick(s):
    return sorted(re.findall(r"(?<![A-Za-z0-9])([A-Z][A-Z0-9]+ AB)(?![A-Za-z0-9])", s))


def commentary(path, lang):
    sh = find_shapes(Presentation(path), lang, strict=False)
    paras = []
    for k in ("mk", "ec", "co", "s2"):
        if sh[k] is None:
            note("INFO", f"{lang.upper()}: no {k} commentary box (results table or no header); skipped.")
            continue
        paras += [(k, p.text.strip()) for p in sh[k].text_frame.paragraphs if p.text.strip()]
    return paras


def check_parity(en_path, ar_path):
    e, a = commentary(en_path, "en"), commentary(ar_path, "ar")
    if len(e) != len(a):
        note("FLAG", f"Commentary: EN has {len(e)} paragraphs, AR {len(a)}. Pairing below may be off.")
    bad = 0
    for i, ((k, te), (_, ta)) in enumerate(zip(e, a)):
        ne, na = numbers(te, "en"), numbers(ta, "ar")
        if ne != na or tick(te) != tick(ta):
            bad += 1
            only_e = sorted(set(ne) - set(na)); only_a = sorted(set(na) - set(ne))
            note("FLAG", f"EN/AR mismatch, paragraph {i + 1} ({k}): \"{te[:60]}...\" "
                         f"EN-only {only_e or '-'} AR-only {only_a or '-'}"
                         + (f"; tickers EN {tick(te)} AR {tick(ta)}" if tick(te) != tick(ta) else ""))
    if not bad:
        note("OK", f"Commentary: {min(len(e), len(a))} EN/AR paragraph pairs, numbers and tickers match.")


def norm(s):
    # bidi marks (RLM, LRM, embeddings) are invisible: a trailing RLM made a matching AR calendar
    # row FLAG on 20 Oct 2026
    return re.sub(r"\s+", " ", re.sub(r"[​-‏‪-‮⁦-⁩﻿]", "", s or "")).strip()


def check_calendar(en_path, ar_path, cal, report_date):
    n0 = len(en.checks)
    events, econ = en.load_calendar(cal, report_date)
    for l, m in en.checks[n0:]:
        if l != "OK":
            note(l, f"Calendar reader: {m}")
    ev_ar, ec_ar = ar.load_calendar_ar(cal)
    for lang, path in (("EN", en_path), ("AR", ar_path)):
        tables = [s for s in Presentation(path).slides[1].shapes if s.has_table]
        cells = [norm(c.text) for t in tables for r in t.table.rows for c in r.cells]
        want = ([norm(e) for _, e, *_ in econ] + [norm(v) for _, _, *pe in econ for v in pe if norm(v) != "-"]
                + [norm(c) for _, c, _ in events] + [norm(x) for _, _, x in events]) \
            if lang == "EN" else ([norm(x) for x in ec_ar] + [norm(x) for pair in ev_ar for x in pair])
        missing = [w for w in want if w and w not in cells]
        for w in missing:
            near = [c for c in cells if c and (c in w or w in c)]
            note("FLAG", f"{lang} calendar: file wording \"{w}\" not in the deck"
                         + (f" (deck has \"{near[0]}\")" if near else "") + ".")
        if not missing:
            note("OK", f"{lang} calendar: every row's wording matches the calendar file.")


def gaps(path, lang, outdir):
    """Visible gap above the slide 1 Economic and Corporate headers, and the bottom of slides 1 and 2."""
    prs = Presentation(path)
    sh = find_shapes(prs, lang, strict=False)
    d = tempfile.mkdtemp(dir=outdir)
    pngs = render(path, d, "c")
    a = np.asarray(Image.open(os.path.join(d, pngs[0])).convert("L")).astype(int)[:, 45:805]
    n = (a < 200).sum(axis=1)
    ink = (n >= 2) & (n < 0.6 * a.shape[1])
    for name, hd in (("Economic", sh["ec_h"]), ("Corporate", sh["co_h"])):
        if hd is None:
            note("FLAG", f"{lang} slide 1: no '{name}' header found."); continue
        y = hd.top / EMU
        i = int(y * 100) - 1
        while i > 0 and not ink[i]:
            i -= 1
        g = y - (i + 1) / 100
        lvl = "OK" if abs(g - SECTION_GAP) <= 0.02 else "FLAG"
        note(lvl, f"{lang} slide 1: gap above the {name} header {g:.2f} in (house gap {SECTION_GAP:.2f}).")
    note(*bottom(ink, lang, 1))
    if len(pngs) > 1:
        a2 = np.asarray(Image.open(os.path.join(d, pngs[1])).convert("L")).astype(int)[:, 45:805]
        n2 = (a2 < 200).sum(axis=1)
        note(*bottom((n2 >= 2) & (n2 < 0.6 * a2.shape[1]), lang, 2))


def bottom(ink, lang, slide):
    """(level, message) for where the content of a slide ends against PAGE_BOTTOM."""
    ys = np.where(ink[: int(round(SCAN_BOTTOM * 100))])[0]
    last = (ys[-1] + 1) / 100 if len(ys) else 0
    over = last > PAGE_BOTTOM + BOTTOM_TOL
    return ("FLAG" if over else "OK",
            f"{lang} slide {slide}: content ends at {last:.2f} in (limit {PAGE_BOTTOM:.2f})"
            + (": it runs into the page-number band." if over else "."))


def check_technicals(en_path, ar_path):
    """Slides 4-5 (Technicals): every level must appear in both decks (13 Oct 2026: EN intraday
    high 10,356.47 vs AR 10,366.47). Compares figures with a
    decimal point or a thousands comma, sign and range dashes ignored, and flags a change in
    points that is not the AR/source figure rounded half-up to one decimal."""
    from decimal import Decimal, ROUND_HALF_UP
    def figs(path):
        p = Presentation(path); txt = []
        for s in list(p.slides)[3:5]:
            for sh in s.shapes:
                if sh.has_table:
                    txt += [c.text for r in sh.table.rows for c in r.cells]
                elif sh.has_text_frame:
                    txt.append(sh.text_frame.text)
        t = " ".join(txt)
        return {m.rstrip(".") for m in re.findall(r"\d{1,3}(?:,\d{3})+(?:\.\d+)?|\d+\.\d+", t)}, t
    (fe, te), (fa, ta) = figs(en_path), figs(ar_path)
    fe -= {"1,000"}; fa -= {"1,000"}   # chart unit labels such as "Volume (1,000)"
    num = lambda x: Decimal(x.replace(",", ""))
    only_e = sorted(x for x in fe - fa if num(x) not in {num(y) for y in fa})
    only_a = sorted(x for x in fa - fe if num(x) not in {num(y) for y in fe})
    # a 1dp EN figure that is the AR 2dp figure rounded half-up is a rounding check, not a mismatch
    rounding = []
    for x in list(only_e):
        if "." in x and len(x.split(".")[1]) == 1:
            near = [y for y in only_a if "." in y and len(y.split(".")[1]) == 2
                    and abs(num(y) - num(x)) < Decimal("0.1")]
            for y in near:
                want = num(y).quantize(Decimal("0.1"), ROUND_HALF_UP)
                if want != num(x):
                    rounding.append(f"EN {x} vs AR {y} (half-up gives {want})")
                only_e.remove(x); only_a.remove(y); break
    for r in rounding:
        note("FLAG", f"Technicals rounding: {r}.")
    if only_e or only_a:
        note("FLAG", f"Technicals EN/AR figures differ (slides 4-5): EN-only {only_e or '-'}, AR-only {only_a or '-'}. "
                     "Ask the Technicals author which is right before the PDFs.")
    elif not rounding:
        note("OK", "Technicals slides 4-5: the same figures in EN and AR.")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--en", required=True); ap.add_argument("--ar", required=True)
    ap.add_argument("--calendar"); ap.add_argument("--outdir", required=True)
    ap.add_argument("--date", help="YYYY-MM-DD, default today")
    ap.add_argument("--no-render", action="store_true")
    ap.add_argument("--technicals-only", action="store_true",
                    help="Task 7: compare slides 4-5 only. The commentary and calendar checks need a filled "
                         "deck and stop on the build's placeholders (20 Oct 2026)")
    a = ap.parse_args()
    os.makedirs(a.outdir, exist_ok=True)
    rd = datetime.strptime(a.date, "%Y-%m-%d").date() if a.date else datetime.now().date()
    if not a.technicals_only:
        check_parity(a.en, a.ar)
        if a.calendar:
            check_calendar(a.en, a.ar, a.calendar, rd)
        else:
            note("FLAG", "No calendar file given: calendar rows not checked.")
    check_technicals(a.en, a.ar)
    if not a.no_render and not a.technicals_only:
        gaps(a.en, "en", a.outdir); gaps(a.ar, "ar", a.outdir)
    lines = [f"# DMW pre-PDF check - {rd:%d %b %Y}", ""] + [f"- {l}: {m}" for l, m in out]
    open(os.path.join(a.outdir, f"DMW pre-PDF check - {rd:%Y%m%d}.md"), "w", encoding="utf-8").write("\n".join(lines) + "\n")
    print("\n".join(lines))


if __name__ == "__main__":
    main()
