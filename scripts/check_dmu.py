"""Final check of a DMU deck against the house rules learnt from the team's 80 decks (Jul to Oct 2026).

python check_dmu.py --deck "Daily Market Update - YYYYMMDD.pptx" --date YYYY-MM-DD
    [--data "DMW - Data.xlsx"] [--prev "<previous DMU>.pptx"] [--dmw "ANB Daily Market Watch - EN.pptx"]
    [--out report.md]
FAIL = must fix before hand-over; WARN = look at it; TEAM = the team's item, still open; OK = passed.
"""
import argparse, datetime, difflib, hashlib, re, sys, os
from pptx import Presentation
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import build_dmu as B

EMU = 914400
BLUE = "0070C0"
out = []


def res(level, msg): out.append(f"{level}: {msg}")


def runs(shape):
    return [r for p in shape.text_frame.paragraphs for r in p.runs if r.text.strip()]


def col(r):
    try:
        return str(r.font.color.rgb) if r.font.color and r.font.color.type == 1 else None
    except Exception:
        return None


def expect_fmt(shape, label, size=None, bold=None, color=None, font="Arial Nova Cond"):
    bad = []
    for r in runs(shape):
        if font and r.font.name not in (font, None): bad.append(f"font {r.font.name}")
        if size and r.font.size and abs(r.font.size.pt - size) > 0.01: bad.append(f"size {r.font.size.pt}")
        if bold is not None and bool(r.font.bold) != bold: bad.append(f"bold {r.font.bold}")
        if color and col(r) not in (color,): bad.append(f"colour {col(r)}")
    if bad: res("FAIL", f"{label}: {sorted(set(bad))[:4]} (spec: {font}, {size} pt, bold={bold}, colour={color})")
    else: res("OK", f"{label}: formatting")


def words(lines): return sum(len(l.split()) for l in lines)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--deck", required=True); ap.add_argument("--date", required=True)
    ap.add_argument("--data"); ap.add_argument("--prev"); ap.add_argument("--dmw"); ap.add_argument("--out")
    a = ap.parse_args()
    day = datetime.date.fromisoformat(a.date)
    P = Presentation(a.deck)
    titles = [B.slide_title(s) for s in P.slides]

    # 1 structure
    core = ["GLOBAL MARKETS", "COMMODITY MARKETS", "SAUDI MARKET"]
    seq = [t for t in titles if t]
    if seq[:3] != core or not seq[-1].startswith("TASI TECHNICAL"):
        res("FAIL", f"slide order {seq}")
    else:
        res("OK", f"slide order ({len(P.slides)} slides: {', '.join(t.title() for t in seq)})")

    # 2 cover date
    want = f"{day.day} {day.strftime('%B %Y')}"
    cov = " ".join(s.text_frame.text for s in B.text_shapes(P.slides[0]))
    res("OK" if want in cov else "FAIL", f"cover date {'is' if want in cov else 'is not'} '{want}'")

    # 3 text rules everywhere
    alltext = "\n".join(s.text_frame.text for sl in P.slides for s in B.text_shapes(sl))
    for sl in P.slides:
        for s in B.flat(sl.shapes):
            if getattr(s, "has_table", False) and s.has_table:
                alltext += "\n" + "\n".join(c.text for r in s.table.rows for c in r.cells)
    if "—" in alltext: res("FAIL", "em-dash found")
    if re.search(r"\bTK\b", alltext): res("FAIL", "'TK' left in the deck")

    g = B.one_slide(P, "GLOBAL MARKETS"); cm = B.one_slide(P, "COMMODITY MARKETS"); sa = B.one_slide(P, "SAUDI MARKET")
    tech = B.one_slide(P, "TASI TECHNICAL")

    # 4 headline boxes: counts, lengths, formatting
    for sl, name, n_rng, w_rng, per in ((g, "global", (4, 6), (204, 347), (30, 105)),
                                        (cm, "commodities", (2, 3), (81, 209), (30, 95)),
                                        (sa, "saudi", (3, 5), (178, 275), (25, 120))):
        box = B.body_box(sl, "HEADLINE UPDATES")
        if box is None: res("FAIL", f"{name}: headline box not found"); continue
        lines = [p.text for p in box.text_frame.paragraphs if p.text.strip()]
        n, w = len(lines), words(lines)
        ok = n_rng[0] <= n <= n_rng[1] and w_rng[0] <= w <= w_rng[1]
        res("OK" if ok else "WARN", f"{name}: {n} items, {w} words (team range {n_rng[0]}-{n_rng[1]} items, "
                                     f"{w_rng[0]}-{w_rng[1]} words)")
        short = [i + 1 for i, l in enumerate(lines) if not per[0] <= len(l.split()) <= per[1]]
        if short: res("WARN", f"{name}: items {short} outside {per[0]}-{per[1]} words (team items are full leads)")
        expect_fmt(box, f"{name}: headline text", size=11, bold=False, color=None)

    # 5 captions: oil on top, gold below, Reuters headline style
    caps = sorted([s for s in B.text_shapes(cm) if s.left > 7 * EMU and s.height < 0.5 * EMU
                   and len(s.text_frame.text.strip()) > 20 and "USD" not in s.text_frame.text], key=lambda s: s.top)
    if len(caps) == 2:
        t0, t1 = caps[0].text_frame.text.lower(), caps[1].text_frame.text.lower()
        ok = re.search(r"oil|brent|crude|wti|opec", t0) and "gold" in t1
        res("OK" if ok else "FAIL", f"captions: top = '{caps[0].text_frame.text.strip()[:50]}', bottom = "
                                    f"'{caps[1].text_frame.text.strip()[:50]}' (top must be oil, bottom gold)")
        for c in caps:
            n = len(c.text_frame.text.split())
            if not 6 <= n <= 15: res("WARN", f"caption has {n} words (team 6-15)")
            expect_fmt(c, "caption", size=12, bold=True, color=BLUE)
    else:
        res("FAIL", f"commodities: expected 2 captions, found {len(caps)}")

    # 6 events
    ev = next((s for s in B.text_shapes(g) if re.match(r"\d{2}-[A-Z][a-z]{2}-\d{2,4}:", s.text_frame.text.strip())), None)
    if ev is not None:
        for p in ev.text_frame.paragraphs:
            t = p.text.strip()
            if not t: continue
            if not re.match(r"^\d{2}-[A-Z][a-z]{2}-\d{2}: ", t):
                res("FAIL", f"events line not 'DD-Mon-YY: ...': '{t[:50]}'")
            if not t.startswith(day.strftime("%d-%b-%y")):
                res("WARN", f"events line not dated today: '{t[:40]}'")
            rr = [r for r in p.runs if r.text.strip()]
            if rr and (not rr[0].font.bold or any(r.font.bold for r in rr[1:])):
                res("FAIL", f"events line formatting: date run must be bold and the rest regular ('{t[:30]}')")
        res("OK", "events: format checked")
    else:
        res("INFO", "events: no events lines today")

    # 7 labels and titles
    for sl in (g, cm, sa, tech):
        for s in B.text_shapes(sl):
            t = s.text_frame.text.strip()
            if t in ("MARKET RECAP", "HEADLINE UPDATES", "COMPARATIVE VALUATIONS", "EVENTS CALENDAR", "STOCK FUTURES",
                     "PERFORMANCE RECAP", "PERFORMANCE HEATMAP", "MARKET OUTLOOK"):
                expect_fmt(s, f"label {t}", size=12, bold=True, color=BLUE)

    # 8 TP column blank or filled, never TK
    for s in B.flat(sa.shapes):
        if getattr(s, "has_table", False) and s.has_table and s.table.rows[0].cells[0].text.strip().upper().startswith(("LEAD", "LAGG")):
            tps = [r.cells[5].text.strip() for r in list(s.table.rows)[2:]]
            blank = sum(1 for x in tps if not x)
            res("INFO", f"{s.table.rows[0].cells[0].text.strip()}: {blank} of {len(tps)} Consensus TP cells blank (team fills)")

    # 9 technical outlook: bold levels, two paragraphs
    box = max((s for s in B.text_shapes(tech) if len(s.text_frame.text) > 200), key=lambda s: len(s.text_frame.text), default=None)
    if box is not None:
        ps = [p for p in box.text_frame.paragraphs if p.text.strip()]
        nb = sum(1 for p in ps for r in p.runs if r.font.bold and r.text.strip())
        ok = len(ps) == 2 and nb >= 2 and ps[-1].text.strip().startswith("Trend is")
        res("OK" if ok else "WARN", f"technical outlook: {len(ps)} paragraphs, {nb} bold level runs "
                                    "(team: brief in one paragraph, then 'Trend is ...', levels bold)")

    # 10 charts against the data file
    if a.data:
        import openpyxl
        wb = openpyxl.load_workbook(a.data, read_only=True, data_only=True)
        for sl, label in ((cm, "oil chart"), (cm, "gold chart"), (sa, "P/E chart")):
            chs = B.charts(sl)
            ch = chs[0] if label != "gold chart" else (chs[1] if len(chs) > 1 else None)
            if ch is None: res("FAIL", f"{label}: not found"); continue
            cs = ch.chart._chartSpace
            f = cs.find(".//" + B.C + "cat").find(".//" + B.C + "f")
            ref = B.parse_ref(f.text)
            sheet, c, s0, e0 = ref
            last = B.last_row(wb[sheet], c, s0)
            vals = B.read_col(wb[sheet], c, s0, last)
            while vals and isinstance(vals[-1], datetime.datetime) and (vals[-1].date() >= day or vals[-1].weekday() >= 5):
                vals.pop()
            want_last = vals[-1]
            got = list(ch.chart.plots[0].categories)[-1]
            if isinstance(want_last, datetime.datetime):
                want_txt = str(B.excel_serial(want_last))
                ok = str(got).split(".")[0] == want_txt
                got_d = datetime.date(1899, 12, 30) + datetime.timedelta(days=int(float(got)))
                res("OK" if ok else "FAIL", f"{label}: last date {got_d} (data file {want_last.date()})")
            else:
                ok = str(got) == B.pe_label(want_last)
                res("OK" if ok else "FAIL", f"{label}: last category {got} (data file {B.pe_label(want_last)})")
            links = [r.target_ref for r in ch.chart.part.rels.values() if r.is_external]
            if links and day.strftime("%Y%m%d") not in links[0]:
                res("FAIL", f"{label}: linked to {re.findall(r'2026\d{4}', links[0])} not today's data file")

    # 11 pictures that must change every day
    if a.prev:
        Q = Presentation(a.prev)

        def pic_hash(prs, prefix, x, y):
            ss = B.find_slides(prs, prefix)
            p = B.pic_at(ss[0], x, y) if ss else None
            return hashlib.md5(p.image.blob).hexdigest() if p is not None else None
        for prefix, x, y, label in (("GLOBAL MARKETS", 0.25, 1.26, "market recap"), ("GLOBAL MARKETS", 7.5, 3.85, "stock futures"),
                                    ("SAUDI MARKET", 0.43, 1.87, "performance heatmap")):
            a_, b_ = pic_hash(P, prefix, x, y), pic_hash(Q, prefix, x, y)
            lvl = "OK" if a_ and a_ != b_ else ("TEAM" if label == "performance heatmap" else "FAIL")
            res(lvl, f"{label} picture {'replaced' if a_ != b_ else 'same as the previous DMU'}")
        tq = B.one_slide(Q, "TASI TECHNICAL")
        hp = {hashlib.md5(p.image.blob).hexdigest() for p in B.pictures(tq)}
        same = [p for p in B.pictures(tech) if p.width > 1.5 * EMU and hashlib.md5(p.image.blob).hexdigest() in hp]
        res("OK" if not same else "TEAM", f"technical pictures: {len(same)} unchanged from the previous DMU "
                                          "(the technical team's summary images: paste today's)")

    # 12 DMW text reused word for word
    if a.dmw:
        D = Presentation(a.dmw)
        dp = [re.sub(r"\s+", " ", p.text).strip() for s in list(D.slides)[:3] for sh in s.shapes
              if sh.has_text_frame for p in sh.text_frame.paragraphs if len(p.text) > 60]
        for sl, name, idx in ((g, "global headline 1", [0]), (sa, "saudi headlines", None)):
            box = B.body_box(sl, "HEADLINE UPDATES")
            lines = [re.sub(r"\s+", " ", p.text).strip() for p in box.text_frame.paragraphs if p.text.strip()]
            for i in (idx if idx is not None else range(len(lines))):
                best = max((difflib.SequenceMatcher(None, lines[i], d).ratio() for d in dp), default=0)
                res("OK" if best >= 0.95 else "FAIL",
                    f"{name} #{i + 1}: {best:.2f} match to the final DMW text (must be copied word for word)")

    rep = "\n".join(f"- {x}" for x in out)
    fails = sum(x.startswith("FAIL") for x in out); warns = sum(x.startswith("WARN") for x in out)
    head = f"# DMU check - {a.date}\n\n{fails} FAIL, {warns} WARN\n\n"
    if a.out: open(a.out, "w", encoding="utf-8").write(head + rep + "\n")
    print(head + rep)
    sys.exit(1 if fails else 0)


if __name__ == "__main__":
    main()
