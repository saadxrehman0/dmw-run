#!/usr/bin/env python3
"""Patch an existing, hand-edited DMW deck without rebuilding it (keeps the team's edits).

  tech   --deck D --technicals T [--lang en|ar] --data X --date YYYY-MM-DD --out O
         swap slides 4-5 from the Technicals deck, cross-check date/close/value traded
  ar-fill --deck AR --json commentary.json --out O
         fill Arabic commentary boxes (keys: markets, economic, corporate_s1, corporate_s2)
  check  --en EN_DECK --json commentary.json
         compare every number and ticker in the EN commentary vs the Arabic JSON
"""
import argparse, copy, json, re, sys
from datetime import datetime
from lxml import etree
from pptx import Presentation
import build_dmw as en
import build_dmw_ar as ar
A = en.A


def en_commentary(path):
    p = Presentation(path); out = []
    for s in list(p.slides)[:2]:
        for sh in sorted([x for x in s.shapes if x.has_text_frame and len(x.text_frame.text) > 60], key=lambda x: x.top):
            out += [pp.text for pp in sh.text_frame.paragraphs if pp.text.strip()]
    return out


def cmd_tech(a):
    rd = datetime.strptime(a.date, "%Y-%m-%d").date()
    tech_date_fix = False
    g, face = en.load_data(a.data)
    tech = Presentation(a.technicals)
    if len(tech.slides) != 2:
        sys.exit("Technicals deck must have 2 slides")
    if a.lang == "en":
        en.cross_checks(g, face, rd, None, tech, None, [], [])
    else:
        # Date style from the deck's own slide 1 date, as build_dmw_ar.py does (until 20 Oct 2026 this
        # was never set here, so the patcher expected the old style).
        for sh in Presentation(a.deck).slides[0].shapes:
            if sh.has_text_frame and re.match(ar.DAY_RE, sh.text_frame.text.strip()):
                ar.AR_DATE_STYLE["arabic_comma"] = "،" in sh.text_frame.text
        td = [sh.text_frame.text.strip() for sh in tech.slides[0].shapes if sh.has_text_frame
              and re.match(ar.DAY_RE, sh.text_frame.text.strip())]
        # Same date in the other punctuation style is fixed on slide 4 after the paste (user, 20 Oct 2026).
        bare = lambda t: re.sub(r"[\s,،]+", " ", t).strip()
        tech_date_fix = bool(td) and td[0] != ar.ar_long_date(rd) and bare(td[0]) == bare(ar.ar_long_date(rd))
        if not td or (td[0] != ar.ar_long_date(rd) and not tech_date_fix):
            en.log("CONFLICT", f"Technicals AR date {td[:1]} vs report date '{ar.ar_long_date(rd)}'.")
        for sh in tech.slides[0].shapes:
            if sh.has_table:
                kv = {sh.table.cell(i, 1).text.strip(): sh.table.cell(i, 0).text.strip() for i in range(len(sh.table.rows))}
                if "الإغلاق" in kv:
                    c = float(kv["الإغلاق"].replace(",", ""))
                    en.log("OK" if abs(c - g["C6"]) < 0.005 else "CONFLICT", f"TASI close: Technicals AR {kv['الإغلاق']} vs data {g['C6']:,.2f}.")
                if "القيمة (مليون ريال)" in kv:
                    v = float(kv["القيمة (مليون ريال)"].replace(",", ""))
                    if abs(v / 1000 - g["F6"]) > 0.05:
                        en.log("CONFLICT", f"Value traded: Technicals AR SAR {v:,.1f} mn vs data file SAR {g['F6']*1000:,.1f} mn.")
    prs = Presentation(a.deck)
    # Task 7 deletes slides 4-5 when the Technicals are late; add them back before patching.
    while len(prs.slides) < 5:
        prs.slides.add_slide(prs.slides[len(prs.slides) - 1].slide_layout)
        en.log("OK", f"Deck had {len(prs.slides) - 1} slides; added slide {len(prs.slides)} for the Technicals.")
    en.replace_with_technicals(prs.slides[3], tech.slides[0]); en.replace_with_technicals(prs.slides[4], tech.slides[1])
    if a.lang != "en":
        ar.ensure_s5_rule(prs.slides[4], en.log)
    if tech_date_fix:
        for sh in prs.slides[3].shapes:
            if sh.has_text_frame and re.match(ar.DAY_RE, sh.text_frame.text.strip()):
                runs = sh.text_frame.paragraphs[0].runs
                runs[0].text = ar.ar_long_date(rd)
                for r in runs[1:]:
                    r._r.getparent().remove(r._r)
        en.log("OK", f"Technicals AR date '{td[0]}' rewritten to the house form '{ar.ar_long_date(rd)}' on slide 4.")
    prs.save(a.out); en.log("OK", f"Slides 4-5 replaced from {a.technicals.split('/')[-1]}.")


def fill_box(sh, paras):
    tb = sh.text_frame._txBody
    ps = tb.findall(f"{{{A}}}p"); ppr = ps[0].find(f"{{{A}}}pPr")
    ppr = copy.deepcopy(ppr) if ppr is not None else None
    for x in ps:
        tb.remove(x)
    for text in paras:
        pe = etree.SubElement(tb, f"{{{A}}}p")
        if ppr is not None:
            pe.append(copy.deepcopy(ppr))
        for seg in re.split(r"([A-Za-z][A-Za-z0-9&.,\-' ]*[A-Za-z0-9.]|[A-Za-z])", text):
            if not seg:
                continue
            r = etree.SubElement(pe, f"{{{A}}}r"); rp = etree.SubElement(r, f"{{{A}}}rPr")
            rp.set("lang", "en-US" if re.match(r"[A-Za-z]", seg) else "ar-SA"); rp.set("sz", "900"); rp.set("dirty", "0")
            etree.SubElement(rp, f"{{{A}}}latin").set("typeface", "+mj-lt")
            etree.SubElement(r, f"{{{A}}}t").text = seg


def cmd_arfill(a):
    d = json.load(open(a.json, encoding="utf-8")); p = Presentation(a.deck); s1, s2 = p.slides[0], p.slides[1]
    hdrs = {"أداء الأسواق/الأخبار": "markets", "أخبار اقتصادية/قطاعية رئيسية": "economic", "أهم أخبار الشركات": "corporate_s1"}
    heads = {sh.text_frame.text.strip(): sh for sh in s1.shapes if sh.has_text_frame and sh.text_frame.text.strip() in hdrs}
    bodies = [sh for sh in s1.shapes if sh.has_text_frame and not sh.has_table and sh.width > 6_500_000 and sh.height > 300_000]
    for h, key in hdrs.items():
        hd = heads.get(h)
        if hd is None:
            sys.exit(f"header '{h}' not found")
        below = sorted([b for b in bodies if b.top > hd.top - 50_000], key=lambda b: b.top)
        fill_box(below[0], d[key])
    big = max((sh for sh in s2.shapes if sh.has_text_frame and not sh.has_table), key=lambda s: s.height)
    fill_box(big, d["corporate_s2"])
    p.save(a.out); print("filled", {k: len(v) for k, v in d.items()})


def cmd_check(a):
    enp = en_commentary(a.en); d = json.load(open(a.json, encoding="utf-8"))
    arp = d["markets"] + d["economic"] + d["corporate_s1"] + d["corporate_s2"]
    num = lambda s: sorted(x.rstrip(".,") for x in re.findall(r"\d[\d,]*\.?\d*", s))
    tick = lambda s: sorted(re.findall(r"\(([A-Z0-9]+ (?:AB|US|UH|LN))\)", s))
    print(f"EN paragraphs {len(enp)}, AR paragraphs {len(arp)}")
    for i, (e, x) in enumerate(zip(enp, arp)):
        if num(e) != num(x) or tick(e) != tick(x):
            print(f"[{i}] numbers EN {num(e)} AR {num(x)}; tickers EN {tick(e)} AR {tick(x)}")


if __name__ == "__main__":
    ap = argparse.ArgumentParser(); sp = ap.add_subparsers(dest="cmd", required=True)
    t = sp.add_parser("tech"); [t.add_argument(f"--{k}", required=True) for k in ("deck", "technicals", "data", "date", "out")]; t.add_argument("--lang", default="en")
    f = sp.add_parser("ar-fill"); [f.add_argument(f"--{k}", required=True) for k in ("deck", "json", "out")]
    c = sp.add_parser("check"); [c.add_argument(f"--{k}", required=True) for k in ("en", "json")]
    a = ap.parse_args()
    {"tech": cmd_tech, "ar-fill": cmd_arfill, "check": cmd_check}[a.cmd](a)
    for l, m in en.checks:
        print(l, m)
