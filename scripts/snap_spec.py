#!/usr/bin/env python3
"""Snap-to-spec for the DMW decks (slides 1-3): put every fixed element back at its spec frame.

Since 29 Oct 2026 Task 7 builds from fixed templates (shared/templates/), so a nudge the team makes
in one day's deck no longer carries over to the next. This script is the backup: it compares a deck
with the spec (shared/data/template_spec.json, read from the templates' own geometry) and moves any
fixed element that has drifted back to its spec position and size (user approved reverting drift,
29 Oct 2026, replacing the 21 Oct "flag but never revert" rule).

Fixed elements: page headers, rules, title, date, logo, the headline cards, the slide 1 and 3 chart
frames, section headers that do not flow with the news (slide 1 Markets header, slide 2 Corporate
header), the Index Performance table and the slide 3 tables. Not touched: the news bodies and the
headers and calendars that flow with them (fill_commentary.py places those), the Results tables,
the DoD box's left edge (placed from the close by build_dmw.place_dod; snap() checks it against the
house rule, DOD_GAP_IN from the close and DOD_TOP_IN top, and re-places it with --fix), table heights (they follow
their rows). Elements are matched by slide and a stable signature (label text, table title cell,
picture position), not by shape id, so they survive the team copying and pasting.

Usage:
  snap_spec.py DECK.pptx --lang EN|AR [--fix] [--out OUT.pptx] [--report report.md]
  (no --fix: report only; exit code 1 when something is off spec)
In the builds: `import snap_spec; snap_spec.snap(prs, "EN", log, fix=True)`.
If the team resizes something on purpose, the user tells Claude, the template gets the change, and
`snap_spec.py --write-spec` re-reads the spec from the two templates.
"""
import argparse, json, re, sys
from pathlib import Path

from pptx import Presentation
from pptx.util import Emu

E = 914400
TOL = 0.011 * E            # moves smaller than this are rounding, not drift
HERE = Path(__file__).resolve().parent
SPEC = HERE.parent / "data" / "template_spec.json"
TEMPLATES = HERE.parent / "templates"
FLOW = re.compile(r"MAJOR ECONOMIC|^s1 TXT MAJOR CORPORATE|أخبار اقتصادية|^s1 TXT أهم أخبار الشركات|"
                  r"EVENT CALENDAR|ECONOMIC CALENDAR|تقويم الأحداث|التقويم الاقتصادي")


def flat(shapes):
    for sh in shapes:
        if sh.shape_type == 6:
            yield from flat(sh.shapes)
        else:
            yield sh


def sig(si, sh):
    """Stable key for a shape, or None when it is content (news, Results, think-cell data)."""
    t = sh.text_frame.text.strip() if sh.has_text_frame else ""
    side = "L" if sh.left < 3.9 * E else "R"
    if sh.has_table:
        c0 = sh.table.cell(0, 0).text.strip()[:30]
        if si == 2 and len(sh.table.columns) >= 9:
            return None
        if re.match(r"(Result|نتائج|إعلانات النتائج)", c0):
            return None
        return f"s{si} TABLE {c0}|{len(sh.table.columns)}c"
    if sh.shape_type == 13:
        return f"s{si} PIC {side} y{round(sh.top / E)}"
    if sh.name.startswith("think-cell"):
        return None
    if re.fullmatch(r"\d", t):
        return f"s{si} PAGENO"
    if si == 1 and 1.1 * E < sh.top < 2.4 * E and sh.has_text_frame and t:
        kind = ("close" if re.match(r"^\s*[\d,]+\.\d{2}", t) else
                "dod" if re.fullmatch(r"\s*\(?-?[\d.]+%\)?\s*", t) else
                "vt" if ("Value" in t or "القيمة" in t) else
                "ytd" if ("YTD" in t or "منذ" in t) else None)
        if kind:
            return f"s1 CARD {side} {kind}"
    if t.startswith("["):              # news placeholder from the build
        return None
    if t and len(t) < 60 and not re.search(r"\d", t):
        return f"s{si} TXT {t[:45]} {side}"
    if not t and sh.shape_type in (1, 5, 9):
        return f"s{si} SHP {sh.name} {side}"
    if re.match(r"^(Monday|Tuesday|Wednesday|Thursday|Sunday|Friday|Saturday)|"
                r"^(الأحد|الإثنين|الاثنين|الثلاثاء|الأربعاء|الخميس)", t):
        return f"s{si} DATE"
    return None


def keyed(prs):
    out, seen = {}, {}
    for si, s in enumerate(list(prs.slides)[:3], 1):
        for sh in flat(s.shapes):
            k = sig(si, sh)
            if not k or FLOW.search(k):
                continue
            seen[k] = seen.get(k, 0) + 1
            out[k if seen[k] == 1 else f"{k} #{seen[k]}"] = sh
    return out


def fields(key, sh):
    if sh.has_table:
        return ("left", "top", "width")
    if key.startswith("s1 CARD") and "dod" in key:
        return ("top", "height")       # left is placed from the close, width grows with the figure
    return ("left", "top", "width", "height")


def load_spec(lang):
    return json.loads(SPEC.read_text(encoding="utf-8"))[lang]


def snap(prs, lang, logf=None, fix=True, spec=None):
    """Compare with the spec; move drifted elements back when fix. Returns the list of moves."""
    spec = spec or load_spec(lang)
    have = keyed(prs)
    moves = []
    for k, frame in spec.items():
        sh = have.get(k)
        if sh is None:
            continue
        for f in fields(k, sh):
            want = int(round(frame[f] * E))
            cur = getattr(sh, f)
            if abs(cur - want) > TOL:
                moves.append(f"{k}: {f} {cur / E:.3f} -> {frame[f]:.3f} in")
                if fix:
                    setattr(sh, f, Emu(want))
    # headline-card DoD boxes: house gap and top (user, 5 Nov 2026), both cards, both decks
    if prs.slides:
        import legacy_build_dmw_v1 as bd
        pairs = bd.card_pairs(prs.slides[0])
        for side, (g, t) in bd.dod_gaps(pairs).items():
            if abs(g - bd.DOD_GAP_IN) > 0.005 or abs(t - bd.DOD_TOP_IN) > 0.005:
                moves.append(f"s1 CARD {side} DoD: gap {g:.3f} / top {t:.3f} -> {bd.DOD_GAP_IN:.3f} / {bd.DOD_TOP_IN:.2f} in")
        if fix and any(m.startswith("s1 CARD") and "DoD:" in m for m in moves):
            bd.place_dod(pairs, None, f"Layout {lang} slide 1", logf or (lambda *a: None))
    missing = [k for k in spec if k not in have]
    if logf:
        verb = "moved back to spec" if fix else "off spec"
        if moves:
            logf("FLAG" if not fix else "OK", f"Layout {lang}: {len(moves)} drift(s) {verb}: " + "; ".join(moves))
        else:
            logf("OK", f"Layout {lang}: every fixed element on spec ({len(spec) - len(missing)} checked).")
        if missing:
            logf("FLAG", f"Layout {lang}: spec element(s) not found in the deck: " + "; ".join(missing))
    return moves


def write_spec():
    spec = {}
    for lang in ("EN", "AR"):
        prs = Presentation(TEMPLATES / f"DMW Template - {lang}.pptx")
        spec[lang] = {k: {f: round(getattr(sh, f) / E, 3) for f in ("left", "top", "width", "height")}
                      for k, sh in keyed(prs).items()}
    SPEC.write_text(json.dumps(spec, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"Wrote {SPEC} ({len(spec['EN'])} EN, {len(spec['AR'])} AR elements)")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("deck", nargs="?")
    ap.add_argument("--lang", choices=("EN", "AR"))
    ap.add_argument("--fix", action="store_true")
    ap.add_argument("--out")
    ap.add_argument("--report")
    ap.add_argument("--write-spec", action="store_true")
    a = ap.parse_args()
    if a.write_spec:
        return write_spec()
    lines = []
    prs = Presentation(a.deck)
    moves = snap(prs, a.lang, lambda lvl, m: lines.append(f"- {lvl}: {m}"), fix=a.fix)
    text = f"# Layout check - {Path(a.deck).name}\n\n" + "\n".join(lines) + "\n"
    print(text)
    if a.report:
        Path(a.report).write_text(text, encoding="utf-8")
    if a.fix and moves:
        prs.save(a.out or a.deck)
        print(f"Saved {a.out or a.deck}")
    sys.exit(1 if moves and not a.fix else 0)


if __name__ == "__main__":
    main()
