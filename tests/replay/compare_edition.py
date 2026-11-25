#!/usr/bin/env python3
"""Compare a replayed edition (run_edition.py --replay) with the edition that was published
(build step 5, 25 Nov 2026). Read only. Writes a Markdown report and a JSON summary.

  python compare_edition.py --run "<runs root>/replay/2026-11-24" --published "<day folder>" \
      [--next "<next day folder>"] --out report.md [--json summary.json]

Sections:
  stages    status, attempts, run-clock and real minutes, alerts (run_state.json, logs)
  news      published Economic and Corporate items (decks) matched to the run's NEWS.json by
            tickers, numbers and words: found, missed, extra, rank
  markets   numbers in the TASI and Global bullets, published vs run
  tables    numbers in the slide 1 to 3 tables other than news (multiset), published vs run
  results   Results Announcements rows by company, published vs run
  email     EMAIL MB paragraphs per heading, text similarity
  pdfs      page counts
  calendar  next edition's Manual File-Calendar.xlsx (S6) vs the one the team used next day
Every count is computed here; the report states which file each side came from.
"""
import argparse, collections, difflib, glob, json, os, re, sys

HERE = os.path.dirname(os.path.abspath(__file__))
SC = os.path.join(HERE, "..", "..", "scripts")
sys.path.insert(0, SC)
sys.path.insert(0, HERE)
from news_text import numbers, tickers  # noqa: E402

EN, AR = "ANB Daily Market Watch - EN.pptx", "ANB Daily Market Watch - AR.pptx"


def newest(pattern_list):
    hits = []
    for p in pattern_list:
        hits += [h for h in glob.glob(p, recursive=True) if not os.path.basename(h).startswith("~$")]
    return max(hits, key=os.path.getmtime) if hits else None


def run_file(run, *names):
    pats = []
    for n in names:
        pats += [os.path.join(run, "practice", "**", n), os.path.join(run, "decks", "**", n)]
    return newest(pats)


# ---------------------------------------------------------------- readers

def is_v2(prs):
    return any(sh.name in ("s1.news", "s2.flow") for s in prs.slides for sh in s.shapes)


def deck_news(path):
    """{"markets": [..], "economic": [..], "corporate": [..]} from a v1 or v2 deck (EN or AR)."""
    from pptx import Presentation
    prs = Presentation(path)
    if is_v2(prs):
        from build_email_mb import read_deck_v2
        news = read_deck_v2(prs)[0]
        return {"markets": news.get("Markets Performance/News", []),
                "economic": news.get("Major Economic/Sector News", []),
                "corporate": news.get("Major Corporate News", [])}
    from legacy_fill_commentary_v1 import find_shapes
    lang = "ar" if "AR" in os.path.basename(path) else "en"
    f = find_shapes(prs, lang, strict=False)
    paras = lambda sh: [p.text.strip() for p in sh.text_frame.paragraphs if p.text.strip()] if sh is not None else []
    return {"markets": paras(f["mk"]), "economic": paras(f["ec"]), "corporate": paras(f["co"]) + paras(f["s2"])}


def run_news(run):
    p = newest([os.path.join(run, "news", "NEWS.json"), os.path.join(run, "practice", "**", "NEWS.json")])
    if not p:
        return None, None
    j = json.load(open(p, encoding="utf-8"))
    return p, {"markets": j.get("markets", {}).get("en", []),
               "economic": [x.get("en", "") for x in j.get("economic", [])],
               "corporate": [x.get("en", "") for x in j.get("corporate", [])]}


def table_numbers(path, skip_news=True):
    from pptx import Presentation
    prs = Presentation(path)
    nums = collections.Counter()
    news_names = ("s1.news", "s2.flow") if skip_news else ()
    for si, s in enumerate(prs.slides, 1):
        if si > 3:
            break
        for sh in s.shapes:
            if not sh.has_table or sh.name in news_names or is_results(sh):
                continue
            for r in sh.table.rows:
                for c in r.cells:
                    t = c.text.strip()
                    if len(t) > 40:       # v1 decks keep news in tables too; skip long text
                        continue
                    for x in numbers(t, "en"):
                        nums[x.replace(",", "")] += 1
    return nums


def is_results(sh):
    """A Results table: v2 names it s1/s2.results; v1 heads it "Result Announcements"."""
    if "result" in sh.name.lower():
        return True
    rows = [[c.text.strip() for c in r.cells] for r in list(sh.table.rows)[:2]]
    return "result" in " ".join(" ".join(r) for r in rows).lower()


def results_rows(path):
    """Company -> list of numbers, from tables whose shape name or first rows say Results."""
    from pptx import Presentation
    prs = Presentation(path)
    out = {}
    for si, s in enumerate(prs.slides, 1):
        if si > 3:
            break
        for sh in s.shapes:
            if not sh.has_table:
                continue
            if not is_results(sh):
                continue
            rows = [[c.text.strip() for c in r.cells] for r in sh.table.rows]
            for r in rows:
                # key on the TASI code: the run and the published deck spell names differently
                # (published v1 "Result Announcements" tables were missed, 4 Nov replay, 26 Nov 2026)
                code = next((c for c in r[1:3] if re.fullmatch(r"\d{4}", c)), None)
                nums = [x for c in r[3:] for x in numbers(c, "en")]
                if code and len(nums) >= 2:
                    out[code] = nums
                elif r and r[0] and len(nums) >= 3 and not re.search(r"(?i)result|company", r[0]):
                    out[re.sub(r"\s+", " ", r[0])] = nums
    return out


def docx_sections(path):
    from docx import Document
    sec, out = "(top)", collections.OrderedDict()
    for p in Document(path).paragraphs:
        t = p.text.strip()
        if not t:
            continue
        if len(t) < 45 and (p.style.name.lower().startswith("heading") or t in (
                "Markets Performance/News", "Major Economic/Sector News", "Major Corporate News",
                "Corporate Calendar")):
            sec = t
            continue
        out.setdefault(sec, []).append(t)
    return out


def pdf_pages(path):
    try:
        from pypdf import PdfReader
        return len(PdfReader(path).pages)
    except Exception as e:  # noqa: BLE001
        return f"unreadable ({e.__class__.__name__})"


def xlsx_cells(path):
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    out = {}
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if c.value not in (None, ""):
                    out[(ws.title, c.coordinate)] = c.value
    return out


# ---------------------------------------------------------------- matching

def words(s):
    return set(w for w in re.findall(r"[a-z]{4,}", s.lower()))


def score(a, b):
    ta, tb = set(tickers(a)), set(tickers(b))
    na, nb = set(numbers(a, "en")), set(numbers(b, "en"))
    wa, wb = words(a), words(b)
    jac = len(wa & wb) / max(1, len(wa | wb))
    return 10 * len(ta & tb) + 2 * len(na & nb) + 10 * jac


def match(pub, mine, floor=6.0):
    """Greedy best-score matching. Returns [(pub_i, mine_i or None, score)], extra mine indices."""
    pairs = sorted(((score(p, m), i, j) for i, p in enumerate(pub) for j, m in enumerate(mine)), reverse=True)
    used_p, used_m, got = set(), set(), {}
    for s, i, j in pairs:
        if s < floor or i in used_p or j in used_m:
            continue
        used_p.add(i); used_m.add(j); got[i] = (j, s)
    res = [(i, got.get(i, (None, 0))[0], round(got.get(i, (None, 0))[1], 1)) for i in range(len(pub))]
    return res, [j for j in range(len(mine)) if j not in used_m]


def short(s, n=90):
    s = re.sub(r"\s+", " ", s)
    return s if len(s) <= n else s[: n - 3] + "..."


# ---------------------------------------------------------------- report

def compare(run, pub_day, next_day=None):
    R = {"run": run, "published": pub_day, "sections": {}}
    L = [f"# Replay vs published: {os.path.basename(os.path.normpath(run))}", "",
         f"- Run folder: `{run}`", f"- Published day folder: `{pub_day}`", ""]

    # stages
    st = json.load(open(os.path.join(run, "run_state.json"), encoding="utf-8")) if \
        os.path.exists(os.path.join(run, "run_state.json")) else {"stages": {}, "alerts": []}
    L += ["## Stages", "", "| Stage | Status | Attempts | Note |", "|---|---|---|---|"]
    for sid, s in st["stages"].items():
        L.append(f"| {sid} | {s['status']} | {s['attempts']} | {short(str(s.get('note') or ''), 120)} |")
    L += ["", f"Alerts: {', '.join(st.get('alerts', [])) or 'none'}", ""]
    R["sections"]["stages"] = {k: v["status"] for k, v in st["stages"].items()}

    # news
    pub_en = os.path.join(pub_day, EN)
    mine_en = run_file(run, EN, "*EN*.pptx")
    npath, rn = run_news(run)
    if os.path.exists(pub_en) and rn is None and mine_en:
        npath, rn = mine_en, deck_news(mine_en)
    L += ["## News", ""]
    if os.path.exists(pub_en) and rn:
        pn = deck_news(pub_en)
        L += [f"Published: `{pub_en}`. Run: `{npath}`.", ""]
        news_sum = {}
        for k in ("economic", "corporate"):
            res, extra = match(pn[k], rn[k])
            found = [r for r in res if r[1] is not None]
            news_sum[k] = {"published": len(pn[k]), "run": len(rn[k]), "found": len(found),
                           "extra": len(extra)}
            L += [f"### {k.title()}: {len(found)} of {len(pn[k])} published items found; "
                  f"run had {len(rn[k])} ({len(extra)} not published)", "",
                  "| Pub rank | Run rank | Score | Published item |", "|---|---|---|---|"]
            for i, j, s in res:
                L.append(f"| {i + 1} | {'' if j is None else j + 1} | {s or ''} | {short(pn[k][i])} |")
            if extra:
                L += ["", "Run items not published:", ""] + [f"- ({j + 1}) {short(rn[k][j])}" for j in extra]
            L.append("")
        R["sections"]["news"] = news_sum
        # markets
        L += ["## Markets bullets (numbers)", ""]
        for i, name in enumerate(("Global", "TASI")):
            a = collections.Counter(numbers(pn["markets"][i], "en")) if len(pn["markets"]) > i else collections.Counter()
            b = collections.Counter(numbers(rn["markets"][i], "en")) if len(rn["markets"]) > i else collections.Counter()
            L.append(f"- {name}: published only {sorted((a - b).elements()) or 'none'}; "
                     f"run only {sorted((b - a).elements()) or 'none'}")
        L.append("")
    else:
        L += ["Not compared: " + ("no run news" if not rn else f"no published deck {pub_en}"), ""]

    # tables, results, pdfs per language
    L += ["## Deck tables, Results and PDFs", ""]
    for lang, name in (("EN", EN), ("AR", AR)):
        pub = os.path.join(pub_day, name)
        mine = run_file(run, name, f"*{lang}*.pptx")
        if not (os.path.exists(pub) and mine):
            L += [f"- {lang}: not compared (published {os.path.exists(pub)}, run deck {mine})"]
            continue
        a, b = table_numbers(pub), table_numbers(mine)
        po, ro = a - b, b - a
        L += [f"### {lang}", "", f"Run deck: `{mine}`", "",
              f"- Table numbers: published {sum(a.values())}, run {sum(b.values())}; "
              f"published only {sum(po.values())}, run only {sum(ro.values())}"]
        if po:
            L.append(f"  - published only (first 25): {sorted(po.elements())[:25]}")
        if ro:
            L.append(f"  - run only (first 25): {sorted(ro.elements())[:25]}")
        if lang == "EN":
            pr, rr = results_rows(pub), results_rows(mine)
            L.append(f"- Results rows: published {len(pr)}, run {len(rr)}")
            for c in sorted(set(pr) | set(rr)):
                if c not in rr:
                    L.append(f"  - published only: {c}")
                elif c not in pr:
                    L.append(f"  - run only: {c}")
                elif pr[c][:4] != rr[c][:4]:
                    L.append(f"  - {c}: published {pr[c]} run {rr[c]}")
            R["sections"]["results"] = {"published": len(pr), "run": len(rr)}
        R["sections"][f"tables_{lang}"] = {"published_only": sum(po.values()), "run_only": sum(ro.values())}
        pp = os.path.join(pub_day, name[:-5] + ".pdf")
        mp = run_file(run, name[:-5] + ".pdf", f"*{lang}*.pdf")
        L.append(f"- PDF pages: published {pdf_pages(pp) if os.path.exists(pp) else 'none'}, "
                 f"run {pdf_pages(mp) if mp else 'none'}")
        L.append("")

    # email
    ymd = os.path.basename(os.path.normpath(pub_day))
    pe = os.path.join(pub_day, f"EMAIL MB {ymd}.docx")
    me = run_file(run, f"EMAIL MB {ymd}.docx", "EMAIL MB*.docx")
    L += ["## EMAIL MB", ""]
    if os.path.exists(pe) and me:
        a, b = docx_sections(pe), docx_sections(me)
        L += [f"Run: `{me}`", "", "| Section | Published paras | Run paras | Similarity |", "|---|---|---|---|"]
        sims = {}
        for sec in list(a) + [s for s in b if s not in a]:
            sa, sb = "\n".join(a.get(sec, [])), "\n".join(b.get(sec, []))
            sims[sec] = round(difflib.SequenceMatcher(None, sa, sb).ratio(), 2)
            L.append(f"| {sec} | {len(a.get(sec, []))} | {len(b.get(sec, []))} | {sims[sec]} |")
        R["sections"]["email"] = sims
        L.append("")
    else:
        L += [f"Not compared (published {os.path.exists(pe)}, run {me})", ""]

    # calendar
    L += ["## Next edition calendar (S6)", ""]
    mc = run_file(run, "Manual File-Calendar.xlsx")
    mc = mc if mc and os.sep + "day" + os.sep not in mc else None
    nc = os.path.join(next_day, "Manual File-Calendar.xlsx") if next_day else None
    if mc and nc and os.path.exists(nc):
        a, b = xlsx_cells(nc), xlsx_cells(mc)
        diff = [k for k in sorted(set(a) | set(b)) if str(a.get(k)) != str(b.get(k))]
        L += [f"Run: `{mc}`; team's next-day file: `{nc}`", "",
              f"- Cells: team {len(a)}, run {len(b)}, different {len(diff)}"]
        for k in diff[:30]:
            L.append(f"  - {k[0]}!{k[1]}: team {short(str(a.get(k)), 40)} | run {short(str(b.get(k)), 40)}")
        R["sections"]["calendar"] = {"different": len(diff)}
    else:
        L += [f"Not compared (run file {mc}, next day file {nc})"]
    L.append("")
    return "\n".join(L), R


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--run", required=True)
    ap.add_argument("--published", required=True)
    ap.add_argument("--next")
    ap.add_argument("--out", required=True)
    ap.add_argument("--json")
    a = ap.parse_args(argv)
    md, summary = compare(a.run, a.published, a.next)
    open(a.out, "w", encoding="utf-8").write(md)
    if a.json:
        json.dump(summary, open(a.json, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(md)
    return 0


if __name__ == "__main__":
    sys.exit(main())
