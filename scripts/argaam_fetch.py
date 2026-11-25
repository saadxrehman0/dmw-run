#!/usr/bin/env python3
"""Argaam news for the unattended run, through a scripted Edge window (build step 5, 25 Nov 2026).

Headless Claude sessions on the runner have no WebFetch and no browser, and Argaam answers plain
scripts with an empty 202 page. A Playwright-driven Edge (or Chrome) gets through, so this script
opens www.argaam.com in one and runs the fetches inside the page (same origin), returning compact
JSON only (rules/argaam.md: never pull whole pages into a session).

  python argaam_fetch.py feed  --out feed.json [--since-id N] [--text ID,ID]
        the live route: browser/argaam_news.js (RSS feeds + list pages, lock and chart rows skipped)
  python argaam_fetch.py range --start "2026-11-23 09:00" --end "2026-11-24 09:00" --out range.json
        every English article published in the window (PKT), found by walking article IDs (they
        rise with time). Used by replays, and by live runs after a long break when page 1 of the
        lists no longer reaches back far enough.
  python argaam_fetch.py text  --ids ID,ID --out text.json
        headline, publish time and body (first --max-body characters) of the given articles

Article pages carry their publish time in Riyadh time (rules/argaam.md, quirks); output times are
given both as "pub_ast" and "pub_pkt" (PKT = AST + 2 h). Rows flagged "lock": true are paywalled
(the article page shows a teaser and a subscribe box) and are skipped by the run, like the lock rows
on the list pages. "special": true only means Argaam's own byline ("Argaam Special"); it is NOT a
paywall flag (25 Nov 2026: 8 of 16 Argaam Special rows on the list page were free).
Browser: Edge on Windows, Chrome on a Mac (no Edge there); --browser or DMW_BROWSER overrides.
Exit 0 on success, 1 when the browser could not reach Argaam.
"""
import argparse, json, os, re, sys
from datetime import datetime, timedelta

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
PKT_MINUS_AST = timedelta(hours=2)

PAGE_JS = r"""
async ({ids, maxBody}) => {
  const clean = s => (s || '').replace(/[‎‏]/g, '').replace(/\s+/g, ' ').trim();
  const one = async id => {
    let h;
    try { h = await (await fetch('/en/article/articledetail/id/' + id, {credentials: 'include'})).text(); }
    catch (e) { return {id, err: String(e)}; }
    const d = new DOMParser().parseFromString(h, 'text/html');
    const h1 = clean(d.querySelector('h1')?.textContent);
    if (!h1 || h1 === 'Page Not Found') return {id, missing: true};
    const pub = (h.match(/(\d{1,2}\/\d{1,2}\/\d{4} \d{1,2}:\d{2}:\d{2} [AP]M)/) || [])[1] || null;
    const art = d.querySelector('div.ck-editor.article');
    const head = clean((d.querySelector('h1')?.closest('section, article, div')?.textContent) || '').slice(0, 600);
    // lock = paywalled for a free reader (the page swaps the body for a teaser and a subscribe box).
    // "Argaam Special" is only Argaam's own byline: half of those rows are free (checked 25 Nov 2026).
    return {id, h: h1, pub, body: clean(art?.textContent).slice(0, maxBody),
            lock: !!d.querySelector('.restricted-article, .subscription-package-article, .locked-article'),
            special: /argaam special/i.test(head)};
  };
  const out = [];
  for (let i = 0; i < ids.length; i += 25) out.push(...await Promise.all(ids.slice(i, i + 25).map(one)));
  return out;
}
"""

FEED_RSS_JS = r"""
async () => {
  const t = await (await fetch('/en/rss/ho-main-news?sectionid=1524')).text();
  const ids = [...t.matchAll(/articledetail\/id\/(\d+)/g)].map(m => +m[1]);
  return ids.length ? Math.max(...ids) : null;
}
"""


def parse_pub(s):
    return datetime.strptime(s, "%m/%d/%Y %I:%M:%S %p") if s else None


class Argaam:
    def __init__(self, channel="msedge", headless=True):
        from playwright.sync_api import sync_playwright
        self._p = sync_playwright().start()
        self.b = self._p.chromium.launch(channel=channel, headless=headless)
        self.pg = self.b.new_page()
        self.pg.goto("https://www.argaam.com/en", timeout=90000, wait_until="domcontentloaded")
        self.pg.wait_for_timeout(4000)
        # The home page reloads itself every few minutes, which kills a long in-page fetch
        # ("Execution context was destroyed", Mac test 25 Nov 2026). Park the tab on the RSS feed
        # (same origin, no scripts) before running fetches there.
        self.pg.goto("https://www.argaam.com/en/rss/ho-main-news?sectionid=1524", timeout=90000,
                     wait_until="domcontentloaded")

    def _eval(self, js, arg=None):
        for attempt in range(3):
            try:
                return self.pg.evaluate(js, arg)
            except Exception as e:  # noqa: BLE001
                if "context was destroyed" not in str(e) or attempt == 2:
                    raise
                self.pg.wait_for_load_state("domcontentloaded")

    def close(self):
        self.b.close()
        self._p.stop()

    def articles(self, ids, max_body=1500):
        rows = self._eval(PAGE_JS, {"ids": [int(i) for i in ids], "maxBody": max_body})
        for r in rows:
            t = parse_pub(r.get("pub"))
            if t:
                r["pub_ast"] = t.strftime("%Y-%m-%d %H:%M")
                r["pub_pkt"] = (t + PKT_MINUS_AST).strftime("%Y-%m-%d %H:%M")
        return rows

    def latest_id(self):
        return self._eval(FEED_RSS_JS)

    def feed(self, since_id=0, text=()):
        js = open(os.path.join(REPO, "browser", "argaam_news.js"), encoding="utf-8").read()
        js = js[: js.rfind("await argaamNews(CFG)")]   # define window.argaamNews only
        self._eval("async () => {" + js + "}")
        return self._eval("cfg => window.argaamNews(cfg)", {"sinceId": since_id, "text": list(text)})

    # ---- ID walk for a time window
    def _time_at(self, i, step=1, tries=12):
        """Publish time (AST) of the first real article at or after id i (IDs have gaps).
        English IDs can have gaps longer than 12 (IDs are shared with Arabic articles): a probe
        that found nothing looks 60 IDs further before giving up. A None here used to be read as
        "before the window" and the search then skipped the whole window (25 Nov 2026, Mac test:
        0 items for 23-24 Nov)."""
        for lo, n in ((i, tries), (i + step * tries, 60)):
            rows = self.articles(range(lo, lo + step * n, step), max_body=0)
            for r in rows:
                if r.get("pub_ast"):
                    return datetime.strptime(r["pub_ast"], "%Y-%m-%d %H:%M"), r["id"]
        return None, None

    def first_id_at_or_after(self, when_ast, hi):
        """Binary search on article IDs for the first one published at or after when_ast."""
        lo = hi - 400
        while True:                     # widen until lo is before the window
            t, _ = self._time_at(lo)
            if t is None or t < when_ast or lo < 1000:
                break
            lo -= 800
        while hi - lo > 15:
            mid = (lo + hi) // 2
            t, _ = self._time_at(mid)
            if t is None or t < when_ast:
                lo = mid
            else:
                hi = mid
        return lo

    def range(self, start_pkt, end_pkt, max_body=1500):
        start_ast, end_ast = start_pkt - PKT_MINUS_AST, end_pkt - PKT_MINUS_AST
        top = self.latest_id() or 0
        if not top:
            raise RuntimeError("no article IDs in the Argaam feed")
        a = self.first_id_at_or_after(start_ast, top)
        b = self.first_id_at_or_after(end_ast, top) + 30
        rows = [r for r in self.articles(range(a, b), max_body) if r.get("pub_ast")]
        keep, seen = [], set()
        for r in rows:   # the same story can sit under two IDs (news and disclosures)
            if start_pkt.strftime("%Y-%m-%d %H:%M") <= r["pub_pkt"] < end_pkt.strftime("%Y-%m-%d %H:%M") \
                    and (r["h"], r["pub_pkt"]) not in seen:
                seen.add((r["h"], r["pub_pkt"]))
                keep.append(r)
        return {"window_pkt": [f"{start_pkt:%Y-%m-%d %H:%M}", f"{end_pkt:%Y-%m-%d %H:%M}"],
                "ids": [a, b], "checked": b - a,
                "maxId": max([int(r["id"]) for r in keep], default=a), "count": len(keep),
                "items": sorted(keep, key=lambda r: r["pub_pkt"])}


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("mode", choices=["feed", "range", "text"])
    ap.add_argument("--out", required=True)
    ap.add_argument("--since-id", type=int, default=0)
    ap.add_argument("--text", default="")
    ap.add_argument("--ids", default="")
    ap.add_argument("--start"); ap.add_argument("--end")
    ap.add_argument("--max-body", type=int, default=1500)
    ap.add_argument("--browser", default=os.environ.get("DMW_BROWSER") or ("chrome" if sys.platform == "darwin" else "msedge"),
                    help="msedge or chrome (default: Edge on Windows, Chrome on a Mac)")
    ap.add_argument("--headed", action="store_true")
    a = ap.parse_args(argv)
    try:
        g = Argaam(a.browser, headless=not a.headed)
    except Exception as e:  # noqa: BLE001
        print(json.dumps({"error": f"browser could not open Argaam: {e}"}))
        return 1
    try:
        if a.mode == "feed":
            out = g.feed(a.since_id, [x for x in a.text.split(",") if x])
        elif a.mode == "text":
            out = {"items": g.articles([x for x in a.ids.split(",") if x], a.max_body)}
        else:
            f = "%Y-%m-%d %H:%M"
            out = g.range(datetime.strptime(a.start, f), datetime.strptime(a.end, f), a.max_body)
    finally:
        g.close()
    with open(a.out, "w", encoding="utf-8") as fh:
        json.dump(out, fh, ensure_ascii=False, indent=1)
    n = out.get("count", len(out.get("items", out.get("feed", []))))
    print(f"argaam {a.mode}: {n} items -> {a.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
