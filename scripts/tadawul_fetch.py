#!/usr/bin/env python3
"""Tadawul (Saudi Exchange) issuer announcements for the unattended run, through a visible browser
window (25 Nov 2026).

saudiexchange.sa sits behind Akamai: plain scripts (curl, requests) get 403 and the site needs a
Saudi route (Saad's home router VPN; the office server needs its own). A real, visible Edge or
Chrome window gets through, so this script opens the issuer-announcements page in one and calls the
page's own list service from inside it (same origin), returning compact JSON only.

  python tadawul_fetch.py range --start "2026-11-23 09:00" --end "2026-11-24 10:10" --out tad.json [--text]
        every announcement filed in the window (PKT), newest pages first until the window is passed
  python tadawul_fetch.py since --since-id 99400 --out tad.json [--text]
        everything above the last anId the previous run logged (maxId), for the S4 rounds
  python tadawul_fetch.py text --ids 99401:9538,99400:7030 --out text.json [--ar]
        title, filing time and body rows of the given announcements (id:symbol; the symbol is
        needed for the title), with --ar the Arabic version too

Times: Tadawul shows Riyadh time (AST); rows carry "pub_ast" and "pub_pkt" (PKT = AST + 2 h).
Body: each detail page holds an "Element List / Explanation" table; it comes back as "rows"
[[label, value], ...] and as "body" (label: value lines, first --max-body characters). Results
filings carry more tables; they come back cell by cell under "tables" (rules: read figures from
the cells, never from run-together text).
Rows are tagged "kind" by title (results, dividend, assembly, contract, fund, other) as a triage
hint only; the run's own news-selection rules decide.
Exit 0 on success, 1 when the browser could not reach Saudi Exchange (blocked or no Saudi route).
"""
import argparse, json, os, re, sys
from datetime import datetime, timedelta

PKT_MINUS_AST = timedelta(hours=2)
SITE = "https://www.saudiexchange.sa"
LIST_PAGE = SITE + "/wps/portal/saudiexchange/newsandreports/issuer-news/issuer-announcements?locale=en"
DETAIL = ("/wps/portal/saudiexchange/newsandreports/issuer-news/issuer-announcements/"
          "issuer-announcements-details/?anId={id}&anCat=1&cs={cs}&locale={loc}")

# The list page's own loader posts to this portlet resource (see the page's
# ns_..._getAnnouncementList). If the site is rebuilt the ids change: the script then reads the
# url from that function's source before giving up.
LIST_JS = r"""
async ({page, size}) => {
  let url = 'p0/IZ7_5A602H80O0HTC060SG6UT81DI1=CZ6_5A602H80O0HTC060SG6UT81D26=NJgetAnnouncementListData=/';
  const fn = Object.keys(window).find(k => /getAnnouncementList$/.test(k));
  if (fn) { const m = String(window[fn]).match(/url:\s*'([^']*NJgetAnnouncementListData[^']*)'/); if (m) url = m[1]; }
  const body = new URLSearchParams({annoucmentType: '1_-1', symbol: '', sectorDpId: '', searchType: '',
    fromDate: '', toDate: '', datePeriod: '', productType: '', advisorsList: '', textSearch: '',
    pageNumberDb: String(page), pageSize: String(size)});
  const r = await fetch(url, {method: 'POST', body, credentials: 'include'});
  const t = await r.text();
  let j; try { j = JSON.parse(t); } catch (e) { return {status: r.status, err: 'not JSON: ' + t.slice(0, 120)}; }
  return {status: r.status, total: j.totalCount,
          rows: (j.announcementList || []).map(x => ({id: x.PRESS_REL_ID, sym: String(x.SYMBOL || '').trim(),
            name: String(x.TITLE || '').trim(), h: String(x.SHORT_DESC || '').replace(/\s+/g, ' ').trim(),
            t: x.newsDateStr, url: x.announcementUrl}))};
}
"""

DETAIL_JS = r"""
async ({items, maxBody, tmpl}) => {
  const clean = s => (s || '').replace(/[‎‏]/g, '').replace(/\s+/g, ' ').trim();
  const skipH = /^(log in|forgot password|prices are delayed|previous announcements|add ?to ?watchlist|.*portlet.*|.*v2$|تسجيل الدخول|نسيت كلمة المرور|تنويه|الإعلانات السابقة)/i;
  const one = async ({id, cs, loc}) => {
    let h;
    const u = tmpl.replace('{id}', id).replace('{cs}', cs || '').replace('{loc}', loc);
    try { h = await (await fetch(u, {credentials: 'include'})).text(); } catch (e) { return {id, loc, err: String(e)}; }
    const d = new DOMParser().parseFromString(h, 'text/html');
    const tables = [...d.querySelectorAll('table')];
    if (!tables.length) return {id, loc, missing: true};
    // the title is the first heading after the details portlet's own name (an h3 on 25 Nov 2026)
    const hs = [...d.querySelectorAll('h1, h2, h3')].map(x => clean(x.textContent));
    const at = hs.findIndex(s => /AnnouncementDetails.*Portlet/i.test(s));
    const title = hs.slice(at + 1).find(s => s && !skipH.test(s)) || '';
    const pub = (h.match(/(\d{2}\/\d{2}\/\d{4} \d{2}:\d{2}:\d{2})/) || [])[1] || null;
    const cells = tr => [...tr.querySelectorAll('th, td')].map(c => clean(c.textContent));
    const main = d.querySelector('table.stacktable') || tables[0];
    const rows = [...main.querySelectorAll('tr')].map(cells).filter(r => r.length >= 2 && r.join('') && !/^element list$/i.test(r[0]) && !/^(ال)?بند$/.test(r[0]));
    const body = rows.map(r => r[0] + ': ' + r.slice(1).join(' | ')).join('\n').slice(0, maxBody);
    const more = tables.filter(t => t !== main && !main.contains(t)).slice(0, 6)
                       .map(t => [...t.querySelectorAll('tr')].map(cells).filter(r => r.join('')).slice(0, 40));
    return {id, loc, h: title, pub, rows, body, tables: more};
  };
  const out = [];
  for (let i = 0; i < items.length; i += 6) out.push(...await Promise.all(items.slice(i, i + 6).map(one)));
  return out;
}
"""

KINDS = [
    ("results", r"financial results|interim financial|annual financial|preliminary|نتائج"),
    ("dividend", r"dividend|distribution|أرباح|توزيع"),
    ("assembly", r"general assembly|assembly meeting|الجمعية"),
    ("contract", r"contract|purchase order|award|agreement|memorandum|mou|عقد|اتفاقية|مذكرة"),
    ("fund", r"\b(fund|reit|etf|sukuk etf)\b|صندوق"),
]


def kind_of(text):
    for k, pat in KINDS:
        if re.search(pat, text or "", re.I):
            return k
    return "other"


def parse_t(s):
    return datetime.strptime(s, "%d/%m/%Y %H:%M:%S") if s else None


def default_browser():
    return os.environ.get("DMW_BROWSER") or ("chrome" if sys.platform == "darwin" else "msedge")


class Tadawul:
    def __init__(self, channel=None, headless=False):
        from playwright.sync_api import sync_playwright
        self._p = sync_playwright().start()
        # Visible window by default: Akamai turns away headless browsers (25 Nov 2026).
        self.b = self._p.chromium.launch(channel=channel or default_browser(), headless=headless,
                                         args=["--disable-blink-features=AutomationControlled"])
        self.ctx = self.b.new_context(locale="en-US", viewport={"width": 1280, "height": 900})
        self.pg = self.ctx.new_page()
        self.pg.goto(LIST_PAGE, timeout=120000, wait_until="domcontentloaded")
        self.pg.wait_for_timeout(5000)
        if "saudiexchange.sa" not in self.pg.url or not self.pg.evaluate(
                "() => Object.keys(window).some(k => /getAnnouncementList$/.test(k))"):
            title = self.pg.title()
            self.close()
            raise RuntimeError(f"announcements page did not load (title {title!r}): blocked, or no Saudi route?")

    def close(self):
        try:
            self.b.close()
        finally:
            self._p.stop()

    def page(self, n, size=50):
        r = self.pg.evaluate(LIST_JS, {"page": n, "size": size})
        if r.get("err") or r.get("status") != 200:
            raise RuntimeError(f"list service: {r}")
        for x in r["rows"]:
            t = parse_t(x.pop("t"))
            if t:
                x["pub_ast"] = t.strftime("%Y-%m-%d %H:%M:%S")
                x["pub_pkt"] = (t + PKT_MINUS_AST).strftime("%Y-%m-%d %H:%M:%S")
            x["kind"] = kind_of(x["h"])
            x["url"] = SITE + x["url"] if x.get("url", "").startswith("/") else x.get("url")
        return r["rows"]

    def walk(self, stop, max_pages=40):
        """List pages newest first until stop(row) is true for a row (that row and older dropped)."""
        out, seen = [], set()
        for n in range(1, max_pages + 1):
            rows = self.page(n)
            if not rows:
                break
            for x in rows:
                if stop(x):
                    return out
                if x["id"] not in seen:
                    seen.add(x["id"])
                    out.append(x)
        return out

    def details(self, pairs, loc="en", max_body=2500):
        items = [{"id": int(i), "cs": cs, "loc": loc} for i, cs in pairs]
        rows = self.pg.evaluate(DETAIL_JS, {"items": items, "maxBody": max_body, "tmpl": DETAIL})
        for r in rows:
            t = parse_t(r.get("pub"))
            if t:
                r["pub_ast"] = t.strftime("%Y-%m-%d %H:%M:%S")
                r["pub_pkt"] = (t + PKT_MINUS_AST).strftime("%Y-%m-%d %H:%M:%S")
        return rows

    def add_text(self, rows, ar=False, max_body=2500):
        pairs = [(x["id"], x["sym"]) for x in rows]
        langs = ["en", "ar"] if ar else ["en"]
        for loc in langs:
            got = {d["id"]: d for d in self.details(pairs, loc, max_body)}
            for x in rows:
                d = got.get(x["id"], {})
                sfx = "" if loc == "en" else "_ar"
                if loc == "ar":
                    x["h_ar"] = d.get("h")
                else:
                    x["h"] = x.get("h") or d.get("h")
                    for k in ("pub_ast", "pub_pkt"):
                        x.setdefault(k, d.get(k))
                x["body" + sfx] = d.get("body", "")
                x["rows" + sfx] = d.get("rows", [])
                if d.get("tables"):
                    x["tables" + sfx] = d["tables"]
        return rows


def summary(rows, extra):
    return {**extra, "maxId": max([x["id"] for x in rows], default=extra.get("sinceId", 0)),
            "count": len(rows), "items": sorted(rows, key=lambda x: (x.get("pub_ast", ""), x["id"]))}


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("mode", choices=["range", "since", "text"])
    ap.add_argument("--out", required=True)
    ap.add_argument("--start"); ap.add_argument("--end")
    ap.add_argument("--since-id", type=int, default=0)
    ap.add_argument("--ids", default="", help="id:symbol,id:symbol (text mode)")
    ap.add_argument("--text", action="store_true", help="also fetch each item's body (range/since)")
    ap.add_argument("--ar", action="store_true", help="also fetch the Arabic version")
    ap.add_argument("--max-body", type=int, default=2500)
    ap.add_argument("--browser", default=None, help="msedge or chrome (default: Edge on Windows, Chrome on Mac)")
    ap.add_argument("--headless", action="store_true", help="for tests only; Saudi Exchange blocks it")
    a = ap.parse_args(argv)
    try:
        g = Tadawul(a.browser, headless=a.headless)
    except Exception as e:  # noqa: BLE001
        print(json.dumps({"error": f"browser could not open Saudi Exchange: {e}"}))
        return 1
    try:
        if a.mode == "range":
            f = "%Y-%m-%d %H:%M"
            s_pkt, e_pkt = datetime.strptime(a.start, f), datetime.strptime(a.end, f)
            lo = (s_pkt - PKT_MINUS_AST).strftime("%Y-%m-%d %H:%M:%S")
            hi = (e_pkt - PKT_MINUS_AST).strftime("%Y-%m-%d %H:%M:%S")
            rows = [x for x in g.walk(lambda x: x.get("pub_ast", "") < lo) if x.get("pub_ast", "") < hi]
            if a.text:
                g.add_text(rows, a.ar, a.max_body)
            out = summary(rows, {"window_pkt": [a.start, a.end]})
        elif a.mode == "since":
            rows = g.walk(lambda x: x["id"] <= a.since_id)
            if a.text:
                g.add_text(rows, a.ar, a.max_body)
            out = summary(rows, {"sinceId": a.since_id})
        else:
            pairs = [p.split(":") + [""] for p in a.ids.split(",") if p]
            rows = [{"id": int(p[0]), "sym": p[1]} for p in pairs]
            g.add_text(rows, a.ar, a.max_body)
            out = {"items": rows}
    finally:
        g.close()
    with open(a.out, "w", encoding="utf-8") as fh:
        json.dump(out, fh, ensure_ascii=False, indent=1)
    print(f"tadawul {a.mode}: {out.get('count', len(out['items']))} items -> {a.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
