#!/usr/bin/env python3
"""Next edition's calendar rows for Manual File-Calendar.xlsx: Saudi Exchange dividends and general
assembly calendars plus the investing.com economic calendar (US and Saudi Arabia), written in the
rows.json format build_cal.py reads (26 Nov 2026).

In the Nov 2026 replays every S6 stage wrote its own one-off Playwright scripts for this (fetch_sx.py,
fetch_inv.py, make_rows.py). This script is those scripts made one tool, with the rules of
rules/calendar.md applied in code.

  python cal_fetch.py --edition 2026-11-25 --out cal.json --rows rows.json \
         [--after 2026-11-24] [--names state/Mapping.xlsx] [--prev-cal "Manual File-Calendar.xlsx"] \
         [--raw raw.json] [--browser chrome|msedge] [--skip sx|inv]
  python cal_fetch.py --edition 2026-11-25 --from-raw raw.json --out cal.json --rows rows.json ...
        rebuild from an earlier fetch with no browser (tests, replays, a second look)

--edition  the edition the calendar is for (the NEXT edition when S6 runs).
--after    the edition before it (default: the previous trading day, Fridays, Saturdays and
           orchestrator/holidays.json skipped). Corporate events dated after it and up to --edition
           belong to the edition, so Friday and Saturday events roll into Sunday's file.

Browser: a visible Chrome (Mac) or Edge (Windows) window, as tadawul_fetch.py and argaam_fetch.py.
saudiexchange.sa turns away headless browsers and non-Saudi routes (Saad's home router VPN works;
the office server needs its own Saudi route). investing.com is read through its calendar data
service from inside its own page, with the page's __NEXT_DATA__ as a fallback.

Output (--out): {"rows": <rows.json>, "sources": [one line per row], "review": [items the S6
agent must settle], "dropped": [...], "days": {...}, "fetch": {...}}. --rows writes the rows part
alone, ready for build_cal.py. A value nobody could source is "TK" and is listed under "review":
the stage fills it from the announcement or the latest final decks, or leaves it TK and gates it.

What the script decides (rules/calendar.md):
- Event days: dividend eligibility/distribution and general assemblies dated in (after, edition].
  None: the next day after the edition that has any (empty-day rule), real dates kept.
- Dividends: one row per eligibility and distribution date that falls on a chosen day, amount to
  two decimals (Aramco 0.3393 -> 0.34, logged), house forms in EN and AR. Distribution rows first,
  then eligibility, then assemblies.
- General assemblies: every assembly on a chosen day has its announcement read. A capital hike or
  cut, or a share buyback in the agenda makes a row (figure read from the announcement, two
  decimals, mn or bn); "Extraordinary - Corporate Actions" with no figure found makes a TK row
  under review; anything else is dropped and logged.
- Corporate actions tab: bonus, capital and rights rows whose eligibility date is a chosen day are
  listed under review as a cross-check (the tab is a log, not a calendar; calendar.md).
- Economic: US and KSA rows dated on the edition date (US rows by their US Eastern date, KSA rows by
  Riyadh date); none: the next day that has any, skipping a borrowed day with only 1-star rows and
  no figures. Highest star tier present (3, else 2, else 1). Prior = the previous figure as first
  published (investing.com's "revised from" when shown, else "previous"); Estimation = forecast.
  The API crude row stands on every Tuesday; API and EIA rows carry no figures. Wording from
  data/cal_wording.json, then the previous calendar file (--prev-cal), else EN by rule and AR TK.
- Not covered here (the stage still checks them by hand): Argaam's calendar for earnings calls,
  listing and index events and KSA releases investing.com lacks (GASTAT, SAMA); a dividend row with a
  blank eligibility or distribution date (its dates come from the announcement); Bloomberg dates.

Exit 0 when the rows were written, 1 when a site could not be read (blocked or no Saudi route).
"""
import argparse, json, os, re, sys
from datetime import date, datetime, timedelta, timezone
from decimal import Decimal, ROUND_HALF_UP
from fractions import Fraction

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
sys.path.insert(0, HERE)

SX = "https://www.saudiexchange.sa"
SX_CAL = SX + "/wps/portal/saudiexchange/newsandreports/issuer-financial-calendars/"
TABS = {"dividends": "", "assemblies": "general-assembly-meetings", "actions": "corporate-actions"}
INV = "https://www.investing.com/economic-calendar/"
INV_COUNTRIES = {"5": "US", "52": "KSA"}
PKT = timezone(timedelta(hours=5))
AST = timezone(timedelta(hours=3))
WORDING = os.path.join(REPO, "data", "cal_wording.json")
HOLIDAYS = os.path.join(REPO, "orchestrator", "holidays.json")
ISO = re.compile(r"^\d{4}-\d{2}-\d{2}$")
DMY = re.compile(r"^(\d{2})/(\d{2})/(\d{4})$")

# ----------------------------------------------------------------------------------------------
# dates


def _holidays(path=HOLIDAYS):
    try:
        h = json.load(open(path, encoding="utf-8"))
    except (OSError, ValueError):
        return set()
    out = set()
    for k, v in h.items():
        if k.isdigit():
            for d in v:
                out.add(d if isinstance(d, str) else d.get("date"))
    return out


def trading_day(d, holidays=None):
    return d.weekday() not in (4, 5) and d.isoformat() not in (holidays or set())


def previous_edition(ed, holidays=None):
    d = ed - timedelta(days=1)
    while not trading_day(d, holidays):
        d -= timedelta(days=1)
    return d


def days(a, b):
    """ISO dates from a to b inclusive."""
    out, d = [], a
    while d <= b:
        out.append(d.isoformat()); d += timedelta(days=1)
    return out


def dmy(iso):
    return f"{iso[8:10]}-{iso[5:7]}-{iso[0:4]}"


def us_eastern(utc):
    """US Eastern wall time for a UTC datetime, DST by the US rule (2nd Sunday of March 02:00 to
    1st Sunday of November 02:00), so Windows needs no tzdata package."""
    y = utc.year
    mar = date(y, 3, 8 + (6 - date(y, 3, 8).weekday()) % 7)      # second Sunday of March
    nov = date(y, 11, 1 + (6 - date(y, 11, 1).weekday()) % 7)     # first Sunday of November
    start = datetime(y, mar.month, mar.day, 7, tzinfo=timezone.utc)   # 02:00 EST
    end = datetime(y, nov.month, nov.day, 6, tzinfo=timezone.utc)     # 02:00 EDT
    off = -4 if start <= utc < end else -5
    return utc.astimezone(timezone(timedelta(hours=off)))


# ----------------------------------------------------------------------------------------------
# numbers and wording


def q2(v):
    return str(Decimal(str(v)).quantize(Decimal("0.01"), ROUND_HALF_UP))


def num(s):
    s = re.sub(r"[^\d.]", "", str(s or ""))
    return Decimal(s) if s and s != "." else None


def money(sar):
    """SAR amount -> ('300.00', 'mn') or ('15.00', 'bn'); two decimals (calendar.md, 27 Sep 2026)."""
    sar = Decimal(sar)
    if sar >= Decimal(10) ** 9:
        return q2(sar / Decimal(10) ** 9), "bn"
    return q2(sar / Decimal(10) ** 6), "mn"


AR_UNIT = {"mn": "مليون", "bn": "مليار"}


def div_text(kind, amt):
    if kind == "dist":
        return f"Cash Dividend Distribution (SAR {amt}/sh)", f"صرف أرباح نقدية (ريال سعودي {amt}/للسهم)"
    return f"Eligibility for Cash Dividend (SAR {amt}/sh)", f"أحقية أرباح نقدية (ريال سعودي {amt}/للسهم)"


def egm_text(act):
    k = act["kind"]
    if k in ("hike", "cut"):
        v, u = money(act["after"])
        if k == "hike":
            return (f"Extraordinary General Meeting (Capital Hike to SAR {v} {u})",
                    f"جمعية عمومية غير عادية (زيادة رأس المال إلى {v} {AR_UNIT[u]} ريال سعودي)")
        return (f"Extraordinary General Meeting (Capital Reduction to SAR {v} {u})",
                f"جمعية عمومية غير عادية (تخفيض رأس المال إلى {v} {AR_UNIT[u]} ريال سعودي)")
    if k == "buyback":
        v = q2(Decimal(act["shares"]) / Decimal(10) ** 6)
        # the 5 Nov 2026 wording (task-03 SKILL.md); the AR is Claude's, named to Saad then
        return (f"Extraordinary General Meeting (Share Buyback of up to {v} mn shares)",
                f"جمعية عمومية غير عادية (شراء ما يصل إلى {v} مليون سهم من أسهم الشركة)")
    return "Extraordinary General Meeting (TK)", "جمعية عمومية غير عادية (TK)"


def _flat(t):
    return re.sub(r"\s+", " ", str(t or "")).replace("’", "'").strip()


def _sar(n, unit=None):
    v = num(n)
    if v is None:
        return None
    if unit and unit.lower().startswith("bil"):
        v *= Decimal(10) ** 9
    elif unit and unit.lower().startswith("mil"):
        v *= Decimal(10) ** 6
    return v


def corporate_action(text):
    """The corporate action in an assembly announcement: {'kind': 'hike'|'cut', 'after': SAR,
    'before': SAR|None} or {'kind': 'buyback', 'shares': n}, else None. Read from the
    announcement's own words; nothing is inferred."""
    t = _flat(text)
    m = re.search(r"capital after the (increase|reduction|decrease)\s*:?\s*(?:SAR|SR)?\s*([\d,]+(?:\.\d+)?)"
                  r"\s*(million|billion)?", t, re.I)
    if m:
        before = re.search(r"capital before the (?:increase|reduction|decrease)\s*:?\s*(?:SAR|SR)?\s*"
                           r"([\d,]+(?:\.\d+)?)\s*(million|billion)?", t, re.I)
        return {"kind": "hike" if m.group(1).lower() == "increase" else "cut",
                "after": _sar(m.group(2), m.group(3)), "before": before and _sar(*before.groups())}
    m = re.search(r"(increas|reduc|decreas)\w*\s+(?:of\s+)?(?:the\s+)?(?:company's\s+)?(?:share\s+)?capital\s+"
                  r"from\s+(?:SAR|SR)\s*([\d,]+(?:\.\d+)?)\s*(million|billion)?.{0,240}?\bto\s+(?:SAR|SR)\s*"
                  r"([\d,]+(?:\.\d+)?)\s*(million|billion)?", t, re.I)
    if m:
        return {"kind": "hike" if m.group(1).lower() == "increas" else "cut",
                "after": _sar(m.group(4), m.group(5)), "before": _sar(m.group(2), m.group(3))}
    m = re.search(r"(?:buy[- ]?back|repurchas\w*|purchas\w*)\b.{0,120}?up to\s*\(?\s*([\d,]{4,})\s*\)?"
                  r"(?=.{0,60}?shares)", t, re.I)
    if m:
        return {"kind": "buyback", "shares": int(m.group(1).replace(",", ""))}
    return None


ACTION_WORDS = re.compile(r"capital (?:increase|reduction|decrease)|(?:increase|reduce|decrease) (?:the )?(?:company's )?"
                          r"capital|bonus shares?|buy[- ]?back|treasury shares|rights issue", re.I)

# ----------------------------------------------------------------------------------------------
# Saudi Exchange tables -> rows


def _cells(r):
    return [_flat(c) for c in (r.get("c") if isinstance(r, dict) else r)]


def _link(r):
    a = (r.get("a") or []) if isinstance(r, dict) else []
    a = [h for h in a if "anId=" in h]
    return (SX + a[0] if a[0].startswith("/") else a[0]) if a else None


def parse_dividends(rows):
    """Dividends tab rows: [symbol, name, announced, eligibility, method, distribution, amount]."""
    out = []
    for r in rows:
        c = _cells(r)
        if len(c) < 7 or not re.fullmatch(r"\d{4}", c[0]):
            continue
        if not ISO.match(c[2] or "") or any(x and not ISO.match(x) for x in (c[3], c[5])):
            raise ValueError(f"dividends tab layout changed: {c}")
        out.append({"sym": c[0], "name": c[1], "announced": c[2], "elig": c[3] or None,
                    "method": c[4], "dist": c[5] or None, "amount": c[6].lstrip("^").replace(",", "")})
    return out


def parse_assemblies(rows):
    """General assembly rows. The cells came as [symbol, name, type, date, time, site, status] on
    24 Nov 2026 while the header row read Symbol, Company Name, Type Of Assembly, Holding, Status,
    Date, Time, Site, so each field is found by its form, not its position."""
    out = []
    for r in rows:
        c = _cells(r)
        if len(c) < 4 or not re.fullmatch(r"\d{4}", c[0]):
            continue
        rest = c[2:]
        typ = next((x for x in rest if re.match(r"(extra)?ordinary", x, re.I)), None)
        day = next((x for x in rest if ISO.match(x)), None)
        if not typ or not day:
            raise ValueError(f"general assembly tab layout changed: {c}")
        status = next((x for x in rest if re.search(r"conven|postpon|cancel|adjourn|held", x, re.I) and len(x) < 40), "")
        tm = next((x for x in rest if re.fullmatch(r"\d{1,2}:\d{2}", x)), "")
        out.append({"sym": c[0], "name": c[1], "type": typ, "date": day, "time": tm, "status": status,
                    "url": _link(r)})
    return out


def _dmy_iso(s):
    m = DMY.match(s or "")
    return f"{m.group(3)}-{m.group(2)}-{m.group(1)}" if m else None


def parse_actions(rows):
    """Corporate actions tab: [symbol, short name, recommended dd/mm/yyyy, type, eligibility
    dd/mm/yyyy, new capital, previous capital]."""
    out = []
    for r in rows:
        c = _cells(r)
        if len(c) < 7 or not re.fullmatch(r"\d{4}", c[0]):
            continue
        out.append({"sym": c[0], "short": c[1], "recommended": _dmy_iso(c[2]), "type": c[3],
                    "elig": _dmy_iso(c[4]), "new": num(c[5]), "prev": num(c[6])})
    return out


def market_ok(rows, market):
    """Did the board switch take? Nomu codes are 9xxx; Main Market lists hold others (sx_04 on
    4 Nov 2026 read Nomu rows for both boards when the switch did not take)."""
    syms = [_cells(r)[0] for r in rows if _cells(r) and re.fullmatch(r"\d{4}", _cells(r)[0])]
    if not syms:
        return True
    nomu = sum(s.startswith("9") for s in syms)
    return nomu == len(syms) if market == "S" else nomu < len(syms)


# ----------------------------------------------------------------------------------------------
# investing.com -> rows

PERIOD = re.compile(r"\s*\((?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec|Q[1-4]|H[12]|\d{4})\)\s*$")


def ev_name(s):
    return PERIOD.sub("", re.sub(r"\s+", " ", s or "").strip()).strip()


def _strip_tags(h):
    import html as _h
    return _h.unescape(re.sub(r"<[^>]+>", "", h or "")).replace("\xa0", " ").strip()


def parse_inv_service(data_html):
    """Rows of investing.com's getCalendarFilteredData HTML (timeZone 77 = PKT)."""
    out = []
    for tr in re.findall(r"<tr\b[^>]*>.*?</tr>", data_html or "", re.S):
        dt_ = re.search(r'data-event-datetime="([^"]+)"', tr)
        if not dt_:
            continue
        def td(cls, raw=False):
            m = re.search(r'<td[^>]*class="[^"]*\b' + cls + r'\b[^"]*"[^>]*>(.*?)</td>', tr, re.S)
            return (m.group(0) if raw else _strip_tags(m.group(1))) if m else ""
        cur = td("flagCur")
        title = re.search(r'<span[^>]*title="([^"]+)"', td("flagCur", True))
        prev_raw = td("prev", True)
        rev = re.search(r'title="Revised From ([^"]+)"', prev_raw)
        pkt = datetime.strptime(dt_.group(1), "%Y/%m/%d %H:%M:%S").replace(tzinfo=PKT)
        eid = re.search(r'event_attr_ID="(\d+)"', tr, re.I)
        out.append({"utc": pkt.astimezone(timezone.utc).isoformat(), "cur": cur.split()[-1] if cur else "",
                    "country": title.group(1) if title else "", "stars": td("sentiment", True).count("grayFullBullishIcon"),
                    "name": ev_name(td("event")), "raw_name": re.sub(r"\s+", " ", td("event")),
                    "actual": td("act"), "forecast": td("fore"), "previous": td("prev"),
                    "revised_from": rev.group(1) if rev else None, "event_id": int(eid.group(1)) if eid else None,
                    "src": "investing.com data service"})
    return out


def parse_inv_next(next_data):
    """Rows of the page's __NEXT_DATA__ (props.pageProps.state.economicCalendarStore), US and KSA."""
    st = next_data["props"]["pageProps"]["state"]["economicCalendarStore"]["calendarEventsByDate"]
    out = []
    for day, evs in st.items():
        for e in evs:
            if e.get("type") != "event" or str(e.get("countryId")) not in INV_COUNTRIES:
                continue
            name = " ".join(x for x in (e.get("event"), e.get("suffix")) if x)
            out.append({"utc": e["time"].replace("Z", "+00:00"), "cur": e.get("currency", ""),
                        "country": e.get("country", ""), "stars": int(e.get("importance") or 0),
                        "name": ev_name(name), "raw_name": " ".join(x for x in (name, e.get("period")) if x),
                        "actual": e.get("actual") or "", "forecast": e.get("forecast") or "",
                        "previous": e.get("previous") or "", "revised_from": e.get("revisedFrom") or None,
                        "event_id": e.get("eventId"), "src": "investing.com __NEXT_DATA__"})
    return out


def local_date(r):
    utc = datetime.fromisoformat(r["utc"]).astimezone(timezone.utc)
    if r["cur"] == "SAR" or "Saudi" in r.get("country", ""):
        return utc.astimezone(AST).date().isoformat(), "KSA"
    return us_eastern(utc).date().isoformat(), "US"


def has_figures(r):
    return bool((r.get("forecast") or "").strip() or (r.get("previous") or "").strip())


def pick_economic_day(rows, edition, horizon=7):
    """The edition date if it has any US/KSA row; else the next date that has one, skipping a day
    whose only rows are 1-star with no figures (calendar.md)."""
    by = {}
    for r in rows:
        by.setdefault(local_date(r)[0], []).append(r)
    d0 = date.fromisoformat(edition)
    for i in range(horizon + 1):
        d = (d0 + timedelta(days=i)).isoformat()
        rs = by.get(d, [])
        if i == 0 and rs:
            return d, rs
        if i and rs and (any(r["stars"] >= 2 for r in rs) or any(has_figures(r) for r in rs)):
            return d, rs
    return None, []


def load_wording(path=WORDING):
    try:
        return json.load(open(path, encoding="utf-8")).get("economic", {})
    except (OSError, ValueError):
        return {}


def prev_cal_pairs(path):
    """EN -> AR economic wording and the blank marker from the previous Manual File-Calendar.xlsx."""
    from dmw_common import cal_sheet, cal_columns
    ws, hr = cal_sheet(path)
    c = cal_columns(ws, hr)
    pairs, blank = {}, None
    for r in range(hr + 1, hr + 30):
        en, ar = ws.cell(r, c["ec_event"]).value, ws.cell(r, c["ec_event_ar"]).value
        if en and ar:
            pairs[_flat(en)] = _flat(ar)
        for k in ("ec_prior", "ec_est"):
            v = ws.cell(r, c[k]).value
            if blank is None and isinstance(v, str) and v.strip() in ("-", "_"):
                blank = v.strip()
    return pairs, blank


def word_economic(r, region, wording, pairs):
    w = wording.get(r["name"])
    if w:
        return w["en"], w["ar"], bool(w.get("nofig")), "data/cal_wording.json"
    en = ("US " if region == "US" else "KSA ") + r["name"]
    if en in pairs:
        return en, pairs[en], False, "previous calendar file"
    return en, "TK", False, None


API = re.compile(r"^API Weekly Crude Oil Stock", re.I)
# client-mandated EN (deck-house-rules.md); AR as ar_glossary.json gives it for that EN
API_ROW = {"en": "US API Weekly Report on Crude Oil Inventories",
           "ar": "تقرير معهد البترول الأمريكي (API) الأسبوعي عن مخزونات النفط الخام"}


def build_economic(rows, edition, wording, pairs):
    """-> (economic rows, sources, review, day)."""
    day, pool = pick_economic_day(rows, edition)
    eco, src, review = [], [], []
    if not day:
        review.append({"what": "economic", "why": f"no US or KSA row on investing.com from {edition} for a week"})
        return eco, src, review, None
    tier = max(r["stars"] for r in pool)
    chosen = [r for r in pool if r["stars"] == tier and not API.match(r["name"])]
    is_tue = date.fromisoformat(day).weekday() == 1
    api = next((r for r in pool if API.match(r["name"])), None)
    src.append(f"investing.com {day} (US Eastern / Riyadh dates): {len(pool)} US+KSA rows, "
               f"tier {tier} used ({len(chosen)} rows)" + (", standing API row added (Tuesday)" if is_tue else ""))
    for r in sorted(chosen, key=lambda r: r["utc"]):
        region = local_date(r)[1]
        en, ar, nofig, wsrc = word_economic(r, region, wording, pairs)
        prior = None if nofig else (r.get("revised_from") or r.get("previous") or None)
        est = None if nofig else (r.get("forecast") or None)
        eco.append({"date": day, "event_en": en, "prior": prior, "estimate": est, "event_ar": ar})
        src.append(f"{en}: prior {prior}, estimate {est} <- {r['src']}, {r['raw_name']} ({r['stars']}-star, "
                   f"{r['utc']} UTC; page previous {r.get('previous') or '-'}, revised from {r.get('revised_from') or '-'})"
                   + ("; no figures by rule (crude rows)" if nofig else ""))
        if ar == "TK":
            review.append({"what": "wording", "row": en, "why": f"no house wording for investing.com '{r['name']}': "
                           "copy the latest final decks' wording into the row and add it to data/cal_wording.json"})
    if is_tue:
        w = wording.get("API Weekly Crude Oil Stock") or API_ROW
        eco.append({"date": day, "event_en": w["en"], "prior": None, "estimate": None, "event_ar": w["ar"]})
        src.append(f"{w['en']}: standing Tuesday row, no figures (calendar.md)" +
                   (f"; investing.com lists it {api['utc']} UTC" if api else "; not on investing.com that day (US holiday?)"))
    return eco, src, review, day


# ----------------------------------------------------------------------------------------------
# assemble


def pick_event_days(cands, after, edition, horizon=14):
    """Dates for the event block: those in (after, edition] that hold a candidate; else the first
    later date that does."""
    own = [d for d in days(date.fromisoformat(after) + timedelta(days=1), date.fromisoformat(edition)) if d in cands]
    if own:
        return own, False
    for d in days(date.fromisoformat(edition) + timedelta(days=1), date.fromisoformat(edition) + timedelta(days=horizon)):
        if d in cands:
            return [d], True
    return [], False


def build(raw, edition, after, names=None, wording=None, pairs=None, blank="-", announced_before=None):
    """Pure: raw fetch (see fetch()) -> the cal.json dict. announced_before (replays): drop dividend
    rows announced on or after that date, which the live run could not have seen."""
    names, pairs = names or {}, pairs or {}
    wording = load_wording() if wording is None else wording
    sources, review, dropped = [], [], []
    sx = raw.get("sx") or {}
    divs = [d for m in ("M", "S") for d in parse_dividends(sx.get("dividends", {}).get(m, {}).get("rows", []))]
    gas = [g for m in ("M", "S") for g in parse_assemblies(sx.get("assemblies", {}).get(m, {}).get("rows", []))]
    acts = [a for m in ("M", "S") for a in parse_actions(sx.get("actions", {}).get(m, {}).get("rows", []))]
    for tab in ("dividends", "assemblies", "actions"):
        for m in ("M", "S"):
            t = sx.get(tab, {}).get(m)
            if t and t.get("market_ok") is False:
                review.append({"what": "fetch", "why": f"Saudi Exchange {tab} tab: the {'Nomu' if m == 'S' else 'Main Market'} "
                               "switch did not take (rows from the other board); read that board by hand"})
            if t and t.get("paged"):
                review.append({"what": "fetch", "why": f"Saudi Exchange {tab} tab ({m}) is paged and only the first page was read"})
    if announced_before:
        late = [d for d in divs if d["announced"] >= announced_before]
        dropped += [f"{d['sym']} {d['name']}: dividend announced {d['announced']}, after the replay cut-off" for d in late]
        divs = [d for d in divs if d["announced"] < announced_before]
    # de-duplicate (the same row can come back for both boards)
    seen = set(); divs = [d for d in divs if not ((k := tuple(d.values())) in seen or seen.add(k))]
    seen = set(); gas = [g for g in gas if not ((k := (g["sym"], g["date"], g["type"])) in seen or seen.add(k))]

    cands = {}
    for d in divs:
        for k in ("dist", "elig"):
            if d[k]:
                cands.setdefault(d[k], []).append(("div", k, d))
        # a blank date could still fall on the edition: flag it when the earliest it could be
        # (its eligibility date, else its announcement) is not after the edition, and it is recent
        low = d["elig"] or d["announced"]
        recent = d["announced"] >= (date.fromisoformat(edition) - timedelta(days=90)).isoformat()
        if (not d["elig"] or not d["dist"]) and low <= edition and recent:
            review.append({"what": "dividend dates", "row": f"{d['sym']} {d['name']}",
                           "why": "blank eligibility or distribution date on the dividends tab: take the dates from its announcement"})
    ann = raw.get("announcements", {})
    for g in gas:
        cands.setdefault(g["date"], []).append(("ga", g["type"], g))
    ev_days, borrowed = pick_event_days(cands, after, edition)
    if borrowed:
        sources.append(f"event calendar: nothing dated {after} < day <= {edition}; borrowed {ev_days[0]} (empty-day rule)")

    def name(sym, fallback):
        n = names.get(sym)
        if n:
            return n["en"], n["ar"]
        review.append({"what": "name", "row": f"{sym} {fallback}", "why": "code not in Mapping.xlsx: Arabic name TK"})
        return fallback, "TK"

    events = {"dist": [], "elig": [], "ga": []}
    for day in ev_days:
        for kind, sub, x in cands.get(day, []):
            if kind == "div":
                amt = q2(x["amount"])
                en, ar = div_text(sub, amt)
                cen, car = name(x["sym"], x["name"])
                events[sub].append({"date": day, "company_en": cen, "event_en": en, "company_ar": car, "event_ar": ar})
                sources.append(f"{x['sym']} {cen}: {en} <- Saudi Exchange dividends tab: announced {x['announced']}, "
                               f"eligibility {x['elig']}, distribution {x['dist']}, amount {x['amount']}"
                               + (f", written {amt} (two decimals)" if amt != x["amount"] else ""))
                continue
            g = x
            a = ann.get(g["url"] or "", {})
            text = a.get("text", "")
            act = corporate_action(text) if text else None
            label = f"{g['sym']} {g['name']} {g['type']} assembly {g['date']} ({g['status'] or 'status ?'})"
            if re.search(r"postpon|cancel|adjourn", g["status"], re.I):
                review.append({"what": "assembly status", "row": label, "why": "postponed or cancelled: check before listing"})
            if act:
                en, ar = egm_text(act)
                cen, car = name(g["sym"], g["name"])
                events["ga"].append({"date": day, "company_en": cen, "event_en": en, "company_ar": car, "event_ar": ar})
                if act["kind"] in ("hike", "cut"):
                    before = "{:,.0f}".format(act["before"]) if act.get("before") else "?"
                    fig = "capital {} -> {:,.0f}".format(before, act["after"])
                else:
                    fig = "buyback of up to {:,} shares".format(act["shares"])
                sources.append(f"{g['sym']} {cen}: {en} <- Saudi Exchange GA tab ({g['type']}, {g['date']} {g['time']}) + "
                               f"announcement {g['url']}: {fig}")
                if act["kind"] in ("hike", "cut") and act.get("before") and act["kind"] == "hike" and act["after"] <= act["before"]:
                    review.append({"what": "figure", "row": label, "why": "capital after is not above capital before"})
            elif g["type"].lower().startswith("extraordinary - corporate") or (text and ACTION_WORDS.search(_flat(text))):
                cen, car = name(g["sym"], g["name"])
                en, ar = egm_text({"kind": "?"})
                events["ga"].append({"date": day, "company_en": cen, "event_en": en, "company_ar": car, "event_ar": ar})
                review.append({"what": "assembly figure", "row": label,
                               "why": ("announcement not read" if not text else "corporate action in the agenda but no figure found")
                               + f": read {g['url']} and write the figure, or drop the row if the agenda is routine"})
                sources.append(f"{g['sym']} {cen}: {en} <- Saudi Exchange GA tab ({g['type']}); figure TK")
            else:
                dropped.append(f"{label}: no corporate action in the agenda" + ("" if text else " (announcement not read)"))
    for a in acts:
        if a["elig"] in ev_days:
            ratio = ""
            if a["new"] and a["prev"] and a["new"] > a["prev"] and re.search(r"bonus", a["type"], re.I):
                f = Fraction(int(a["new"] - a["prev"]), int(a["prev"]))
                ratio = f" (bonus ratio {f.numerator}:{f.denominator}; Eligibility for Bonus Shares ({f.numerator}:{f.denominator}))"
            review.append({"what": "corporate action", "row": f"{a['sym']} {a['short']}: {a['type']} eligibility {a['elig']}",
                           "why": f"on the Corporate Actions tab, capital {a['prev']} -> {a['new']}{ratio}: add a row if the "
                                  "event calendar lacks it (cross-check, calendar.md)"})

    eco, esrc, erev, eday = build_economic(raw.get("inv", {}).get("rows", []), edition, wording, pairs) \
        if raw.get("inv") else ([], [], [{"what": "economic", "why": "investing.com not fetched"}], None)
    sources += esrc; review += erev
    for e in raw.get("errors", []):
        review.append({"what": "fetch", "why": e})
    review.append({"what": "by hand", "why": "Argaam calendar for earnings calls, listing and index events and KSA "
                   "releases investing.com lacks (GASTAT, SAMA); Bloomberg dates when sources disagree (calendar.md)"})
    rows = {"blank": blank, "events": events["dist"] + events["elig"] + events["ga"], "economic": eco}
    return {"edition": edition, "after": after, "days": {"events": ev_days, "events_borrowed": borrowed, "economic": eday},
            "rows": rows, "sources": sources, "review": review, "dropped": dropped,
            "fetch": {k: v for k, v in raw.items() if k in ("fetched_at", "errors")}}


# ----------------------------------------------------------------------------------------------
# browser

ROWS_JS = r"""
() => {
  const clean = s => (s || '').replace(/[‎‏]/g, '').replace(/\s+/g, ' ').trim();
  const t = document.querySelector('#issuerTable');
  if (!t) return {err: 'no #issuerTable'};
  let trs = [...t.querySelectorAll('tbody tr')], via = 'dom';
  try {                                   // a DataTables table hides its other pages from the DOM
    const $ = window.jQuery;
    if ($ && $.fn.dataTable && $.fn.dataTable.isDataTable(t)) { trs = [...$(t).DataTable().rows().nodes()]; via = 'datatables'; }
  } catch (e) {}
  const pager = document.querySelector('.dataTables_paginate .next:not(.disabled), .pagination .next:not(.disabled)');
  return {via, paged: via === 'dom' && !!pager,
          head: [...t.querySelectorAll('thead th')].map(c => clean(c.textContent)),
          rows: trs.map(tr => ({c: [...tr.cells].map(c => clean(c.textContent)),
                                a: [...tr.querySelectorAll('a[href]')].map(a => a.getAttribute('href')).filter(h => /anId=/i.test(h))}))};
}
"""

# set the board (M Main, S Nomu) and the date range, then press the visible search button
# (the 23 Nov 2026 replay's fetch_sx.py; a bare change event did not refresh the table on 4 Nov)
SET_JS = r"""
([m, s, e]) => {
  const $ = window.jQuery; if (!$) return {err: 'no jQuery'};
  if (s) { $('#drp_start').val(s); $('#drp_end').val(e); }
  $('#marketsList').val(m); $('#markethid').val(m); $('#marketsList').trigger('change');
  const btn = [...document.querySelectorAll('button, input[type=button], input[type=submit], a')].find(b =>
    /search|apply|filter/i.test((b.id || '') + ' ' + (b.className || '') + ' ' + (b.textContent || b.value || '')) &&
    !/mobile|reset|searchButton/i.test(b.id || '') && b.offsetParent);
  if (btn) btn.click();
  return {market: $('#marketsList').val(), start: $('#drp_start').val(), end: $('#drp_end').val(), button: !!btn};
}
"""

ANN_JS = r"""
async (urls) => {
  const clean = s => (s || '').replace(/[‎‏]/g, '').replace(/[ \t]+/g, ' ').trim();
  const one = async u => {
    try {
      const h = await (await fetch(u, {credentials: 'include'})).text();
      const d = new DOMParser().parseFromString(h, 'text/html');
      const rows = [...d.querySelectorAll('table tr')].map(tr => [...tr.querySelectorAll('th, td')].map(c => clean(c.textContent)))
                     .filter(r => r.join(''));
      return [u, {rows, text: rows.map(r => r.join(': ')).join('\n').slice(0, 20000)}];
    } catch (e) { return [u, {err: String(e)}]; }
  };
  return Object.fromEntries(await Promise.all(urls.map(one)));
}
"""

INV_JS = r"""
async ([a, b, countries]) => {
  const fd = new URLSearchParams(); countries.forEach(c => fd.append('country[]', c));
  fd.append('dateFrom', a); fd.append('dateTo', b); fd.append('timeZone', '77');
  fd.append('timeFilter', 'timeOnly'); fd.append('currentTab', 'custom'); fd.append('limit_from', '0');
  const r = await fetch('/economic-calendar/Service/getCalendarFilteredData', {method: 'POST',
    headers: {'X-Requested-With': 'XMLHttpRequest', 'Content-Type': 'application/x-www-form-urlencoded'}, body: fd.toString()});
  if (!r.ok) return {status: r.status};
  const j = await r.json();
  return {status: r.status, data: j.data, rows_num: j.rows_num, more: !!j.bind_scroll_handler};
}
"""


def default_browser():
    return os.environ.get("DMW_BROWSER") or ("chrome" if sys.platform == "darwin" else "msedge")


class Browser:
    def __init__(self, channel=None, headless=False, wait_ms=9000):
        from playwright.sync_api import sync_playwright
        self._p = sync_playwright().start()
        # visible window by default: Akamai turns away headless browsers (25 Nov 2026)
        self.b = self._p.chromium.launch(channel=channel or default_browser(), headless=headless,
                                         args=["--disable-blink-features=AutomationControlled"])
        self.ctx = self.b.new_context(locale="en-US", viewport={"width": 1280, "height": 900})
        self.pg = self.ctx.new_page()
        self.wait = wait_ms

    def close(self):
        try:
            self.b.close()
        finally:
            self._p.stop()

    def sx_tab(self, tab, market, start=None, end=None):
        for attempt in (1, 2):
            self.pg.goto(SX_CAL + TABS[tab] + "?locale=en", timeout=120000, wait_until="domcontentloaded")
            self.pg.wait_for_timeout(self.wait)
            if "saudiexchange.sa" not in self.pg.url or not self.pg.query_selector("#issuerTable"):
                raise RuntimeError(f"Saudi Exchange {tab} page did not load (title {self.pg.title()!r}): blocked, or no Saudi route?")
            info = self.pg.evaluate(SET_JS, [market, start, end])
            self.pg.wait_for_timeout(self.wait)
            got = self.pg.evaluate(ROWS_JS)
            if got.get("err"):
                raise RuntimeError(f"Saudi Exchange {tab}: {got['err']}")
            got["set"] = info
            got["market_ok"] = market_ok(got["rows"], market)
            if got["market_ok"]:
                break
        return got

    def announcements(self, urls):
        urls = [u for u in urls if u]
        return self.pg.evaluate(ANN_JS, urls) if urls else {}

    def investing(self, a, b):
        self.pg.goto(INV, timeout=120000, wait_until="domcontentloaded")
        self.pg.wait_for_timeout(min(self.wait, 7000))
        r = self.pg.evaluate(INV_JS, [a, b, list(INV_COUNTRIES)])
        if r.get("status") == 200 and r.get("data"):
            return {"via": "service", "rows": parse_inv_service(r["data"]), "more": r.get("more"), "rows_num": r.get("rows_num")}
        nd = self.pg.evaluate("() => document.getElementById('__NEXT_DATA__')?.textContent || ''")
        rows = parse_inv_next(json.loads(nd)) if nd else []
        return {"via": "__NEXT_DATA__", "service_status": r.get("status"), "rows": rows}


def fetch(edition, after, channel=None, skip=(), headless=False):
    """Open the sites and return the raw dict build() reads."""
    ed, af = date.fromisoformat(edition), date.fromisoformat(after)
    raw = {"fetched_at": datetime.now(PKT).strftime("%Y-%m-%d %H:%M PKT"), "errors": []}
    g = Browser(channel, headless)
    try:
        if "sx" not in skip:
            try:
                sx = {}
                # dividends: the range filters the announcement date; half a year back covers
                # payouts announced long before they fall due
                rng = {"dividends": (dmy((ed - timedelta(days=183)).isoformat()), dmy((ed + timedelta(days=1)).isoformat())),
                       "assemblies": (dmy((af + timedelta(days=1)).isoformat()), dmy((ed + timedelta(days=14)).isoformat())),
                       "actions": (None, None)}
                for tab, (s, e) in rng.items():
                    sx[tab] = {m: g.sx_tab(tab, m, s, e) for m in ("M", "S")}
                raw["sx"] = sx
                gas = [x for m in ("M", "S") for x in parse_assemblies(sx["assemblies"][m]["rows"])]
                urls = sorted({x["url"] for x in gas if x["url"] and after < x["date"] <= (ed + timedelta(days=14)).isoformat()})
                raw["announcements"] = g.announcements(urls)
            except Exception as e:  # noqa: BLE001
                raw["errors"].append(f"Saudi Exchange: {e}")
        if "inv" not in skip:
            try:
                raw["inv"] = g.investing((ed - timedelta(days=1)).isoformat(), (ed + timedelta(days=8)).isoformat())
                if raw["inv"].get("more"):
                    raw["errors"].append("investing.com returned more rows than one call holds; later days may be missing")
            except Exception as e:  # noqa: BLE001
                raw["errors"].append(f"investing.com: {e}")
    finally:
        g.close()
    return raw


def load_names(path):
    import openpyxl, mapcols
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = mapcols.sheet(wb); ix = mapcols.idx(ws)
    out = {}
    for r in ws.iter_rows(min_row=2, values_only=True):
        code = mapcols.get(r, ix, "code")
        if code is None:
            continue
        code = str(code).strip().split(".")[0]
        en, ar = mapcols.get(r, ix, "en"), mapcols.get(r, ix, "ar")
        if en and ar:
            out[code.zfill(4)] = {"en": _flat(en), "ar": _flat(ar)}
    return out


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--edition", required=True, help="YYYY-MM-DD, the edition the calendar is for")
    ap.add_argument("--after", help="YYYY-MM-DD, the edition before it (default: previous trading day)")
    ap.add_argument("--out", required=True)
    ap.add_argument("--rows", help="also write the rows.json build_cal.py reads")
    ap.add_argument("--raw", help="save the raw fetch here (default: next to --out)")
    ap.add_argument("--from-raw", help="rebuild from an earlier raw fetch, no browser")
    ap.add_argument("--names", help="Mapping.xlsx (state copy) for official EN and AR names")
    ap.add_argument("--prev-cal", help="previous Manual File-Calendar.xlsx: blank marker and economic AR wording")
    ap.add_argument("--blank", help="no-value marker (default: the previous file's, else '-')")
    ap.add_argument("--announced-before", help="replays: ignore dividends announced on or after this date")
    ap.add_argument("--browser", default=None, help="msedge or chrome (default: Edge on Windows, Chrome on Mac)")
    ap.add_argument("--skip", action="append", default=[], choices=["sx", "inv"])
    ap.add_argument("--headless", action="store_true", help="for tests only; Saudi Exchange blocks it")
    a = ap.parse_args(argv)
    ed = date.fromisoformat(a.edition)
    after = a.after or previous_edition(ed, _holidays()).isoformat()
    if a.from_raw:
        raw = json.load(open(a.from_raw, encoding="utf-8"))
    else:
        try:
            raw = fetch(a.edition, after, a.browser, a.skip, a.headless)
        except Exception as e:  # noqa: BLE001
            print(json.dumps({"error": f"browser could not start: {e}"}))
            return 1
        rp = a.raw or os.path.splitext(a.out)[0] + ".raw.json"
        json.dump(raw, open(rp, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    names = load_names(a.names) if a.names else {}
    pairs, blank = {}, None
    if a.prev_cal:
        try:
            pairs, blank = prev_cal_pairs(a.prev_cal)
        except (Exception, SystemExit) as e:  # noqa: BLE001  (a missing or re-laid-out file: go on without it)
            raw.setdefault("errors", []).append(f"previous calendar file not read ({a.prev_cal}): {e}")
    out = build(raw, a.edition, after, names=names, pairs=pairs, blank=a.blank or blank or "-",
                announced_before=a.announced_before)
    json.dump(out, open(a.out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    if a.rows:
        json.dump(out["rows"], open(a.rows, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    r = out["rows"]
    print(f"cal_fetch {a.edition}: {len(r['events'])} event rows ({', '.join(out['days']['events']) or 'none'}), "
          f"{len(r['economic'])} economic rows ({out['days']['economic'] or 'none'}), {len(out['review'])} review items -> {a.out}")
    blocked = [e for e in raw.get("errors", []) if e.startswith("Saudi Exchange") or e.startswith("investing.com:")]
    return 1 if blocked else 0


if __name__ == "__main__":
    sys.exit(main())
