#!/usr/bin/env python3
"""Monthly Saudi Exchange check of the team's names sheet, Mapping.xlsx (Task 0; added 29 Oct 2026).

Two steps around one browser call, so the exchange's 2,000-row directory never passes through the chat:
  1. python refresh_mapping.py digest Mapping.xlsx
     prints a short digest (code:board:hash of the names and ISIN) to paste into SX.diff() in
     shared/browser/sx_names.js, run in the built-in browser on saudiexchange.sa.
  2. Save the JSON SX.diff() returns as diff.json (and SX.sectors() for the new or moved codes as
     sectors.json), then
     python refresh_mapping.py apply Mapping.xlsx diff.json [--sectors sectors.json] --out NEW.xlsx --report R.md
     writes a proposed copy and a change report. The user approves before the shared file is replaced.
Optional cross-check against a Bloomberg EQS export (CSV: Ticker, Name, ISIN, Tkr & Exch, as in
Claude Outputs/Workflow Automation/2026-10-08/TASI master names):
     python refresh_mapping.py bbg Mapping.xlsx export.csv
What changes on apply: new listings get a row (names, code, RIC, market, sector, ISIN; Bloomberg
ticker left for the bbg check or the team); board moves update Market; renames update the long and
short names (EN and AR) and the ISIN; companies no longer listed are reported, never deleted. Every
checked row gets today's date in "Saudi Exchange Last Check". Sector labels follow the "appendix"
sector lookup and Arabic spellings. Columns are found by header name (row 1), so moved columns are fine;
new listings are added at the bottom (the team re-sorts the sheet if they want it alphabetical).
"""
import argparse, csv, datetime, json, re, sys
from copy import copy
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent))
import mapcols  # noqa: E402

# Columns are found by their row-1 header through mapcols.py, which knows the old and the new
# header names (1 Nov 2026: headers renamed, columns reordered, the sector lookup moved from
# columns S:T of "mapping" to the "appendix" sheet). COLS is filled by sheet().
COLS = {}
ALL = ("no", "en", "short", "code", "ar_short", "ar", "ric", "bbg", "bbg2", "pref", "market",
       "sec_en", "sec_ar", "sec_pref", "isin", "checked")


def sheet(wb):
    ws = mapcols.sheet(wb)
    COLS.clear()
    COLS.update(mapcols.cols(ws, ALL))
    return ws


def norm(s):
    s = str(s or "").replace("ی", "ي").replace("ک", "ك")
    s = re.sub(r"\s*\*\s*$", "", s)
    return re.sub(r"\s+", " ", s).strip().lower()


def fnv(s):
    h = 0x811C9DC5
    for b in s.encode("utf-8"):
        h ^= b
        h = (h * 0x01000193) & 0xFFFFFFFF
    return f"{h:08x}"[:6]


def rows(ws):
    for r in range(2, ws.max_row + 1):
        if ws.cell(r, COLS["code"]).value is not None:
            yield r


def digest(path):
    ws = sheet(openpyxl.load_workbook(path, read_only=False, data_only=True))
    out = []
    for r in rows(ws):
        v = lambda k: ws.cell(r, COLS[k]).value
        board = "N" if str(v("market") or "").strip().lower().startswith("nomu") else "M"
        h = fnv("|".join(norm(x) for x in (v("en"), v("short"), v("ar_short"), v("isin"))))
        out.append(f"{int(v('code'))}:{board}:{h}")
    return ",".join(out)


def sector_lookup(ws):
    """Exchange sector name -> (sheet EN label, AR label, preferred label), from the sheet itself."""
    st = {}
    wb = ws.parent
    src, c1, r0 = ws, 19, 2                     # old layout: columns S:T of "mapping"
    if "appendix" in wb.sheetnames:             # since 1 Nov 2026: "appendix", found by header
        apx = wb["appendix"]
        for r in range(1, 6):
            for c in range(1, apx.max_column + 1):
                if str(apx.cell(r, c).value or "").strip().lower() == "saudi exchange sector":
                    src, c1, r0 = apx, c, r + 1
    for r in range(r0, r0 + 60):
        s, t = src.cell(r, c1).value, src.cell(r, c1 + 1).value
        if s and t:
            st[re.sub(r"\W", "", str(s)).lower()] = (str(s).strip(), str(t).strip())
    ar = {}
    for r in rows(ws):
        en, a = ws.cell(r, COLS["sec_en"]).value, ws.cell(r, COLS["sec_ar"]).value
        if en and a:
            ar.setdefault(re.sub(r"\W", "", str(en)).lower(), str(a).strip())
    def look(name):
        k = re.sub(r"\W", "", name).lower()
        en, pref = st.get(k, (name, name))
        return en, ar.get(re.sub(r"\W", "", en).lower()), pref
    return look


def apply(a):
    wb = openpyxl.load_workbook(a.mapping)
    ws = sheet(wb)
    d = json.loads(Path(a.diff).read_text(encoding="utf-8"))
    secs = json.loads(Path(a.sectors).read_text(encoding="utf-8")) if a.sectors else {}
    look = sector_lookup(ws)
    today = datetime.date.today().strftime("%d-%b-%Y")
    by_code = {str(int(ws.cell(r, COLS["code"]).value)): r for r in rows(ws)}
    log, flags = [], []

    def setv(r, k, new, why):
        old = ws.cell(r, COLS[k]).value
        if (str(old).strip() if old is not None else None) != new:
            ws.cell(r, COLS[k]).value = new
            log.append(f"| {r} | {ws.cell(r, COLS['code']).value} | {k} | {old if old not in (None, '') else '(blank)'} | {new} | {why} |")

    def set_sector(r, code):
        if code in secs:
            en, ar, pref = look(secs[code])
            setv(r, "sec_en", en, "Saudi Exchange market watch")
            if ar:
                setv(r, "sec_ar", ar, "Saudi Exchange market watch")
            else:
                flags.append(f"- {code}: sector '{en}' has no Arabic label in the sheet yet; add it.")
            setv(r, "sec_pref", pref, "appendix sector lookup")
        else:
            flags.append(f"- {code}: sector not read (run SX.sectors on the market watch page).")

    for code, sheet_board, board in d.get("market", []):
        r = by_code[code]
        setv(r, "market", board, f"board move ({'Nomu to Main' if board == 'Main' else board})")
        set_sector(r, code)
    for code, isin, short, long_en, ar_short, ar_long in d.get("changed", []):
        r = by_code[code]
        setv(r, "en", long_en, "Saudi Exchange name")
        setv(r, "short", short, "Saudi Exchange trading name")
        if ar_short:
            setv(r, "ar_short", ar_short, "Saudi Exchange Arabic trading name")
        if ar_long:
            setv(r, "ar", ar_long, "Saudi Exchange Arabic name")
        if isin:
            setv(r, "isin", isin, "Saudi Exchange directory")
    last = max(rows(ws))
    nos = [ws.cell(r, 1).value for r in rows(ws) if isinstance(ws.cell(r, 1).value, (int, float))]
    nxt = int(max(nos)) + 1 if nos else 1
    for code, board, isin, short, long_en, ar_short, ar_long in d.get("new", []):
        last += 1
        for c in range(1, ws.max_column + 1):
            ws.cell(last, c)._style = copy(ws.cell(last - 1, c)._style)
        for k, v in dict(no=nxt, en=long_en, short=short, code=int(code), ar_short=ar_short, ar=ar_long,
                         ric=f"{code}.SE", market=board, isin=isin).items():
            ws.cell(last, COLS[k]).value = v
        nxt += 1
        by_code[code] = last
        set_sector(last, code)
        log.append(f"| {last} | {code} | (new row) | | {long_en} ({short}), {board} | new listing |")
        flags.append(f"- {code} {long_en}: new row; Bloomberg ticker and Preferred Name left for the team "
                     "(or run the bbg check).")
    for code in d.get("gone", []):
        r = by_code.get(code)
        flags.append(f"- {code} {ws.cell(r, COLS['en']).value if r else ''}: not listed on Saudi Exchange today "
                     "(suspended, delisted or an old code). Row kept; ask before removing.")
    for r in rows(ws):
        if str(int(ws.cell(r, COLS["code"]).value)) not in d.get("gone", []):
            ws.cell(r, COLS["checked"]).value = datetime.datetime.combine(datetime.date.today(), datetime.time())
            ws.cell(r, COLS["checked"]).number_format = "dd-mmm-yyyy"
    wb.save(a.out)
    text = [f"# Mapping.xlsx monthly check - {today}", "",
            f"Saudi Exchange: {d.get('checked', '?')} listed companies checked. New {len(d.get('new', []))}, "
            f"board moves {len(d.get('market', []))}, name/ISIN changes {len(d.get('changed', []))}, "
            f"not listed {len(d.get('gone', []))}.", "",
            "Proposed copy: " + str(a.out) + " (the shared file is unchanged until the user approves).", ""]
    if log:
        text += ["## Changes", "", "| Row | Code | Field | Was | Now | Why |", "|---|---|---|---|---|---|"] + log + [""]
    if flags:
        text += ["## For the user", ""] + flags + [""]
    if not log and not flags:
        text += ["No changes: the sheet matches Saudi Exchange."]
    Path(a.report).write_text("\n".join(text), encoding="utf-8")
    print("\n".join(text))


def bbg(a):
    ws = sheet(openpyxl.load_workbook(a.mapping, read_only=True, data_only=True))
    by_isin = {r["ISIN"]: re.sub(r"\s+", " ", r["Tkr & Exch"]).strip() for r in csv.DictReader(open(a.csv))}
    g = lambda row, k: row[COLS[k] - 1] if len(row) >= COLS[k] else None
    for row in ws.iter_rows(min_row=2, values_only=True):
        if g(row, "code") is None:
            continue
        isin = g(row, "isin")
        t = by_isin.get(isin)
        cur = re.sub(r"\s+", " ", str(g(row, "bbg2") or "")).strip()
        if t and cur != t:
            print(f"{int(g(row, 'code'))} {str(g(row, 'en')).strip()}: sheet '{cur or '(blank)'}', Bloomberg '{t}'")
        elif isin and not t:
            print(f"{int(g(row, 'code'))} {str(g(row, 'en')).strip()}: ISIN {isin} not in the Bloomberg export")


def main():
    ap = argparse.ArgumentParser()
    sp = ap.add_subparsers(dest="cmd", required=True)
    p = sp.add_parser("digest"); p.add_argument("mapping")
    p = sp.add_parser("apply"); p.add_argument("mapping"); p.add_argument("diff")
    p.add_argument("--sectors"); p.add_argument("--out", required=True); p.add_argument("--report", required=True)
    p = sp.add_parser("bbg"); p.add_argument("mapping"); p.add_argument("csv")
    a = ap.parse_args()
    if a.cmd == "digest":
        print(digest(a.mapping))
    elif a.cmd == "apply":
        apply(a)
    else:
        bbg(a)


if __name__ == "__main__":
    main()
