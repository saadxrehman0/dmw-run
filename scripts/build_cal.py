"""Refresh the data rows of Manual File-Calendar.xlsx by direct XML edit, so the workbook's
images, charts and web extensions survive (saving through openpyxl drops them).

Handles both layouts seen in 2026:
  - old: sheet 'Calender', headers in row 3, data rows 4-26
  - new (from 10 Oct 2026): one sheet (e.g. 'Sheet1'), headers in row 2, data from row 3
The header row is found by looking for 'Date' in column B; each block's dated rows below it (at
most 23) are cleared and rewritten, and anything below them (the team's wording list) is kept. Columns: events EN B date, C company, D event; AR G event, H company, I date.
Economic EN N date, O event, P prior, Q estimate; AR S estimate, T prior, U event, V date.

Usage: python build_cal.py <previous Manual File-Calendar.xlsx> <output .xlsx> <rows.json>

rows.json:
{"blank": "-",                          # the previous file's no-value marker ("-" or "_"); copy it
 "events":   [{"date": "2026-10-13", "company_en": "...", "event_en": "...",
               "company_ar": "...", "event_ar": "..."}],
 "economic": [{"date": "2026-10-13", "event_en": "...", "prior": "3.98M", "estimate": null,   # written 4.0M
               "event_ar": "..."}]}
Every value must be sourced. A missing prior/estimate is written with the blank marker.
AR date cells are written the way the previous file wrote them: a date value if it held a number,
or Arabic-digit text ('١٣ أكتوبر ٢٠٢٦') if it held text.
"""
import re, sys, json, zipfile, datetime
from lxml import etree

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
N = {"m": NS}
EVENT_COLS = list("BCDGHI"); ECO_COLS = list("NOPQSTUV")
CAPACITY = 23


def one_dp(v):
    """Economic prior/estimate to ONE decimal, half-up, suffix kept: 3.98M -> 4.0M, 0.095M -> 0.1M,
    -0.25% -> -0.3%, 38K -> 38.0K (user, 13 Oct 2026: the 13 Oct file went out at 3.98M / 3.96M)."""
    from decimal import Decimal, ROUND_HALF_UP
    if v is None or not isinstance(v, (str, int, float)):
        return v
    m = re.fullmatch(r"\s*(-?\d+(?:\.\d+)?)\s*([KMBT%]?)\s*", str(v))
    if not m:
        return v
    q = Decimal(m.group(1)).quantize(Decimal("0.1"), ROUND_HALF_UP)
    return f"{q}{m.group(2)}"
AR_MONTHS = ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو", "أغسطس", "سبتمبر",
             "أكتوبر", "نوفمبر", "ديسمبر"]
AR_DIGITS = str.maketrans("0123456789", "٠١٢٣٤٥٦٧٨٩")


def col_index(col):
    n = 0
    for ch in col: n = n * 26 + ord(ch) - 64
    return n


def serial(iso):
    return (datetime.date.fromisoformat(iso) - datetime.date(1899, 12, 30)).days


def ar_date_text(iso):
    d = datetime.date.fromisoformat(iso)
    return f"{d.day} {AR_MONTHS[d.month - 1]} {d.year}".translate(AR_DIGITS)


def sheet_path(z):
    wb = etree.fromstring(z.read("xl/workbook.xml"))
    rels = etree.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    sheets = wb.findall(".//m:sheets/m:sheet", N)
    pick = next((s for s in sheets if s.get("name") == "Calender"), sheets[0])
    rid = pick.get("{%s}id" % RNS)
    target = next(r.get("Target") for r in rels if r.get("Id") == rid)
    return pick.get("name"), "xl/" + target.lstrip("/").replace("xl/", "", 1)


def shared_strings(z):
    try:
        sst = etree.fromstring(z.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    return ["".join(si.itertext()) for si in sst.findall("m:si", N)]


class Sheet:
    def __init__(self, xml, sst):
        self.root = etree.fromstring(xml)
        self.data = self.root.find("m:sheetData", N)
        self.sst = sst

    def row(self, r, create=True):
        for el in self.data.findall("m:row", N):
            n = int(el.get("r"))
            if n == r: return el
            if n > r:
                if not create: return None
                new = etree.Element("{%s}row" % NS, r=str(r)); el.addprevious(new); return new
        if not create: return None
        new = etree.SubElement(self.data, "{%s}row" % NS, r=str(r)); return new

    def cell(self, ref, create=True):
        col, r = re.match(r"([A-Z]+)(\d+)", ref).groups()
        row = self.row(int(r), create)
        if row is None: return None
        for c in row.findall("m:c", N):
            if c.get("r") == ref: return c
            if col_index(re.match(r"[A-Z]+", c.get("r")).group()) > col_index(col):
                if not create: return None
                new = etree.Element("{%s}c" % NS, r=ref); c.addprevious(new); return new
        if not create: return None
        return etree.SubElement(row, "{%s}c" % NS, r=ref)

    def value(self, ref):
        c = self.cell(ref, create=False)
        if c is None: return None
        v = c.find("m:v", N)
        if c.get("t") == "s" and v is not None: return self.sst[int(v.text)]
        if c.get("t") == "inlineStr": return "".join(c.find("m:is", N).itertext())
        return v.text if v is not None else None

    def is_text(self, ref):
        c = self.cell(ref, create=False)
        return c is not None and c.get("t") in ("s", "inlineStr", "str")

    def style(self, ref):
        c = self.cell(ref, create=False)
        return None if c is None else c.get("s")

    def set(self, ref, val, style=None):
        c = self.cell(ref)
        for ch in list(c): c.remove(ch)
        if "t" in c.attrib: del c.attrib["t"]
        # a number (a date serial) takes the first data row's style even where the cell has its
        # own, or a new row below the previous file's last one shows 46345 (General); 18 Nov replay
        if style and (not c.get("s") or isinstance(val, (int, float))): c.set("s", style)
        if val is None: return
        if isinstance(val, (int, float)):
            etree.SubElement(c, "{%s}v" % NS).text = str(val)
        else:
            c.set("t", "inlineStr")
            t = etree.SubElement(etree.SubElement(c, "{%s}is" % NS), "{%s}t" % NS)
            t.text = val; t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


def main(src, out, rows_path):
    rows = json.load(open(rows_path, encoding="utf-8"))
    blank = rows.get("blank", "-")
    ev, eco = rows.get("events", []), rows.get("economic", [])
    if len(ev) > CAPACITY or len(eco) > CAPACITY:
        raise SystemExit(f"more than {CAPACITY} rows in one block")
    with zipfile.ZipFile(src) as z:
        name, path = sheet_path(z)
        sh = Sheet(z.read(path), shared_strings(z))
    header = next((r for r in range(1, 10) if (sh.value(f"B{r}") or "").strip() == "Date"), None)
    if header is None: raise SystemExit("no 'Date' header in column B rows 1-9: layout changed, stop and ask")
    first = header + 1
    # AR dates follow the previous file; if its block was empty, use the Arabic-digit text form
    # (the house form since 10 Oct 2026), so one empty day does not switch the column to numbers.
    ar_ev_text = sh.is_text(f"I{first}") if sh.value(f"I{first}") not in (None, "") else True
    ar_eco_text = sh.is_text(f"V{first}") if sh.value(f"V{first}") not in (None, "") else True
    styles = {c: sh.style(f"{c}{first}") for c in EVENT_COLS + ECO_COLS}
    # Clear each block only down to its last dated row: the team keeps a wording list below the
    # events (D14:G21 in Nov 2026) that a blanket 23-row clear wiped (18 Nov replay, 26 Nov 2026).
    def extent(date_col):
        r = first
        while r < first + CAPACITY and sh.value(f"{date_col}{r}") not in (None, ""):
            r += 1
        return r
    for cols, date_col, text_col, n in ((EVENT_COLS, "B", "D", len(ev)), (ECO_COLS, "N", "O", len(eco))):
        end = extent(date_col)
        for r in range(end, first + n):
            if sh.value(f"{text_col}{r}") not in (None, ""):
                raise SystemExit(f"{n} rows would overwrite {text_col}{r} ('{sh.value(f'{text_col}{r}')}'), "
                                 "which is not a dated calendar row: stop and ask")
        for r in range(first, end):
            for c in cols:
                if sh.cell(f"{c}{r}", create=False) is not None: sh.set(f"{c}{r}", None)
    for i, e in enumerate(ev):
        r = first + i; d = serial(e["date"])
        vals = dict(B=d, C=e["company_en"], D=e["event_en"], G=e["event_ar"], H=e["company_ar"],
                    I=ar_date_text(e["date"]) if ar_ev_text else d)
        for c, v in vals.items(): sh.set(f"{c}{r}", v, styles[c])
    for i, e in enumerate(eco):
        r = first + i; d = serial(e["date"])
        pr, es = (blank if e.get(k) in (None, "") else one_dp(e.get(k)) for k in ("prior", "estimate"))
        vals = dict(N=d, O=e["event_en"], P=pr, Q=es, S=es, T=pr, U=e["event_ar"],
                    V=ar_date_text(e["date"]) if ar_eco_text else d)
        for c, v in vals.items(): sh.set(f"{c}{r}", v, styles[c])
    xml = etree.tostring(sh.root, xml_declaration=True, encoding="UTF-8", standalone=True)
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            zout.writestr(item, xml if item.filename == path else zin.read(item.filename))
    print(f"written {out}: sheet '{name}', data from row {first}, {len(ev)} event rows, {len(eco)} economic rows")


if __name__ == "__main__":
    main(*sys.argv[1:4])
