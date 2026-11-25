"""Find Mapping.xlsx columns by their row-1 header (1 Nov 2026).

The team's names sheet (Mapping.xlsx, sheet "mapping") had its headers renamed and its columns
reordered on 1 Nov 2026. Every script that reads it finds columns here, by header name, so a
moved column or a renamed header (old and new names both listed) never breaks a script.
cols(ws) returns {key: 1-based column}; idx(ws) the same as 0-based indexes for values_only rows.
"""
import re

ALIASES = dict(
    no=("No",),
    code=("Code",),
    market=("Market Type", "Market"),
    en=("Long Name - English", "Long name English"),
    short=("Short Name - English", "Short Name English"),
    pref=("Preferred Name", "Prefered Name"),
    ar=("Long Name - Arabic", "Long name arabic"),
    ar_short=("Short Name - Arabic", "Short name arabic"),
    sec_en=("Sector Name - English", "Sector - EN"),
    sec_ar=("Sector Name - Arabic", "Sector - AR"),
    sec_pref=("Preferred Sector", "Prefered Sector"),
    ric=("Reuters",),
    bbg2=("Bloomberg Ticker", "Bloomberg2"),
    bbg=("Bloomberg Name", "Bloomberg"),
    isin=("ISIN",),
    checked=("Saudi Exchange Last Check", "Saudi Exchange check"),
)
REQUIRED = ("code", "en", "ar")


def _n(s):
    return re.sub(r"\s+", " ", str(s or "")).strip().lower()


def sheet(wb):
    return wb["mapping"] if "mapping" in wb.sheetnames else wb.worksheets[0]


def cols(ws, required=REQUIRED):
    hdr = {}
    for c, cell in enumerate(next(ws.iter_rows(min_row=1, max_row=1)), start=1):
        if _n(cell.value):
            hdr.setdefault(_n(cell.value), c)
    out = {}
    for k, names in ALIASES.items():
        c = next((hdr[_n(h)] for h in names if _n(h) in hdr), None)
        if c is not None:
            out[k] = c
    missing = [ALIASES[k][0] for k in required if k not in out]
    if missing:
        raise SystemExit(f"Mapping.xlsx: no {', '.join(repr(m) for m in missing)} header in row 1")
    return out


def idx(ws, required=REQUIRED):
    return {k: c - 1 for k, c in cols(ws, required).items()}


def get(row, ix, key):
    i = ix.get(key)
    return row[i] if i is not None and i < len(row) else None
