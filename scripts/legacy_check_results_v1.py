"""Check the Results Announcements tables in the EN and AR DMW decks against DMW - Results.xlsx.

Usage: python check_results.py EN_PPTX AR_PPTX RESULTS_XLSX [NAMES_XLSX]

Added 29 Oct 2026 after two misses in a row on the team's results tables:
- 28 Oct: BAHRI and AMERICANA consensus and variance wrong in both decks.
- 29 Oct: SAUDI CABLE row missing from the AR deck; ARMAH's TASI code 9590 (old Nomu code)
  in both decks while the Results file and the exchange have 6022.

Every row in either deck is matched by Bloomberg ticker. Prints one FLAG line per problem:
rows on one side only, TASI code differences, and any figure that differs from the Results file
after 1dp rounding (halves up). Read only.
"""
import sys
from decimal import Decimal, ROUND_HALF_UP
import openpyxl
from pptx import Presentation

def hu(v):
    return Decimal(repr(v)).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)

def num(x):
    if x is None: return None
    if isinstance(x, (int, float)): return float(x)
    x = str(x).strip().replace(',', '').replace('%', '')
    if x in ('', '-', 'NM', 'nm', 'N/A'): return None
    try: return float(x)
    except ValueError: return None

LABELS = ['current', 'prior year', 'change SAR mn', 'change %', 'consensus', 'variance %']

def deck_rows(path, ar=False):
    rows = {}
    for si, s in enumerate(Presentation(path).slides, 1):
        if si > 3: break
        for sh in s.shapes:
            if not sh.has_table: continue
            for r in sh.table.rows:
                c = [x.text.strip() for x in r.cells]
                if len(c) < 9: continue
                if ar: c = c[::-1]          # AR table runs right to left
                if not c[1].isdigit() or not c[2].endswith(' AB'): continue
                rows[c[2]] = {'slide': si, 'name': c[0], 'code': c[1], 'figs': [num(v) for v in c[3:9]]}
    return rows

def xlsx_rows(path):
    ws = openpyxl.load_workbook(path, data_only=True)['Results']
    rows = {}
    for r in range(6, 45):
        v = [ws.cell(r, c).value for c in range(1, 11)]
        if v[3] and str(v[3]).strip().endswith(' AB') and v[2] not in (None, ''):
            rows[str(v[3]).strip()] = {'name': v[1], 'code': str(v[2]).strip(), 'figs': [num(x) for x in v[4:10]]}
    return rows

def main(en, ar, res, names=None):
    E, A, X = deck_rows(en), deck_rows(ar, ar=True), xlsx_rows(res)
    flags = []
    for b in sorted(set(E) | set(A) | set(X)):
        where = {k for k, d in (('EN', E), ('AR', A), ('Results file', X)) if b in d}
        if len(where) < 3:
            flags.append(f'{b}: only in {", ".join(sorted(where))}')
        codes = {k: d[b]['code'] for k, d in (('EN', E), ('AR', A), ('Results file', X)) if b in d}
        if len(set(codes.values())) > 1:
            flags.append(f'{b}: TASI code differs, ' + ', '.join(f'{k} {v}' for k, v in codes.items()))
        if b not in X: continue
        ref = X[b]['figs']
        for k, d in (('EN', E), ('AR', A)):
            if b not in d: continue
            for lab, got, want in zip(LABELS, d[b]['figs'], ref):
                if want is None and got is None: continue
                if (want is None) != (got is None) or hu(got) != hu(want):
                    flags.append(f'{b} {k} slide {d[b]["slide"]} {lab}: deck {got} vs Results file '
                                 f'{"-" if want is None else hu(want)}')
    if names:   # TASI code in the names sheet, so a stale sheet is caught too
        ws = openpyxl.load_workbook(names, read_only=True, data_only=True).worksheets[0]
        by_bbg = {}
        for r in ws.iter_rows(min_row=2, values_only=True):
            if r[3] is not None and len(r) > 8 and r[8]: by_bbg[str(r[8]).strip()] = str(r[3])
        for b, d in X.items():
            if b in by_bbg and by_bbg[b] != d['code']:
                flags.append(f'{b}: names sheet has TASI code {by_bbg[b]}, Results file {d["code"]}')
    print('\n'.join('FLAG: ' + f for f in flags) or 'OK: EN, AR and the Results file agree on every results row')
    return flags

if __name__ == '__main__':
    main(*sys.argv[1:5])
