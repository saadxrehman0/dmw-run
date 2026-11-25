#!/usr/bin/env python3
"""Check the two reference workbooks for code and RIC problems (open items 9 and 14, 29 Oct 2026).

Usage: python check_mapping.py [TASI_Mapping.xlsx] [Mapping.xlsx]
       (Mapping.xlsx, the team's names sheet, replaced names_codes.xlsx on 29 Oct 2026)
       (defaults: the copies in ../data/ next to this script)

Flags, per file:
  RIC     the Reuters RIC's number is not the company code (a Nomu-to-Main move that kept
          its old RIC: Armah 6022 with 9590.SE, 29 Oct 2026)
  FORMAT  the RIC has the right number but is not written CODE.SE (e.g. "SE.4018")
  DUP     the same code on two rows (often a renamed company whose old row was kept)
Across the two files:
  CODE    the same English name carries a different code in each file (Armah 9590 vs 6022)
Blank RICs are not flagged: Nomu rows carry none by design.
Exit code 1 when any RIC or CODE problem is found, so make_skill.py stops on it.
"""
import re
import sys
from collections import defaultdict
from pathlib import Path

import openpyxl

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import mapcols  # noqa: E402


def norm(s):
    return re.sub(r"\s+", " ", str(s or "")).strip().lower()


def rows_mapping(path):
    ws = openpyxl.load_workbook(path, read_only=True, data_only=True)["Companies"]
    out = []
    for r, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if row[0] is None:
            continue
        out.append(dict(row=r, code=str(row[0]).strip(), name=row[1], ric=row[5],
                        status=(row[6] if len(row) > 6 else None)))
    return out


def rows_names(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb["mapping"] if "mapping" in wb.sheetnames else wb["TASI CODE"]
    ix = mapcols.idx(ws, ("code", "en", "ric"))   # columns by header name (1 Nov 2026)
    out = []
    for r, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        code, name = mapcols.get(row, ix, "code"), mapcols.get(row, ix, "en")
        if code is None or name is None:
            continue
        out.append(dict(row=r, code=str(code).strip(), name=name, ric=mapcols.get(row, ix, "ric"), status=None))
    return out


def check_rows(label, rows):
    found = []
    by_code = defaultdict(list)
    for x in rows:
        by_code[x["code"]].append(x)
        if x.get("status") == "Archive only" or x["ric"] in (None, ""):
            continue
        ric = str(x["ric"]).strip()
        digits = re.sub(r"\D", "", ric)
        if digits != x["code"]:
            found.append(("RIC", f"{label} row {x['row']}: {x['name']} code {x['code']} but RIC {ric} "
                                 f"(should be {x['code']}.SE)"))
        elif not re.fullmatch(r"\d{4}\.SE", ric):
            found.append(("FORMAT", f"{label} row {x['row']}: {x['name']} RIC written '{ric}' "
                                    f"(should be {x['code']}.SE)"))
    for code, xs in by_code.items():
        if len(xs) > 1:
            names = "; ".join(f"row {x['row']} {str(x['name']).strip()}" for x in xs)
            found.append(("DUP", f"{label} code {code} on {len(xs)} rows: {names}"))
    return found


def cross(mapping, names):
    found = []
    m = {norm(x["name"]): x for x in mapping if x.get("status") != "Archive only"}
    for x in names:
        y = m.get(norm(x["name"]))
        if y and y["code"] != x["code"]:
            found.append(("CODE", f"{x['name'].strip()}: Mapping row {x['row']} has {x['code']}, "
                                  f"TASI Mapping row {y['row']} has {y['code']}"))
    return found


def main():
    tasi = Path(sys.argv[1]) if len(sys.argv) > 1 else HERE.parent / "data" / "TASI Mapping.xlsx"
    names = Path(sys.argv[2]) if len(sys.argv) > 2 else HERE.parent / "data" / "Mapping.xlsx"
    a, b = rows_mapping(tasi), rows_names(names)
    found = check_rows("TASI Mapping", a) + check_rows("Mapping", b) + cross(a, b)
    for level in ("RIC", "CODE", "FORMAT", "DUP"):
        items = [m for l, m in found if l == level]
        if items:
            print(f"## {level} ({len(items)})")
            print("\n".join("- " + m for m in items))
    if not found:
        print("OK: codes and RICs agree in both workbooks.")
    sys.exit(1 if any(l in ("RIC", "CODE") for l, _ in found) else 0)


if __name__ == "__main__":
    main()
