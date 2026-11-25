#!/usr/bin/env python3
"""Set the Arabic company names in Manual File-Calendar.xlsx from the names sheet (open item 8).

Pre-approved by the user on 29 Oct 2026 for this one kind of fix only: each event-calendar row
whose English company name matches a row of the team's names sheet (Mapping.xlsx, sheet "mapping";
since 29 Oct 2026, replacing names_codes.xlsx / TASI CODE; column
"Long Name - English", found by header) exactly, apart from spacing, gets that row's "Long Name - Arabic" when its
Arabic company name differs. Dates, figures, event wording, rows and English names are never
touched. A row with no exact match is left as it is and listed. Runs locally, no internet.

Usage:
  python fix_cal_ar_names.py "Manual File-Calendar.xlsx" [--names Mapping.xlsx]
         [--out FILE] [--dry-run] [--report report.md]
Without --out the calendar file is updated in place (work on a staged copy, then commit it with
the deck open-file routine). Prints one line per change and a one-line note for the calendar owner.
Exit code 0 always; 2 if the calendar file cannot be read.
"""
import argparse
import re
import sys
from datetime import datetime
from pathlib import Path

import openpyxl

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from dmw_common import cal_sheet, cal_columns  # noqa: E402  (same header finding as the build)
import mapcols  # noqa: E402


def clean(s):
    return re.sub(r"\s+", " ", str(s or "")).strip()


def code_sheet(path):
    """English long name -> set of Arabic long names (a name on two rows may disagree)."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb["mapping"] if "mapping" in wb.sheetnames else wb["TASI CODE"]
    ix = mapcols.idx(ws)   # columns by header name (1 Nov 2026)
    out, lower = {}, {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        en, ar = clean(mapcols.get(row, ix, "en")), clean(mapcols.get(row, ix, "ar"))
        if en and ar:
            out.setdefault(en, set()).add(ar)
            lower.setdefault(en.lower(), set()).add(en)
    return out, lower


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("calendar")
    ap.add_argument("--names", default=str(HERE.parent / "data" / "Mapping.xlsx"))
    ap.add_argument("--out")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--report")
    a = ap.parse_args()

    try:
        ws, hr = cal_sheet(a.calendar)
    except SystemExit as e:
        print(f"Calendar file not read: {e}"); sys.exit(2)
    k = cal_columns(ws, hr)
    names, lower = code_sheet(a.names)

    changed, nomatch, ambiguous, same = [], [], [], 0
    r = hr + 1
    while isinstance(ws.cell(r, k["ev_date"]).value, datetime):
        en = clean(ws.cell(r, k["ev_company"]).value)
        cell = ws.cell(r, k["ev_company_ar"])
        cur = clean(cell.value)
        if en and en not in ("-", "_"):
            want = names.get(en)
            if not want:
                near = lower.get(en.lower())
                nomatch.append(f"row {r}: '{en}'" + (f" (case differs from '{sorted(near)[0]}')" if near else ""))
            elif len(want) > 1:
                ambiguous.append(f"row {r}: '{en}' has {len(want)} Arabic names in the code sheet: " + " / ".join(sorted(want)))
            elif cur != next(iter(want)):
                new = next(iter(want))
                changed.append((r, en, cur, new))
                if not a.dry_run:
                    cell.value = new
            else:
                same += 1
        r += 1

    lines = [f"# Calendar Arabic company names - {Path(a.calendar).name}", ""]
    lines.append(f"{len(changed)} changed, {same} already matched the code sheet, {len(nomatch)} without an exact "
                 f"English match, {len(ambiguous)} ambiguous." + (" DRY RUN: file not written." if a.dry_run else ""))
    if changed:
        lines += ["", "## Changed (Arabic company name set from the names sheet, Mapping.xlsx)", ""]
        lines += [f"- row {r}: {en}: {old or '(blank)'} -> {new}" for r, en, old, new in changed]
    if nomatch:
        lines += ["", "## No exact English match (left as is)", ""] + [f"- {x}" for x in nomatch]
    if ambiguous:
        lines += ["", "## Ambiguous in the code sheet (left as is)", ""] + [f"- {x}" for x in ambiguous]
    if changed:
        lines += ["", "## Note for the calendar owner", "",
                  "Arabic company names now follow the names sheet (Mapping.xlsx): " +
                  "; ".join(f"{old or '(blank)'} -> {new}" for _, _, old, new in changed) + "."]
    text = "\n".join(lines) + "\n"
    print(text)
    if a.report:
        Path(a.report).write_text(text, encoding="utf-8")
    if changed and not a.dry_run:
        out = a.out or a.calendar
        ws.parent.save(out)
        print(f"Saved {out}")


if __name__ == "__main__":
    main()
