#!/usr/bin/env python3
"""Find the technical team's inputs in an edition's day folder (config.md rules).

  python find_inputs.py --day "<day folder>" [--need data|technicals|graphs|all] [--json out.json]

Exit 0 when everything --need asks for is there, 1 when something is missing (the orchestrator's
`ready` check for S2), 2 on a bad folder. Prints one JSON object with what was found.

- Data file: found by its contents, not its name (the team renames it; config.md, 1 Nov 2026): any
  `DMW - *.xlsx` other than `DMW - Results.xlsx` and anything with "format" in its name, whose
  sheets include Face, PPT-ENG and PPT-AR. Two matches: the newest wins and `note` says so.
- Graphs: `DMW Graphs - EN.pptx` and `DMW Graphs - AR.pptx`.
- Technicals: `Technicals -ENG-*.pptx` and `Technicals -AR-*.pptx` (with or without the space).
Lock files (`~$...`) and Drive temp files are skipped.
"""
import argparse, glob, json, os, sys, zipfile
import re

SHEETS = {"Face", "PPT-ENG", "PPT-AR"}


def sheet_names(path):
    """Sheet names straight from the workbook XML (no openpyxl needed; fast on a 10 MB file)."""
    try:
        with zipfile.ZipFile(path) as z:
            xml = z.read("xl/workbook.xml").decode("utf-8", "replace")
    except (zipfile.BadZipFile, KeyError, OSError):
        return set()
    return {re.sub(r"&amp;", "&", m) for m in re.findall(r'<sheet\b[^>]*\bname="([^"]+)"', xml)}


def _files(day, pattern):
    return [p for p in glob.glob(os.path.join(day, pattern))
            if not os.path.basename(p).startswith(("~$", ".")) and os.path.isfile(p)]


def find(day):
    out = {"day": day, "data": None, "graphs": {}, "technicals": {}, "notes": []}
    cands = [p for p in _files(day, "DMW - *.xlsx")
             if os.path.basename(p) != "DMW - Results.xlsx" and "format" not in os.path.basename(p).lower()
             and SHEETS <= sheet_names(p)]
    cands.sort(key=os.path.getmtime, reverse=True)
    if cands:
        out["data"] = cands[0]
        if len(cands) > 1:
            out["notes"].append(f"{len(cands)} data files match; took the newest, {os.path.basename(cands[0])}")
    for lang in ("EN", "AR"):
        g = _files(day, f"DMW Graphs - {lang}.pptx")
        if g:
            out["graphs"][lang.lower()] = g[0]
    for lang, tag in (("en", "ENG"), ("ar", "AR")):
        # the space before the dash is not always there ("Technicals-AR- 20261125.pptx", 25 Nov 2026)
        t = sorted(set(_files(day, f"Technicals -{tag}-*.pptx") + _files(day, f"Technicals-{tag}-*.pptx")),
                   key=os.path.getmtime, reverse=True)
        if t:
            out["technicals"][lang] = t[0]
    return out


def missing(found, need):
    m = []
    if need in ("data", "all") and not found["data"]:
        m.append("data file")
    if need in ("graphs", "all"):
        m += [f"Graphs {l.upper()}" for l in ("en", "ar") if l not in found["graphs"]]
    if need in ("technicals", "all"):
        m += [f"Technicals {l.upper()}" for l in ("en", "ar") if l not in found["technicals"]]
    return m


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--day", required=True)
    ap.add_argument("--need", default="data", choices=["data", "graphs", "technicals", "all"])
    ap.add_argument("--json")
    a = ap.parse_args(argv)
    if not os.path.isdir(a.day):
        print(json.dumps({"error": f"no day folder {a.day}"}))
        return 2
    f = find(a.day)
    f["missing"] = missing(f, a.need)
    s = json.dumps(f, ensure_ascii=False, indent=1)
    if a.json:
        open(a.json, "w", encoding="utf-8").write(s)
    print(s)
    return 1 if f["missing"] else 0


if __name__ == "__main__":
    sys.exit(main())
