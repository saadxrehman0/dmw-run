import os, sys
import openpyxl
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "..", "scripts"))
import place_results as pr  # noqa: E402


def test_bloomberg_from_substitute_formula(tmp_path):
    """A workbook saved by openpyxl has no cached values: =SUBSTITUTE(H, " Equity", "") in TASI
    CODE column I reads as None, and the row was dropped (5 Nov replay)."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Results"
    for _ in range(5):
        ws.append([])
    ws.append(["3Q26", "FLYNAS", "=XLOOKUP(B6,...)", "=XLOOKUP(B6,...)", 27.0, 120.2, None, None, None, None])
    tc = wb.create_sheet("TASI CODE")
    tc.append([None, None, "FLYNAS", 4264, None, None, None, "FLYNAS AB Equity", '=SUBSTITUTE(H1, " Equity", "")'])
    p = tmp_path / "r.xlsx"
    wb.save(p)
    rows = pr.read_rows(str(p), "3Q26")
    assert [(r["name"], r["code"], r["bbg"]) for r in rows] == [("FLYNAS", "4264", "FLYNAS AB")]
