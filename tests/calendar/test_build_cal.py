"""build_cal keeps the team's wording list below the calendar and dates new rows (26 Nov 2026)."""
import json, os, subprocess, sys
import openpyxl
import pytest

SCRIPT = os.path.join(os.path.dirname(__file__), "..", "..", "scripts", "build_cal.py")


def _template(path, dated_rows=2):
    wb = openpyxl.Workbook(); ws = wb.active
    ws["B3"] = "Date"; ws["D3"] = "Event"; ws["N3"] = "Date"; ws["O3"] = "Event"
    for r in range(4, 4 + dated_rows):
        ws[f"B{r}"] = 46000 + r; ws[f"B{r}"].number_format = "dd-mmm-yy"; ws[f"D{r}"] = f"old {r}"
        ws[f"N{r}"] = 46000 + r; ws[f"N{r}"].number_format = "dd-mmm-yy"; ws[f"O{r}"] = f"old eco {r}"
    ws["D14"] = "Earning Call"; ws["D15"] = "Ordinary Assembly Meeting"
    wb.save(path)


def _rows(path, n_ev):
    ev = [{"date": "2026-11-19", "company_en": f"Co {i}", "event_en": f"Event {i}",
           "company_ar": "ش", "event_ar": "ح"} for i in range(n_ev)]
    eco = [{"date": "2026-11-19", "event_en": "US CPI", "event_ar": "م", "prior": "", "estimate": ""}]
    json.dump({"blank": "-", "events": ev, "economic": eco}, open(path, "w"))


def _run(tmp_path, n_ev, dated_rows=2):
    src, out, rows = tmp_path / "in.xlsx", tmp_path / "out.xlsx", tmp_path / "rows.json"
    _template(src, dated_rows); _rows(rows, n_ev)
    return subprocess.run([sys.executable, SCRIPT, str(src), str(out), str(rows)],
                          capture_output=True, text=True), out


def test_wording_list_kept_and_dates_formatted(tmp_path):
    r, out = _run(tmp_path, 3)
    assert r.returncode == 0, r.stderr + r.stdout
    ws = openpyxl.load_workbook(out).active
    assert [ws[f"D{i}"].value for i in (4, 5, 6)] == ["Event 0", "Event 1", "Event 2"]
    assert ws["D14"].value == "Earning Call" and ws["D15"].value == "Ordinary Assembly Meeting"
    assert ws["D7"].value is None and ws["O5"].value is None
    assert "yy" in ws["B6"].number_format        # a new row below the old last one still shows a date


def test_refuses_to_overwrite_wording_list(tmp_path):
    r, _ = _run(tmp_path, 11)                    # rows 4-14 reach D14
    assert r.returncode != 0 and "D14" in (r.stderr + r.stdout)
