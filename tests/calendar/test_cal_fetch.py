"""Offline checks for cal_fetch.py (no browser), against what the Nov 2026 Mac replays read from
Saudi Exchange and investing.com (fixtures/), plus small made-up cases for the date rules."""
import json, os, subprocess, sys
from datetime import date, datetime, timezone

import openpyxl
import pytest

HERE = os.path.dirname(__file__)
FIX = os.path.join(HERE, "fixtures")
sys.path.insert(0, os.path.join(HERE, "..", "..", "scripts"))
sys.path.insert(0, HERE)
import cal_fetch as cf  # noqa: E402


def fx(name):
    with open(os.path.join(FIX, name), encoding="utf-8") as fh:
        return fh.read() if name.endswith(".html") else json.load(fh)


# ---- the two replay editions ----------------------------------------------------------------

def test_25_nov_matches_the_replay_rows():
    """Edition 25 Nov 2026: the S6 agent's hand-made rows.json (replay of 24 Nov). The one row the
    script cannot give is Derayah, which was on Argaam and not on the exchange's dividends tab."""
    out = cf.build(fx("raw_20261125.json"), "2026-11-25", "2026-11-24", names=fx("names_20261125.json"))
    want = fx("replay_rows_20261125.json")
    assert out["rows"]["events"] == [e for e in want["events"] if e["company_en"] != "Derayah Financial Co."]
    key = lambda e: e["event_en"]
    assert sorted(out["rows"]["economic"], key=key) == sorted(want["economic"], key=key)
    assert out["days"] == {"events": ["2026-11-25"], "events_borrowed": False, "economic": "2026-11-25"}
    assert any("0.3393, written 0.34" in s for s in out["sources"])            # Aramco, two decimals
    assert any("291,000,000 -> 300,000,000" in s for s in out["sources"])      # ACIG from its announcement


def test_5_nov_rows_and_the_failed_board_switch():
    out = cf.build(fx("raw_20261105.json"), "2026-11-05", "2026-11-04", announced_before="2026-11-04")
    ev = [(e["company_en"], e["event_en"]) for e in out["rows"]["events"]]
    assert ev == [("East Pipes Integrated Company for Industry", "Cash Dividend Distribution (SAR 2.00/sh)"),
                  ("Arabian United Float Glass Co.", "Eligibility for Cash Dividend (SAR 1.00/sh)"),
                  ("Miral Dental Clinics Co.", "Eligibility for Cash Dividend (SAR 2.00/sh)")]
    assert out["rows"]["economic"] == [{"date": "2026-11-05", "event_en": "US Initial Jobless Claims", "prior": "194K",
                                        "estimate": "197K", "event_ar": "الولايات المتحدة – طلبات إعانة البطالة الأولية"}]
    why = " ".join(r["why"] for r in out["review"])
    assert "assemblies tab: the Main Market switch did not take" in why
    assert "Arabic name TK" in why                                    # no names sheet given
    # blank-date rows: only those that could still fall on the edition are flagged
    flagged = {r["row"].split()[0] for r in out["review"] if r["what"] == "dividend dates"}
    assert flagged == {"4291", "9598"}
    assert any("2320" in d and "after the replay cut-off" in d for d in out["dropped"])


# ---- parsers ---------------------------------------------------------------------------------

def test_service_html_parse():
    rows = cf.parse_inv_service(fx("inv_service_20261105.html"))
    assert len(rows) == 12 and {r["cur"] for r in rows} == {"USD"}
    ijc = next(r for r in rows if r["name"] == "Initial Jobless Claims")
    assert (ijc["stars"], ijc["forecast"], ijc["previous"], ijc["revised_from"], ijc["event_id"]) == (3, "197K", "195K", "194K", 294)
    assert ijc["utc"] == "2026-11-05T13:30:00+00:00"                 # 18:30 PKT
    assert any(r["name"] == "Nonfarm Productivity (QoQ)" for r in rows)   # trailing (Q3) dropped


def test_next_data_parse_keeps_us_and_ksa_only():
    rows = cf.parse_inv_next(fx("inv_next_20261125.json"))
    assert rows and all(r["cur"] in ("USD", "SAR") for r in rows)
    nhs = next(r for r in rows if r["name"] == "New Home Sales")
    assert (nhs["stars"], nhs["previous"], nhs["revised_from"]) == (3, "614K", "676K")


def test_assembly_fields_found_by_form_and_layout_change_stops():
    g = cf.parse_assemblies([{"c": ["8150", "Allied", "Extraordinary - Corporate Actions", "2026-11-25", "19:00", "Riyadh",
                                    "ConveningAnn"], "a": ["/wps/x?anId=98897&anCat=53&cs=8150&locale=en"]}])[0]
    assert (g["type"], g["date"], g["time"], g["status"]) == ("Extraordinary - Corporate Actions", "2026-11-25", "19:00", "ConveningAnn")
    assert g["url"].startswith("https://www.saudiexchange.sa/wps/x?anId=98897")
    with pytest.raises(ValueError):
        cf.parse_assemblies([{"c": ["8150", "Allied", "Meeting", "25/11/2026"], "a": []}])
    with pytest.raises(ValueError):
        cf.parse_dividends([{"c": ["2222", "Aramco", "03/11/2026", "16/11/2026", "x", "25/11/2026", "^0.3393"]}])


def test_market_check():
    nomu = [{"c": ["9577", "x"]}, {"c": ["9553", "y"]}]
    assert cf.market_ok(nomu, "S") and not cf.market_ok(nomu, "M")
    assert cf.market_ok([{"c": ["2222", "x"]}, {"c": ["9553", "y"]}], "M")


# ---- assembly announcements -------------------------------------------------------------------

def test_capital_hike_from_the_acig_announcement():
    text = fx("raw_20261125.json")["announcements"]
    act = cf.corporate_action(next(iter(text.values()))["text"])
    assert act["kind"] == "hike" and act["after"] == 300_000_000 and act["before"] == 291_000_000


@pytest.mark.parametrize("text,en", [
    # made up, in the wording of the 5 Nov 2026 Saudi Azm invitation (agenda item 2)
    ("General Assembly Meeting Agenda: 1) Voting on amending the bylaws. 2) Voting on the purchase of up to (3,600,000) "
     "of the Company's shares and keeping them as treasury shares.",
     "Extraordinary General Meeting (Share Buyback of up to 3.60 mn shares)"),
    ("Voting on the Board recommendation to reduce the company's capital from SAR 100,000,000 to SAR 76,520,000 to "
     "offset accumulated losses.", "Extraordinary General Meeting (Capital Reduction to SAR 76.52 mn)"),
    ("Capital before the increase: SAR 10,000,000,000 Capital after the increase: SAR 15,000,000,000",
     "Extraordinary General Meeting (Capital Hike to SAR 15.00 bn)"),
    ("increase the capital from SAR 400 million to SAR 401.92 million through bonus shares",
     "Extraordinary General Meeting (Capital Hike to SAR 401.92 mn)"),
])
def test_assembly_wording(text, en):
    assert cf.egm_text(cf.corporate_action(text))[0] == en


def test_routine_assembly_dropped_and_unread_corporate_action_is_tk():
    raw = {"sx": {"assemblies": {"M": {"rows": [
        {"c": ["1111", "Routine Co.", "Ordinary", "2026-11-25", "19:00", "x", "Convening"], "a": ["/a?anId=1"]},
        {"c": ["2222", "Action Co.", "Extraordinary - Corporate Actions", "2026-11-25", "19:00", "x", "Convening"], "a": ["/a?anId=2"]},
        {"c": ["3333", "Bylaws Co.", "Extraordinary", "2026-11-25", "19:00", "x", "Convening"], "a": ["/a?anId=3"]}]}}},
        "announcements": {cf.SX + "/a?anId=1": {"text": "Agenda: 1) Voting on the financial statements."},
                          cf.SX + "/a?anId=3": {"text": "Agenda: 1) Voting on amending Article (12) of the bylaws."}}}
    out = cf.build(raw, "2026-11-25", "2026-11-24", wording={})
    assert [e["event_en"] for e in out["rows"]["events"]] == ["Extraordinary General Meeting (TK)"]
    assert len(out["dropped"]) == 2
    assert any(r["what"] == "assembly figure" and "announcement not read" in r["why"] for r in out["review"])


# ---- dates -----------------------------------------------------------------------------------

def _div(sym, elig, dist, amt="1.00", announced="2026-11-01"):
    return {"c": [sym, f"Co {sym}", announced, elig, "Account Transfer", dist, "^" + amt], "a": []}


def test_friday_and_saturday_events_roll_into_sunday():
    assert cf.previous_edition(date(2026, 11, 29)) == date(2026, 11, 26)          # Sunday -> Thursday
    raw = {"sx": {"dividends": {"M": {"rows": [_div("1001", "2026-11-27", "2026-12-10"),
                                               _div("1002", "2026-11-20", "2026-11-29"),
                                               _div("1003", "2026-11-26", "2026-12-10")]}}}}
    out = cf.build(raw, "2026-11-29", "2026-11-26", wording={})
    assert out["days"]["events"] == ["2026-11-27", "2026-11-29"]
    assert [(e["company_en"], e["date"]) for e in out["rows"]["events"]] == [("Co 1002", "2026-11-29"), ("Co 1001", "2026-11-27")]


def test_empty_day_borrows_the_next_day_with_items():
    raw = {"sx": {"dividends": {"M": {"rows": [_div("1001", "2026-12-02", "2026-12-20")]}}}}
    out = cf.build(raw, "2026-11-30", "2026-11-29", wording={})
    assert out["days"]["events"] == ["2026-12-02"] and out["days"]["events_borrowed"]
    assert out["rows"]["events"][0]["date"] == "2026-12-02"


def test_holiday_list_is_skipped():
    assert cf.previous_edition(date(2026, 12, 1), {"2026-11-30"}) == date(2026, 11, 29)       # Monday off
    assert cf.previous_edition(date(2026, 12, 1), {"2026-11-29", "2026-11-30"}) == date(2026, 11, 26)


def test_us_eastern_dst_edges():
    u = lambda *a: datetime(*a, tzinfo=timezone.utc)
    assert cf.us_eastern(u(2026, 3, 8, 6, 59)).utcoffset().total_seconds() == -5 * 3600
    assert cf.us_eastern(u(2026, 3, 8, 7, 0)).utcoffset().total_seconds() == -4 * 3600
    assert cf.us_eastern(u(2026, 11, 1, 5, 59)).utcoffset().total_seconds() == -4 * 3600
    assert cf.us_eastern(u(2026, 11, 1, 6, 0)).utcoffset().total_seconds() == -5 * 3600


# ---- economic block --------------------------------------------------------------------------

def _ev(utc, name, stars, fore="", prev="", rev=None, cur="USD"):
    return {"utc": utc, "cur": cur, "country": "", "stars": stars, "name": name, "raw_name": name, "actual": "",
            "forecast": fore, "previous": prev, "revised_from": rev, "event_id": None, "src": "test"}


def test_tuesday_api_row_and_us_dates():
    rows = [_ev("2026-11-24T15:00:00+00:00", "CB Consumer Confidence", 3, "93.0", "94.6"),
            _ev("2026-11-24T21:30:00+00:00", "API Weekly Crude Oil Stock", 2, "-1.0M", "4.4M"),   # 02:30 PKT Wednesday
            _ev("2026-11-24T19:00:00+00:00", "FOMC Meeting Minutes", 3),                         # 00:00 PKT Wednesday
            _ev("2026-11-25T13:30:00+00:00", "GDP (QoQ)", 3, "3.4%", "2.2%")]
    eco, src, review, day = cf.build_economic(rows, "2026-11-24", cf.load_wording(), {"US CB Consumer Confidence": "س"})
    assert day == "2026-11-24"
    assert [(e["event_en"], e["prior"], e["estimate"], e["event_ar"]) for e in eco] == [
        ("US CB Consumer Confidence", "94.6", "93.0", "س"),                    # AR from the previous file
        ("US FOMC Meeting Minutes", None, None, "TK"),
        ("US API Weekly Report on Crude Oil Inventories", None, None, "تقرير معهد البترول الأمريكي (API) الأسبوعي عن مخزونات النفط الخام")]
    assert [r["row"] for r in review if r["what"] == "wording"] == ["US FOMC Meeting Minutes"]


def test_borrowed_day_skips_one_star_rows_without_figures():
    rows = [_ev("2026-11-30T14:00:00+00:00", "IMF Meetings", 1),
            _ev("2026-12-01T15:00:00+00:00", "ISM Manufacturing PMI", 3, "48.6", "48.7")]
    eco, src, review, day = cf.build_economic(rows, "2026-11-29", {}, {})        # Sunday: no US row
    assert day == "2026-12-01" and eco[0]["event_en"] == "US ISM Manufacturing PMI"
    assert eco[-1]["event_en"] == "US API Weekly Report on Crude Oil Inventories"   # 1 Dec is a Tuesday


def test_ksa_rows_pool_with_us_rows_and_take_the_top_tier():
    rows = [_ev("2026-11-25T09:00:00+00:00", "Saudi CPI (YoY)", 2, "", "2.2%", cur="SAR"),
            _ev("2026-11-25T13:30:00+00:00", "Durable Goods Orders (MoM)", 2, "0.7%", "2.0%", "1.8%")]
    eco, *_ = cf.build_economic(rows, "2026-11-25", cf.load_wording(), {})
    assert [e["event_en"] for e in eco] == ["KSA Saudi CPI (YoY)", "US Durable Goods Orders (MoM)"]
    assert eco[1]["prior"] == "1.8%"                                              # as first published


# ---- end to end: CLI from a raw file into build_cal.py ---------------------------------------

def test_cli_from_raw_into_build_cal(tmp_path):
    from test_build_cal import _template
    raw, out, rows, names = tmp_path / "raw.json", tmp_path / "cal.json", tmp_path / "rows.json", tmp_path / "Mapping.xlsx"
    raw.write_text(json.dumps(fx("raw_20261125.json"), ensure_ascii=False), encoding="utf-8")
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "mapping"
    ws.append(["Code", "Long Name - English", "Long Name - Arabic"])
    for code, n in fx("names_20261125.json").items():
        ws.append([int(code), n["en"], n["ar"]])
    wb.save(names)
    script = os.path.join(HERE, "..", "..", "scripts", "cal_fetch.py")
    r = subprocess.run([sys.executable, script, "--edition", "2026-11-25", "--after", "2026-11-24", "--from-raw", str(raw),
                        "--out", str(out), "--rows", str(rows), "--names", str(names)], capture_output=True, text=True)
    assert r.returncode == 0, r.stderr + r.stdout
    assert "7 event rows" in r.stdout and "7 economic rows" in r.stdout
    got = json.loads(rows.read_text(encoding="utf-8"))
    assert got["events"][0]["company_ar"] == "شركة جرير للتسويق"
    src, xl = tmp_path / "in.xlsx", tmp_path / "out.xlsx"
    _template(src)
    b = subprocess.run([sys.executable, os.path.join(HERE, "..", "..", "scripts", "build_cal.py"), str(src), str(xl), str(rows)],
                       capture_output=True, text=True)
    assert b.returncode == 0, b.stderr + b.stdout
    ws = openpyxl.load_workbook(xl).active
    assert ws["D4"].value == "Cash Dividend Distribution (SAR 0.30/sh)" and ws["O4"].value == "US Core PCE Price Index (MoM)"
    assert ws["D14"].value == "Earning Call"                                       # the team's wording list kept


def test_previous_file_gives_wording_and_blank_marker(tmp_path):
    p = tmp_path / "prev.xlsx"
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Sheet1"
    hdr = {"B": "Date", "C": "Company", "D": "Event", "G": "الحدث", "H": "الشركة", "I": "التاريخ",
           "N": "Date", "O": "Event", "P": "Prior", "Q": "Estimation", "S": "التوقعات", "T": "الرقم السابق",
           "U": "الحدث", "V": "التاريخ"}
    for c, v in hdr.items():
        ws[f"{c}2"] = v
    ws["N3"] = 46350; ws["O3"] = "US CB Consumer Confidence"; ws["P3"] = "94.6"; ws["Q3"] = "_"
    ws["U3"] = "الولايات المتحدة – ثقة المستهلك"
    wb.save(p)
    pairs, blank = cf.prev_cal_pairs(str(p))
    assert pairs == {"US CB Consumer Confidence": "الولايات المتحدة – ثقة المستهلك"} and blank == "_"


# ---- the page scripts, on stand-in pages (Playwright's bundled Chromium, no network) -----------

def test_page_scripts_on_stand_in_pages():
    pw = pytest.importorskip("playwright.sync_api")
    table = ("<table id='issuerTable'><thead><tr><th>Symbol</th><th>Company Name</th></tr></thead><tbody>"
             "<tr><td>8150</td><td>Allied</td><td>Extraordinary - Corporate Actions</td><td>2026-11-25</td><td>19:00</td>"
             "<td>Riyadh</td><td>ConveningAnn</td><td><a href='/ann?anId=98897&amp;cs=8150&amp;locale=en'>x</a></td></tr>"
             "</tbody></table>")
    ann = "<table><tr><td>Meeting Agenda</td><td>Capital after the increase: SAR 300,000,000</td></tr></table>"
    inv = json.dumps({"data": fx("inv_service_20261105.html"), "rows_num": 12, "bind_scroll_handler": False})
    with pw.sync_playwright() as p:
        try:
            b = p.chromium.launch(executable_path=os.environ.get("DMW_TEST_CHROMIUM") or None)
        except Exception as e:  # noqa: BLE001
            pytest.skip(f"no Chromium: {e}")
        pg = b.new_page()
        pg.route("https://www.saudiexchange.sa/**", lambda r: r.fulfill(
            body=ann if "/ann" in r.request.url else f"<html><body>{table}</body></html>", content_type="text/html"))
        pg.route("https://www.investing.com/**", lambda r: r.fulfill(
            body=inv if r.request.method == "POST" else "<html><body></body></html>",
            content_type="application/json" if r.request.method == "POST" else "text/html"))
        pg.goto(cf.SX_CAL + "general-assembly-meetings?locale=en")
        got = pg.evaluate(cf.ROWS_JS)
        assert got["via"] == "dom" and got["rows"][0]["c"][2] == "Extraordinary - Corporate Actions"
        g = cf.parse_assemblies(got["rows"])[0]
        assert g["url"] == cf.SX + "/ann?anId=98897&cs=8150&locale=en"
        assert pg.evaluate(cf.SET_JS, ["M", None, None]) == {"err": "no jQuery"}
        texts = pg.evaluate(cf.ANN_JS, [g["url"]])
        assert cf.corporate_action(texts[g["url"]]["text"])["after"] == 300_000_000
        pg.goto(cf.INV)
        r = pg.evaluate(cf.INV_JS, ["2026-11-05", "2026-11-05", ["5", "52"]])
        assert r["status"] == 200 and len(cf.parse_inv_service(r["data"])) == 12
        b.close()
