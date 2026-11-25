"""Tests for the edition orchestrator and the gate mailbox.  Run: python3 -m pytest tests/orchestrator -q

Stages are stub_stage.py (S1 opens the real pick gate from the 24 Nov fixture); the clock is fake
and fast, so a whole 8:54 to 10:25 run takes about 10 seconds.
"""
import json, os, sys, threading, time, zipfile
from datetime import date, datetime

import pytest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "../../scripts"))
import gates, run_edition, find_inputs  # noqa: E402
from gates import PKT  # noqa: E402

SPEC = json.load(open(os.path.join(HERE, "../../orchestrator/stages.json"), encoding="utf-8"))
ED = date(2026, 11, 26)  # a Thursday


def at(hhmm):
    h, m = map(int, hhmm.split(":"))
    return datetime(2026, 11, 26, h, m, tzinfo=PKT)


def cfg(tmp_path):
    return dict(run_edition.DEFAULTS, runs_root=str(tmp_path), day_root=str(tmp_path / "day"),
                holidays=str(tmp_path / "none.json"), mode_file=str(tmp_path / "none.json"),
                alert_email=None)


@pytest.fixture(autouse=True)
def fast_stubs(monkeypatch):
    monkeypatch.setenv("DMW_STUB_SLEEP", "0.05")
    monkeypatch.delenv("DMW_STUB_FAIL", raising=False)


def run(tmp_path, start="08:54", speed=600, answer_with=None):
    r = run_edition.Runner(ED, cfg(tmp_path), SPEC, run_edition.Clock(ED, start, speed), stub=True,
                           out=lambda *_: None)
    if answer_with:
        def saad():
            for _ in range(400):
                if os.path.exists(os.path.join(r.run, "gates", "pick.json")):
                    gates.answer(r.run, "pick", answer_with)
                    return
                time.sleep(0.02)
        threading.Thread(target=saad, daemon=True).start()
    assert r.loop(max_minutes=1) == 0
    return r


def outbox(r):
    return sorted(f for f in os.listdir(os.path.join(r.run, "outbox")) if f.count(".") == 1)


# ---------------------------------------------------------------- gate mailbox

def test_pick_gate_clock(tmp_path):
    rn = str(tmp_path)
    g = gates.open_gate(rn, "pick", "pick", {"blocks": []}, "Pick today's news", now=at("09:08"))
    assert g["remind"] == "09:18" and g["deadline"] == "09:25"
    gates.open_gate(rn, "pick", "pick", {"blocks": []}, "again", now=at("09:09"))  # retried stage
    assert gates.pending_outbox(rn) == ["001-pick-card.json"]
    assert gates.tick(rn, at("09:17")) == [] and len(gates.pending_outbox(rn)) == 1
    gates.tick(rn, at("09:18")); gates.tick(rn, at("09:20"))
    assert [f for f in gates.pending_outbox(rn) if "reminder" in f] == ["002-pick-reminder.json"]
    assert not gates.resolved(rn, "pick")
    assert gates.tick(rn, at("09:25")) == ["pick"]
    a = gates.get_answer(rn, "pick")
    assert (a["choice"], a["by"]) == ("go_proposed", "run")
    # Saad's press after the default does not overwrite it
    assert gates.answer(rn, "pick", "doc_edited")["by"] == "run"


def test_answer_before_deadline_wins(tmp_path):
    rn = str(tmp_path)
    gates.open_gate(rn, "pick", "pick", None, "Pick", now=at("09:08"))
    gates.answer(rn, "pick", "words", text="drop C7", now=at("09:12"))
    assert gates.tick(rn, at("09:30")) == []
    assert gates.get_answer(rn, "pick")["text"] == "drop C7"
    assert not any("reminder" in f for f in gates.pending_outbox(rn))


def test_conflict_has_no_default_and_late_news_holds(tmp_path):
    rn = str(tmp_path)
    gates.open_gate(rn, "conflict-tasi", "conflict", None, "TASI close differs", now=at("09:05"))
    gates.open_gate(rn, "late-sabic", "late_news", None, "SABIC late filing", now=at("10:02"))
    gates.tick(rn, at("11:00"))
    assert gates.get_answer(rn, "conflict-tasi") is None
    assert gates.get_answer(rn, "late-sabic")["choice"] == "hold"


def test_outbox_marks(tmp_path):
    rn = str(tmp_path)
    gates.ping(rn, "alert", "Data file late")
    gates.ping(rn, "alert", "Technicals late")
    gates.mark_sent(rn, "001-alert-alert.json", "cmsg_x")
    open(os.path.join(rn, "outbox", "002-alert-alert.emailed.json"), "w").write("{}")
    assert gates.pending_outbox(rn) == ["002-alert-alert.json"]


# ---------------------------------------------------------------- orchestrator

def test_full_run_with_the_runs_own_pick(tmp_path):
    r = run(tmp_path)
    st = r.st["stages"]
    assert {k: v["status"] for k, v in st.items()} == {k: "done" for k in st}
    assert st["S3"]["note"] == "stub S3: pick answered by run: go_proposed"
    gate = json.load(open(os.path.join(r.run, "gates", "pick.json"), encoding="utf-8"))
    assert at("09:00") <= datetime.fromisoformat(gate["opened"]) < at("09:08")  # stage on the run's clock
    assert datetime.fromisoformat(st["S3"]["started"]) >= at("09:25")
    assert st["S4"]["round"] >= 5          # rounds every 10 min from 9:00 until 10:10
    assert datetime.fromisoformat(st["S5"]["started"]) >= at("10:10")
    assert datetime.fromisoformat(st["S6"]["started"]) >= at("10:20")
    names = outbox(r)
    assert names[:3] == ["001-pick-card.json", "002-pick-reminder.json", "003-pick-closed.json"]
    card = json.load(open(os.path.join(r.run, "outbox", "001-pick-card.json"), encoding="utf-8"))
    assert card["card"]["blocks"][0]["text"] == "Pick today's news by 9:25 AM"
    saved = json.load(open(os.path.join(r.run, "run_state.json"), encoding="utf-8"))
    assert saved["heartbeat"] and saved["ended"]


def test_saads_pick_starts_commentary_early(tmp_path):
    r = run(tmp_path, answer_with="go_proposed")
    s3 = r.st["stages"]["S3"]
    assert "answered by Saad" in s3["note"]
    assert datetime.fromisoformat(s3["started"]) < at("09:18")
    assert not any("reminder" in f for f in outbox(r))


def test_failed_stage_retries_once_then_blocks_dependents(tmp_path, monkeypatch):
    monkeypatch.setenv("DMW_STUB_FAIL", "S2")
    r = run(tmp_path)
    st = r.st["stages"]
    assert (st["S2"]["status"], st["S2"]["attempts"]) == ("failed", 2)
    assert st["S3"]["status"] == "blocked" and st["S5"]["status"] == "blocked"
    assert st["S6"]["status"] == "done"     # needs S5 finished, not done
    alerts = [json.load(open(os.path.join(r.run, "outbox", f), encoding="utf-8"))["text"]
              for f in outbox(r) if "-alert" in f]
    assert any(t.startswith("S2 Decks") and "failed twice" in t for t in alerts)
    assert any("S5" in t and "cut-off" in t for t in alerts)
    # the run emails what the relay never posted (stub: written to RUN/mail)
    assert os.listdir(os.path.join(r.run, "mail"))


def test_restart_counts_the_interrupted_attempt(tmp_path):
    c = cfg(tmp_path)
    r = run_edition.Runner(ED, c, SPEC, run_edition.Clock(ED, "09:01"), stub=True, out=lambda *_: None)
    r.st["stages"]["S1"].update(status="running", attempts=1, pid=None)
    run_edition.save(r.run, r.st)
    r2 = run_edition.Runner(ED, c, SPEC, run_edition.Clock(ED, "09:02"), stub=True, out=lambda *_: None)
    s1 = r2.st["stages"]["S1"]
    assert (s1["status"], s1["attempts"]) == ("waiting", 1) and "restart" in s1["note"]


def test_edition_days(tmp_path):
    hol = tmp_path / "hol.json"
    hol.write_text(json.dumps({"2026": ["2026-12-03"]}))
    assert run_edition.edition_day(date(2026, 11, 27), str(hol)) == "Friday: no edition"
    assert run_edition.edition_day(date(2026, 11, 28), str(hol)) == "Saturday: no edition"
    assert run_edition.edition_day(date(2026, 12, 3), str(hol)) == "Tadawul holiday"
    assert run_edition.edition_day(date(2026, 11, 29), str(hol)) is None
    assert run_edition.day_folder("G:\\root", date(2026, 11, 26)).replace("\\", "/").endswith(
        "11. Nov/20261126")


def _xlsx(path, sheets):
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("xl/workbook.xml", "<workbook><sheets>" + "".join(
            f'<sheet name="{s}" sheetId="{i}"/>' for i, s in enumerate(sheets, 1)) + "</sheets></workbook>")


def test_find_inputs_by_contents(tmp_path):
    _xlsx(tmp_path / "DMW - Main.xlsx", ["Face", "PPT-ENG", "PPT-AR", "x"])
    _xlsx(tmp_path / "DMW - Results.xlsx", ["Face", "PPT-ENG", "PPT-AR"])
    _xlsx(tmp_path / "DMW - format.xlsx", ["Face", "PPT-ENG", "PPT-AR"])
    _xlsx(tmp_path / "DMW - Other.xlsx", ["Face"])
    (tmp_path / "~$DMW - Main.xlsx").write_text("lock")
    f = find_inputs.find(str(tmp_path))
    assert os.path.basename(f["data"]) == "DMW - Main.xlsx"
    assert find_inputs.missing(f, "data") == []
    assert find_inputs.missing(f, "technicals") == ["Technicals EN", "Technicals AR"]
    assert find_inputs.main(["--day", str(tmp_path), "--need", "all"]) == 1


# ---------------------------------------------------------------- replay (build step 5)

def _day(root, ymd, published=True):
    d = root / "11. Nov" / ymd
    d.mkdir(parents=True)
    _xlsx(d / "DMW - Data.xlsx", ["Face", "PPT-ENG", "PPT-AR"])
    for n in ("DMW Graphs - EN.pptx", "DMW Graphs - AR.pptx", f"Technicals -ENG-{ymd}.pptx",
              f"Technicals-AR- {ymd}.pptx", "Manual File-Calendar.xlsx", "DMW - Results.xlsx"):
        (d / n).write_text("x")
    if published:
        for n in ("ANB Daily Market Watch - EN.pptx", "ANB Daily Market Watch - AR.pptx",
                  f"EMAIL MB {ymd}.docx", "ANB Daily Market Watch - EN.pdf"):
            (d / n).write_text("published")
    return d


def test_replay_runs_unattended_on_a_jumping_clock(tmp_path):
    _day(tmp_path / "day", "20261125")
    _day(tmp_path / "day", "20261126")
    c = cfg(tmp_path)
    r = run_edition.Runner(ED, c, SPEC, run_edition.Clock(ED, "08:55", 1), stub=True,
                           out=lambda *_: None, replay=True)
    t0 = time.monotonic()
    assert r.loop(max_minutes=1) == 0
    assert time.monotonic() - t0 < 30            # real time is only the stages' own time
    st = r.st["stages"]
    assert {k: v["status"] for k, v in st.items()} == {k: "done" for k in st}
    assert r.run == os.path.join(str(tmp_path), "replay", "2026-11-26")
    assert r.st["mode"] == "practice"
    assert "answered by run" in st["S3"]["note"]
    assert datetime.fromisoformat(st["S3"]["started"]) >= at("09:25")
    assert st["S4"]["round"] == 0                # one round in a replay
    assert datetime.fromisoformat(st["S5"]["started"]) >= at("10:10")
    # the sandbox day folder has the inputs and the previous edition's decks, never this edition's finals
    day = sorted(os.listdir(r.day))
    assert "DMW - Data.xlsx" in day and "Technicals-AR- 20261126.pptx" in day and "prev" in day
    assert not any(n.startswith(("ANB Daily", "EMAIL MB", "DMW - Results")) for n in day)
    assert sorted(os.listdir(os.path.join(r.day, "prev"))) == ["ANB Daily Market Watch - AR.pptx",
                                                              "ANB Daily Market Watch - EN.pptx"]
    assert not os.path.exists(os.path.join(r.run, "mail"))


def test_replay_prompt_has_the_cutoffs():
    ctx = {"EDITION": "2026-11-24", "RUN": "R", "DAY": "D", "REPO": "X", "STAGE": "S1", "PY": "py",
           "MODE": "practice", "RESULT": "r.json", "ROUND": 0}
    _, prompt = run_edition.command(run_edition.DEFAULTS and SPEC["stages"][1], ctx,
                                    dict(run_edition.DEFAULTS), stub=False, replay=True)
    assert "Replay of a past edition" in prompt and "9:00 AM PKT on 2026-11-24" in prompt
    _, live = run_edition.command(SPEC["stages"][1], ctx, dict(run_edition.DEFAULTS), stub=False)
    assert "Replay" not in live
