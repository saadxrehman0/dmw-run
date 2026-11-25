"""Tests for the news Doc writer, reader and pick gate.  Run: python3 -m pytest tests/news_doc -q

Fixtures: published_<day>.json is the news of a published edition read back from its v1 decks;
candidates_<day>.json is that news shaped as the run's sweep output (make_fixtures.py). The core
check is the round trip: candidates -> Doc markdown -> reader -> NEWS.json gives fill_commentary.py
exactly the text that was published.
"""
import json, os, re, sys
import pytest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "../../scripts"))
sys.path.insert(0, HERE)
import news_doc  # noqa: E402
from news_doc import DocError, parse, render  # noqa: E402
import pick_card, read_news_doc, write_news_doc  # noqa: E402
from make_fixtures import candidates  # noqa: E402

DAYS = ("20261124", "20261118")


def pub(day):
    return json.load(open(os.path.join(HERE, "fixtures", f"published_{day}.json"), encoding="utf-8"))


def all_in(c):
    for it in c["economic"] + c["corporate"]:
        it["tier"] = "Include"
    return c


def edit(md, old, new, count=1):
    assert old in md, old
    return md.replace(old, new, count)


def row(md, iid):
    return next(l for l in md.splitlines() if re.match(rf"^\| [^|]* \| {iid} · ", l))


# ------------------------------------------------------------------ round trip

@pytest.mark.parametrize("day", DAYS)
def test_round_trip_matches_published(day):
    news, info = parse(render(all_in(candidates(day)), chips=False))
    p = pub(day)
    assert news["markets"] == p["markets"]
    assert [x["en"] for x in news["corporate"]] == [x["en"] for x in p["corporate"]]
    assert [x["ar"] for x in news["corporate"]] == [x["ar"] for x in p["corporate"]]
    assert [x["en"] for x in news["economic"]] == [x["en"] for x in p["economic"]]
    assert [x["ar"] for x in news["economic"]] == [x["ar"] for x in p["economic"]]
    assert not info["pending"]


def test_mockup_default_pick_leaves_maybes_out():
    news, info = parse(render(candidates("20261124"), chips=False))
    assert info["order"]["corporate"] == ["C1", "C2", "C3", "C4", "C5", "C6", "C7", "C8", "C11"]
    assert {"C9", "C10"} <= set(info["out"])
    assert len(news["corporate"]) == 9
    e1 = news["economic"][0]
    assert e1["short_en"].startswith("Saudi Arabia’s construction cost index (CCI) rose 2.3%")
    assert "104.1" in e1["short_ar"]


def test_export_byline_and_link_placeholder_are_ignored():
    """The real export renders chips as text ("Nov 24, 2026 · @Saad Rehman") and the mock-up's source
    lines read "Source: Argaam (the run puts the article link here)"."""
    md = render(candidates("20261124"), chips=False)
    md = edit(md, "# DMW news · Tue 24 Nov 2026\n\n", "# DMW news · Tue 24 Nov 2026\n\nNov 24, 2026 · @Saad Rehman\n\n")
    md = md.replace("Source: Argaam\n", "Source: Argaam (the run puts the article link here)\n")
    news, info = parse(md)
    assert info["sources"]["C1"] == {"source": "Argaam", "url": ""}
    assert len(news["corporate"]) == 9


# ------------------------------------------------------------------ Saad's edits win

def test_rank_edit_reorders():
    md = render(candidates("20261124"), chips=False)
    md = edit(md, row(md, "C6"), row(md, "C6").replace("| 6 |", "| 1.5 |", 1))
    _, info = parse(md)
    assert info["order"]["corporate"][:3] == ["C1", "C6", "C2"]


def test_pick_switch_in_and_out():
    md = render(candidates("20261124"), chips=False)
    md = edit(md, row(md, "C9"), row(md, "C9").replace("| Out |", "| In |"))
    md = edit(md, row(md, "C7"), row(md, "C7").replace("| In |", "| Out |"))
    _, info = parse(md)
    assert "C9" in info["order"]["corporate"] and "C7" not in info["order"]["corporate"]


def test_text_edit_wins():
    md = render(candidates("20261124"), chips=False)
    md = edit(md, "has a duration of 22 months.", "has a duration of 22 months (Saad's edit).")
    news, info = parse(md)
    c2 = news["corporate"][info["order"]["corporate"].index("C2")]
    assert "(Saad's edit)" in c2["en"]


def test_source_link_read():
    c = candidates("20261124")
    c["corporate"][1]["url"] = "https://www.saudiexchange.sa/x?anId=1"
    _, info = parse(render(c, chips=False))
    assert info["sources"]["C2"] == {"source": "Tadawul", "url": "https://www.saudiexchange.sa/x?anId=1"}


def test_left_out_switched_in_is_pending(tmp_path):
    md = render(candidates("20261124"), chips=False)
    x = next(l for l in md.splitlines() if l.startswith("| X2 |"))
    md = edit(md, x, x.replace("| Out |  |", "| In | 3 |"))
    p = tmp_path / "doc.md"; p.write_text(md, encoding="utf-8")
    rc = read_news_doc.main(["--doc", str(p), "--out", str(tmp_path / "n.json"), "--info", str(tmp_path / "i.json")])
    info = json.loads((tmp_path / "i.json").read_text(encoding="utf-8"))
    assert rc == 2 and info["pending"] == [{"id": "X2", "headline": "A board appoints a new CEO", "rank": 3.0}]
    assert (tmp_path / "n.json").exists()


def test_split_paragraph_is_joined_and_flagged():
    md = render(candidates("20261124"), chips=False)
    md = edit(md, "The contract is valued at", "\n\nThe contract is valued at")
    news, info = parse(md)
    assert any("C2" in f and "joined" in f for f in info["flags"])
    assert "Oman. The contract" in news["corporate"][1]["en"]


def test_bad_pick_and_rank_cells():
    md = render(candidates("20261124"), chips=False)
    md = edit(md, row(md, "C3"), row(md, "C3").replace("| In |", "| maybe? |"))
    md = edit(md, row(md, "C4"), row(md, "C4").replace("| 4 |", "| top |", 1))
    _, info = parse(md)
    assert "C3" in info["out"] and "C4" in info["order"]["corporate"]
    assert any("C3" in f for f in info["flags"]) and any("C4" in f for f in info["flags"])


def test_markdown_escapes_unescaped():
    md = render(candidates("20261124"), chips=False)
    md = md.replace("S&P 500", r"S&P 500").replace("(ARAMCO AB)", r"\(ARAMCO AB\)")
    news, _ = parse(md)
    assert "(ARAMCO AB)" in news["corporate"][0]["en"]


# ------------------------------------------------------------------ stops

def test_missing_arabic_on_an_in_item_is_pending():
    # Maybe items carry no Arabic until picked In (S1 time audit, 25 Nov 2026); S3 drafts it
    c = candidates("20261124")
    md = render(c, chips=False)
    md = edit(md, c["corporate"][1]["ar"] + "\n\n", "")
    news, info = parse(md)
    assert info["pending"] == [{"id": "C2", "headline": c["corporate"][1]["headline"], "rank": 2.0, "need": "AR"}]
    assert "C2" not in info["order"]["corporate"]


def test_maybe_item_without_arabic_renders_and_goes_pending_when_picked(tmp_path):
    c = candidates("20261124")
    c["corporate"][8]["ar"] = ""          # C9, Maybe
    assert write_news_doc.validate(c) == []
    md = render(c, chips=False)
    parse(md)                              # Out: fine
    md = edit(md, row(md, "C9"), row(md, "C9").replace("| Out |", "| In |"))
    _, info = parse(md)
    assert [p["id"] for p in info["pending"]] == ["C9"]
    c["corporate"][0]["ar"] = ""          # an Include item still needs its Arabic up front
    assert "C1: ar missing" in write_news_doc.validate(c)


def test_missing_arabic_on_an_out_item_is_fine():
    c = candidates("20261124")
    md = edit(render(c, chips=False), c["corporate"][8]["ar"] + "\n\n", "")
    parse(md)


def test_missing_markets_stops(tmp_path):
    c = candidates("20261124")
    md = edit(render(c, chips=False), "### TASI", "### TASI write-up")
    p = tmp_path / "d.md"; p.write_text(md, encoding="utf-8")
    assert read_news_doc.main(["--doc", str(p), "--out", str(tmp_path / "n.json")]) == 1
    assert not (tmp_path / "n.json").exists()


def test_pick_row_without_text_stops():
    md = render(candidates("20261124"), chips=False)
    md = edit(md, "### C5 · ", "### C55 · ")
    with pytest.raises(DocError, match="C5 is in the Pick table but has no item text"):
        parse(md)


# ------------------------------------------------------------------ writer

def test_writer_calls(tmp_path):
    cp = tmp_path / "c.json"; cp.write_text(json.dumps(candidates("20261124"), ensure_ascii=False), encoding="utf-8")
    assert write_news_doc.main(["--cand", str(cp), "--batch", str(tmp_path / "b.json"), "--md", str(tmp_path / "d.md")]) == 0
    b = json.loads((tmp_path / "b.json").read_text(encoding="utf-8"))
    birth = b["call1_birth"]["container"]["create"]
    assert birth["name"] == "DMW news · Tue 24 Nov 2026" and "## Pick" not in birth["doc"]["markdown"]
    op = b["call2_fill"]["batch"][1]["payload"]["ops"][0]
    body, chips = op["source"]["from"]["content"], op["source"]["blocks"]
    tokens = set(re.findall(r"<\?claude block (\w+)\?>", body))
    assert tokens == set(chips) and len(tokens) == 13 + 4
    assert chips["pc9"]["index"] == 1 and chips["pc1"]["index"] == 0 and chips["px1"]["index"] == 1
    # The chip version, with each dropdown shown as its option, reads the same as the plain one.
    shown = re.sub(r"<\?claude block (\w+)\?>", lambda m: news_doc.PICK_OPTIONS[chips[m.group(1)]["index"]], body)
    assert parse(birth["doc"]["markdown"].split("\n\n<?claude")[0] + "\n\n" + shown) == parse((tmp_path / "d.md").read_text(encoding="utf-8"))


def test_writer_needs_short_versions(tmp_path):
    c = candidates("20261124"); del c["economic"][1]["short_ar"]
    cp = tmp_path / "c.json"; cp.write_text(json.dumps(c, ensure_ascii=False), encoding="utf-8")
    assert write_news_doc.main(["--cand", str(cp), "--md", str(tmp_path / "d.md")]) == 1


def test_item_order_en_ar_source():
    md = render(candidates("20261124"), chips=False)
    block = md.split("### C2 · ")[1].split("\n\n### ")[0].split("\n\n")
    assert news_doc.lang_of(block[1]) == "en" and news_doc.lang_of(block[2]) == "ar"
    assert block[3].startswith("Source: ")


# ------------------------------------------------------------------ pick gate

@pytest.mark.parametrize("now,answered,reminded,want", [
    ("09:08", False, False, "wait"), ("09:17", False, False, "wait"), ("09:18", False, False, "remind"),
    ("09:20", False, True, "wait"), ("09:25", False, True, "go"), ("09:12", True, False, "go"),
    ("09:40", False, False, "go")])
def test_state(now, answered, reminded, want):
    assert pick_card.state(now, answered, reminded) == want


def test_card_fits_limits():
    res = pick_card.card(render(candidates("20261124"), chips=False), "https://claude.ai/artifact/x")
    b = res["card"]["blocks"]
    table = next(x for x in b if x["type"] == "table")
    assert len(table["rows"]) <= 5 and all(len(c) <= 60 for r in table["rows"] for c in r)
    assert [e["id"] for e in b[-1]["elements"]] == ["go_proposed", "doc_edited"]
    assert len(b[0]["text"]) <= 60 and len(b[1]["text"]) <= 1000
    assert table["rows"][0][0] == "E1, E2" and table["rows"][1][0] == "C1"
    assert table["rows"][-1] == ["C9, C10", "Meyar; Arab Sea", "Out (Maybe)"]
