#!/usr/bin/env python3
"""The news pick gate: the card posted to DMW Ops at about 9:08 PKT, and what the run does next.

  python3 pick_card.py card  --doc DOC_EXPORT.md --url <Doc link> --out card.json
  python3 pick_card.py state --sent 09:08 --now 09:19 [--answered] [--reminded]

`card` reads the Doc's markdown export (the run's own proposal, before Saad touches it) and writes
the card: a header, one summary line, a table of at most 5 rows, the Doc link and two buttons,
"Go with proposed" (action go_proposed) and "I edited the doc" (action doc_edited). Both buttons,
and a reply in words, end the wait; the pick itself is always read from the Doc, so a press
never needs its own data.

`state` prints one word for the run's loop (times are PKT clock times):
  wait    keep waiting
  remind  post the reminder (once, at 9:18)
  go      read the Doc now: Saad answered, or 9:25 passed (the Doc still holds the run's own
          picks, so the run goes ahead with its Include tier in its proposed order; the Done ping
          says the pick was the run's, decided 25 Nov 2026)
"""
import argparse, json, os, sys
from datetime import datetime
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from news_doc import parse  # noqa: E402

REMIND, DEADLINE = "09:18", "09:25"


def _t(s):
    return datetime.strptime(s, "%H:%M").time()


def state(now, answered=False, reminded=False, remind=REMIND, deadline=DEADLINE):
    if answered or _t(now) >= _t(deadline):
        return "go"
    if not reminded and _t(now) >= _t(remind):
        return "remind"
    return "wait"


def _ids(xs):
    if not xs:
        return "-"
    return ", ".join(xs) if len(xs) <= 3 else f"{xs[0]} to {xs[-1]}"


def _names(heads, ids, cap=60):
    """Short names for a table cell: the part of each headline before its colon, cut on a word."""
    s = "; ".join(heads.get(i, i).split(":")[0].strip() for i in ids)
    if len(s) <= cap:
        return s
    return s[:cap - 1].rsplit(" ", 1)[0].rstrip(";,") + "…"


def card(md, url, deadline="9:25 AM"):
    news, info = parse(md)
    o, out = info["order"], [x for x in info["out"] if not x.startswith("X")]
    n_left = sum(1 for x in info["out"] if x.startswith("X"))
    heads = {}
    for line in md.splitlines():
        if line.startswith("### "):
            iid, _, h = line[4:].partition(" · ")
            heads[iid.strip()] = h.strip()
    co = o["corporate"]
    rows = [[_ids(o["economic"]), _names(heads, o["economic"]), "In, Economic"]]
    if co:
        rows.append([co[0], _names(heads, co[:1]), "In, slide 1"])
    if len(co) > 1:
        s2 = co[1:8]
        rows.append([_ids(s2), f"{len(s2)} items for slide 2", "In, slide 2"])
    if len(co) > 8:
        rows.append([_ids(co[8:]), "past slide 2's 7-item room, dropped", "In, no room"])
    if out:
        rows.append([_ids(out), _names(heads, out), "Out (Maybe)"])
    summary = (f"Proposed: {len(o['economic'])} Economic and {len(co)} Corporate In, "
               f"{len(out)} Maybe Out, {n_left} left out. Change Pick or Rank in the Doc, or reply in "
               f"words such as drop C7. With no answer, the run's In items go at {deadline}.")
    blocks = [
        {"type": "header", "text": f"Pick today's news by {deadline}"},
        {"type": "section", "text": summary},
        {"type": "table", "headers": ["Items", "What", "Pick"], "rows": rows[:5]},
        {"type": "section", "label": "News doc", "text": {"type": "link", "text": "Open the news Doc", "url": url}},
        {"type": "actions", "elements": [
            {"type": "button", "id": "go_proposed", "label": "Go with proposed", "style": "primary",
             "done_label": "Going with proposed"},
            {"type": "button", "id": "doc_edited", "label": "I edited the doc", "done_label": "Reading your pick"}]},
    ]
    text = f"News pick for today: {len(o['economic'])} Economic and {len(co)} Corporate items proposed, due {deadline}."
    return {"text": text, "card": {"blocks": blocks}}


def main(argv=None):
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="cmd", required=True)
    c = sub.add_parser("card")
    c.add_argument("--doc", required=True); c.add_argument("--url", required=True); c.add_argument("--out", required=True)
    s = sub.add_parser("state")
    s.add_argument("--now", required=True); s.add_argument("--sent")
    s.add_argument("--answered", action="store_true"); s.add_argument("--reminded", action="store_true")
    a = ap.parse_args(argv)
    if a.cmd == "card":
        res = card(open(a.doc, encoding="utf-8").read(), a.url)
        json.dump(res, open(a.out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
        print(res["text"])
    else:
        print(state(a.now, a.answered, a.reminded))
    return 0


if __name__ == "__main__":
    sys.exit(main())
