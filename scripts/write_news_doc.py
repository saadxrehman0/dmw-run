#!/usr/bin/env python3
"""Render the run's news candidates into the edition's news Doc.

  python3 write_news_doc.py --cand candidates.json --batch doc_create.json [--md preview.md]

--batch writes the two Claude Docs `batch` calls that make the Doc: call 1 creates it with the
title, byline and lead; call 2 creates the Pick enum (In / Out) and adds the rest of the Doc with
its dropdowns (the enum has to exist in the Doc first). The run's session sends both, opens
nothing (the pick card links the Doc), and saves the link to the day folder.
--md writes the same Doc as plain markdown, in the shape of the Doc's own markdown export (what
read_news_doc.py parses), for logs and tests.

candidates.json (written by the S1 news sweep, items in proposed rank order):
  {"edition": "2026-11-24",
   "markets": {"global": {"en", "ar"}, "tasi": {"en", "ar"}},
   "economic":  [{"id": "E1", "headline", "tier": "Include|Maybe", "why", "en", "ar",
                  "short_en", "short_ar", "source", "url"}],
   "corporate": [{"id": "C1", "headline", "tier", "why", "en", "ar", "source", "url"}],
   "left_out":  [{"id": "X1", "headline", "source", "reason"}]}
Include items start as In, Maybe as Out: that is the pick the run goes with at 9:25.
"""
import argparse, json, os, re, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from news_doc import PICK_OPTIONS, PREFIX, render, title  # noqa: E402


def split_head(md):
    """Title, byline and lead (born with the Doc) and the rest (filled in the second call, where the
    Pick enum exists)."""
    i = md.index("\n\n## Pick")
    return md[:i] + "\n", md[i + 2:]


def validate(c):
    errs = []
    for k in ("global", "tasi"):
        for lang in ("en", "ar"):
            if not c.get("markets", {}).get(k, {}).get(lang):
                errs.append(f"markets.{k}.{lang} missing")
    ids = set()
    for sec in ("economic", "corporate", "left_out"):
        for i, it in enumerate(c.get(sec, [])):
            iid = it.get("id", "")
            if not re.fullmatch(PREFIX[sec] + r"\d+", iid):
                errs.append(f"{sec}[{i}] id {iid!r} should be {PREFIX[sec]}<n>")
            if iid in ids:
                errs.append(f"duplicate id {iid}")
            ids.add(iid)
            need = ("headline",) if sec == "left_out" else ("headline", "en", "ar")
            errs += [f"{iid}: {k} missing" for k in need if not it.get(k)]
    for i, it in enumerate(c.get("economic", [])[:2]):
        if not (it.get("short_en") and it.get("short_ar")):
            errs.append(f"{it.get('id')}: top-2 Economic items need short_en and short_ar (Saad, 25 Nov)")
    return errs


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--cand", required=True)
    ap.add_argument("--batch", help="JSON with the Claude Docs batch call's container and batch")
    ap.add_argument("--md", help="plain markdown preview, in the export's shape")
    a = ap.parse_args(argv)
    c = json.load(open(a.cand, encoding="utf-8"))
    errs = validate(c)
    if errs:
        for e in errs:
            print(f"STOP {e}")
        return 1
    if a.md:
        open(a.md, "w", encoding="utf-8").write(render(c, chips=False))
    if a.batch:
        md, blocks = render(c, chips=True)
        head, body = split_head(md)
        hb = {k: v for k, v in blocks.items() if k in ("asof", "me")}
        bb = {k: v for k, v in blocks.items() if k not in hb}
        calls = {
            "_how": "Send call 1 (birth). From its ack take frame.slug as DOC_ID and lids.n as BODY_ID, "
                    "put them in call 2 and send it. Save frame.url to the day folder as the Doc link.",
            "call1_birth": {"container": {"kind": "project", "create": {
                "name": title(c["edition"]), "doc": {"markdown": head, "blocks": hb}}}, "batch": []},
            "call2_fill": {"container": {"kind": "project", "id": "<DOC_ID>"}, "batch": [
                {"$lid": "pick", "verb": "create", "object": "enum",
                 "payload": {"value": {"name": "Pick", "options": [{"name": o} for o in PICK_OPTIONS]}}},
                {"verb": "update", "ref": {"object": "node", "id": "<BODY_ID>"}, "engine": "prose",
                 "payload": {"ops": [{"op": "insert", "target": {"kind": "root"}, "side": "end",
                                      "source": {"from": {"kind": "inline", "content": body},
                                                 "as": "markdown", "blocks": bb}}]}}]}}
        json.dump(calls, open(a.batch, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    n = {k: len(c.get(k, [])) for k in ("economic", "corporate", "left_out")}
    print(f"OK {title(c['edition'])}: {n['economic']} Economic, {n['corporate']} Corporate, {n['left_out']} left out")
    return 0


if __name__ == "__main__":
    sys.exit(main())
