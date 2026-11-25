#!/usr/bin/env python3
"""Read the edition's news Doc (its markdown export) and write NEWS.json for fill_commentary.py.

  python3 read_news_doc.py --doc "DMW news - 2026-11-24.md" --out NEWS.json [--info pick.json]

The run gets the export with the Claude Docs connector (export, format markdown, the Doc's one tab)
and saves the decoded text. Picked items come out in Rank order; Saad's text edits win because the
Doc is read as it stands. See news_doc.py for the layout.

Exit codes: 0 NEWS.json written; 2 NEWS.json written, but Saad switched Left out items to In and
they still need drafting (listed in --info and on stdout); 1 the Doc cannot be read safely (nothing
written).
"""
import argparse, json, os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from news_doc import DocError, parse  # noqa: E402


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--doc", required=True, help="markdown export of the news Doc")
    ap.add_argument("--out", required=True, help="NEWS.json to write")
    ap.add_argument("--info", help="optional JSON with the pick order, sources, pending and flags")
    a = ap.parse_args(argv)
    try:
        news, info = parse(open(a.doc, encoding="utf-8").read())
    except DocError as e:
        print(f"STOP {e}")
        return 1
    with open(a.out, "w", encoding="utf-8") as f:
        json.dump(news, f, ensure_ascii=False, indent=1)
    if a.info:
        with open(a.info, "w", encoding="utf-8") as f:
            json.dump(info, f, ensure_ascii=False, indent=1)
    o = info["order"]
    print(f"OK economic {' '.join(o['economic']) or '-'}; corporate {' '.join(o['corporate']) or '-'}; "
          f"out {' '.join(info['out']) or '-'}")
    for fl in info["flags"]:
        print(f"FLAG {fl}")
    if info["pending"]:
        for p in info["pending"]:
            print(f"PENDING {p['id']} {p['headline']} (rank {p['rank'] if p['rank'] is not None else 'bottom'})")
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main())
