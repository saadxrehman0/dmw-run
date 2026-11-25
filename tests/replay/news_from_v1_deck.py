#!/usr/bin/env python3
"""Replay helper: the news JSON for fill_commentary.py, read back from a published v1 edition
(EN and AR decks). Used only to replay past editions; live runs write the JSON from the Claude Doc.

  python3 news_from_v1_deck.py --en "<day>/ANB Daily Market Watch - EN.pptx" \
      --ar "<day>/ANB Daily Market Watch - AR.pptx" --out news.json
"""
import argparse, json, os, sys
from pptx import Presentation
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "../../scripts"))
from legacy_fill_commentary_v1 import find_shapes   # noqa: E402


def paras(sh):
    if sh is None:
        return []
    return [p.text.strip() for p in [x for x in sh.text_frame.paragraphs] if p.text.strip()]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--en", required=True); ap.add_argument("--ar", required=True); ap.add_argument("--out", required=True)
    a = ap.parse_args()
    got = {}
    for lang, path in (("en", a.en), ("ar", a.ar)):
        f = find_shapes(Presentation(path), lang, strict=False)
        got[lang] = {"mk": paras(f["mk"]), "ec": paras(f["ec"]),
                     "co": paras(f["co"]) + paras(f["s2"]), "s1co": len(paras(f["co"]))}
    for k in ("mk", "ec", "co"):
        if len(got["en"][k]) != len(got["ar"][k]):
            print(f"WARN {k}: EN has {len(got['en'][k])} paragraphs, AR {len(got['ar'][k])}", file=sys.stderr)
    n = lambda k: min(len(got["en"][k]), len(got["ar"][k]))
    news = {"markets": {"en": got["en"]["mk"][:2], "ar": got["ar"]["mk"][:2]},
            "economic": [{"en": got["en"]["ec"][i], "ar": got["ar"]["ec"][i]} for i in range(n("ec"))],
            "corporate": [{"en": got["en"]["co"][i], "ar": got["ar"]["co"][i]} for i in range(n("co"))],
            "_published_s1_corporate": got["en"]["s1co"]}
    json.dump(news, open(a.out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(f"markets {len(news['markets']['en'])}, economic {len(news['economic'])}, corporate {len(news['corporate'])} "
          f"(published slide 1 had {got['en']['s1co']} EN / {got['ar']['s1co']} AR)")


if __name__ == "__main__":
    main()
