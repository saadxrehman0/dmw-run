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


def pair(en, ar):
    """Match each EN item to its AR item: the published AR decks do not always keep the EN order
    (11 Nov 2026). Score = shared tickers x10 + shared numbers; ties keep the position."""
    from news_text import numbers, tickers
    left, out = list(range(len(ar))), []
    for i, e in enumerate(en):
        te, ne = set(tickers(e)), set(numbers(e, "en"))
        best = max(left, key=lambda j: (10 * len(te & set(tickers(ar[j]))) + len(ne & set(numbers(ar[j], "ar"))),
                                        -abs(j - i)), default=None)
        if best is None:
            print(f"WARN no AR item left for: {e[:60]}", file=sys.stderr)
            continue
        left.remove(best)
        out.append({"en": e, "ar": ar[best]})
    for j in left:
        print(f"WARN AR item with no EN match: {ar[j][:60]}", file=sys.stderr)
    return out


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
    news = {"markets": {"en": got["en"]["mk"][:2], "ar": got["ar"]["mk"][:2]},
            "economic": pair(got["en"]["ec"], got["ar"]["ec"]),
            "corporate": pair(got["en"]["co"], got["ar"]["co"]),
            "_published_s1_corporate": got["en"]["s1co"]}
    json.dump(news, open(a.out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(f"markets {len(news['markets']['en'])}, economic {len(news['economic'])}, corporate {len(news['corporate'])} "
          f"(published slide 1 had {got['en']['s1co']} EN / {got['ar']['s1co']} AR)")


if __name__ == "__main__":
    main()
