#!/usr/bin/env python3
"""Stand-in for a stage's Claude session, for orchestrator and gate tests (run_edition.py --stub).

  python stub_stage.py S1        (reads DMW_RUN, DMW_RESULT, DMW_STAGE from the environment)

S1 does the real gate work with the 24 Nov 2026 news fixture: renders the news Doc as markdown,
builds the pick card with pick_card.py (the Doc link is the 24 Nov mock-up Doc) and opens the
`pick` gate. S3 reads the pick answer and reports it. Every other stage sleeps briefly.
DMW_STUB_FAIL=S2,S4 makes those stages exit 1 (retry and alert tests); DMW_STUB_SLEEP sets the
sleep in seconds (default 1).
"""
import json, os, sys, time

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
sys.path.insert(0, HERE)
import gates  # noqa: E402

MOCK_DOC = "https://claude.ai/artifact/8EEXKGnoshg7bzf3JUW2wg"


def main():
    sid = sys.argv[1]
    run, res = os.environ["DMW_RUN"], os.environ["DMW_RESULT"]
    time.sleep(float(os.environ.get("DMW_STUB_SLEEP", "1")))
    if sid in os.environ.get("DMW_STUB_FAIL", "").split(","):
        print(f"stub {sid}: failing on purpose")
        return 1
    out = {"ok": True, "summary": f"stub {sid}"}
    if sid == "S1":
        from news_doc import render
        import pick_card
        cand = json.load(open(os.path.join(REPO, "tests", "news_doc", "fixtures", "candidates_20261124.json"),
                              encoding="utf-8"))
        md = render(cand, chips=False)
        os.makedirs(os.path.join(run, "news"), exist_ok=True)
        open(os.path.join(run, "news", "doc.md"), "w", encoding="utf-8").write(md)
        c = pick_card.card(md, MOCK_DOC)
        g = gates.open_gate(run, "pick", "pick", c["card"], c["text"])
        out["summary"] = f"stub S1: pick gate open ({g['opened']}), 24 Nov fixture"
    elif sid == "S3":
        a = gates.get_answer(run, "pick")
        out["summary"] = f"stub S3: pick answered by {a['by']}: {a['choice']}" + \
            (f" ({a['text']})" if a.get("text") else "")
        out["pick"] = a
    print(out["summary"])
    with open(res, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False)
    return 0


if __name__ == "__main__":
    sys.exit(main())
