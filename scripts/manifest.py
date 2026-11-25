#!/usr/bin/env python3
"""Script hashes for the S0 preflight check (plan doc: "Script differs from master").

  python manifest.py --write     after a change is merged: rewrite scripts/MANIFEST.json
  python manifest.py --check     S0: exit 1 and list files whose sha256 differs from the manifest
Covers scripts/*.py, scripts/*.ps1, browser/*.js and orchestrator/ (stages and prompts).
"""
import argparse, glob, hashlib, json, os, sys
from datetime import datetime, timezone

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PATH = os.path.join(REPO, "scripts", "MANIFEST.json")
GLOBS = ("scripts/*.py", "scripts/*.ps1", "browser/*.js", "orchestrator/stages.json",
         "orchestrator/prompts/*.md")


def hashes():
    out = {}
    for g in GLOBS:
        for p in sorted(glob.glob(os.path.join(REPO, g))):
            # text files hash with LF line ends, so a Windows checkout matches the repo
            data = open(p, "rb").read().replace(b"\r\n", b"\n")
            out[os.path.relpath(p, REPO).replace("\\", "/")] = hashlib.sha256(data).hexdigest()
    return out


def main(argv=None):
    ap = argparse.ArgumentParser()
    g = ap.add_mutually_exclusive_group(required=True)
    g.add_argument("--write", action="store_true"); g.add_argument("--check", action="store_true")
    a = ap.parse_args(argv)
    now = hashes()
    if a.write:
        json.dump({"written": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC"), "files": now},
                  open(PATH, "w", encoding="utf-8", newline="\n"), indent=1)
        print(f"{len(now)} files")
        return 0
    want = json.load(open(PATH, encoding="utf-8"))["files"]
    bad = sorted(k for k in set(want) | set(now) if want.get(k) != now.get(k))
    print(json.dumps({"ok": not bad, "differ": bad}, indent=1))
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
