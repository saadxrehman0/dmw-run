#!/usr/bin/env python3
"""ANB Rough position map (open item 3, 29 Oct 2026).

Keeps a small model of the Draft tab (each paragraph's text and style, so every index is known)
plus the doc's revisionId, so Claude can write straight away with
writeControl.requiredRevisionId instead of a full read_doc before every write. Google refuses a
write whose required revision is not the latest ("does not match the latest revision"), and then
nothing is applied: re-read, rebuild the map, rebuild the requests, retry.

Tested 29 Oct 2026 on a scratch doc: an accepted write returns writeControl.requiredRevisionId
(the new revision), so Claude's own consecutive writes chain without any read.

  build   READ_DOC.json [--tab t.0] --out MAP.json     map from a saved read_doc result
  apply   MAP.json REQUESTS.json --revision REV        replay an accepted batch on the map
                                                       (REQUESTS = the requests list, or the
                                                       whole update_doc arguments)
  show    MAP.json [SECTION ...]                       compact index listing per section
  verify  MAP.json READ_DOC.json                       does the map match a fresh read?

Paragraph text in the map uses U+FFFC for non-text elements (chips, images) and for tables, so
lengths stay exact. Requests the model cannot replay (tables, images, page breaks) mark the map
stale: the next write must start from a fresh read.
"""
import argparse
import json
import re
import sys
from datetime import datetime, timezone
from pathlib import Path

OBJ = "￼"
SECTIONS = ("Economic News", "Corporate News", "Global", "TASI")


def u16(s):
    return len(s.encode("utf-16-le")) // 2


def now():
    return datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")


# ---------------------------------------------------------------- build
def _tab(d, tab_id):
    d = d.get("content", d)
    for t in d["tabs"]:
        if t["tabProperties"]["tabId"] == tab_id:
            return d, t
    raise SystemExit(f"tab {tab_id} not in the doc")


def paras_from_doc(path, tab_id):
    d, t = _tab(json.load(open(path, encoding="utf-8")), tab_id)
    out = []
    for el in t["documentTab"]["body"]["content"]:
        if "paragraph" in el:
            txt = ""
            for e in el["paragraph"]["elements"]:
                n = e["endIndex"] - e["startIndex"]
                c = e.get("textRun", {}).get("content")
                txt += c if c is not None and u16(c) == n else OBJ * n
            out.append({"style": el["paragraph"]["paragraphStyle"].get("namedStyleType", ""), "text": txt})
        elif "startIndex" in el:
            out.append({"style": "BLOCK", "text": OBJ * (el["endIndex"] - el["startIndex"])})
    base = next(el["startIndex"] for el in t["documentTab"]["body"]["content"] if "startIndex" in el)
    return d, base, out


def build(a):
    d, base, paras = paras_from_doc(a.doc, a.tab)
    m = {"doc": d.get("documentId"), "tab": a.tab, "revisionId": d.get("revisionId"), "saved": now(),
         "stale": False, "base": base, "paras": paras}
    Path(a.out).write_text(json.dumps(m, ensure_ascii=False, indent=0), encoding="utf-8")
    print(f"Map saved: {len(paras)} paragraphs, revision {m['revisionId'][:12]}...")


# ---------------------------------------------------------------- model edits
def _chars(m):
    """Per UTF-16 unit: (char, style) where a paragraph's style sits on its closing newline and
    table/object blocks carry 'BLOCK'."""
    out = []
    for p in m["paras"]:
        for ch in p["text"]:
            units = u16(ch)
            st = "BLOCK" if p["style"] == "BLOCK" else (p["style"] if ch == "\n" else None)
            out.append((ch, st))
            out.extend([("", None)] * (units - 1))   # astral chars take two units
    return out


def _paras(chars):
    paras, cur, block = [], "", False
    for ch, st in chars:
        if st == "BLOCK":
            if cur and not block:
                paras.append({"style": "NORMAL_TEXT", "text": cur}); cur = ""
            block = True; cur += ch
            continue
        if block:
            paras.append({"style": "BLOCK", "text": cur}); cur = ""; block = False
        cur += ch
        if ch == "\n":
            paras.append({"style": st or "NORMAL_TEXT", "text": cur}); cur = ""
    if cur:
        paras.append({"style": "BLOCK" if block else "NORMAL_TEXT", "text": cur})
    return paras


def _next_style(chars, i):
    for ch, st in chars[i:]:
        if ch == "\n":
            return st
    return "NORMAL_TEXT"


def insert(chars, i, text):
    st = _next_style(chars, i)
    new = []
    for ch in text:
        new.append((ch, st if ch == "\n" else None))
        new.extend([("", None)] * (u16(ch) - 1))
    chars[i:i] = new


def apply(a):
    m = json.loads(Path(a.map).read_text(encoding="utf-8"))
    req = json.loads(Path(a.requests).read_text(encoding="utf-8"))
    req = req.get("requests", req) if isinstance(req, dict) else req
    base, chars, tab = m["base"], _chars(m), m["tab"]
    for r in req:
        (kind, body), = r.items()
        loc_tab = (body.get("location") or body.get("range") or body.get("endOfSegmentLocation") or {}).get("tabId", tab)
        if loc_tab != tab:
            continue
        if kind == "insertText":
            i = body["location"]["index"] - base if "location" in body else len(chars) - 1
            insert(chars, i, body["text"])
        elif kind == "deleteContentRange":
            s, e = body["range"]["startIndex"] - base, body["range"]["endIndex"] - base
            del chars[s:e]
        elif kind == "replaceAllText":
            tabs = body.get("tabsCriteria", {}).get("tabIds")
            if tabs and tab not in tabs:
                continue
            crit = body["containsText"]
            pat = crit["text"] if crit.get("searchByRegex") else re.escape(crit["text"])
            flat = "".join(c for c, _ in chars)   # fine for BMP text, which the doc is
            hits = list(re.finditer(pat, flat, 0 if crit.get("matchCase") else re.I))
            for h in reversed(hits):
                del chars[h.start():h.end()]
                insert(chars, h.start(), body["replaceText"])
        elif kind == "updateParagraphStyle" and "namedStyleType" in body.get("fields", "") \
                and "namedStyleType" in body.get("paragraphStyle", {}):
            s, e = body["range"]["startIndex"] - base, body["range"]["endIndex"] - base
            k = s
            while k < len(chars):
                if chars[k][0] == "\n":
                    chars[k] = ("\n", body["paragraphStyle"]["namedStyleType"])
                    if k >= e - 1:
                        break
                k += 1
        elif kind in ("updateTextStyle", "updateParagraphStyle", "createParagraphBullets",
                      "deleteParagraphBullets"):
            pass
        else:
            m["stale"] = True
            print(f"Map marked stale: cannot replay '{kind}'. Re-read before the next write.")
    m["paras"] = _paras(chars)
    m["revisionId"], m["saved"] = a.revision, now()
    Path(a.out or a.map).write_text(json.dumps(m, ensure_ascii=False, indent=0), encoding="utf-8")
    print(f"Map updated to revision {a.revision[:12]}...: {len(m['paras'])} paragraphs.")


# ---------------------------------------------------------------- show / verify
def indexed(m):
    i = m["base"]
    for n, p in enumerate(m["paras"]):
        L = u16(p["text"])
        yield n, i, i + L, p
        i += L


def show(a):
    m = json.loads(Path(a.map).read_text(encoding="utf-8"))
    print(f"revision {m['revisionId']}  saved {m['saved']}" + ("  STALE: re-read first" if m.get("stale") else ""))
    want = set(a.sections) if a.sections else None
    sec = None
    for n, s, e, p in indexed(m):
        t = p["text"].rstrip("\n")
        if p["style"].startswith("HEADING") and t.strip():
            sec = next((x for x in SECTIONS if t.strip().startswith(x)), t.strip()[:30])
        if want and sec not in want:
            continue
        kind = "H" if p["style"].startswith("HEADING") else ("T" if p["style"] == "BLOCK" else
               ("-" if not t.strip() else ("L" if t.startswith("http") else
               ("A" if re.match(r"\s*[؀-ۿ]", t) else "E"))))
        print(f"{s:6d}-{e:<6d} {kind} {t[:60]}")


def verify(a):
    m = json.loads(Path(a.map).read_text(encoding="utf-8"))
    d, base, paras = paras_from_doc(a.doc, m["tab"])
    same = base == m["base"] and [(p["style"], p["text"]) for p in paras] == [(p["style"], p["text"]) for p in m["paras"]]
    print(("MATCH" if same else "DIFFERENT") + f": map rev {m['revisionId'][:12]}..., doc rev {d.get('revisionId', '')[:12]}...")
    if not same:
        for k, (x, y) in enumerate(zip(m["paras"], paras)):
            if (x["style"], x["text"]) != (y["style"], y["text"]):
                print(f"first difference at paragraph {k}: map {x['style']} {x['text'][:50]!r} vs doc {y['style']} {y['text'][:50]!r}")
                break
        else:
            print(f"paragraph count: map {len(m['paras'])}, doc {len(paras)}")
    sys.exit(0 if same else 1)


def main():
    ap = argparse.ArgumentParser(); sub = ap.add_subparsers(dest="cmd", required=True)
    b = sub.add_parser("build"); b.add_argument("doc"); b.add_argument("--tab", default="t.0"); b.add_argument("--out", required=True)
    p = sub.add_parser("apply"); p.add_argument("map"); p.add_argument("requests"); p.add_argument("--revision", required=True); p.add_argument("--out")
    s = sub.add_parser("show"); s.add_argument("map"); s.add_argument("sections", nargs="*")
    v = sub.add_parser("verify"); v.add_argument("map"); v.add_argument("doc")
    a = ap.parse_args()
    {"build": build, "apply": apply, "show": show, "verify": verify}[a.cmd](a)


if __name__ == "__main__":
    main()
