"""The edition's news Doc (a Claude Doc, one per edition, in place of ANB Rough): its layout,
the writer that renders the run's picks into it, and the parser that reads it back.

Layout approved by Saad on 25 Nov 2026 (mock-up of the 24 Nov edition):

  # DMW news · Tue 24 Nov 2026
  <date chip> · <author chip>
  <lead: pick by 9:25 AM PKT ...>
  ## Pick
  Economic: ...            | Rank | Item | Pick | Run says | Why |   (Item = "E1 · <headline>")
  Corporate: ...           | Rank | Item | Pick | Run says | Why |   (Item = "C1 · <headline>")
  ## Markets
  ### Global   EN paragraph, AR paragraph
  ### TASI     EN paragraph, AR paragraph
  ## Economic
  ### E1 · <headline>   EN, AR, "Source: [Argaam](url)", "**Short version** (...)", short EN, short AR
  ## Corporate
  ### C1 · <headline>   EN, AR, "Source: [Tadawul](url)"
  ## Left out
  | ID | Headline | Source | Reason | Pick | Rank |                   (ID = X1, X2, ...)

Item order inside an item is EN, then AR, then the source line (Saad, 25 Nov 2026).
Pick is a dropdown (In / Out) in the live Doc; the markdown export renders it as plain "In" / "Out".
The item ids never change; the order lives only in the Rank column, so Saad reorders by typing
a number and the reader sorts by it. Edits Saad makes to an item's text win: the reader takes the
text as it stands in the Doc.
"""
import re
from datetime import date

PICK_OPTIONS = ("In", "Out")
SECTIONS = ("economic", "corporate")
PREFIX = {"economic": "E", "corporate": "C", "left_out": "X"}
DOT = " · "
SHORT_MARK = "**Short version** (used only if slide 1 overflows)"
LEAD = ("Pick by 9:25 AM PKT: set each item to In or Out and change Rank to reorder. The decks "
        "follow these tables; after 9:25 the run goes with its own In items in this order. Edits "
        "you make to any item's text below also win over the run's.")
INTRO = {"economic": "Economic: the deck runs the top 2; short versions are ready if slide 1 overflows.",
         "corporate": "Corporate: rank 1 goes on slide 1, ranks 2 to 8 on slide 2; lower ranks are "
                      "cut first if slide 2 overflows."}
LEFT_INTRO = ("Switch an item to In to bring it back: the run drafts its EN and AR text and adds it "
              "to Corporate at the bottom, or at the rank you type.")
ARABIC = re.compile(r"[؀-ۿ]")
LATIN = re.compile(r"[A-Za-z]")


class DocError(Exception):
    pass


def title(day):
    d = date.fromisoformat(day) if isinstance(day, str) else day
    return f"DMW news{DOT}{d.strftime('%a')} {d.day} {d.strftime('%b %Y')}"


ECON_MAX = 2   # Economic items in the deck


def default_pick(item):
    """The run's own pick: Include goes in, Maybe and anything else stays out."""
    return "In" if item.get("tier", "Include") == "Include" else "Out"


def _cell(s):
    return str(s).replace("|", "/").replace("\n", " ").strip()


def _source_line(it):
    src = it.get("source", "")
    return f"Source: [{src}]({it['url']})" if it.get("url") else f"Source: {src}"


# ---------------------------------------------------------------- writer

def render(cand, chips=True):
    """Markdown for the edition's Doc from the run's candidates.

    chips=True returns (markdown, blocks) for the Claude Docs batch create: the byline and every
    Pick cell are chip tokens, and `blocks` maps each token to its chip (the Pick dropdowns point
    at the enum "$lid:pick", created in the same batch). chips=False returns plain markdown in the
    shape of the Doc's markdown export, which is what read_news_doc.py parses."""
    blocks, md = {}, []

    def pick(key, value):
        if not chips:
            return value
        blocks[key] = {"type": "dropdown", "enum": "$lid:pick", "index": PICK_OPTIONS.index(value)}
        return f"<?claude block {key}?>"

    md.append(f"# {title(cand['edition'])}")
    if chips:
        blocks["asof"] = {"type": "date", "value": cand["edition"]}
        blocks["me"] = {"type": "mention", "user": "me"}
        md.append("<?claude block asof?> · <?claude block me?>")
    md.append(LEAD)
    md.append("## Pick")
    for sec in SECTIONS:
        md.append(INTRO[sec])
        rows = ["| Rank | Item | Pick | Run says | Why |", "| --- | --- | --- | --- | --- |"]
        n_in = 0
        for r, it in enumerate(cand[sec], 1):
            dp = it.get('pick') or default_pick(it)
            # the house runs exactly two Economic items (news-selection.md): a third Include starts
            # Out (4 Nov replay, 26 Nov 2026)
            if sec == "economic" and dp == "In":
                n_in += 1
                if n_in > ECON_MAX and not it.get('pick'):
                    dp = "Out"
            rows.append(f"| {r} | {it['id']}{DOT}{_cell(it['headline'])} | "
                        f"{pick('p' + it['id'].lower(), dp)} | "
                        f"{_cell(it.get('tier', 'Include'))} | {_cell(it.get('why', ''))} |")
        md.append("\n".join(rows))
    md.append("## Markets")
    for key, head in (("global", "Global"), ("tasi", "TASI")):
        m = cand["markets"][key]
        md += [f"### {head}", m["en"], m["ar"]]
    for sec, head in (("economic", "Economic"), ("corporate", "Corporate")):
        md.append(f"## {head}")
        for it in cand[sec]:
            # Maybe items carry no Arabic until picked In (S1 time audit, 25 Nov 2026)
            md += [f"### {it['id']}{DOT}{it['headline']}", it["en"]] + ([it["ar"]] if it.get("ar") else []) + \
                [_source_line(it)]
            if it.get("short_en") or it.get("short_ar"):
                md += [SHORT_MARK, it.get("short_en", ""), it.get("short_ar", "")]
    md.append("## Left out")
    md.append(LEFT_INTRO)
    rows = ["| ID | Headline | Source | Reason | Pick | Rank |", "| --- | --- | --- | --- | --- | --- |"]
    for it in cand.get("left_out", []):
        rows.append(f"| {it['id']} | {_cell(it['headline'])} | {_cell(it.get('source', ''))} | "
                    f"{_cell(it.get('reason', ''))} | {pick('p' + it['id'].lower(), 'Out')} |  |")
    md.append("\n".join(rows))
    text = "\n\n".join(x for x in md if x is not None) + "\n"
    return (text, blocks) if chips else text


# ---------------------------------------------------------------- parser

_ESC = re.compile(r"\\([\\`*_{}\[\]()#+\-.!|>~&<])")


def unescape(s):
    return _ESC.sub(r"\1", s).strip()


def lang_of(p):
    """'ar' when Arabic letters outnumber Latin ones (AR items carry tickers and index names)."""
    a, l = len(ARABIC.findall(p)), len(LATIN.findall(p))
    return "ar" if a > l else "en"


def _split_row(line):
    cells = line.strip().strip("|").split("|")
    return [unescape(c) for c in cells]


def blocks_of(md):
    """The export as a list of ('h', level, text) / ('p', text) / ('t', [rows]) blocks."""
    out, para, table = [], [], []

    def flush():
        if para:
            out.append(("p", unescape(" ".join(para))))
            para.clear()
        if table:
            rows = [r for r in table if not re.fullmatch(r"\|?\s*:?-{3,}:?\s*(\|\s*:?-{3,}:?\s*)*\|?", r.strip())]
            out.append(("t", [_split_row(r) for r in rows]))
            table.clear()

    for raw in md.replace("\r\n", "\n").split("\n"):
        line = raw.rstrip()
        h = re.match(r"^(#{1,6})\s+(.*)$", line)
        if h:
            flush()
            out.append(("h", len(h.group(1)), unescape(h.group(2))))
        elif line.strip().startswith("|"):
            if para:
                flush()
            table.append(line)
        elif not line.strip():
            flush()
        else:
            if table:
                flush()
            para.append(line.strip())
    flush()
    return out


def _sections(blocks):
    """{h2 title (lower): [blocks under it]}"""
    secs, cur = {}, None
    for b in blocks:
        if b[0] == "h" and b[1] <= 2:
            cur = b[2].strip().lower()
            secs[cur] = []
        elif cur is not None:
            secs[cur].append(b)
    return secs


def _items(blocks):
    """[(h3 text, [paragraphs])] under one h2."""
    items = []
    for b in blocks:
        if b[0] == "h":
            items.append((b[2], []))
        elif b[0] == "p" and items:
            items[-1][1].append(b[1])
    return items


ID_RE = re.compile(r"^\s*([ECX])(\d+)\b\s*[·:.\-]?\s*(.*)$")
LINK = re.compile(r"\[([^\]]+)\]\((https?://[^)\s]+)\)")


def _parse_item(head, paras, flags):
    m = ID_RE.match(head)
    iid = f"{m.group(1)}{m.group(2)}" if m else head
    it = {"id": iid, "headline": m.group(3) if m else head, "en": [], "ar": [],
          "short_en": [], "short_ar": [], "source": "", "url": ""}
    short = False
    for p in paras:
        if re.match(r"^\**\s*short version", p, re.I):
            short = True
            continue
        if re.match(r"^source\s*:", p, re.I):
            link = LINK.search(p)
            if link:
                it["source"], it["url"] = link.group(1).strip(), link.group(2)
            else:
                it["source"] = re.sub(r"\(.*?\)", "", p.split(":", 1)[1]).strip()
            continue
        lang = lang_of(p)
        it[("short_" if short else "") + lang].append(p)
    for k in ("en", "ar", "short_en", "short_ar"):
        if len(it[k]) > 1:
            flags.append(f"{iid}: {len(it[k])} {k.upper().replace('_', ' ')} paragraphs joined into one")
        it[k] = " ".join(it[k])
    return it


def _rank(s, fallback):
    try:
        return float(s.strip())
    except (ValueError, AttributeError):
        return None if not str(s).strip() else fallback


def _pick_rows(table, flags):
    """Rows of a Pick or Left out table as dicts keyed by the header."""
    if not table:
        return []
    head = [h.strip().lower() for h in table[0]]
    rows = []
    for pos, r in enumerate(table[1:]):
        r = (r + [""] * len(head))[:len(head)]
        d = dict(zip(head, r))
        item = d.get("item") or d.get("id") or ""
        m = ID_RE.match(item)
        if not m:
            flags.append(f"Pick row without an item id skipped: {item[:50]!r}")
            continue
        iid = f"{m.group(1)}{m.group(2)}"
        pick = d.get("pick", "").strip().lower()
        if pick not in ("in", "out"):
            flags.append(f"{iid}: Pick is {d.get('pick', '')!r}, read as Out")
        rank_s = d.get("rank", "")
        rank = _rank(rank_s, None)
        if rank_s.strip() and rank is None:
            flags.append(f"{iid}: Rank {rank_s!r} is not a number; it keeps its row position")
        rows.append({"id": iid, "pick": pick == "in", "rank": rank, "pos": pos,
                     "headline": m.group(3) if d.get("item") else d.get("headline", "")})
    return rows


def parse(md):
    """Read the Doc's markdown export. Returns (news, info).

    news is the NEWS.json fill_commentary.py reads. info holds what the run logs: the picked ids
    in order, every item's source and link, items Saad moved in from Left out that still need
    drafting (pending), and flags. Raises DocError for anything that would put a wrong or missing
    text in a deck: a missing Markets bullet, a picked item with no EN or AR, an unknown id."""
    flags = []
    secs = _sections(blocks_of(md))
    for need in ("pick", "markets", "economic", "corporate"):
        if need not in secs:
            raise DocError(f"The Doc has no '## {need.title()}' section.")

    markets = {"en": [], "ar": []}
    mk = {h.strip().lower(): ps for h, ps in _items(secs["markets"])}
    for key in ("global", "tasi"):
        ps = mk.get(key)
        if ps is None:
            raise DocError(f"Markets has no '### {key.upper() if key == 'tasi' else 'Global'}' bullet.")
        en = [p for p in ps if lang_of(p) == "en" and not re.match(r"^source\s*:", p, re.I)]
        ar = [p for p in ps if lang_of(p) == "ar"]
        if not en or not ar:
            raise DocError(f"Markets {key}: needs an EN and an AR paragraph (found {len(en)} EN, {len(ar)} AR).")
        if len(en) > 1 or len(ar) > 1:
            flags.append(f"Markets {key}: several paragraphs joined into one")
        markets["en"].append(" ".join(en))
        markets["ar"].append(" ".join(ar))

    bodies = {}
    for sec in SECTIONS:
        for head, ps in _items(secs[sec]):
            it = _parse_item(head, ps, flags)
            if it["id"] in bodies:
                raise DocError(f"Item {it['id']} appears twice in the Doc.")
            it["section"] = sec
            bodies[it["id"]] = it

    rows = []
    for b in secs["pick"]:
        if b[0] == "t":
            rows += _pick_rows(b[1], flags)
    left = []
    for b in secs.get("left out", []):
        if b[0] == "t":
            left += _pick_rows(b[1], flags)

    news = {"markets": markets, "economic": [], "corporate": []}
    info = {"order": {"economic": [], "corporate": []}, "out": [], "pending": [], "flags": flags,
            "sources": {}}
    seen = set()
    for sec in SECTIONS:
        pre = PREFIX[sec]
        srows = [r for r in rows if r["id"].startswith(pre)]
        for r in srows:
            if r["id"] in seen:
                raise DocError(f"{r['id']} has two rows in the Pick tables.")
            seen.add(r["id"])
            if r["id"] not in bodies:
                raise DocError(f"{r['id']} is in the Pick table but has no item text in the Doc.")
        # Rank wins; a blank or bad rank keeps the row's place; ties keep table order.
        srows.sort(key=lambda r: (r["rank"] if r["rank"] is not None else r["pos"] + 1, r["pos"]))
        for r in srows:
            if not r["pick"]:
                info["out"].append(r["id"])
                continue
            it = bodies[r["id"]]
            if not it["en"]:
                raise DocError(f"{r['id']} is picked In but has no EN text.")
            if not it["ar"]:   # a Maybe item picked In: S3 drafts its Arabic, then reads again
                info["pending"].append({"id": r["id"], "headline": it["headline"], "rank": r["rank"],
                                        "need": "AR"})
                continue
            entry = {"en": it["en"], "ar": it["ar"]}
            if sec == "economic":
                for k in ("short_en", "short_ar"):
                    if it[k]:
                        entry[k] = it[k]
                if bool(it["short_en"]) != bool(it["short_ar"]):
                    flags.append(f"{r['id']}: short version has only {'EN' if it['short_en'] else 'AR'}")
            news[sec].append(entry)
            info["order"][sec].append(r["id"])
            info["sources"][r["id"]] = {"source": it["source"], "url": it["url"]}
    for iid, it in bodies.items():
        if iid not in seen:
            flags.append(f"{iid} has item text but no Pick row; left out")
    for r in left:
        if r["pick"]:
            info["pending"].append({"id": r["id"], "headline": r["headline"], "rank": r["rank"]})
        else:
            info["out"].append(r["id"])
    if len(news["economic"]) != 2:
        flags.append(f"{len(news['economic'])} Economic items picked (the deck runs 2)")
    if not news["corporate"]:
        flags.append("No Corporate items picked")
    return news, info
