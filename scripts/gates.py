#!/usr/bin/env python3
"""Gates and pings: a file mailbox between the run and the DMW Ops relay (build step 4).

A headless `claude -p` stage cannot post into a Claude project, so the run and the DMW Ops project
meet in the run folder (RUN = the edition's run folder, beside run_state.json):

  RUN/outbox/<n>-<gate>-<what>.json   the run writes a message to post (card, reminder, alert, done)
  RUN/outbox/<same name>.sent.json    the relay writes it once posted: {"message_id", "at"}
  RUN/gates/<gate>.json               an open gate: kind, times, default, the card
  RUN/inbox/<gate>.json               the answer: the relay writes Saad's, the run writes its default

The relay is a DMW Ops thread linked to the runner (rules/dmw-ops-relay.md). It never decides
anything: it posts what is in the outbox and writes back what Saad pressed or typed.

Gate kinds (plan doc, "Approval gates and pings", signed off 25 Nov 2026):
  pick       news pick. Card about 9:08, one reminder at 9:18, at 9:25 the run's own pick
             (default answer go_proposed, by "run").
  conflict   two inputs disagree. No default: the deck keeps a TK placeholder until answered,
             and S5 lists it again.
  late_news  a top-3 late item, 10:00 to 10:10. Default at 10:10: hold to the next edition.
Pings with no answer: alert (input late, stage failed, run down) and done.

Command line (what a stage session calls; times are PKT clock times):
  python gates.py open  --run RUN --id pick --kind pick --card card.json [--remind 09:18 --deadline 09:25]
  python gates.py ping  --run RUN --kind alert --text "Data file not in at 9:10; S1 goes ahead"
  python gates.py answer --run RUN --id pick --choice go_proposed [--text ...] [--by Saad]
  python gates.py status --run RUN [--id pick]      prints JSON
"""
import argparse, json, os, re, sys, time
from datetime import datetime, timedelta, timezone

PKT = timezone(timedelta(hours=5))
KINDS = {
    "pick":      {"remind": "09:18", "deadline": "09:25", "default": {"choice": "go_proposed"}},
    "conflict":  {"remind": None,    "deadline": None,    "default": None},
    "late_news": {"remind": None,    "deadline": "10:10", "default": {"choice": "hold"}},
}
PINGS = ("alert", "done", "info")


def now_iso(now=None):
    """Now on the run's clock. A stage started by a test run on a fake clock gets the offset in
    DMW_CLOCK_OFFSET (seconds), so its gate times match the orchestrator's."""
    if now is None:
        now = datetime.now(PKT) + timedelta(seconds=float(os.environ.get("DMW_CLOCK_OFFSET", "0")))
    return now.isoformat(timespec="seconds")


def _t(s):
    return datetime.strptime(s, "%H:%M").time()


def _write(path, obj):
    """Write JSON atomically (temp file then rename), so a reader never sees half a file."""
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=1)
    for i in range(5):  # Drive for desktop can hold a lock for a moment
        try:
            os.replace(tmp, path)
            return
        except PermissionError:
            time.sleep(0.5 * (i + 1))
    os.replace(tmp, path)


def _read(path):
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return None


def _dirs(run):
    return {k: os.path.join(run, k) for k in ("outbox", "gates", "inbox")}


def _next_n(run):
    box = _dirs(run)["outbox"]
    os.makedirs(box, exist_ok=True)
    ns = [int(f.split("-", 1)[0]) for f in os.listdir(box) if f[:3].isdigit()]
    return max(ns, default=0) + 1


def post(run, gate, what, msg, now=None):
    """Queue one message for the relay. msg: {"text", "card"?}. Returns the outbox path."""
    n = _next_n(run)
    path = os.path.join(_dirs(run)["outbox"], f"{n:03d}-{gate}-{what}.json")
    _write(path, dict(msg, gate=gate, what=what, queued=now_iso(now)))
    return path


def ping(run, kind, text, now=None, card=None):
    if kind not in PINGS:
        raise ValueError(f"ping kind {kind!r}, want one of {PINGS}")
    return post(run, kind, kind, {"text": text, **({"card": card} if card else {})}, now)


def open_gate(run, gid, kind, card, text, now=None, remind=None, deadline=None, default=None):
    """Open a gate and queue its card. A gate already open keeps its first card (idempotent, so a
    retried stage does not post twice)."""
    if kind not in KINDS:
        raise ValueError(f"gate kind {kind!r}, want one of {list(KINDS)}")
    path = os.path.join(_dirs(run)["gates"], f"{gid}.json")
    if _read(path):
        return _read(path)
    k = KINDS[kind]
    g = {"id": gid, "kind": kind, "opened": now_iso(now), "text": text,
         "remind": remind or k["remind"], "deadline": deadline or k["deadline"],
         "default": default if default is not None else k["default"], "reminded": False}
    _write(path, g)
    post(run, gid, "card", {"text": text, "card": card}, now)
    return g


def answer(run, gid, choice, by="Saad", text=None, now=None):
    """Record the answer. The first answer stands: a later press or the default never overwrites it."""
    path = os.path.join(_dirs(run)["inbox"], f"{gid}.json")
    old = _read(path)
    if old:
        return old
    a = {"id": gid, "choice": choice, "by": by, "at": now_iso(now)}
    if text:
        a["text"] = text
    _write(path, a)
    return a


def get_answer(run, gid):
    return _read(os.path.join(_dirs(run)["inbox"], f"{gid}.json"))


def gates(run):
    d = _dirs(run)["gates"]
    if not os.path.isdir(d):
        return {}
    out = {}
    for f in sorted(os.listdir(d)):
        if f.endswith(".json"):
            g = _read(os.path.join(d, f))
            if g:
                g["answer"] = get_answer(run, g["id"])
                out[g["id"]] = g
    return out


def tick(run, now):
    """The gate clock, called by the orchestrator each loop. Posts the one reminder, and at the
    deadline writes the default answer (by "run") and pings what was decided. Returns the ids
    resolved by default on this tick."""
    done = []
    t = now.time()
    for gid, g in gates(run).items():
        if g["answer"]:
            continue
        if g["remind"] and not g["reminded"] and t >= _t(g["remind"]) and \
                (not g["deadline"] or t < _t(g["deadline"])):
            post(run, gid, "reminder", {"text": f"Reminder: {g['text']}", "reply_to": gid}, now)
            g2 = dict(g); g2.pop("answer"); g2["reminded"] = True
            _write(os.path.join(_dirs(run)["gates"], f"{gid}.json"), g2)
        if g["deadline"] and g["default"] and t >= _t(g["deadline"]):
            a = answer(run, gid, g["default"]["choice"], by="run", now=now)
            if a["by"] == "run":
                post(run, gid, "closed", {"text": f"No answer by {g['deadline']}: the run went "
                                          f"ahead with its own choice ({a['choice']}).",
                                          "reply_to": gid}, now)
                done.append(gid)
    return done


def resolved(run, gid):
    return get_answer(run, gid) is not None


def pending_outbox(run):
    """Outbox messages the relay has not marked sent, oldest first (what the relay posts)."""
    box = _dirs(run)["outbox"]
    if not os.path.isdir(box):
        return []
    names = sorted(f for f in os.listdir(box) if re.fullmatch(r"\d{3}-[^.]+\.json", f))
    return [f for f in names if not os.path.exists(os.path.join(box, f[:-5] + ".sent.json"))]


def mark_sent(run, name, message_id, now=None):
    _write(os.path.join(_dirs(run)["outbox"], name[:-5] + ".sent.json"),
           {"message_id": message_id, "at": now_iso(now)})


def main(argv=None):
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="cmd", required=True)
    o = sub.add_parser("open")
    o.add_argument("--run", required=True); o.add_argument("--id", required=True)
    o.add_argument("--kind", required=True, choices=list(KINDS))
    o.add_argument("--card", required=True, help="JSON file: {text, card} as pick_card.py writes")
    o.add_argument("--remind"); o.add_argument("--deadline")
    p = sub.add_parser("ping")
    p.add_argument("--run", required=True); p.add_argument("--kind", required=True, choices=PINGS)
    p.add_argument("--text", required=True)
    a = sub.add_parser("answer")
    a.add_argument("--run", required=True); a.add_argument("--id", required=True)
    a.add_argument("--choice", required=True); a.add_argument("--text"); a.add_argument("--by", default="Saad")
    s = sub.add_parser("status")
    s.add_argument("--run", required=True); s.add_argument("--id")
    x = ap.parse_args(argv)
    if x.cmd == "open":
        c = json.load(open(x.card, encoding="utf-8"))
        r = open_gate(x.run, x.id, x.kind, c.get("card"), c["text"], remind=x.remind, deadline=x.deadline)
    elif x.cmd == "ping":
        r = {"queued": ping(x.run, x.kind, x.text)}
    elif x.cmd == "answer":
        r = answer(x.run, x.id, x.choice, by=x.by, text=x.text)
    else:
        r = gates(x.run)
        if x.id:
            r = r.get(x.id)
        r = {"gates": r, "outbox_pending": pending_outbox(x.run)} if not x.id else r
    print(json.dumps(r, ensure_ascii=False, indent=1))
    return 0


if __name__ == "__main__":
    sys.exit(main())
