#!/usr/bin/env python3
"""The edition run: one script the runner's scheduler starts at 8:55 AM PKT, Sunday to Thursday
(build step 3 of the plan doc, signed off 25 Nov 2026).

  python run_edition.py [--edition 2026-11-26] [--config orchestrator/config.local.json]
                        [--stub] [--clock 09:07 --speed 1] [--max-minutes N]

It only starts and tracks stages. Each stage (orchestrator/stages.json) is a headless Claude Code
session (`claude -p`, permission mode dontAsk, that stage's allowed tools) given its prompt file
(orchestrator/prompts/<id>.md). A stage starts when its start time has passed, the stages it needs
are done (or finished), its gate is answered and its input is ready. It ends when the session exits
0 AND it wrote RUN/stages/<id>.result.json with "ok": true. A failed or timed-out attempt is retried
once; the second failure pings DMW Ops and blocks the stages that need it done.

Everything the run knows is in RUN/run_state.json (RUN = <runs root>/<edition>/), rewritten every
loop, so a restart picks up where it stopped and the DMW Ops relay can read the heartbeat.
Gates and pings go through the gate mailbox (gates.py).

--stub runs scripts/stub_stage.py instead of Claude for every stage (orchestrator and gate tests).
--clock HH:MM starts a fake PKT clock at that time on the edition date; --speed N runs it N times
faster than real time. Both are for tests only.

--replay runs a past edition (build step 5, 25 Nov 2026): a fake clock from 8:55 that jumps to the
next start time, round or gate deadline whenever no stage is running; RUN = <runs root>/replay/
<edition>/; DAY = RUN/day, a copy of only the technical team's inputs and the calendar file (never
the published finals, which the comparison reads later); S4 runs one round; the pick gate takes the
run's own pick at 9:25; no email fallback. Each stage also gets orchestrator/prompts/_replay.md with
the cut-off times for sources.
"""
import argparse, json, os, shutil, signal, subprocess, sys, time
from datetime import date, datetime, timedelta

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
sys.path.insert(0, HERE)
import gates  # noqa: E402
from gates import PKT, _read, _write  # noqa: E402

TERMINAL = ("done", "failed", "blocked")
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
DEFAULTS = {
    "runs_root": os.path.join(REPO, "out", "runs"),
    "day_root": r"G:\.shortcut-targets-by-id\18rkHxDXGFRnp_kbzgyFhbwCjVH4-XOpc\2026\03. Submission",
    "claude": "claude",
    "python": sys.executable,
    "model": None,
    "holidays": os.path.join(REPO, "orchestrator", "holidays.json"),
    "mode_file": os.path.join(REPO, "orchestrator", "mode.json"),
    "tick_seconds": 15,
    "alert_email": None,      # Saad's address, set only in config.local.json on the runner
    "mail_after_min": 5,      # email a card or alert the relay has not posted after this long
}


# ---------------------------------------------------------------- clock

class Clock:
    """PKT wall clock, or a fake one for tests: starts at `start` on the edition date and runs
    `speed` times faster than real time."""

    def __init__(self, edition, start=None, speed=1.0):
        self.fake = start is not None
        self.speed = speed
        self.t0 = time.monotonic()
        self.offset = 0.0     # seconds added by jump() (replay clock)
        if self.fake:
            hh, mm = map(int, start.split(":"))
            self.base = datetime(edition.year, edition.month, edition.day, hh, mm, tzinfo=PKT)

    def now(self):
        if not self.fake:
            return datetime.now(PKT)
        return self.base + timedelta(seconds=(time.monotonic() - self.t0) * self.speed + self.offset)

    def jump(self, to):
        """Replay clock: move a fake clock forward to `to` (never back)."""
        if self.fake and to > self.now():
            self.offset += (to - self.now()).total_seconds()

    def at(self, hhmm, edition):
        hh, mm = map(int, hhmm.split(":"))
        return datetime(edition.year, edition.month, edition.day, hh, mm, tzinfo=PKT)


# ---------------------------------------------------------------- setup

def load_config(path):
    cfg = dict(DEFAULTS)
    if path and os.path.exists(path):
        cfg.update(json.load(open(path, encoding="utf-8-sig")))  # Notepad and PowerShell 5 add a BOM
    return cfg


def day_folder(day_root, ed):
    """<root>\\<MM. Mon>\\<YYYYMMDD> (config.md)."""
    return os.path.join(day_root, f"{ed.month:02d}. {MONTHS[ed.month - 1]}", ed.strftime("%Y%m%d"))


def edition_day(ed, holidays_file):
    """None if the run should go ahead, else the reason to skip (Friday, Saturday, a Tadawul
    holiday from Saad's yearly list; plan doc open questions, 25 Nov 2026)."""
    if ed.weekday() in (4, 5):
        return f"{ed:%A}: no edition"
    hol = _read(holidays_file) or {}
    for d in hol.get(str(ed.year), []):
        if d == ed.isoformat() or (isinstance(d, dict) and d.get("date") == ed.isoformat()):
            return "Tadawul holiday"
    return None


def previous_edition(ed, holidays_file):
    d = ed - timedelta(days=1)
    while edition_day(d, holidays_file):
        d -= timedelta(days=1)
    return d


def read_mode(path):
    m = (_read(path) or {}).get("mode", "practice")
    return m if m in ("practice", "live") else "practice"


def fill(s, ctx):
    for k, v in ctx.items():
        s = s.replace("{" + k + "}", str(v))
    return s


REPLAY_INPUTS = ("Manual File-Calendar.xlsx",)


def prepare_replay_day(src, dst):
    """Copy only what the run may read on the morning (the technical team's data file, Graphs and
    Technicals, and the calendar file the previous edition's S6 would have made) into dst. The
    published decks, PDFs, EMAIL MB and the team's Results file stay out: the comparison reads them
    from src afterwards."""
    import find_inputs
    os.makedirs(dst, exist_ok=True)
    if not os.path.isdir(src):
        return []
    f = find_inputs.find(src)
    paths = [f["data"]] + list(f["graphs"].values()) + list(f["technicals"].values()) + \
        [os.path.join(src, n) for n in REPLAY_INPUTS]
    prev = previous_day(src)
    if prev:
        for n in ("ANB Daily Market Watch - EN.pptx", "ANB Daily Market Watch - AR.pptx"):
            if os.path.isfile(os.path.join(prev, n)):
                paths.append((os.path.join(prev, n), os.path.join("prev", n)))
    res = previous_day(src, need="DMW - Results.xlsx")   # S4 starts its workbook from the last one
    if res:
        paths.append((os.path.join(res, "DMW - Results.xlsx"), os.path.join("prev", "DMW - Results.xlsx")))
    copied = []
    for p in paths:
        p, rel = p if isinstance(p, tuple) else (p, os.path.basename(p) if p else None)
        if p and os.path.isfile(p):
            q = os.path.join(dst, rel)
            os.makedirs(os.path.dirname(q), exist_ok=True)
            if not os.path.exists(q):
                shutil.copy2(p, q)
            copied.append(rel)
    return copied


def previous_day(src, need="ANB Daily Market Watch - EN.pptx"):
    """The latest YYYYMMDD day folder before src holding `need`, in src's month folder or the
    month before."""
    name = os.path.basename(os.path.normpath(src))
    month = os.path.dirname(os.path.normpath(src))
    cands = []
    for m in (month, os.path.join(os.path.dirname(month), _prev_month_name(os.path.basename(month)))):
        if m and os.path.isdir(m):
            cands += [os.path.join(m, d) for d in os.listdir(m)
                      if len(d) == 8 and d.isdigit() and d < name and
                      os.path.isfile(os.path.join(m, d, need))]
    return max(cands, key=os.path.basename) if cands else None


def _prev_month_name(m):
    try:
        n = int(m.split(".")[0])
    except ValueError:
        return ""
    n = 12 if n == 1 else n - 1
    return f"{n:02d}. {MONTHS[n - 1]}"


# ---------------------------------------------------------------- state

def new_state(ed, mode, spec, run, day):
    st = {"edition": ed.isoformat(), "mode": mode, "run": run, "day": day, "started": None,
          "heartbeat": None, "ended": None, "alerts": [], "stages": {}}
    for s in spec["stages"]:
        st["stages"][s["id"]] = {"name": s["name"], "status": "waiting", "attempts": 0,
                                 "started": None, "ended": None, "exit": None, "pid": None,
                                 "log": None, "note": None, "round": 0, "next_at": None}
    return st


def save(run, st):
    _write(os.path.join(run, "run_state.json"), st)


def alert_once(run, st, key, text, now):
    if key in st["alerts"]:
        return
    st["alerts"].append(key)
    gates.ping(run, "alert", text, now)


# ---------------------------------------------------------------- processes

def kill_tree(pid):
    try:
        if os.name == "nt":
            subprocess.run(["taskkill", "/T", "/F", "/PID", str(pid)], capture_output=True, timeout=30)
        else:
            os.killpg(os.getpgid(pid), signal.SIGKILL)
    except Exception:  # already gone
        pass


def allowed_tools(tools, python):
    """stages.json lists Bash rules. On Windows, Claude Code's shell tool is PowerShell (no Git
    Bash on the runner; replay pilot, 25 Nov 2026), and the prompts call python by its full path
    ({PY}), which "python *" does not match. So every Bash(x) rule also goes in as PowerShell(x),
    and the python rules also go in for the configured interpreter's path."""
    out = []
    for t in tools:
        rules = [t]
        m = t.startswith("Bash(") and t[5:-1]
        if m:
            rules.append(f"PowerShell({m})")
            if m.split(" ")[0] in ("python", "python3", "py") and python:
                for q in (python, f'"{python}"', f"& {python}", f'& "{python}"'):
                    rules += [f"Bash({q} *)", f"PowerShell({q} *)"]
        for r in rules:
            if r not in out:
                out.append(r)
    return out


def command(stage, ctx, cfg, stub, replay=False):
    """The argv and stdin for one stage attempt."""
    if stub:
        return [cfg["python"], os.path.join(HERE, "stub_stage.py"), stage["id"]], None
    pdir = os.path.join(REPO, "orchestrator", "prompts")
    parts = ["_common.md"] + (["_replay.md"] if replay else []) + [f"{stage['id']}.md"]
    prompt = fill("\n\n".join(open(os.path.join(pdir, f), encoding="utf-8").read() for f in parts), ctx)
    exe = shutil.which(cfg["claude"]) or cfg["claude"]
    argv = [exe, "-p", "--output-format", "json", "--permission-mode", "dontAsk",
            "--allowedTools", ",".join(allowed_tools(stage["tools"], cfg["python"])), "--add-dir", ctx["RUN"]]
    if os.path.isdir(ctx["DAY"]):
        argv += ["--add-dir", ctx["DAY"]]
    if cfg.get("model"):
        argv += ["--model", cfg["model"]]
    return argv, prompt


class Runner:
    def __init__(self, ed, cfg, spec, clock, stub=False, out=print, replay=False):
        self.ed, self.cfg, self.spec, self.clock, self.stub, self.out = ed, cfg, spec, clock, stub, out
        self.replay = replay
        if replay:
            self.run = os.path.join(cfg["runs_root"], "replay", ed.isoformat())
            self.src_day = day_folder(cfg["day_root"], ed)
            self.day = os.path.join(self.run, "day")
            prepare_replay_day(self.src_day, self.day)
            self.spec = dict(spec, stages=[{k: v for k, v in x.items() if k not in ("repeat", "late")}
                                           for x in spec["stages"]])
        else:
            self.run = os.path.join(cfg["runs_root"], ed.isoformat())
            self.day = day_folder(cfg["day_root"], ed)
        for d in ("logs", "stages"):
            os.makedirs(os.path.join(self.run, d), exist_ok=True)
        mode = "practice" if replay else read_mode(cfg["mode_file"])
        self.st = _read(os.path.join(self.run, "run_state.json")) or \
            new_state(ed, mode, spec, self.run, self.day)
        self.procs = {}       # stage id -> (Popen, started monotonic, log file)
        self.last_ready = {}  # stage id -> monotonic of the last ready check
        self._resume()

    # A restart finds stages marked running whose process belongs to the old orchestrator: count
    # that attempt as failed (kill it if it is still there) and let the retry rule decide.
    def _resume(self):
        for sid, s in self.st["stages"].items():
            if s["status"] == "running":
                if s.get("pid"):
                    kill_tree(s["pid"])
                s.update(status="waiting", note="interrupted by an orchestrator restart", pid=None)

    def ctx(self, sid, rnd=0):
        return {"EDITION": self.ed.isoformat(), "EDITION_COMPACT": self.ed.strftime("%Y%m%d"),
                "RUN": self.run, "DAY": self.day, "REPO": REPO, "MODE": self.st["mode"],
                "STAGE": sid, "ROUND": rnd, "PY": self.cfg["python"],
                "RESULT": self.result_path(sid, rnd), "REPLAY": "yes" if self.replay else "no",
                "SRC_DAY": getattr(self, "src_day", self.day),
                "PREV_EDITION": previous_edition(self.ed, self.cfg["holidays"]).isoformat()}

    def result_path(self, sid, rnd=0):
        return os.path.join(self.run, "stages", f"{sid}{'.r%d' % rnd if rnd else ''}.result.json")

    def stage(self, sid):
        return next(s for s in self.spec["stages"] if s["id"] == sid)

    # ---- one loop
    def tick(self):
        now = self.clock.now()
        st = self.st
        st["started"] = st["started"] or gates.now_iso(now)
        st["heartbeat"] = gates.now_iso(now)
        gates.tick(self.run, now)
        self._mail_fallback(now)
        for spec in self.spec["stages"]:
            sid = spec["id"]
            s = st["stages"][sid]
            if s["status"] == "running":
                self._poll(spec, s, now)
            elif s["status"] in ("waiting", "next_round"):
                self._maybe_start(spec, s, now)
            if spec.get("cutoff") and s["status"] != "done" and \
                    now >= self.clock.at(spec["cutoff"], self.ed):
                alert_once(self.run, st, f"cutoff-{sid}",
                           f"{sid} {spec['name']} is not done at the {spec['cutoff']} cut-off "
                           f"(status {s['status']}). Files may need sending by hand.", now)
        if now >= self.clock.at(self.spec["hard_stop"], self.ed):
            for sid, s in st["stages"].items():
                if s["status"] not in TERMINAL:
                    if sid in self.procs:
                        kill_tree(self.procs.pop(sid)[0].pid)
                    s.update(status="failed", note=f"stopped at the {self.spec['hard_stop']} hard stop")
            alert_once(self.run, st, "hard-stop", f"Run stopped at {self.spec['hard_stop']} with stages unfinished.", now)
        save(self.run, st)
        return all(s["status"] in TERMINAL for s in st["stages"].values())

    def _needs(self, spec, now):
        """'go', 'wait' or 'blocked: <why>'."""
        for d, need in spec.get("needs", {}).items():
            ds = self.st["stages"][d]["status"]
            if need == "done" and ds in ("failed", "blocked"):
                return f"blocked: {d} {ds}"
            if (need == "done" and ds != "done") or (need == "finished" and ds not in TERMINAL):
                return "wait"
        g = spec.get("gate")
        if g:
            if not os.path.exists(os.path.join(self.run, "gates", f"{g}.json")):
                return f"blocked: no {g} gate was opened" if all(
                    self.st["stages"][d]["status"] == "done" for d in spec.get("needs", {})) else "wait"
            if not gates.resolved(self.run, g):
                return "wait"
        return "go"

    def _ready(self, spec, now):
        if not spec.get("ready") or self.stub:
            return True
        last = self.last_ready.get(spec["id"], -1e9)
        if time.monotonic() - last < 60 / max(self.clock.speed, 1):
            return False
        self.last_ready[spec["id"]] = time.monotonic()
        argv = [fill(a, self.ctx(spec["id"])) for a in spec["ready"]]
        try:
            return subprocess.run(argv, capture_output=True, timeout=120).returncode == 0
        except Exception:
            return False

    def _maybe_start(self, spec, s, now):
        sid = spec["id"]
        if now < self.clock.at(spec["start"], self.ed):
            return
        if s["status"] == "next_round" and now < datetime.fromisoformat(s["next_at"]):
            return
        n = self._needs(spec, now)
        if n.startswith("blocked"):
            s.update(status="blocked", note=n[9:], ended=gates.now_iso(now))
            alert_once(self.run, self.st, f"blocked-{sid}", f"{sid} {spec['name']} cannot run: {n[9:]}.", now)
            return
        if n == "wait":
            return
        if not self._ready(spec, now):
            late = spec.get("late")
            if late and now >= self.clock.at(late["at"], self.ed):
                alert_once(self.run, self.st, f"late-{sid}", late["text"], now)
            return
        self._launch(spec, s, now)

    def _launch(self, spec, s, now):
        sid = spec["id"]
        rnd = s["round"] if spec.get("repeat") else 0
        ctx = self.ctx(sid, rnd)
        try:
            os.remove(ctx["RESULT"])
        except FileNotFoundError:
            pass
        argv, prompt = command(spec, ctx, self.cfg, self.stub, self.replay)
        s["attempts"] += 1
        log = os.path.join(self.run, "logs", f"{sid}{'.r%d' % rnd if rnd else ''}.a{s['attempts']}.log")
        env = dict(os.environ, **{f"DMW_{k}": str(v) for k, v in ctx.items()})
        env["PYTHONIOENCODING"] = "utf-8"
        env["DMW_CLOCK_OFFSET"] = str((now - datetime.now(PKT)).total_seconds())
        f = open(log, "w", encoding="utf-8")
        kw = {"creationflags": subprocess.CREATE_NEW_PROCESS_GROUP} if os.name == "nt" else {"start_new_session": True}
        p = subprocess.Popen(argv, stdin=subprocess.PIPE if prompt else subprocess.DEVNULL,
                             stdout=f, stderr=subprocess.STDOUT, cwd=REPO, env=env, **kw)
        if prompt:
            p.stdin.write(prompt.encode("utf-8"))
            p.stdin.close()
        self.procs[sid] = (p, time.monotonic(), f)
        s.update(status="running", started=gates.now_iso(now), pid=p.pid, log=log, exit=None, note=None)
        self.out(f"{now:%H:%M:%S} start {sid} (attempt {s['attempts']}{', round %d' % rnd if rnd else ''})")

    def _poll(self, spec, s, now):
        sid = spec["id"]
        if sid not in self.procs:  # should not happen after _resume
            s.update(status="waiting")
            return
        p, t0, f = self.procs[sid]
        code = p.poll()
        if code is None:
            limit = spec.get("timeout_min", 20) * 60 / max(self.clock.speed, 1)
            if time.monotonic() - t0 < limit:
                return
            kill_tree(p.pid)
            p.wait(timeout=30)
            code = "timeout"
        f.close()
        del self.procs[sid]
        rnd = s["round"] if spec.get("repeat") else 0
        res = _read(self.result_path(sid, rnd)) or {}
        ok = code == 0 and res.get("ok") is True
        s.update(exit=code, pid=None, ended=gates.now_iso(now))
        self.out(f"{now:%H:%M:%S} {sid} exit {code} ok={ok} {res.get('summary', '')}")
        if ok:
            s["note"] = res.get("summary")
            s["attempts"] = 0 if spec.get("repeat") else s["attempts"]
            self._finish(spec, s, now)
            return
        why = res.get("error") or (f"timed out after {spec.get('timeout_min', 20)} min" if code == "timeout"
                                    else f"exit {code}, no ok result")
        if s["attempts"] < self.spec["max_attempts"]:
            s.update(status="next_round" if spec.get("repeat") else "waiting", note=f"retrying: {why}",
                     next_at=gates.now_iso(now))
            return
        if spec.get("repeat"):  # a failed round pings, and the next round still runs
            alert_once(self.run, self.st, f"failed-{sid}-r{rnd}", f"{sid} {spec['name']} round {rnd} failed twice: {why}.", now)
            s["attempts"] = 0
            self._finish(spec, s, now, failed_round=True)
            return
        s.update(status="failed", note=why)
        alert_once(self.run, self.st, f"failed-{sid}", f"{sid} {spec['name']} failed twice: {why}. Log: {s['log']}", now)

    def _finish(self, spec, s, now, failed_round=False):
        rep = spec.get("repeat")
        if not rep:
            s["status"] = "done"
            return
        nxt = now + timedelta(minutes=rep["every_min"])
        if nxt < self.clock.at(rep["until"], self.ed):
            s.update(status="next_round", round=s["round"] + 1, next_at=gates.now_iso(nxt))
        else:
            s["status"] = "done"

    def next_event(self, now):
        """The next fake time at which something can happen with no stage running: a stage's start
        or next round, a gate's reminder or deadline, or the hard stop."""
        ev = [self.clock.at(self.spec["hard_stop"], self.ed)]
        for spec in self.spec["stages"]:
            s = self.st["stages"][spec["id"]]
            if s["status"] == "waiting":
                ev.append(self.clock.at(spec["start"], self.ed))
            elif s["status"] == "next_round" and s.get("next_at"):
                ev.append(datetime.fromisoformat(s["next_at"]))
            if spec.get("cutoff"):
                ev.append(self.clock.at(spec["cutoff"], self.ed))
        for g in gates.gates(self.run).values():
            if not g["answer"]:
                for k in ("remind", "deadline"):
                    if g.get(k):
                        ev.append(self.clock.at(g[k], self.ed))
        fut = [e for e in ev if e > now]
        return min(fut) if fut else None

    def _jump(self):
        now = self.clock.now()
        nxt = self.next_event(now)
        if nxt and (nxt - now).total_seconds() > 1:
            self.clock.jump(nxt)
            self.out(f"{now:%H:%M:%S} replay clock -> {nxt:%H:%M}")

    def _mail_fallback(self, now):
        """The relay is down or slow: email what it has not posted (Saad, 25 Nov 2026). One mail
        per outbox message, sent by a short headless session with the Gmail connector. Stub runs
        write the mail to RUN/mail/ instead."""
        to = self.cfg.get("alert_email")
        if self.replay or (not to and not self.stub):
            return
        box = os.path.join(self.run, "outbox")
        for name in gates.pending_outbox(self.run):
            mark = os.path.join(box, name[:-5] + ".emailed.json")
            msg = _read(os.path.join(box, name)) or {}
            if os.path.exists(mark) or msg.get("what") == "info" or not msg.get("queued"):
                continue
            age = (now - datetime.fromisoformat(msg["queued"])).total_seconds() / 60
            if age < self.cfg["mail_after_min"]:
                continue
            subject = f"DMW {self.ed:%a %d %b}: {msg.get('what')} {msg.get('gate')}"
            body = msg.get("text", "") + "\n\nThe DMW Ops relay has not posted this. Reply in DMW Ops, " \
                "or answer in the news Doc for the pick."
            if self.stub:
                os.makedirs(os.path.join(self.run, "mail"), exist_ok=True)
                open(os.path.join(self.run, "mail", name[:-5] + ".txt"), "w", encoding="utf-8").write(
                    f"To: {to or '(alert_email not set)'}\nSubject: {subject}\n\n{body}\n")
            else:
                exe = shutil.which(self.cfg["claude"]) or self.cfg["claude"]
                prompt = (f"Send one email with the Gmail connector, then stop. To: {to}. Subject: "
                          f"{subject}. Body, exactly as given:\n\n{body}")
                log = open(os.path.join(self.run, "logs", f"mail-{name[:-5]}.log"), "w", encoding="utf-8")
                p = subprocess.Popen([exe, "-p", "--permission-mode", "dontAsk", "--allowedTools",
                                      "mcp__claude_ai_Gmail"], stdin=subprocess.PIPE, stdout=log,
                                     stderr=subprocess.STDOUT, cwd=REPO)
                p.stdin.write(prompt.encode("utf-8")); p.stdin.close()
            _write(mark, {"at": gates.now_iso(now), "to": to})

    def loop(self, max_minutes=None):
        t_end = time.monotonic() + max_minutes * 60 if max_minutes else None
        tick = self.cfg["tick_seconds"] / max(self.clock.speed, 1)
        tick = max(tick, 0.2)
        if self.replay:   # stage ends are noticed at once, so the clock can jump straight on
            tick = min(tick, 1.0)
        while True:
            if self.tick():
                now = self.clock.now()
                self.st["ended"] = gates.now_iso(now)
                summary = "; ".join(f"{k} {v['status']}" for k, v in self.st["stages"].items())
                gates.ping(self.run, "info", f"Run for {self.ed:%a %d %b} ended: {summary}.", now)
                save(self.run, self.st)
                return 0
            if self.replay and not self.procs:
                self._jump()
            if t_end and time.monotonic() > t_end:
                self.out("stopping: --max-minutes reached")
                for sid, (p, _, f) in list(self.procs.items()):
                    kill_tree(p.pid)
                return 3
            time.sleep(tick)


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--edition", help="YYYY-MM-DD, default today (PKT)")
    ap.add_argument("--config", default=os.path.join(REPO, "orchestrator", "config.local.json"))
    ap.add_argument("--stages", default=os.path.join(REPO, "orchestrator", "stages.json"))
    ap.add_argument("--stub", action="store_true")
    ap.add_argument("--clock", help="fake clock start HH:MM (tests)")
    ap.add_argument("--speed", type=float, default=1.0)
    ap.add_argument("--max-minutes", type=float)
    ap.add_argument("--replay", action="store_true", help="replay a past edition (see the module doc)")
    a = ap.parse_args(argv)
    cfg = load_config(a.config)
    ed = date.fromisoformat(a.edition) if a.edition else datetime.now(PKT).date()
    skip = edition_day(ed, cfg["holidays"])
    if skip:
        print(f"{ed}: {skip}; no run.")
        return 0
    spec = json.load(open(a.stages, encoding="utf-8"))
    if a.replay and not a.edition:
        ap.error("--replay needs --edition")
    clock = Clock(ed, a.clock or ("08:55" if a.replay else None), a.speed)
    r = Runner(ed, cfg, spec, clock, stub=a.stub, replay=a.replay)
    print(f"run {ed} mode={r.st['mode']} RUN={r.run} DAY={r.day}{' (stub stages)' if a.stub else ''}")
    return r.loop(a.max_minutes)


if __name__ == "__main__":
    sys.exit(main())
