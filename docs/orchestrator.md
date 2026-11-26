# The edition run: orchestrator and gates (build steps 3 and 4, 25 Nov 2026)

One script, `scripts/run_edition.py`, runs an edition. The runner's Windows Task Scheduler starts
it at 8:55 AM PKT, Sunday to Thursday (`orchestrator/install_task.ps1`). It skips Fridays,
Saturdays and the days in `orchestrator/holidays.json`.

## What it does

- Reads `orchestrator/stages.json` (S0 to S6, as in the plan doc) and keeps
  `RUN/run_state.json` (RUN = `<runs_root>\<YYYY-MM-DD>\`), rewritten every 15 seconds with a
  heartbeat.
- Starts each stage as a headless Claude Code session:
  `claude -p --output-format json --permission-mode dontAsk --allowedTools <that stage's list>`,
  with `orchestrator/prompts/_common.md` and `<stage>.md` on stdin. Saad approved dontAsk over
  bypassPermissions on 25 Nov 2026.
- A stage starts once its start time has passed, the stages it needs are `done` (or `finished`),
  its gate is answered and its `ready` check passes. It is done when the session exits 0 and wrote
  `RUN/stages/<id>.result.json` with `"ok": true`.
- A failed or timed-out attempt is retried once. A second failure pings DMW Ops and blocks the
  stages that need it done. S4 runs in rounds every 10 minutes until 10:10, and a failed round
  pings but does not stop the next round.
- Late-input alerts (data file at 9:10) and the 10:20 cut-off alert are in `stages.json`.
- A restart reads `run_state.json` and counts an interrupted attempt as a failure.
- `orchestrator/mode.json` holds practice or live (practice until Saad switches it).

## Gates

The gate mailbox (`scripts/gates.py`) and the DMW Ops relay are described in
`rules/dmw-ops-relay.md`. The pick gate's clock (card about 9:08, reminder 9:18, the run's own pick
at 9:25) is kept by the orchestrator. If the relay has not posted a card or alert after 5 minutes,
the run emails it to `alert_email` with the Gmail connector (Saad, 25 Nov 2026).

## Set-up on a runner

1. Copy `orchestrator/config.example.json` to `orchestrator/config.local.json` and set the paths and
   `alert_email`.
2. `python scripts/manifest.py --write` after each merge; S0 checks it.
3. `powershell -ExecutionPolicy Bypass -File orchestrator\install_task.ps1` as the DMW user.

## Running on a Mac (26 Nov 2026)

The same scripts run on Saad's Mac; the Windows runner and the office server are unchanged.

- `orchestrator/config.local.json` on the Mac: `runs_root` =
  `~/Library/CloudStorage/GoogleDrive-<personal>/My Drive/Claude Outputs/DMW Tools/runs`,
  `day_root` = `~/Library/CloudStorage/GoogleDrive-<work>/.shortcut-targets-by-id/18rkHxDXGFRnp_kbzgyFhbwCjVH4-XOpc/2026/03. Submission`
  (read only: outputs never go to the work drive in practice or replay), `python` = the
  python.org 3.15 interpreter, `claude` = `/opt/homebrew/bin/claude`.
- Fetchers use Chrome on a Mac by default (`--browser` or `DMW_BROWSER` override). Saudi Exchange
  needs the home router's Saudi route.
- PDFs and PowerPoint's layout: `scripts/export_measure.py` runs export_measure.ps1 on Windows and
  PowerPoint for Mac through AppleScript on a Mac, staging each deck in Office's group container so
  PowerPoint never asks for file access. Same `ppt_layout.csv` columns.
- Python packages the scripts need: python-pptx, openpyxl, playwright, python-docx, fonttools,
  uharfbuzz, numpy, pypdf (pytest for the tests).
- The headless stages need the `claude` CLI signed in (`claude` then `/login` in a terminal). When it
  is not, `run_edition.py --replay --edition <date> --prompts-only` sets up the run folder and writes
  each stage's exact prompt to `RUN/prompts/<id>.md`, so the stages can be run by another driver.

## Tests

- `python3 -m pytest tests/orchestrator -q` runs the gate clock, retries, blocking, restart,
  holidays and a whole stub run on a fast fake clock (about 10 seconds a run).
- By hand on a PC: `python scripts\run_edition.py --edition <date> --stub --clock 09:06` runs stub
  stages on a fake clock. S1 opens the real pick gate from the 24 Nov 2026 fixture.
  The end-to-end test on 25 Nov 2026 (DESKTOP-3S8HFHF) worked this way. The card went from the
  outbox to the project thread, Saad pressed Go with proposed, and the answer went back to
  `inbox\pick.json`. S3 then started at 9:07 on the test clock, without waiting for 9:25.

## Still open

- The stage prompts S1 to S6 are first drafts. The replay test (build step 5) runs them for real.
- The DMW Ops project needs its instructions and a scheduled task at 8:58 (rules/dmw-ops-relay.md).
- `holidays.json` is empty until Saad sends the yearly Tadawul list.
