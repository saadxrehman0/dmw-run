# DMW Ops relay (gates and pings, build step 4)

Decided with Saad on 25 Nov 2026. The run on the runner (run_edition.py) cannot post into a
Claude project, so gates travel as files in the run folder and a thread in the **DMW Ops** project
(a separate Claude project, created by Saad 25 Nov 2026) relays them. The relay never decides
anything: it posts what the run queued and writes back what Saad pressed or typed.

## The mailbox (scripts/gates.py)

RUN = `<runs root>\<YYYY-MM-DD>\` on the runner (practice on Saad's PC:
`H:\My Drive\Claude Outputs\DMW Tools\runs\<date>\`).

| File | Written by | Meaning |
|---|---|---|
| `outbox\NNN-<gate>-<what>.json` | run | a message to post: `text`, optional `card` (blocks), `gate`, `what` (card, reminder, closed, alert, done, info), `reply_to` |
| `outbox\NNN-<gate>-<what>.sent.json` | relay | posted: `{"message_id", "at"}` |
| `outbox\NNN-<gate>-<what>.emailed.json` | run | the relay had not posted it after 5 min, so the run emailed it |
| `gates\<gate>.json` | run | an open gate: kind (pick, conflict, late_news), remind and deadline times, default |
| `inbox\<gate>.json` | relay (Saad's answer) or run (default at the deadline) | `{"id", "choice", "by", "at", "text"?}`; the first answer stands |
| `run_state.json` | run, every 15 s | stages and `heartbeat` |

## What the relay does

1. **Post.** List `RUN\outbox\` (device_list_dir). For each `NNN-*.json` with no `.sent.json`
   beside it, oldest first: stage it and post it in the DMW Ops thread, as a card (`post_widget`,
   widget `card`, input = its `card`, text = its `text`) when it has one, else as a message.
   A `reminder` or `closed` message goes in the thread of its gate's card. Then write the
   `.sent.json` beside it (save under /mnt/project-files/, then device_commit_files).
2. **Answer.** When Saad presses a card button, write `RUN\inbox\<gate>.json` with
   `{"id": "<gate>", "choice": "<the button's action_id>", "by": "Saad", "at": "<now, ISO>"}`.
   A reply in words becomes `"choice": "words", "text": "<his words>"`. Never write an answer
   Saad did not give. If `inbox\<gate>.json` is already there (the run's default at 9:25 came
   first), tell him the run already went ahead and that the pick can still change in the Doc until
   10:00. Update the card so it shows what happened.
3. **Watch.** If `run_state.json`'s heartbeat is more than 10 minutes old on an edition day
   between 8:58 and 10:30, or there is no run folder by 9:00, post an alert: the run is down, run
   it by hand (`python scripts\run_edition.py` on the runner).
4. **Pace.** Check the outbox at 8:58, then every 3 minutes until 10:30 (send_later), and at once
   when Saad writes. Stop after the run's `info` message (run ended) has been posted.

## Set-up in DMW Ops (once)

- Link the runner (the PC for practice) to the project, with the run folder's parent approved.
- A scheduled task in DMW Ops, Sunday to Thursday at 8:58 AM PKT, starts the day's relay thread
  with this prompt: "Relay today's DMW run: follow rules/dmw-ops-relay.md in
  github.com/saadxrehman0/dmw-run, run folder <runs root>\<today's date>."
- The Windows scheduled task on the runner starts the run itself at 8:55
  (orchestrator/install_task.ps1). The Claude scheduled task only wakes the relay: it cannot
  promise a desktop session for PowerPoint.
