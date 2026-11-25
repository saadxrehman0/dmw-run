---
name: anb-dmw-daily
description: The ANB Capital Daily Market Watch (DMW) cycle in one skill, Tasks 0 to 12. Evening news, TASI write-up, Manual File-Calendar, write-up hand-off, Global write-up, morning news, DMW deck build (EN and AR), commentary and Technicals, DMU, Tadawul news and Results, final checks and PDFs, Sales Email (EMAIL MB). Use for "daily run", "what's next", "DMW", any task by number, or any of these outputs by name.
---

# ANB DMW Daily

The whole DMW cycle: a conductor (Task 0) and 12 numbered tasks. This file is the live entry.
The installed skill is a thin loader (`LOADER.md` here) that reads this file, `config.md` and
the task file from the project at the start of every run (29 Oct 2026).

## Where everything is

| What | Where | How |
|---|---|---|
| This entry, task files, DMU references | project `claude/anb-dmw-daily/` (`SKILL.md`, `skills/<task folder>/SKILL.md`, `skills/task-09-dmu/references/dmu-*.md`) | `project_read` |
| Rule files (`SHARED/rules/*.md`) | project `claude/anb-dmw-daily/shared/rules/` | `project_read` |
| JSON seed files (`SHARED/data/*.json`) | project `claude/anb-dmw-daily/shared/data/` | `project_read`, first run only; the live copies are in `WORK/DMW Tools/state/` (`config.md`) |
| Scripts (`SHARED/scripts/*.py`), `font_widths.json`, `template_spec.json`, `Mapping.xlsx`, `TASI Mapping.xlsx`, the fixed deck templates (`SHARED/templates/`), browser scripts (`SHARED/browser/`) | the installed skill, copied to `~/dmw/` by the loader | run from `~/dmw/scripts/` |
| Script hashes | project `shared/scripts/MANIFEST.json` (master) vs `~/dmw/BUILD.json` (installed) | the loader compares them |
| Change history, team setup | project `claude/anb-dmw-daily/README.md` | only when maintaining |
| Apps Script sidebar | project `claude/anb-dmw-daily/shared/sidebar/` | only when maintaining |

The task files were written for the plugin layout. Where one says `SHARED` or `../../shared/`,
read rules and JSON seeds from the project path above and run scripts from `~/dmw/`. Where one
says "load the next task's skill", `project_read` the next task file. "The plugin" or "the task
skill" means this skill and that task file.

Helper agents cannot read project docs (only the main chat has the `Projects` tool; found 29 Oct
2026). When a step is handed to a helper, put the text it needs in its prompt or on disk.

## Tasks (Karachi time)

Windows are the usual working times, not deadlines. The manager's padded per-task cut-offs were
dropped on 4 Nov 2026 (user); the hard deadlines are the late-input ones in Task 0 (Asia story
9:15; data, Technicals and sequence 9:30; Results rows and late Tadawul filings 10:30). Task 0
shows only the user's tasks for the week.

| # | Task folder | Usual window | Output |
|---|---|---|---|
| 0 | `task-00-daily-run` | any time | What is done, next, late; monthly names check |
| 1 | `task-01-evening-news` | 5:30 to 8:00 PM | Economic and Corporate news (EN + AR) in ANB Rough |
| 2 | `task-02-tasi-writeup` | after 5:30 PM | TASI bullet (EN + AR) in ANB Rough |
| 3 | `task-03-calendar-file` | evening | Next edition's `Manual File-Calendar.xlsx` |
| 4 | `task-04-writeup-handoff` | evening, after 1 to 3 | `DMW Writeup - YYYYMMDD.docx` |
| 5 | `task-05-global-writeup` | 8:30 to 9:15 | Global bullet (EN + AR) in ANB Rough |
| 6 | `task-06-morning-news` | 8:50 to 9:00 | Morning items in ANB Rough |
| 7 | `task-07-deck-build` | from about 9:00, when the data lands | DMW EN and AR decks, slides 1 to 3, from the fixed templates |
| 8 | `task-08-commentary-technicals` | straight after Task 7; Technicals about 9:30 | Commentary filled; slides 4-5 |
| 9 | `task-09-dmu` | after Task 8 and the Technicals | `Daily Market Update - YYYYMMDD.pptx` (rules: `references/dmu-*.md`) |
| 10 | `task-10-tadawul-news-results` | 9:45 to 10:30 | Late Tadawul news; Results table (`place_results.py`) |
| 11 | `task-11-final-checks-pdfs` | after the last deck change | Pre-PDF check, layout drift; PDFs |
| 12 | `task-12-sales-email` | after Task 11 | `EMAIL MB YYYYMMDD.docx` |

Where a task file still quotes an old cut-off time (Tasks 5, 9, 11, 12), the table above and
Task 0's deadlines win.

## What is in `shared/`

- `rules/`: config, core style, news selection, Arabic rules, Global and TASI, ANB Rough doc,
  calendar, deck house rules (fixed templates and layout spec), open-file safety, Sales Email,
  Argaam (`argaam.md`).
- `scripts/`: the build and check scripts the tasks call, plus `check_mapping.py` (codes and
  RICs in the two workbooks), `fix_cal_ar_names.py` (calendar Arabic company names),
  `snap_spec.py` (layout spec: puts drifted fixed elements back), `refresh_mapping.py`
  (monthly Saudi Exchange check of `Mapping.xlsx`, Task 0) and `place_results.py` (Results tables
  on slides 1 and 2 of both decks, Task 10; added 4 Nov 2026).
- `data/`: `Mapping.xlsx` (the team's names sheet, sheet `mapping`: names EN/AR, code, RIC,
  Bloomberg, market, sectors EN/AR, ISIN; the one names file since 29 Oct 2026, bundled from the
  project file `Mapping.xlsx`; the builds read the state copy, `--names`), `TASI Mapping.xlsx`
  (Claude's Saudi Exchange list from the project file `TASI Mapping_2.xlsx`: sector index codes and
  names, heatmap labels, status notes; kept for what the team file lacks), `template_spec.json`
  (layout spec), name overrides (EN, AR, DMU), Arabic glossary, font widths, chart-state seed.
- `templates/`: `DMW Template - EN.pptx` and `- AR.pptx`, the fixed Task 7 templates (kept in
  `WORK/DMW Tools/templates/`, bundled into the zip; binary files cannot live in the project).
- `browser/`: in-page scripts for the built-in browser (Argaam feeds and list pages; `sx_names.js`
  for the Saudi Exchange names check).
- `sidebar/`: the Apps Script sidebar for the ANB Rough doc, with offline tests (project only).

## Standing rules for every task

- Ask every question a run needs at the start, in one batch, then finish without stopping.
- Keep a dated source log for each run.
- Never write em-dashes.
- Arabic: follow `shared/rules/arabic-rules.md` and run `check_ar_text.py` before posting.
- On doubt between the team's text and what Claude has checked, Claude's checked sources and the
  house rules win; say in one line what was overridden (Task 8; user, 4 Nov 2026).
- When the user corrects a run, fix the output, then write the fix into the one task file, rule
  file or script it belongs to in the project master copy (`claude/anb-dmw-daily/`), and say in
  one line what changed. Task and rule fixes are live at once. For a script fix, see below.
- **Run scripts only from `~/dmw/scripts/`** (the installed skill). Never run a `.py` found in
  `WORK` (Claude Outputs): day folders hold one-off scripts and older copies, and
  `WORK/_Archive/` holds retired ones (clean-up, 10 Nov 2026). The only run copies kept on Drive
  are `WORK/DMW Tools/export_pdf.ps1` and `WORK/DMW Tools/make_skill.py`, both the same as the
  master. A script written for one run stays in that run's folder and is never reused from there;
  if it is worth keeping, it goes into the master as a script fix.

## Changing a script (maintainer steps)

1. Edit the script on disk (start from `~/dmw/scripts/`, which matches the master when the
   loader found no drift).
2. Upload it with `project_write` and `local_path` (never retype a script into `content`).
3. `python3 make_skill.py --manifest ~/dmw` writes the new `scripts/MANIFEST.json`; upload it the
   same way to `claude/anb-dmw-daily/shared/scripts/MANIFEST.json`.
4. Rebuild the zip (`make_skill.py`, see its header: it now also takes `Mapping.xlsx` and the
   templates folder), send it to the user to re-upload, and add a line under "Fixes not yet sent"
   in `claude/Open items - DMW plugin.md`. The run copy is `WORK/DMW Tools/make_skill.py`
   (sha256 `b9479bb1...` on 10 Nov 2026, same as the project file); if the project's
   `make_skill.py` changes, replace that copy in the same step. Build straight to
   `WORK/DMW Tools/anb-dmw-daily.zip` (one zip, overwritten), not to dated copies; any other zip
   there goes to `WORK/_Archive/<date>/` once the new one is installed.

## Keeping token use low (user, 29 Oct 2026)

The user wants the cycle as token-efficient as possible. Every step re-reads the whole chat, so:

- **One chat per block.** When the user's last task of a block is done (the evening block ends
  with their last evening task, the morning block with their last morning task; Task 0 says which
  tasks are theirs), end the reply with one line: "Block done. Start a new chat for the next
  block and paste:" followed by a one-line hand-off (edition date, tasks done, where the logs
  are, next task). Claude cannot open a chat itself. Also suggest a new chat mid-block when a
  chat has run long (about 40 tool steps), with the same one-line hand-off.
- **Batch steps.** Make independent reads and checks in one step (several files in one call,
  several checks in one script), never one per step.
- **Keep bulk data on disk.** Never pull a whole web page or a whole document into the chat when
  a script can return only the fields the task needs. Save large tool results to files and read
  only the part needed. Browser scripts return only differences (the names check sends a short
  digest in and gets only changes back; 29 Oct 2026: a full directory pull cost about 10k tokens).
