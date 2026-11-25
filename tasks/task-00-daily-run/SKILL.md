---
name: task-00-daily-run
description: Shows where the ANB Daily Market Watch cycle stands and which numbered task (1 to 12) to run next, with its prerequisites and deadline. Also runs the monthly Saudi Exchange check of the team's names sheet (Mapping.xlsx). Use when someone says "daily run", "what's next", "start the morning run", "start the evening run", "DMW status" or asks which DMW task to do now.
---

# Task 0: Daily run (the conductor)

This skill runs no task itself. It looks at the clock and the day folder, says which task is next,
checks that task's prerequisites, and hands over to it. Shared files live in `../../shared/`
relative to this skill's base directory (call it SHARED). Read `SHARED/rules/config.md` first.

## The cycle (Karachi time)

The times are the usual working windows, not deadlines. The manager's padded cut-offs (9:15,
9:45, 9:50, 10:15, 10:20, 10:30 per task) were dropped on 4 Nov 2026 (user). The only hard
deadlines are the late-input ones below, set by the user one by one on 4 Nov 2026.

| # | Task | Usual window | Needs |
|---|---|---|---|
| 1 | Evening news (Economic + Corporate) and doc refresh | 5:30 to 8:00 PM | Argaam, Tadawul |
| 2 | TASI write-up | after 5:30 PM (3:00 PM Riyadh close + 30 min) | Saudi Exchange pages |
| 3 | Manual File-Calendar (next edition) | evening | previous calendar file |
| 4 | Write-up hand-off docx | after 1 to 3 are approved | ANB Rough; approved `Manual File-Calendar.xlsx` in the next edition's folder; previous proofread write-up |
| 5 | Global write-up | 8:30 to 9:15 | Reuters, Yahoo |
| 6 | Morning news top-up | 8:50 to 9:00 | Tadawul, Argaam |
| 7 | DMW deck build (numbers) | from about 9:00, when the data lands | `DMW - Data.xlsx` (about 9:00), Graphs decks, calendar file |
| 8 | Commentary fill + Technicals | straight after Task 7; Technicals about 9:30 | the sequence (Claude's own on the user's Pasting weeks); Technicals |
| 9 | Daily Market Update (DMU) | after Task 8 and the Technicals | Task 7 data, the final DMW EN deck (Task 8), Technicals, Bloomberg capture on the user's PC |
| 10 | Tadawul news + Results table | 9:45 to 10:30 | Tadawul; `DMW - Results.xlsx` (after 10:00) |
| 11 | Final checks + PDFs | after the last deck change | team-saved decks, calendar file |
| 12 | Sales Email (EMAIL MB) | after Task 11 | final EN deck |

## Whose tasks today (weekly schedule, user 29 Oct 2026)

The team splits the DMW and DMU jobs in a weekly schedule (Sunday to Thursday edition days).
Claude runs only the user's jobs, unless the user says they are covering someone. The schedule
lives in `WORK/DMW Tools/state/schedule.json`, one entry per week keyed by the week's Sunday
(keep the file's format as it is; user, 4 Nov 2026):

```json
{"2026-10-25": {"jobs": ["News", "Sales Email"], "reviewer": [], "covering": {}}}
```

| Schedule job | Tasks |
|---|---|
| Calendar & Writeups | 2, 3, 4, 5 (TASI and Global write-ups, calendar file, hand-off docx) |
| News | 1, 6, and the news part of 10 |
| Results | the Results part of 10 |
| DMU | 9 |
| Sales Email | 12 |
| English Pasting / Arabic Pasting | 8 (that deck's commentary) |
| Daily Market Note | not mapped; the user has never had it (to learn: how it is made, what is in it, who uses it) |
| Data Check | always Kumail; never Claude's |

- **Task 7 (deck build) is always the user's**, outside the schedule, unless they say they are
  out that day (then the team builds the decks). Task 11's checks run whenever the user built
  the decks; ask once who makes the PDFs if the schedule does not say.
- Thursday evening prepares Sunday's edition, so it follows the **new** week's schedule.
- Where the user is **reviewer** of a job, run Claude's checks on the teammate's output (no
  rewriting, no internet lookups beyond the checks) and report to the user.
- A teammate's job is not Claude's to do. The one standing exception is the morning calendar
  Arabic-name fix in Task 7 (local files only).
- If the edition's week has no entry, ask the user for that week's schedule (a photo or a list of
  their jobs and reviewer roles) in the run's first batch of questions, and save it. On Thursday
  evenings, ask for the coming week's schedule even before it is needed (it is posted Thursday
  toward the end of the day).

## Times come from the week's assigned tasks (user, 4 Nov 2026)

When saying what is next or late, show only the user's tasks for the edition (from the schedule)
plus Task 7, with their usual window and, where one applies, the late-input deadline below. Do not
list a teammate's task as late, and do not quote the old per-task cut-offs.

## Monthly names check (Mapping.xlsx; user, 29 Oct 2026)

The team's `Mapping.xlsx` (sheet `mapping`) is the one names file: names EN/AR, code, RIC,
Bloomberg, market, sectors EN/AR, ISIN. Saudi Exchange is the routine source; Bloomberg (an EQS
export, when the user has one) the occasional cross-check.

**When:** on the first daily run of each month (`WORK/DMW Tools/state/names_check.json` holds the
last check date; the first check was 29 Oct 2026), and after any listing news (a new listing, a
Nomu-to-Main move, a rename, a delisting). Run it in a quiet window (evening, or before 8:30), never
inside the 9:00 to 10:30 stretch. It costs one browser step and a small diff.

1. Stage the team's current `Mapping.xlsx` (path in `config.md`) and run
   `python SHARED/scripts/refresh_mapping.py digest Mapping.xlsx` (prints a short digest).
2. Built-in browser on any `www.saudiexchange.sa` page (Chrome with the Saudi VPN if blocked): paste
   `SHARED/browser/sx_names.js` and `await SX.diff("<digest>")` in one `javascript_tool` call. It
   returns only new listings, board moves, name or ISIN changes and companies no longer listed.
3. If anything is new or moved: open `ourmarkets/main-market-watch?locale=en` (and
   `ourmarkets/nomuc-market-watch?locale=en` for Nomu), wait for the table, paste the script again
   and call `SX.sectors(["code", ...])`.
4. Save the results as `diff.json` (and `sectors.json`) in `WORK/DMW/<date>/` and run
   `python SHARED/scripts/refresh_mapping.py apply Mapping.xlsx diff.json --sectors sectors.json --out "<WORK>/DMW/<date>/Mapping (proposed).xlsx" --report "<WORK>/DMW/<date>/Mapping check - <date>.md"`.
5. Show the user the report. Nothing replaces the team's shared file without their OK; companies no
   longer listed are only reported. On OK, save the proposed copy over the shared file (open-file
   check) and over `SHARED/data/Mapping.xlsx` at the next skill rebuild; ask the user to re-upload
   it to the project. Update `names_check.json`.
6. When the user has a fresh Bloomberg EQS export (ticker, name, ISIN), also run
   `python SHARED/scripts/refresh_mapping.py bbg Mapping.xlsx export.csv` and report ticker
   differences.

## Steps

1. Get the current time (Karachi) and the edition date. Evening runs (after 3:00 PM) prepare the
   **next** edition (Thursday evening prepares Sunday's). Read `schedule.json` and work out the
   user's tasks for that edition (table above); only those count as "next" in step 3. If the
   monthly names check is due, list it as well (run it only in a quiet window).
2. List the edition's day folder and note which inputs and outputs exist and when they were saved.
3. Say in three lines or fewer: what is done, what is next, and anything past a late-input
   deadline below (for example "Technicals not in yet at 9:30: delete slides 4-5 and build the DMU
   without them; paste them when they land").
4. Check the next task's prerequisites from the table. If one is missing, say what the fallback is
   (the table below and each task skill) rather than waiting silently.
5. Hand over: load the next task's skill and follow it.

## Late-input fallbacks (deadlines set by the user one by one, 4 Nov 2026)

| Late input | Deadline | Fallback |
|---|---|---|
| Reuters Asia story | 9:15 | Asia line from Yahoo quotes with the plain "Asian stocks rose/fell/were mixed on <day> morning" opener |
| Data team files | 9:30 | Tell the user at once; build the moment data lands, then run Task 8's fill straight after (never fill a deck that will be rebuilt) |
| Technicals | 9:30 | DMW: delete the template's old slides 4-5. DMU: build without `--technicals` (the technical slide stays the previous day's) and say so; paste or rebuild when they arrive. 4 Nov 2026: they landed at 9:30 and a 9:36 save missed them, so Task 7 re-lists the day folder right before saving |
| Sequence not final (other weeks; on the user's Pasting weeks Claude sets it) | 9:30 | Ask the user to call the sequence final; fill only on their explicit say-so |
| Results rows | 10:30 | Rows saved after 10:30 wait for the next edition (4 Nov 2026: SAICO and AYYAN at 10:38 left out). Results sentences can go by hand into the EMAIL MB on the user's say-so |
| Late Tadawul filing | 10:30 | Next edition, unless it is a top-3 item and the user says to add it |

Never write em-dashes. Ask all questions for a run in one batch at the start.
