---
name: task-08-commentary-technicals
description: Task 8 of the ANB DMW cycle. Fills the EN and AR DMW commentary boxes (Markets, Economic, Corporate) from the ANB Rough doc under the house spacing rule, and pastes the Technicals into slides 4 and 5. Runs straight after Task 7; Technicals land about 9:30. On the user's Pasting weeks Claude sets the sequence itself and owns the decks' correctness. The Arabic deck is always filled by Claude; a teammate may paste the English. Use for "sequence is final", "fill commentary", "English pasting", "Arabic pasting", "paste the technicals", "Task 8".
---

# Task 8: Commentary fill and Technicals (straight after Task 7)

Target: both decks filled within minutes of Task 7 on the user's Pasting weeks, or about 2 minutes
after the user says the sequence is final on other weeks. Late inputs (sequence, Technicals): see
the deadlines in Task 0 (9:30 each, user 4 Nov 2026).

SHARED = `../../shared/` relative to this skill's base directory. Read
`SHARED/rules/config.md` first (day folder, `WORK` folder, fonts), then the spacing section of
`SHARED/rules/deck-house-rules.md` and `SHARED/rules/deck-safety.md`.

## Who sets the sequence (user, 3 Nov 2026)

- **On the user's English and Arabic Pasting weeks, Claude sets the sequence.** Do not wait for
  anyone to call it final, and do not ask. Read the live ANB Rough doc, pick and order the items
  yourself under `SHARED/rules/news-selection.md` (exactly two Economic, official data first;
  Corporate in that file's order after the flows item), fill both decks, and commit them. Then, as
  the team adds news or reorders the doc up to 10:10 / 10:30, patch the saved decks with only those
  changes. Never rebuild a deck the team has edited. Say in one line which items went in, in what
  order, and what was left out. The user can still swap an item ("replace #1 eco news with this",
  5 Nov 2026): their pick wins over the official-data-first order.
- To fill in Claude's order, write a reordered copy of the doc JSON (blocks moved, never the live
  doc) and pass it to `--doc`; pick Economic with `--economic`. The doc read is large (about 200k
  characters): the connector saves it to a file when too large; work on that file (its JSON sits
  under a top-level `content` key) and never pull it into the chat. Print only an outline
  (paragraph index, style, first 100 characters).
- **Misfiled items:** a company action filed under Economic (5 Nov 2026: the AYYAN capital-cut
  withdrawal) moves to Corporate in the reordered copy; say so in the override line.
- Em-dashes in a doc bullet: replace them with commas in the reordered copy (deck only), and say so.
- A write-up missing from the Draft tab does not hold up the fill: fill what is there, report the
  FLAG, and add the write-up when it lands. When the user pastes a write-up or bullet into the chat
  (a proofreader's version, a new item), inject it into the reordered copy and refill. The
  proofreader's Arabic replaces the team's whole; keep the source tag (add it if the proofread text
  dropped it: 5 Nov 2026 Global, (رويترز) added) and run `check_ar_text.py`.
- **Keep base copies for refills (5 Nov 2026):** keep the decks as they were before the first fill
  (Task 7 output plus any calendar-row removals and team edits that Claude reproduced, such as a
  slide 3 short name) and refill from them. On a results day refill from those and re-place the
  Results tables (Task 10), because `--force` stacks the slide 2 table with the calendars.
- On other weeks (a teammate pastes or owns the sequence), the "sequence final" rule below applies.

## Our decks come first (user, 3 Nov 2026)

On the user's Pasting weeks the two DMW decks are the user's and Claude's work. When a teammate's
input that feeds the decks is wrong (the calendar file, the Results workbook, a write-up) and the
owner has not fixed it:

- Fix it **in the decks** from the primary source, log the source, and tell the user which team file
  still carries the error and who owns it.
- Never edit the team's own files (calendar, Results workbook, doc).
- Keep checking the team's files on every pass: when the owner fixes it, the decks should match.

## When in doubt, Claude's checked sources and house rules win (user, 4 Nov 2026)

When the team's text and what Claude has checked disagree, use the version Claude can stand behind:
the primary source, the names sheet, and the house rules. Do not ask first. Say in **one line** per
run what was overridden. Examples: 4 Nov 2026 "Aramco 'Global' -> 'global'; Retal AR tag ->
(أرقام); EAT -> Etihad GO Telecom Co."; 5 Nov 2026 "TASI 'rose 1.7%' -> 1.6% (162.8 points on
9,876.09 = 1.65%; data file 1.648%); AYYAN moved Economic -> Corporate; EN Global '(Reuters)' tag
added". Always recompute the TASI % from points and close before filling. Keep the team's text only
where Claude cannot check it. A figure Claude cannot source stays as the team wrote it and is listed
as unverified (5 Nov: the Global index moves; the user said not to chase them).

## First questions (ask in the run's opening batch, every day)

**The Arabic deck is always filled by Claude.** Never offer an Arabic paste route. Only the English
deck has a choice:

1. **"English news (Economic and Corporate): should I fill it, or will a teammate paste it from
   the English - Paste tab?"**
2. Only when a teammate pastes: **"Global and TASI (the Markets box) in the English deck: should
   I fill them, or will the teammate paste them from the Draft tab?"**
3. **"Is today a results day?"** (decides `--results`). A `DMW - Results.xlsx` saved this morning
   or a team results message answers it without asking.

On the user's Pasting weeks the answer to 1 and 2 is "Claude fills everything"; ask only 3.

Then:

- **Claude fills everything:** run the steps below as written.
- **Teammate pastes the English news, Claude fills the English Markets box:** `--en-skip ec,co`.
- **Teammate pastes all the English:** `--en-skip mk,ec,co`.
- On both teammate routes, run the fill **before** the teammate starts pasting and pass them the
  slide 1 / slide 2 break. Never fill a box the teammate is pasting into.
- If the user later asks Claude to fill the English after all, the script leaves a box the team
  already filled alone unless `--force`; ask before using it.
- Never read the paste tab as a source. The fill reads only the Draft tab.

## Prerequisites

- [ ] **Not one of the user's Pasting weeks: the user has said the sequence is final.**
- [ ] Today's decks from Task 7 in the day folder
- [ ] Fonts staged and the Arial Nova Cond measuring copy installed (`SHARED/rules/config.md`)
- [ ] At least one write-up (Global or TASI, EN and AR) is in the Draft tab
- [ ] The script runs where LibreOffice (`soffice`), `pdftoppm` and `fc-match` exist (the cloud
      workspace)

## Steps (one pass, no exploring)

1. **In parallel:** read the ANB Rough doc with the Docs connector's full read and keep it on disk;
   stage both decks (note each mtime).
2. Fonts (once per session), then:
   `python fill_commentary.py --doc <doc json> --en <EN deck> --ar <AR deck> --outdir <out> --render [--economic 1,2] [--en-skip ec,co] [--results]`
   `--results` on a results day: every Corporate item goes to slide 2. If the doc holds more or
   fewer than two Economic items, the script stops: on the user's Pasting weeks pick the two
   yourself; on other weeks ask. EN/AR "mismatch" FLAGs where the AR writes a number in words
   ("Tier 2", "1Q 2027") are expected; others need a look. A "bullet without a closing source tag"
   FLAG: add the tag in the reordered copy (append to the paragraph's last text run) and refill.
3. Look once at the renders: headers clear, calendars directly under the news, nothing past
   10.60 in (measure the real bottom, not the page number; crude rows are never cut), no
   placeholder left, bullets on every item. Any spacing FLAG goes to the user.
4. Open-file check, then commit both decks with `expectedMtimeMs`, then the content check
   (`deck-safety.md`). Append the report to the day's check reports in `WORK`.
5. **Technicals** (about 9:30): if Task 7 built without them, patch them in without rebuilding:
   `python dmw_patch.py tech --deck <deck> --technicals <Technicals deck> --lang en|ar --data "DMW - Data.xlsx" --date YYYY-MM-DD --out <deck>`
   It swaps slides 4-5, cross-checks date, close and value traded, and (AR) adds the slide 5 navy
   header rule when the Technicals AR deck lacks it (5 Nov 2026). If nobody has saved the decks yet,
   a full Task 7 rebuild with `--technicals` is cleaner (5 Nov: rebuilt at 9:46 with the calendar
   and Technicals together). Run `check_final.py --technicals-only` on the pair.
6. **Calendar file changes after the build:** compare the file with the decks and patch the slide 2
   calendar tables in both decks (remove dropped rows and shrink the table frame by the row height
   before refilling, so the calendars re-flow). On a results day, `place_results.py --drop-cal` can
   do the removal in the same pass.
7. Tell the user in a few lines: items per box and slide, anything left out, any FLAG, and the
   one-line list of team text overridden.

## Rules that still apply

- Ignore the yellow highlight when filling. If the user says which items go in, follow them.
- Every bullet keeps its source tag in both languages; add a missing one.
- AR company names must match the state `Mapping.xlsx`; AR sector names match slide 3. When the
  user changes EN text, translate the change into the AR bullet and say which sentence.
- If the team has pasted text into a box, keep it and make only the change asked, except where
  "Claude's checked sources and house rules win" applies.
- A news bullet asked for from a Tadawul link: open it, draft per the spec, show it, ask where it
  goes before inserting; Arabic counterparty names from the `locale=ar` filing.

## Hand-off check

- [ ] Routes settled (Pasting weeks: Claude fills all); Arabic deck filled by Claude
- [ ] Filled from Claude's own sequence (Pasting weeks) or after "sequence final"; overflow listed; every FLAG reported
- [ ] Team text overridden from checked sources or house rules, listed in one line
- [ ] Renders looked at; decks committed with the mtime guard and checked by content
- [ ] Technicals in slides 4-5 of both decks (AR slide 5 rule present), or pending, said clearly
- [ ] Pasting weeks: doc watched for additions and reorders until 10:10 (10:30 on results days), changes patched in
- [ ] Pasting weeks: every known error in a team input fixed in the decks from a primary source, logged, and the owner named
