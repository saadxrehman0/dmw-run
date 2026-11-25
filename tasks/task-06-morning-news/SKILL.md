---
name: task-06-morning-news
description: Task 6 of the ANB DMW cycle. Adds overnight and early-morning Tadawul and Argaam news to the ANB Rough Draft tab, highlighted yellow in English and Arabic, and folds proofread Arabic names into the glossary. Use for "morning news", "morning top-up", "new Tadawul news this morning", "Task 6".
---

# Task 6: Morning news top-up (8:50 to 9:00)

SHARED = `../../shared/` relative to this skill's base directory. Read
`SHARED/rules/config.md` first (day folder, `WORK` folder, approvals). Same writing rules as Task 1:
`SHARED/rules/news-selection.md`, `core-style.md`, `arabic-rules.md`, `anb-rough-doc.md`.

## Who sequences (user, 11 Nov 2026)

**On the user's news weeks (the weeks the user does the news, per `schedule.json`), Claude is the
sequencer**, evening and morning. Task 1 posts the evening items already in sequence order; in the
morning Claude ranks each new item into that sequence and updates the order of the Draft tab as it
adds them (re-ranking existing items only when a new item changes their order). Do not ask whether
the team is sequencing on those weeks. On other weeks (a teammate does the news and Claude is only
covering an item), ask once whether the team is already sequencing in the Draft tab and, if so,
tell the user in one line which items are new and where they rank.

## Prerequisites

- [ ] Last night's post log (to know what is already in the doc, in sequence order)
- [ ] Last night's source log (`Source log - <date>.md`), for the user's triage decisions
- [ ] The timestamp to sweep from (the user names it, or the last Task 1 sweep)
- [ ] Whose news week it is (`schedule.json`): the user's week means Claude sequences (above)

## Steps

1. Sweep Tadawul issuer announcements from the timestamp, then Argaam page 1.
2. Triage (`SHARED/scripts/triage.py` plus the checklist), overlap check against the doc, last
   night's post log and the latest published decks. Show the user the short list.
   Also read the "User decisions" section of last night's source log: never re-raise an item the
   user excluded there, and never call it a miss (the post log lists only what was posted; on
   29 Oct 2026 the Watani Steel and NBM tariff filings the user had cut were wrongly flagged).
3. Draft EN + link + AR for the survivors; second pass; `check_ar_text.py` OK. Give each item its
   proposed rank in the sequence (section and position) in the review.
4. After approval, insert each item at its ranked position in the Draft tab, **highlighted yellow
   in EN and AR**, so the Draft tab reads in final sequence order. On the user's news weeks this
   is the sequence (Claude is the sequencer); say in one line which items are new and where they
   rank.
5. **Glossary:** for each Arabic name Claude wrote itself last night (listed in the post log), take
   the proofreader's final form and add it to `WORK/DMW Tools/state/ar_glossary.json`. Do not adopt
   proofreader edits that break house rules (list them to the user).
6. Write a morning post log (`ANB Rough post log - <date> (morning).md`), with the final sequence.

## Hand-off check

- [ ] Only new items; overlap checked (post log, source log decisions, decks); triage shown
- [ ] New items ranked into the sequence (Claude sequences on the user's news weeks)
- [ ] Yellow highlight on every morning item, EN and AR; new items named to the user
- [ ] `check_ar_text.py` OK; glossary updated with proofread names
- [ ] Post log written, with the final sequence

Next: Task 7 at about 9:00 when the data file lands.
