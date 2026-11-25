---
name: task-01-evening-news
description: Task 1 of the ANB DMW cycle. Sweeps Tadawul and Argaam, triages and drafts the ranked Economic and Corporate news (English and Arabic) for the next edition, refreshes the ANB Rough Draft tab and posts after approval. Use for "evening news", "evening run", "Task 1", "draft the economic and corporate news".
---

# Task 1: Evening news and doc refresh

Window: from about 5:30 PM Karachi (the user says when it starts), nominally done by 8:00 PM.
Output: ranked Economic and Corporate items, EN + link + AR each, in the ANB Rough Draft tab, for
the **next** edition. The edition covers the previous Saudi trading session (Sun to Thu week).

**Sequencing (user, 11 Nov 2026):** on the user's news weeks Claude is the sequencer. The items are
posted to the Draft tab in final sequence order (Economic and Corporate, house ranking rules in
`news-selection.md`), and Task 6 updates that sequence in the morning as new items come in.

SHARED = `../../shared/` relative to this skill's base directory.

## Prerequisites

- [ ] `SHARED/rules/config.md` read; the user's `WORK` folder known
- [ ] Google Docs and Drive connectors on (ask the user to turn them on if not)
- [ ] Browser available for `saudiexchange.sa` (it blocks plain fetching)
- [ ] Previous post log in `WORK/DMW/<previous date>/` (if none, first-run rules apply)

Read before drafting: `SHARED/rules/news-selection.md`, `SHARED/rules/core-style.md`,
`SHARED/rules/arabic-rules.md`, `SHARED/rules/anb-rough-doc.md`.

## Steps

1. **Sweep headlines.** Tadawul issuer announcements first (from the timestamp the user names, or
   since the last run), then Argaam page 1. Skip lock and bar-graph rows. Collect headlines only.
2. **Triage before fetching anything.** Run `SHARED/scripts/triage.py` on the headline list, then
   apply the checklist in `news-selection.md`, and the overlap check against the previous post log,
   the latest published DMW decks (day folder, read slide text with python-pptx) and the live ANB
   Rough Draft tab (Drive `read_file_content`). The Draft tab catches items the team added that are
   in no post log and were cut from the deck: treat those as already used and leave them out unless
   the user wants them back (Walaa Moody's rating, 27 Oct 2026). Show the user
   the triage (include / maybe / exclude, one-line reason each, plus items already in the deck) and
   ask every question in that one message.
3. **Fetch the survivors only** (Argaam links via WebFetch; Tadawul details in-page, `locale=en`
   and `locale=ar`). Never substitute a source or a remembered figure; an unsourced number is `TK`.
4. **Draft and sequence.** Economic: at least three where the day allows, never fewer than two.
   Corporate: flows / ownership item first (Mon/Tue/Wed), then the house order. Arabic drafted as
   final text. The order drafted is the sequence (Claude sequences on the user's news weeks).
5. **Second pass**, one item at a time: plain short sentences · section placement · ranking ·
   mn/bn/tn · `XQYY` quarters · currency codes · one decimal (two on ownership %) · tickers on first
   mention in every bullet · EN and AR match sentence for sentence · every bullet has its tag and URL
   line. Run `SHARED/scripts/check_ar_text.py` on the Arabic: it must print OK.
6. **Review with the user** in a Claude Doc they can comment on (sections Economic, Corporate, review
   notes; each item EN, link, AR). Apply their inline edits to EN and AR together; answer each
   comment in its thread. Name the Arabic names you wrote yourself.
7. **Refresh the doc** (after approval): back up, compare with post logs, ask about anything not
   Claude's, clear old Economic and Corporate items (never the TASI or Global write-ups), post in
   sequence order.
   Before posting, re-read the whole review doc so the user's inline edits reach the Arabic, and
   post the final text from that doc, not the first draft. Justify every posted paragraph (EN,
   link line and AR: `updateParagraphStyle` alignment `JUSTIFIED`; user, 27 Oct 2026) and check
   it on the read-back.
8. **Write the post log** to `WORK/DMW/<date>/ANB Rough post log - <date>.md`, items listed in
   sequence order.

## Hand-off check

- [ ] Triage shown before any article was fetched
- [ ] Flows / ownership item is Corporate #1, never Economic
- [ ] Each bullet 1 to 3 plain sentences (a 4th only when needed, never 5), one article per bullet, past tense
- [ ] Every figure sourced or `TK`; one decimal; currency codes; mn/bn/tn; `XQYY`
- [ ] Tickers from the mapping files, never inferred; no unlisted company given a listed name
- [ ] EN and AR match; `check_ar_text.py` OK; self-written Arabic names listed
- [ ] Every bullet has a source tag and a URL line
- [ ] Overlap checked against the previous post log, published decks and the live Draft tab
- [ ] Doc backed up, others' content checked, user approved, posted to Draft in sequence order, justified
- [ ] Post log written

Next: Task 2 (TASI write-up) and Task 3 (calendar), same evening.
