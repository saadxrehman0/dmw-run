---
name: task-05-global-writeup
description: Task 5 of the ANB DMW cycle. Writes the morning Global bullet (US close plus that morning's Asia session) in English then Arabic from Reuters and Yahoo quotes, and posts it to the ANB Rough Draft tab after approval. Usual window 8:30 to 9:15 Karachi; the Reuters Asia story has a 9:15 late-input deadline. Use for "global write-up", "global bullet", "Asia update", "Task 5".
---

# Task 5: Global write-up (morning, usual window 8:30 to 9:15)

> Times are usual windows, not deadlines (per-task cut-offs dropped 4 Nov 2026; late-input deadlines in Task 0). Reuters Asia story not in by 9:15: Asia line from Yahoo quotes (Task 0 fallback).

Target: drafted by 8:40, approved and posted by 8:50 (25 minutes of slack).

SHARED = `../../shared/` relative to this skill's base directory. Read
`SHARED/rules/config.md` first (day folder, `WORK` folder, approvals), then
`SHARED/rules/global-tasi.md` (Global section) and the Global part of `SHARED/rules/arabic-rules.md`.

## Prerequisites

- [ ] Edition date and weekday known (Sunday: US Friday close + Asia Friday close; Monday: Asia
      only; Tue to Thu: US close + Asia morning)
- [ ] Reuters reachable through the browser, or the user pastes the articles
- [ ] Market holidays checked for the indices in the line (e.g. Japan's Sports Day, mainland
      Golden Week): a closed market drops out of the line

## Steps

1. Read the Reuters US close article (if the weekday needs it) and that morning's Asia-session
   story. Lift each lede, keeping all of its clauses (including Saudi or Middle East context);
   add nothing, and trim only what `global-tasi.md` allows.
2. Get index levels. Compute every move with `SHARED/scripts/global_calc.py` from the underlying
   figures (prior close from the daily series, never the 1-day chart API), half-up to one decimal.
   Fill any index the article omits from Yahoo, and refresh the others from the same source.
3. Write the English: "U.S. stocks ..." then "Asian stocks ... on <day> morning ..." (on Sunday:
   "Asian stocks ... on Friday", no "morning", since Friday's session is closed) with the plain
   verb, "mixed" when indices split, largest move first, ties with "each". Tag `(Reuters)` or
   `(Reuters, Bloomberg)`.
4. Write the Arabic (house index names, بنسبة, صباح + day; on Sunday يوم الجمعة). Run `SHARED/scripts/check_ar_text.py`.
5. Refresh the quotes once more just before showing the user; update any figure that moved.
6. Log sources and quote timestamps in `WORK/DMW/<date>/Source log - <date>.md`.
7. Post to the Draft tab under the Global heading after approval. **Never highlight it.**

## If the Reuters Asia story is not up by 8:55

Write the Asia sentence from Yahoo quotes with the plain opener ("Asian stocks rose / fell / were
mixed on <day> morning") and no narrative clause, and tell the user.

## Hand-off check

- [ ] Nothing added beyond Reuters' own wording; "Asian stocks" opener with "morning"
- [ ] No closed market reported as a move; no index silently dropped
- [ ] Every % computed in code from underlying figures, one decimal, half-up; quotes refreshed
- [ ] Arabic follows the house names and constructions; `check_ar_text.py` OK
- [ ] Source log written; approved; posted without highlight

Next: Task 6 (morning news top-up).
