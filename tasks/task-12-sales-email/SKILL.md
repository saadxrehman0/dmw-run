---
name: task-12-sales-email
description: Task 12 of the ANB DMW cycle. Builds the Sales Email document "EMAIL MB YYYYMMDD.docx" from the final EN DMW deck (markets, economic and corporate news with results sentences first, corporate calendar without EGMs) after Task 11 (usually 10:17 to 10:30 Karachi). Use for "sales email", "EMAIL MB", "MB email", "Task 12".
---

> **Template v2 (25 Nov 2026):** the commands below are the v1 ones. For v2 decks use the run order and flags in `docs/v2-pipeline.md`; the v1 scripts remain as `scripts/legacy_*` for v1 decks only.

# Task 12: Sales Email / EMAIL MB (after Task 11, usually 10:17 to 10:30)

> Times are usual windows, not deadlines (per-task cut-offs dropped 4 Nov 2026; late-input deadlines in Task 0).

SHARED = `../../shared/` relative to this skill's base directory. Read
`SHARED/rules/config.md` first (day folder, `WORK` folder, approvals), then `SHARED/rules/email-mb.md`.

## Prerequisites

- [ ] The **final** EN deck (its PDF saved next to it is the sign it is final)
- [ ] The previous `EMAIL MB <date>.docx` as template
- [ ] `SHARED/data/Mapping.xlsx` (the team's names sheet) for full names and tickers
- [ ] The day's `Daily Market Update - YYYYMMDD.pptx` (coverage paragraphs and the results cross-check)
- [ ] On results days, the AR deck and `DMW - Results.xlsx` (deck results check, step 2)

To save time, a draft can be built at about 9:45 from the pasted deck; at 10:17 rebuild from the
final deck and check only the differences. If the decks are saved again after the build, rebuild
and diff before saving.

## Steps

1. Stage the final EN and AR decks, the previous EMAIL MB, the DMU deck and (results days)
   `DMW - Results.xlsx`.
2. Deck check first (29 Oct 2026, the user asks for it every time): run `check_final.py` as in
   Task 11 step 3 and, on results days, `check_results.py` as in Task 11 step 4, even when Task 11
   already ran, because the team saves again before export. Pass every FLAG to the user.
3. Run:
   `python SHARED/scripts/build_email_mb.py <folder with the final EN deck> "<previous EMAIL MB>.docx" SHARED/data/Mapping.xlsx "EMAIL MB YYYYMMDD.docx" "Daily Market Update - YYYYMMDD.pptx"`
4. Read its checks: dropped calendar rows (EGMs only), results table mismatches over 0.15, names
   not in the sheet, empty sections, and every "DMU vs DMW Results table" line (figures, TASI
   codes, and rows on one side only). For a mismatch, go back to `DMW - Results.xlsx`
   (unrounded figures) to say which deck is wrong. Pass them to the user.
5. Open the docx and read it once: results sentences first in Major Corporate News, full change
   clause, `(Tadawul, Bloomberg)` tag (coverage names: `(Tadawul, Bloomberg, anbc research)`),
   consensus wording bands right. Coverage names: each sentence carries its own company's DMU
   paragraph only (29 Oct 2026: two analyst paragraphs in one DMU box were both pasted into each
   sentence; the script now splits the box and matches by ticker).
6. Open-file check, then save to the day folder.
7. If the user edits the email, diff their version against yours, fix the script and rerun until
   it reproduces their version; send the fix to the plugin maintainer.

## Hand-off check

- [ ] Deck checks (`check_final.py`, and `check_results.py` on results days) run and passed to the user
- [ ] Built from the final EN deck; checks passed to the user
- [ ] DMU results table cross-checked against the DMW Results table
- [ ] Corporate Calendar: every event row except EGMs
- [ ] Results sentences: full clause, recomputed figures, right consensus band, tag present
- [ ] Coverage names: one DMU analyst paragraph each, Bloomberg ticker from the table, 1dp figures
- [ ] Saved to the day folder
