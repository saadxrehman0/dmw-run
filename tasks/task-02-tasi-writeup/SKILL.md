---
name: task-02-tasi-writeup
description: Task 2 of the ANB DMW cycle. Writes the evening TASI bullet (English then Arabic) from the Saudi Exchange market summary, indices performance and historical reports pages, and posts it to the ANB Rough Draft tab after approval. Use for "TASI write-up", "TASI bullet", "Task 2".
---

# Task 2: TASI write-up (evening)

Window: after the 3:00 PM Riyadh close, from about 5:30 PM Karachi. The bullet runs in the **next**
edition (written 31 Aug, published 1 Sep).

SHARED = `../../shared/` relative to this skill's base directory. Read
`SHARED/rules/config.md` first (day folder, `WORK` folder, approvals), then
`SHARED/rules/global-tasi.md` (TASI section and sector map) and the TASI part of
`SHARED/rules/arabic-rules.md`.

## Prerequisites

- [ ] Browser available for `saudiexchange.sa`
- [ ] The TASI write-up is not a teammate's turn tonight. When a teammate writes it, do not write
      or edit it; proofreading fixes go to the teammate directly.

## Steps

1. Open the three exchange pages (market summary; Indices Performance → Key Parameters for today's
   value traded; Historical Reports for the previous session's value traded). Use only these.
2. Pick the three companies from the page's own gainers or losers table (TASI up → gainers; down
   → losers; flat → gainers) and the top true sector from the heatmap (exclude TASI50, size
   indices, IPO and MSCI Tadawul 30).
3. Round company moves from the displayed two-decimal %, half-up; report sector ties as ties.
4. Write the English bullet in the house pattern, then the Arabic (المؤشر العام, sector names from the
   map, لكل منهما for ties). Run `SHARED/scripts/check_ar_text.py` on the Arabic.
5. **Frame check, phrase by phrase (user, 21 Oct 2026).** Put the drafted Arabic next to the
   published frame for the day's direction in `global-tasi.md` (up day or down day) and compare
   each sentence in turn: opening verb and بنسبة; the points in brackets; `ليغلق عند مستوى … نقطة`;
   the trading-value sentence; the sector sentence (`الأكثر ارتفاعًا` / `الأكثر تراجعاً`);
   `ومن بين الأسهم …`; then `تلاه سهما … و…` with the dual and the و between the two companies, and
   `على التوالي` or `لكل منهما`. Any phrase that differs from the frame is either today's data
   (names, figures) or a fix to make before showing the user. Why: the 20 Oct bullet went out with
   `يغلق` for `ليغلق` and `تلاه سهم … (YANSAB AB) شركة …` with no و and no dual; the checker could
   not see grammar and the proofreader missed it.
6. Log every figure's source link in `WORK/DMW/<date>/Source log - <date>.md`.
7. Show the user EN and AR; post to the Draft tab under the TASI heading after approval (no
   highlight), **replacing any TASI pair already there** (a teammate's version included; the deck
   fill reads only the Draft tab). Never touch the news sections while posting.
8. If the user pastes into the Google Doc by hand, the Arabic paragraph lands left-to-right and the
   %, brackets and numbers scramble. Tell them up front: File > Language > العربية (this only adds
   the direction buttons; the old Tools > Preferences option is gone), then click inside the Arabic
   paragraph and press the right-to-left ¶ button (under ⋮ on a narrow toolbar). Paste without
   formatting is greyed out for plain text copied from the chat and is not the fix.

## Hand-off check

- [ ] Only the two/three exchange pages used; nothing fabricated (`TK` if missing)
- [ ] Exactly three companies, from the right table; tickers from the mapping file
- [ ] "each" trails the percentage; ties reported as ties; "Index" for Household & Personal Products
- [ ] One decimal on every percentage
- [ ] Arabic says المؤشر العام; `check_ar_text.py` OK; frame check done phrase by phrase
- [ ] Source log written; user approved before posting
