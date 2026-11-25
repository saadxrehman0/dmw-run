---
name: task-10-tadawul-news-results
description: Task 10 of the ANB DMW cycle. Sweeps late Tadawul filings into the DMW decks up to the 10:30 Karachi cut-off, and on results days builds the Results Announcements table from DMW - Results.xlsx checked against Tadawul and places it with place_results.py. Use for "Tadawul news", "late news", "morning news cut-off", "results table", "Task 10".
---

# Task 10: Tadawul news and Results (9:45 to 10:30)

SHARED = `../../shared/` relative to this skill's base directory. Read
`SHARED/rules/config.md` first (day folder, `WORK` folder, approvals), then
`SHARED/rules/news-selection.md`, `core-style.md`, `arabic-rules.md`, `deck-house-rules.md`
(Results days) and `deck-safety.md`.

## Prerequisites

- [ ] Decks filled (Task 8) and saved by the team; stage their **latest** copies
- [ ] Browser available for `saudiexchange.sa`
- [ ] **News cut-off 10:30 (user, 4 Nov 2026):** anything filed later goes into the next edition
      unless it is a top-3 item and the user says to add it
- [ ] **Results cut-off 10:30 (user, 4 Nov 2026):** rows saved to the workbook after 10:30 wait
      for the next edition. A correction to a row that was in before 10:30 (a period label, a
      figure) is applied
- [ ] **Who places the table:** Results is a schedule job. On the user's Pasting weeks, when the
      user says so ("put the results as they come in the deck, don't wait on me", 5 Nov 2026),
      Claude places and re-places the table on every Results save until 10:30 without asking again

## Steps: late news

1. Sweep Tadawul issuer announcements at 9:45, again at about 10:10, and a last time just before
   10:30, each from the last swept timestamp.
2. Triage and overlap-check; show the user what qualifies.
3. Draft EN and AR (Arabic counterparty names from `locale=ar`), show them, and ask where each item
   goes (rank, slide). Then patch it into the team's latest deck copies (never rebuild), keeping the
   house spacing; `check_ar_text.py` OK; commit with the mtime guard. When the user pastes an
   updated doc section into the chat, take the new items from it (5 Nov 2026: Bank Aljazira Tier 2
   sukuk and AlSaif tariff impact added this way), refill from the pre-Results copies and re-place.

## Steps: Results days (`DMW - Results.xlsx` appears after 10:00)

0. If the user asks you to build `DMW - Results.xlsx` yourself, start from the **most recent day
   folder's copy for the same quarter**, never an older quarter's. Keep its formulas; fill only A
   (period), B (short name), E, F, I, and leave I blank and yellow when Bloomberg consensus is not
   yet known (user, 20 Oct 2026).
   **Consensus source:** Bloomberg `<ticker> AB Equity EEO` (Consensus Overview, Quarters, BEst
   Standard, SAR), column for the reported quarter. Default row: **Net Income Adj+**.
   **Exception, banks: use Net Income, GAAP.** Read the reported quarter's **Est** column, never the
   year-earlier Est column next to the Act column (3 Nov 2026 CATRION).
   **Consensus far from the result** (opposite sign, or more than 10x off): keep what the workbook
   shows and mention it.
   **The user's Bloomberg screenshots (5 Nov 2026):** the user drops EEO screenshots (WhatsApp
   images) into a PC folder and connects it (`C:\Users\Saad Rehman\Pictures\results from yesterday`
   on 5 Nov). Stage the new ones, make one contact sheet per batch of 3 or 4, and tie each
   consensus (and any Act column shown) to the workbook. "No data available for this security"
   means no consensus: the workbook should show "-". Report the tie-out in a small table.
   **Team results messages** (pasted by the user): tie net profit and YoY to the workbook. Their
   "our estimate" is ANB's own estimate, not the Bloomberg consensus in the table, so the two
   beats differ (5 Nov: Aldawaa ANB 34.4 vs consensus 40.3; Fakeeh ANB 37.7 vs consensus 39.2).
   Say so; do not change the table. The messages are not deck news unless the user asks.
1. Fill company, TASI code, Bloomberg code, current, prior, change SAR mn, change %, consensus,
   variance. Plain black, 1dp, half-up. EN source line: `Source: Tadawul, Bloomberg`.
2. **The workbook does not auto-update.** Check period labels and every figure against the Tadawul
   announcement; recompute change and variance; ask the user if labels disagree. Never edit the
   team's workbook (the team keeps it open). When the user says to skip the Tadawul check (5 Nov
   2026, time), report which rows were checked against Bloomberg actuals and which come straight
   from the workbook.
   **Read the Tadawul table cell by cell** (3 Nov 2026): pull the `tr`/`td` cells of the "Net Profit
   (Loss) Attributable to Shareholders of the Issuer" row and check current / prior - 1 against the
   filing's own % change before using the figures.
   **Units:** read the line under each Tadawul table (Millions, Thousands, Actual). The workbook is
   in SAR mn. Check the Bloomberg unit too (M vs K vs B).
   **Verify every save** by content (sha256 or the rows themselves), per `deck-safety.md`.
   **Wrong period label in the workbook** (5 Nov 2026: Abo Moati arrived as 3Q26; it is 2Q27 on its
   April-March fiscal year, user): copy the workbook locally (`openpyxl`, `data_only=True`), fix
   column A in the copy only, and pass the copy to `place_results.py`. Tell the user the team's file
   still carries the old label. **Leftover labels** (2 Nov 2026): if a label does not fit the
   figure's size, ask the user rather than starring the row.
   **Results the owner has not added yet (user's Pasting weeks):** add the row to the decks from the
   Tadawul filing and the EEO screenshot, log it, and tell the user the workbook still lacks it.
   **Rows without an order number** in the workbook's last column (5 Nov: Aldawaa, Dar Alarkan,
   Entaj) are still today's rows; `place_results.py` sorts by its own rule, not that column.
3. Point out halves rounded down by floating point; the team's figure wins. **Exact ties:** compute
   change with exact decimals (`Decimal`, ROUND_HALF_UP). **Loss rows keep the workbook's sign**
   (client, 22 Oct 2026). **Fiscal-year quarters:** use the company's own quarter label (East Pipes
   and Abo Moati, April-March year: the Sep 2026 quarter is 2Q27). **Banks with no GAAP estimate:**
   use Adj+ and tell the user. **Bloomberg codes:** the deck uses the real ticker.
4. Estimated (unaudited management) results never go in.
5. **Row order (user, 1 Nov 2026):** ANB-covered companies first, largest current-quarter net
   profit first; then uncovered companies, largest net profit first (losses last). Coverage = the
   "ANB Universe" sheet of the newest `ANB Coverage Universe Market Cap - <date>.xlsx` in
   `G:\.shortcut-targets-by-id\18rkHxDXGFRnp_kbzgyFhbwCjVH4-XOpc\2026\ANB Coverage Universe\` (13 Oct
   2026 file, still the newest on 5 Nov). Different-period rows go **at the very end**, starred, with
   the period footnote under the source line. Re-sort the whole table on both slides on every new
   Results save. `place_results.py` applies this order.
6. **Placement: run `place_results.py`.** On every Results save up to 10:30, run it on the
   **pre-Results** filled decks (keep them; Task 8's fill output), so a re-run never stacks a second
   table:
   ```
   python place_results.py --results "DMW - Results.xlsx" --names state/Mapping.xlsx \
     --prev "<last results day's day folder>" --en <pre-Results EN deck> --ar <pre-Results AR deck> \
     --quarter 3Q26 --universe "<newest ANB Coverage Universe file>" \
     [--skip "<rows after 10:30>"] [--drop-cal "<calendar rows the owner dropped>"] \
     [--ar-short state/name_overrides_results_ar.json] --outdir <out>
   ```
   It clones the previous results day's table frames, recomputes change, % and variance with exact
   decimals, sorts, splits the rows between slide 1 and slide 2 the same way in EN and AR (the most
   rows that fit above 10.60 in on slide 1 in both decks), removes the empty slide 1 Corporate box,
   moves the slide 2 news and calendars down by the table's height, takes Arabic names from the
   names sheet (the Short Name - Arabic when the long one would wrap), blanks the figures on the
   crude oil inventories rows, and drops the calendar rows named in `--drop-cal`.
   **Since 5 Nov 2026 it also** stars the AR name of a different-period row (the short name used to
   lose the star), writes the period footnote under the source line in both languages and grows the
   source row so the news moves with it (never add the footnote by hand; the hand edit on 5 Nov left
   the news overlapping the table), and measures the slide 2 bottom without the page-number box (its
   "ends at 10.90 in" warning on 5 Nov was false; the real bottom was 7.75 to 9.8 in).
   Then patch the result onto the team's latest saved decks (or commit it if nobody has saved since)
   with the mtime guard and the content check. Read its report: "Missing AR names" means a TK in the
   AR table.
   The team's layout it follows (all 47 results days, 15 Jul to 21 Oct 2026):
   - The table sits under the slide 1 Corporate header and runs to 10.60 in. Slide 1 carries **no**
     Corporate item on a results day (Task 8 runs `fill_commentary.py --results`).
   - Slide 1 holds about 1 to 5 company rows (5 Nov: 2, then 1 after a longer Economic item). The
     rest continue in a second table under the slide 2 Corporate header (EN about 0.89 in; AR
     1.08 in since the 5 Nov AR template fix), same columns and header rows, with the Corporate news
     under it at the template gap (EN 0.068 in). The source line goes under the last table only.
   - **Refilling after the tables are placed:** `fill_commentary.py --force` treats the slide 2
     Results table as a calendar and stacks it under the news. Refill from the pre-Results copies and
     run `place_results.py` again.
   - Row heights: data rows 0.16 in, source row 0.12 in per line (the script sets them).
   - AR mirrors EN row for row (the AR tables carry one extra header row).
   - After placing, check that slide 2 still ends above 10.60 in. If it does not, move the
     lowest-ranked Corporate items out of both decks and tell the user which. **Never cut the crude
     oil inventories rows**, even past 10.60 in; say how far past it runs.
   - Results sentences are not added to the deck news unless the user asks (table only).
   - The build (Task 7) removes yesterday's tables on both slides, so a table on either slide after
     Task 7 is today's.

## Hand-off check

- [ ] Only items filed before 10:30 (or approved top-3 exceptions); EN and AR match
- [ ] Results figures checked against Tadawul cell by cell (or the user said to skip, and the report says which rows were checked), periods confirmed
- [ ] Consensus taken from the reported quarter's Est column; tied to the user's screenshots when given
- [ ] Rows saved after 10:30 left out (`--skip`) and named to the user
- [ ] Tables placed with `place_results.py` from the pre-Results decks; re-run on every Results save
- [ ] Different-period rows last, starred in EN and AR, footnote written by the script
- [ ] Slide 2 ends above 10.60 in (real bottom), or past it only because the crude rows were kept
- [ ] Decks patched on the team's latest copies with the mtime guard and the content check
