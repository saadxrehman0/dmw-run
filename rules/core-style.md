# Writing spec: core style (Economic and Corporate news)

Derived from the published editions (Aug 2026 onward) and the desk's confirmed decisions.

## The red line: nothing critical of Saudi Arabia (team, 29 Oct 2026)

**Never include an article, a sentence or a quote that criticizes Saudi Arabia, its government, its
leadership or its policies**, in any output (DMW, DMU, Sales Email, Daily Bulletin, SAMA Monthly,
Global and TASI write-ups). The team's words: it is "the fastest way to lose our job". This
overrides every selection and length rule. When in doubt, leave the item or the clause out and tell
the user.

- Straight reporting of events that hit Saudi Arabia (Houthi attacks on facilities, pipeline
  closures, export disruptions) is not criticism: it ran in 22 of 96 DMU oil paragraphs, 1 Jul to
  21 Oct 2026, and in the 29 Oct oil caption.
- Leave out unconfirmed claims of damage to Saudi facilities. On 29 Oct the team cut "an attack on a
  refinery in Saudi Arabia" from the oil paragraph, where the source said Saudi authorities had not
  confirmed the strike (my reading of why; the team did not say).
- Drop analyst or official comments that blame or second-guess Saudi policy (OPEC+ decisions,
  fiscal choices, the conduct of the war), even inside an otherwise neutral story.

## Decisions

| Question | Decision |
|---|---|
| Economic bullet count | At least three where the day allows, never fewer than two; the deck carries two |
| Corporate bullet count | Over-supply, the team trims from the bottom |
| `(Arab News)` | Approved tag |
| Decimals | **One decimal**, with one carve-out (ownership percentages, below) |
| Arabic version | Each item carries its Arabic version in the ANB Rough doc; rules in `arabic-rules.md` |
| Source URL line | **Required**, own line under each bullet |
| Length | Brief, no wordy bullets |
| Debt-only issuers | **TASI code alone** where no Bloomberg equity ticker exists |
| Banks' pre-Zakat profit item | **Economic**, not Corporate |
| `pre-Zakat` | Always this form, capital Z (both cases appeared in print) |
| Trillion | **`tn`**, lowercase like `mn` and `bn` |
| Oil benchmarks | **`Brent` and `WTI`**, never "Brent futures", "Brent crude futures" or "(US) West Texas Intermediate crude (futures)" (management rule, 29 Oct 2026; 78 of 96 DMU oil paragraphs use bare Brent, 77 use WTI) |

## Global style rules

- **Currency code first, always, including dollars.** `SAR 301.2 mn` · `USD 8.0 bn` · `EUR 135.3 mn`.
  Never `28.0 mn`. Never `SAR 14.7mn` without the space. Never `$`, "riyals" or "SR".
  (The DMU's agency items on slides 2 and 3 keep `$`: `dmu-writing.md`.)
- **Magnitudes lowercase: `mn`, `bn`, `tn`.** Spelled out only in Arabic (مليون / مليار / تريليون).
- **One decimal on every figure**, ownership percentages excepted. `3.57%` → `3.6%`.
  `SAR 106.93 mn` → `SAR 106.9 mn`. `USD 1 bn` → `USD 1.0 bn`. Argaam's `6%` → `6.0%`.
  SAR 1,152.5 mn → SAR 1.2 bn. Never add precision the source does not have.
- **Whole-number targets and counts stay whole.** A target or a count the source gives as a whole
  number takes no added `.0`, in EN and AR: `80% of properties registered by 2028`, `2 mn property
  records`, `5 mn properties` (user, 29 Oct 2026, Real Estate Registry item). Currency values and
  percentage changes keep the one-decimal rule above.
- **Figures reconcile:** if a total is broken into parts, the parts sum. Keep the reconciling item
  (14.0 mn tons less 11.3 mn classified needs the 2.7 mn unclassified).
- **`%`, never "percent".**
- **Past tense throughout.** `announced`, `signed`.
- **Company form:** `Long Name English (TICKER AB)` on first mention in each bullet (the team
  reorders and trims, so every bullet carries its own first mention), short name thereafter.
  Company name plus ticker only, nothing else in the brackets.
  **Debt-only issuers:** `Bidaya Finance Co. (5025)`.
  A listed bank inside a product or index name takes no ticker ("Riyad Bank Saudi Arabia PMI").
  Unlisted entities take no ticker.
- **Source tags, exactly these strings:** `(Tadawul)` · `(Argaam)` · `(Zawya)` · `(Arab News)` ·
  `(Reuters)` · `(Reuters, Bloomberg)`. Every bullet carries one, after the closing period:
  `… starting from 3Q26. (Tadawul)`
- **Source URL on its own line** under each bullet. A second source used to check a figure goes on
  a following line labelled `Corroborating:` or `Conflicting figure:`.
- Quarters `3Q26`, halves `1H26`, ranges `3Q26 to 3Q29`; comparative quarters too ("the same
  quarter of 2025" → `2Q25`).
- Dates `Aug. 27`, `Sept. 30, 2026`. Drop the date an announcement was made and a bare effective
  date; keep dates that are terms of the deal (start and end of a duration, eligibility dates).
- American spelling (center, neighborhood, authorize); proper names and URLs keep their own.
- Abbreviations follow the source: keep an expansion the source wrote, add none it did not.
- Where the exchange labels an index "Index" (Household & Personal Products Index), write Index.

## Tickers

Look the ticker up in the names sheet `shared/data/Mapping.xlsx` (sheet `mapping`, column I
"Bloomberg", keyed on the 4-digit code in column D; Main Market and Nomu). Since 29 Oct 2026 it is
the one names file: names, codes and markets checked against Saudi Exchange (monthly, Task 0) and
tickers against Bloomberg by ISIN. **Never infer a ticker from the pattern of others** (ALRAJHI is
RJHI, ANB is ARNB, SENAAT is ZIIC, TASNEE is NIC). If the cell is blank (a new listing the team has
not filled yet), try `shared/data/TASI Mapping.xlsx` (Companies), then Bloomberg; if none has it,
omit the ticker (no `TK`) and tell the user so the row can be added.

**Foreign (non-Saudi) listed companies** take the Bloomberg ticker of their **primary listing**, in
Bloomberg's `CODE EXCH` form: `Citigroup (C US)`, `JPMorgan Chase (JPM US)`, `HSBC (HSBA LN)`,
`MUFG (8306 JP)`. Chinese banks with H and A shares take the Hong Kong H-share: `Bank of China
(3988 HK)`, `Industrial and Commercial Bank of China (1398 HK)`. Confirm each ticker against
Bloomberg before use (the quote page `bloomberg.com/quote/CODE:EXCH`, which can be found by
search when the site shows a robot check, or the Terminal); never write one from memory alone.
The Arabic carries the same ticker after the Arabic name (user, 27 Oct 2026).

**Look it up yourself before asking the user (user, 8 Nov 2026: "stop asking me for stuff you can
easily look up online").** A foreign ticker, a names-sheet ticker that looks odd, a revenue or GWP
denominator for an undisclosed contract value, or any other fact that is online gets looked up
before the review doc, never handed to the user as a question or a Terminal job. Search with
`mode: "extended"` when a standard search comes back thin. Routes that worked on 8 Nov 2026:

- Foreign ticker: WebSearch `"<CODE> <EXCH>" Equity Bloomberg <name>` finds the quote page
  (`bloomberg.com/quote/ONGC:IN` → `ONGC IN`).
- Saudi ticker check: a broker report's "Bloomberg ticker" line (WebFetch the PDF; Riyadh Cement
  → `SAWCEM AB`, matching the names sheet).
- Denominators: Argaam's results pages, fetched in the built-in browser from an argaam.com tab:
  `/en/financial-reports/company-report/<Argaam company id>/<year>/<period>?isNoHeaderFooter=true`,
  period `1` = full year, `70` = Q1, `71` = H1, `72` = 9M (Al Sagr is company id 963). Read the
  figure from the table and check it against the page's own % change before use. Insurers'
  contract thresholds use gross written premiums (GWP); others use revenue.

Ask the user only when every route fails, and say which routes were tried.

**Match on the entity, not a similar name.** Red Sea Global (unlisted PIF developer) is not Red
Sea International Co. (REDSEA AB). Naseej for Technology (NASEEJTE AB) is not Naseej International
Trading (NASEEJ AB). Al-Rashid Trading & Contracting (unlisted) is not Al Rashid Industrial Co.
An "...Ltd." suffix often marks an operating subsidiary, not the listed parent.
`SECO AB` is Saudi Energy Co. (formerly Saudi Electricity Co.; the names sheet carries the new name).
Al-Jouf Agricultural Development Co. (6070) is `JADCO AB`, not `ALJOUF AB` (the exchange short
name is the trap; 6 Sep 2026).
Watani Iron Steel Co. moved from Nomu (9513, WATANI AB) to TASI as 1325 on 6 Oct 2026. Reconfirm
the Bloomberg ticker before use.

### Known name inconsistencies to settle

Seen both ways in published editions; use one form per edition and flag the pair to the user.

| Ran as | Also ran as |
|---|---|
| Saudi Ceramic Co. | Saudi Ceramics Co. |
| Arabian Centers Co. | Arabian Centres Co. |
| Al Moammar Information Systems Co. | Al-Moammar Information Systems Co. |
| Axelerated Solutions for Information Technology and Communications Co. | … for Information and Communication Technology Co. |
| Middle East Pharmaceutical Industries Co. | Middle East Pharmaceutical Industries Co (no period) |

Settled: 4014 is written **Scientific and Medical Equipment House Co.** (with "and", as the
company is officially named; user, 24 Oct 2026). Not "&", not "Company".

## The carve-out: ownership percentages

Week-on-week ownership and market-share percentages in the weekly flow bullet carry **two
decimals** (`from 11.27% to 11.30%`), including the GCC-investor line. Buy/sell transaction shares
in the same bullet are single levels and take one decimal (`34.61%` → `34.6%`).

> Non-institutional foreign investors increased their ownership in Tadawul-listed equities,
> excluding Saudi Aramco (ARAMCO AB), from 11.27% to 11.30%, or SAR 353.1 bn, in the week
> ended Aug. 13. (Argaam)

## When one-decimal rounding collapses a move

12.82% and 12.81% both read 12.8%. Do not print the same figure twice as though it were a move:
give the current level and let the verb carry the direction, or drop the sentence, and flag it.

## Reading Tadawul filings directly

```
https://www.saudiexchange.sa/wps/portal/saudiexchange/newsandreports/issuer-news/
issuer-announcements/issuer-announcements-details/?anId=NNNNN&anCat=1&cs=CODE&locale=en
```

`anId` is sequential, so the day's range can be walked. Each page has a structured table
(Counterparty, Contract Value, Date of the Award, Contract Details, Related Parties) that maps onto
the Corporate bullet template, and lists the issuer's three previous announcements.

**Corrective announcements: read the body, never the title.** The title names the figure being
retracted (Anmat, 3 Sep: titled SAR 113.7 mn, corrected to **SAR 98.9 mn** in the body).

**Sweep order:** Tadawul first, then Argaam. The Tadawul filing date governs, not the Argaam
publication date.

## Sentences

- **Plain declarative sentences, like Argaam's.** One fact each, in the source's order. No stacked
  clauses, no dashed asides.
- **Short:** 1 to 3 sentences; a 4th only when the item needs it; never 5 (Sep and Oct 2026 decks: 83% of 358 items had 1 to 3 sentences, 13% had 4, 3% had 5). Cut attribution and procedural boilerplate ("in a statement
  to Tadawul", "will announce any material developments in due course").
- **Lead with the news.** Recurring items open with the figure that changed, then the dates.
- **Stay in the source's words.** To shorten, delete a whole sentence rather than rewording the
  ones that stay. Never merge two source sentences into one.
- **Attributed comments stay attributed:** name the speaker once, then "he added", "he cautioned".
  Unconfirmed reports keep "Reuters reported, citing sources". One "X said" per bullet at most.
- **One article, one bullet.** Never combine several filings about one deal.
- Keep "in addition to ..." clauses beside the figure they qualify.
