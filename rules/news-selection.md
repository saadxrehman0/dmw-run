# News selection, sections and bullet shapes

Derived from 88 published EN decks (1 Jun to 5 Oct 2026, 913 items) and approved by the desk.
`shared/scripts/triage.py` encodes the triage order. A cut rate seen only in drafts is not enough to
exclude a type: check it against the published decks.

## Triage checklist (run first, on the full headline list, before fetching any article)

First match decides.

| # | If the headline is... | Verdict | Section |
|---|---|---|---|
| 1 | Already in the latest published deck (any source) | Exclude | - |
| 2 | A market report: index close, 52-week / all-time lows, negotiated deals, weekly market cap or traded value, major ownership changes, index weights, "what you need to know", ex-dividend, commodity or crypto prices | Exclude | - |
| 3 | A fund ranking or screen | Exclude | - |
| 4 | Foreign news with no Saudi angle | Exclude | - |
| 5 | Weekly foreign flows (Mon) or non-institutional foreign ownership (Tue/Wed) | Include, #1 | Corporate |
| 6 | Official macro data (PMI, GDP, CPI, budget, SAMA monthly series, GASTAT) | Include, top | Economic |
| 7 | Ministry or authority sector statistics | Include | Economic |
| 8 | A board, CEO or committee appointment or resignation | Exclude | - |
| 9 | A dividend declaration, any company, large-cap banks and Aramco included (none ran in the 40 published decks, 26 Aug to 21 Oct 2026; user, 22 Oct 2026) | Exclude | - |
| 10 | A sukuk procedure or default / non-payment notice | Exclude | - |
| 11 | A buyback, ESIP allocation or loss-restructuring plan | Exclude | - |
| 12 | An interview or outlook piece without news | Exclude | - |
| 13 | A results-date notice ("X will announce its Q3 results on 20 Oct") | Drop | Not used: no Sep or Oct 2026 deck carried one in the news or the event calendar (only earnings calls) |
| 14 | Estimated financial results (management's unaudited figures, "announces the estimated financial results") | Exclude, and tell the user the company has reported | - |
| 15 | A contract, award or work order | Include, rank by value | Corporate |
| 16 | M&A, a stake, or a giga-project / PIF company | Include | Corporate |
| 17 | Large-cap strategy or CEO comments (Aramco, SABIC, the banks) | Include | Corporate |
| 18 | A credit facility, insurance product approval, showroom, branch or opening | Include, bottom | Corporate |
| 19 | A tax or regulatory penalty on a listed name | Maybe | Corporate |
| 20 | An MoU or cooperation deal without a value | Maybe | Corporate |
| 21 | A rebrand or new trading name of a listed company | Include, bottom | Corporate |

Also excluded: lawsuit updates and appeals; capital hikes, rights issues and bonus-share increases
(they go to the event calendar); dividend policies, even for large caps; routine exchange
administration; broker notes and analyst assessments; plain results; corrective announcements as
items of their own. Capital reductions are excluded too (none ran in the 40 published decks, 26 Aug
to 21 Oct 2026; user, 22 Oct 2026).

Also in scope: license grants and renewals; credit-rating actions; the MSCI periodic review, which
covers both Global Standard and Small Cap (report additions and deletions in both tiers).

**Block sales are not "negotiated deals".** A bookbuilt block sale executed through the CMA's
negotiated-transactions mechanism (named bookrunner, struck price, a stake sold) runs as a
Corporate item. The exclusion covers only the exchange's routine daily negotiated-deals screen
(ACDIMA selling its 10.5% SPIMACO stake via GIB Capital, 6 Sep 2026).

Same-event follow-ups: keep one (keep the listing debut bullet, drop the earlier approval).

**Wars and conflict: a tone rule, not a topic ban.** Name the war and the parties only as the
backdrop to an economic or company story, attributed to its source, in neutral wording. Standalone
war-event reports with no market or company angle (strike claims, interceptions) stay out.

### Precedents

- Accepted (4 Oct 2026): Tadawul market-maker approvals for large caps (SNB Capital on Almarai,
  Al Rajhi, SABIC); Tadawul trading suspensions for missed financial-statement deadlines (Amwaj,
  Rawasi Albina); a commercial property auction (Royal Mall, Riyadh).
- Skipped (4 Oct 2026): a Nomu company republishing its board report ahead of a TASI move (Saudi
  Networkers).
- Placed (5 Oct 2026): the CMA chairman's comments ran at the bottom of Corporate.

## Economic or Corporate

- **Economic:** macro and statistical data, regulator or ministry action, sector-wide statistics
  (including the monthly banks' pre-Zakat profit item and Argaam's monthly cement sales report),
  OPEC+ decisions. No listed-company story. When unsure about a national-level statement, include
  it and let the team decide.
- Ministry and authority items stay Economic even when a ministry signs or awards something:
  industrial licenses, building permits, contractor-sector award totals, Mawani concessions.
- **Corporate:** any company or project story, listed or not. The weekly flows / ownership item
  opens Corporate and is never Economic.
- Where past editions placed a kind of item differently, follow the house precedent.

Do not propose in Economic: bilateral MoUs, royal calls and visits, foreign reserves, remittances,
port throughput, weekly POS, OPEC / IEA demand forecasts.

## Economic section

Draft several, ranked: at least three where the day allows, **never fewer than two**. The deck
carries exactly two. Official national data leads: GASTAT, SAMA monthly series, PMI, budget, GDP,
IMF and World Bank outlooks.

Authorities seen in use: GASTAT, SAMA, MOMAH, CST, Monshaat, SCA, NDMC, MISA, HRSD, Mawani, SFDA,
Insurance Authority, CMA.

Argaam's monthly cement sales report (URL path `financial-reports/cement-monthly-report`) carries
no bar-graph icon and is in scope as a sector statistic (6 Oct 2026).

Shape: headline statistic with period and comparison basis → issuing authority in full with
acronym on first use → one or two supporting breakdowns.

> Saudi Arabia's inflation rate recorded an increase of 1.8% YoY in July 2026. According to the
> General Authority for Statistics (GASTAT), the Consumer Price Index (CPI) rose to 105.6 points,
> compared with 103.8 points in July 2025. The annual increase was mainly due to a 4.2% rise in
> housing, water, electricity, gas, and other fuels, driven by a 4.3% increase in actual rentals
> for housing. (Argaam)

## Corporate section

Median 9 items per edition (5 to 16); fewer when slide 1 carries a Results table. Over-supply and
rank; the bottom is cut first.

Ordering after slot 1 (flows / ownership):
1. Index and market-structure changes (MSCI review, TASI transfers, new listings)
2. Sovereign-linked entities (PIF, Aramco, Maaden)
3. Contract and project awards, largest disclosed value first (no size floor)
4. M&A, stake sales and acquisitions
5. Financing, credit facilities and sukuk
6. MoUs and non-binding agreements
7. Management commentary and strategy updates
8. Nomu filings (in scope on the same footing as Main Market)

Bottom fillers, cut first: credit facilities, insurance product approvals, openings, showrooms,
branches, rebrands.

### Bullet template

Three sentences, a 4th only when needed, never 5 (house limit in `core-style.md`). Sentence 1: `Long Name (TICKER AB)` + past-tense verb + named
counterparty + subject + value + VAT treatment. Sentence 2: duration or term. Sentence 3:
financial-impact timing as a quarter, or "There is no material immediate financial impact at this
stage."

> Al Moammar Information Systems Co. (MIS AB) has received two purchase orders from the Ministry
> of Municipalities and Housing, with a total value of SAR 301.2 mn (including VAT). The contract
> duration is 3 years, with financial impact expected to start from 3Q26 to 3Q29. (Tadawul)

Verbs: signed · announced · completed · received · renewed · won · obtained · was awarded ·
entered into · secured. VAT: `including VAT` / `excluding VAT`, one form per edition.

**Keep conditionality:** *expected to be signed on* · *contingent upon* · *subject to* ·
*non-binding* · *potential transaction* · *upon mutual agreement*. A conditional purchase is not a
completed one.

**Undisclosed values:** % of revenue, with the SAR denominator from Bloomberg:

> … at a value exceeding 30% of its total revenues for 2025 (VAT inclusive). MIS reported revenue
> of approximately SAR 1.3 bn for 2025, with 30% amounting to around SAR 381.7 mn. (Argaam)

**Management commentary** (a named executive on the company's own outlook) runs as a normal
bullet; third-party analyst opinion does not.

## Sources, in sweep order

1. **Tadawul issuer announcements** (primary). `saudiexchange.sa` blocks plain fetching: use the
   browser, read the list's page text, pull each row's detail URL from the DOM (`anId`, `anCat`,
   `cs`), and fetch details in-page (`locale=en` and `locale=ar`). If the user names a timestamp to
   start from, take that row and everything newer. Tag `(Tadawul)`.
2. **Argaam** main-news feed:
   `https://www.argaam.com/en/article/sectionarticles/sectionid/138/marketid/3/pageno/1/main-news-saudi`
   (the old `/en/main-news/sectionid/138/marketid/3/pageno/1` path returns Page Not Found since
   26 Oct 2026; if this one moves too, open `argaam.com/en` and follow the "Main News" link). Page 2
   needs a login. Skip rows with a lock icon (paywalled) or a bar-graph icon (financial reports).
   Tag `(Argaam)`.
3. **Zawya / Arab News**: allowed on any day, with one condition (Saad, 2 Nov 2026): **the news must
   be new.** Both outlets sometimes run stories days after the event, so for every item check that
   the event itself (the announcement, the data release, the deal) happened since the previous
   edition, not just that the article is recent: find the date of the underlying announcement or
   release (Tadawul, the issuing authority, the company), and drop the item if it is older or was
   already run in an earlier deck or post log. Zawya's site search returns only 20 mixed-relevance
   results and cannot be paged, and its article IDs are not chronological (a low ID is not an old
   article), so check each date. No other restriction applies to these two sources (the earlier
   "only if the team decides that day" rule was removed on 2 Nov 2026). Tag `(Zawya)` /
   `(Arab News)`.
4. Bloomberg for fields nobody publishes (consensus, revenue denominators).

If a fetch fails, ask the user to paste the content. Never retry in a loop, never substitute a
different source or a remembered figure.

## Overlap check (before posting any item)

Check each new bullet against the previous day's post log and against the news slides of the
latest published DMW decks (read with python-pptx). If a deck is cloud-only on the Mac, stage it
with `device_stage_files` rather than copying it in the shell. Drop the same event from a second
source. The counterparty's own filing of the same deal is its own bullet, but flag the overlap.
Flag any figure or date that disagrees between two filings.
