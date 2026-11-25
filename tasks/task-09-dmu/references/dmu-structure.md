# DMU deck structure (v2, 21 Oct 2026)

Rebuilt from all 80 team DMUs from 1 Jul to 21 Oct 2026. It replaces the 12 Oct version. Writing rules are in `dmu-writing.md`. Review scripts and log: Claude Outputs/DMW/2026-10-21/DMU review/ on the personal Drive.

Back-tested 21 Oct 2026 on 4 Aug, 13 Sep, 19, 20 and 21 Oct. Rebuilt from the previous team DMU plus that day's data, with the team's own text and pictures as inputs: 2,086 table cells compared and 3 differ (company short names, where the team itself is inconsistent, e.g. Alarabia / ALARABIA). Charts match. Formatting matches apart from one extra price level in bold. The Sukuk slide is the team's.

29 Oct 2026: Bloomberg terminal sources added for the heatmap, Consensus TP and Results consensus, from the team's 22 Oct training video and Saad's answers (`claude/DMU video walkthrough - 2026-10-29.md`). 1 Nov 2026: slide 4 source (ANB Rough doc) and the Results slide rows, order, source note and layout, from the 29 Oct slide-by-slide test. 2 Nov 2026 (first from-scratch practice run): stock futures from the Futures page, captions on one line, consensus TPs from an EQS export, the technical slide taken whole from the team, the heatmap from the previous session, no template leftovers. 3 Nov 2026: Areeb's correction session with Saad (audio and screen recording, `claude/DMU corrections - Areeb - 2026-11-03.md`): consensus TPs checked on `ANR`, picture colours, and the 11 to 29 Oct Stocks-page futures logged as a training error. The DMU rotates weekly between team members, so where the team's decks disagree, these files set the standard (Saad, 2 Nov 2026).

## Slide order

Widescreen 13.33 x 7.5 in. The deck runs 6 to 8 slides. **Find each slide by its title, never by its position.**

| Order | Title | Present |
|---|---|---|
| 1 | Cover: "Daily Market Update", date, FOR INTERNAL USE ONLY | always |
| 2 | GLOBAL MARKETS | always |
| 3 | COMMODITY MARKETS | always |
| 4 | SAUDI MARKET | always |
| 5 (and 6) | Result Announcements | 46 of 80 decks. Two slides on 2, 4, 5, 6 and 9 Aug |
| next | Sukuk Issuances | 5 of 80 decks (19 and 26 Jul, 16 Aug, 6 and 13 Sep) |
| next | TASI TECHNICAL ANALYSIS | always |
| last | Back cover | always |

Deck sizes seen: 6 slides (33 decks), 7 (38), 8 (9). Since 22 Sep the commodities slide has a "USD/bbl" label box sitting off-slide at (-7.83, 0.16). Ignore it.

## Picture frames (Saad, 29 Oct 2026)

The team pasted each picture by hand, so the frames drifted (84 decks, 1 Jul to 29 Oct: recap 6.75 to 7.07 x 1.95 to 2.20 in, futures 5.16 to 5.82 x 1.72 to 2.12 in) and the images were stretched by up to 8% (recap) and 16% (futures). **The build now sets two fixed frames every run**, shaped like the Bloomberg snips so nothing is stretched (`build_dmu.py` `FRAMES`):

| Picture | Frame (x, y, w, h in) | Why |
|---|---|---|
| Market recap | 0.25, 1.17, 6.92 x 2.04 | Full left-column width; ratio 3.39 = the October snip median. The MARKET RECAP label moves up from 0.97 to 0.88 so the bottom edge stays at 3.21 and the headlines do not move (pushing them down 0.09 in ran 310 words into the footer) |
| Stock futures | 7.50, 3.85, 5.43 x 1.79 | Same width as the valuations table; ratio 3.04 = the October snip median; bottom at 5.64, above "Source: Bloomberg" at 5.67 |

Each snip is padded to the frame's shape with its own edge colour (recap Bloomberg dark grey, futures near-white), centred, then fitted. More than 10% padding is flagged as a likely bad crop. Data: `Claude Outputs/Workflow Automation/2026-10-29/dmu_frames.csv` (`dmu_frames.ps1`, `frames.py`). Best source (2 Nov 2026): Saad's full-size Retina Mac screenshots, cropped by Claude (futures 1760 x 575 px against the team's usual 1098 x 359).

**Colours (3 Nov 2026).** Areeb saw the Mac snips as looking different in colour from the team's. Saad: the Mac snips are sharper, and the team's laptop screens show colours more vividly than his Windows screen, so this is a display difference, not a file problem. No change to the capture. The team snips at 100% browser zoom, crops off the white margin, then aligns the picture with the table above it (select both, Align); the fixed frame does the alignment.

## Commodities right column (Saad, 29 Oct 2026)

The two charts, their captions and the USD/bbl label were placed by hand and drifted every day (84 decks, 1 Jul to 29 Oct: chart x 7.67 to 8.58, widths 4.6 to 5.1 in; 28 Oct 8.58, 29 Oct 8.08). **The build sets them to one fixed layout every run** (`build_dmu.py` `COMMODITY_LAYOUT`, the medians of those decks):

| Element | x, y (in) | w x h (in) |
|---|---|---|
| Oil chart | 7.83, 1.15 | 5.07 x 3.10 |
| Gold chart | 7.83, 4.40 | 5.07 x 3.10 |
| Oil caption | 7.83, 0.87 | 5.50 wide |
| Gold caption | 7.83, 4.25 | 5.50 wide |
| USD/bbl label | 7.92, 1.24 | as is |

Captions stay on one line at 12 pt (`dmu-writing.md`, Captions). If one still wraps, the build moves the chart below it down 0.20 in and shortens it by the same, keeping its bottom edge (2 Nov 2026).

Data: `Claude Outputs/Workflow Automation/2026-10-29/dmu_commod_pos.csv` (`dmu_commod_pos.ps1`, `commod_pos.py`).

## Results slide layout (29 Oct 2026 test)

- Table at (0.79, 1.15), 12.09 in wide; header rows 0.50 and 0.38 in, data rows 0.21 in, source row 0.55 in. The build sets the table frame to the sum of its rows.
- Source row: `Source: Tadawul, Bloomberg`, then a second line for a starred row's note (`*FY26`), 8 pt.
- Commentary box (`TextBox 4`, 12.26 in wide, 12 pt) sits 0.07 in above the bottom of the last row, about 0.48 in tall per paragraph. The build moves it; a box running past 7.0 in is flagged (use `split_at`).
- The 29 Oct team table held 14 rows on one slide with two commentary paragraphs, ending at about 6.4 in. 2 Nov 2026: 14 rows and three paragraphs also fitted on one slide (render before splitting).

## Element map

| Slide | Element | Source | Who | Changes |
|---|---|---|---|---|
| Cover | Date `D Month YYYY` (no leading zero: "1 October 2026") | build | script | daily |
| Global | Market recap picture (`Picture 4`, fixed frame 0.25/1.17, 6.92 x 2.04 in) | bloomberg.com Markets dropdown, Top Securities, first two rows | Saad's Mac screenshot, cropped by Claude | 78 of 79 days |
| Global | Headline updates (5 items, 11 pt) | `dmu-writing.md` | Claude | daily |
| Global | Comparative valuations (10 rows: Last Close, %∆, Forward P/E, YTD %∆) | `DMW - Data.xlsx` PPT-ENG D109:H118 | script | daily |
| Global | "Stock futures" picture (`Picture 8`, fixed frame 7.50/3.85, 5.43 x 1.79 in) | **bloomberg.com Markets > Futures**, Today in the Markets, Popular Indices tab: the 10 index-futures cards (`xx1:IND`) (Saad, 2 Nov 2026). Not the Stocks page's tab of the same name (index cards SHSZ300 to SPTSX). The team's decks from 11 to 29 Oct used the Stocks page because of a **training error**: the trainer confused the two pages and taught it on (team, 3 Nov 2026). The Futures page ran on 1 to 8 Oct and 1 to 3 Nov. Evidence: `Claude Outputs/Workflow Automation/2026-11-03/DMU stock futures - *.png` | Saad's Mac screenshot, cropped by Claude | every day |
| Global | Events calendar (10 pt; date run bold, rest regular, black) | Manual File-Calendar / investing.com | Claude | daily |
| Commodities | Performance recap table (Last Close, %∆, MTD, YTD; Bloomberg 2026e, Avg 2026e) | PPT-ENG D122:H133; forecast columns from the week's `DMU Commodity Data` file, passed every day | script | daily; forecasts weekly |
| Commodities | Oil chart (line, Arab Light + WTI) and gold chart (line) | Linked to the day's `DMW - Data.xlsx`, `Sheet1!T:V` (oil), `Sheet1!X:Y` (gold) | **script** (was "team") | daily |
| Commodities | Two captions above the charts (12 pt bold, #0070C0, one line) | `dmu-writing.md` (Reuters headline, trimmed to one line when long) | Claude | daily |
| Commodities | Headline updates (oil, gold, 11 pt) | `dmu-writing.md` | Claude | daily |
| Saudi | TASI and NOMU strips | PPT-ENG C6:T8 | script | daily |
| Saudi | Performance heatmap picture (`Picture 9`, about 0.43/1.87, 7.32 x 2.42 in, 1580 x 618 px) | Bloomberg terminal: load `SASEIDX Index`, run `IMAP` (Intraday Market Map). Group by **Sector** (GICS), Market Cap (USD) **All**, view **Pie Chart & Table**, **previous session** (Historical Performance when snipping after the Saudi open; "All Securities" should match the TASI DoD on the slide), mouse off the chart. Snip the pie and the sector table together | **Saad** (shared terminal over AnyDesk); Claude places it, padded to the frame | 77 of 79 days |
| Saudi | Leaders and laggards (5 each: Last Close, %∆, Vol (mn), Val (SAR mn), Consensus TP) | PPT-ENG D85:O89; names via name_overrides | script | daily |
| Saudi | Consensus TP column | Saad's **EQS export** (Tadawul All Share Index members, "BEst Target Px:D-1", tickers), mapped by Bloomberg ticker, one decimal, `-` where none (2 Nov 2026: 9 of 10 matched the team's deck; Sumou 50.5 vs the team's `-`). **The EQS field is a day old (`D-1`) and can differ from `ANR`** (Areeb, 3 Nov 2026: Sumou Real Estate has no target price, so the team writes `-`). Check the 10 leaders and laggards on `ANR <ticker>` before handing over; if the terminal is not free, list them as a TEAM check. Without an export: blank for the team, source `ANR <ticker>`. The team writes a number or `-` (403 numbers, 397 `-` across 80 decks) | Saad (terminal), Claude maps | daily |
| Saudi | P/E column chart | Linked to `DMW - Data.xlsx` `'Graph PE'!O6:S24`; categories Dec-11 to Dec-27e (Dec-26e until late Sep) | **script** | daily |
| Saudi | Headline updates (3 or 4 items, 11 pt) | **ANB Rough doc** at build time (Saad, 29 Oct 2026); selection rule in `dmu-writing.md` (economic item first when strong, contracts by size, then the other corporate items in the DMW's order; 3 Nov 2026) | Claude | daily |
| Results | Table (every result in `DMW - Results.xlsx` at build time; coverage names first, then by net profit) plus commentary | Task 10, `DMW - Results.xlsx`. Consensus from Bloomberg `EEO <ticker>`. anbc estimates and commentary from the covering analyst's WhatsApp bullet, pasted by Saad (a results-preview file as the fallback source is pending, see `dmu-writing.md`) | Claude (consensus: terminal; estimates and commentary: analyst) | results days |
| Sukuk | Table: Company, TASI Code, Bloomberg Code, Size (SAR mn), Rate, Type, Maturity (2 to 4 rows) | announcements | team, for now | rare |
| Technical | **The whole slide** (big chart, title, outlook box, three right-hand images: Today's Summary, Market Summary, Intraday Index Levels) | **The data/technical team's DMU in the day folder**, taken whole (Saad, 2 Nov 2026), as Task 8 takes slides 4-5 for the DMW. Fallback when it is not there yet: chart, title and outlook from the Technicals EN deck (`Picture 3`; `Text 5`; `Content Placeholder 2`, Market Brief joined into one paragraph, levels bold, "Trend is ..." second) and TO UPDATE placeholders for the three images | team (fallback: script) | daily |

## Bloomberg terminal

The terminal is a shared PC reached over **AnyDesk**, one user at a time, so terminal steps can wait on someone else (22 Oct video: about six minutes). I do not drive it. Functions used for the DMU:

| Function | For |
|---|---|
| `SASEIDX Index` then `IMAP` | Slide 4 heatmap |
| `EQS` (saved screen: Tadawul All Share Index members, "BEst Target Price" as "No condition, display only", Ticker), copied to Excel | Consensus TP, all names at once (a day old, `D-1`) |
| `ANR <ticker>` | Consensus TP, one name at a time; the check for the 10 leaders and laggards (3 Nov 2026) |
| `EEO <ticker>` | Results slide consensus net profit |
| `TOP` <GO>, **More »** on **Eco & Gov**, date **Last Day** | Macro stories for slide 2 when Reuters has none |

## Charts (all three are regular PowerPoint charts linked to Excel, not think-cell)

- The think-cell object on the commodities slide ("think-cell data - do not delete", 0 x 0) only holds data. Leave it alone.
- Chart shape names change from day to day (Chart 4/8, 8/10, 10/11). **Find charts by slide and position**: oil above gold on the commodities slide (both then set to the fixed layout above), P/E at about (8.16, 5.05).
- Auto-update is off, so the chart shows its cached values. Rebuild the cache (categories and values) from the data file, and point the external link at the day's data file. 214 of 243 team charts linked to the same day's file. The other 29 were left pointing at an older day, which is a team slip and not a rule. Keeping the template's charts and updating their data (rather than pasting charts in) keeps formatting fixed and lets the data be checked in code (Saad agreed, 2 Nov 2026).
- The data file carries weekend rows that repeat Friday's close. The build ends the chart at the **last weekday row before the edition**: the previous day from Tuesday to Thursday, Friday on Sunday and Monday. The team's decks sometimes end on the weekend carry-forward row instead (same values, different last date). `check_dmu.py` compares the chart's last date with the data file.
- **Sunday and Monday show the same oil and gold charts** (Areeb, 3 Nov 2026: "on Monday we don't change" them): both end on Friday's close. The team's Sunday/Monday pairs of 11/12, 18/19 and 25/26 Oct had identical chart data. On 1/2 Nov the Monday chart ended on a later weekend carry-forward row with the same values. On 4 Oct the Sunday chart was stale (it ended on Wed 30 Sep), a team slip.
- The P/E categories roll forward one year in late September (Dec-26e became Dec-27e on 29 Sep).

## Formatting spec (held in 98 to 100% of decks)

| Box | Font | Size | Bold | Colour |
|---|---|---|---|---|
| Slide titles (GLOBAL MARKETS etc.) | Arial Nova Cond | 20 | yes | inherit |
| Section labels (MARKET RECAP, HEADLINE UPDATES, ...) | Arial Nova Cond | 12 | yes | #0070C0 |
| Headline bodies (slides 2, 3, 4) | Arial Nova Cond | 11 | no | inherit (no fill) |
| Chart captions | Arial Nova Cond | 12 | yes | #0070C0 |
| Events lines | Arial Nova Cond | 10 | date run bold, text run regular | #000000 |
| Footnotes ("Source: Bloomberg", "*Sourced from Bloomberg ...") | Arial Nova Cond | 7.5 | no | inherit |
| FOR INTERNAL USE ONLY | Arial Nova Cond | 10 (14 on cover) | no | inherit |
| Cover title | Arial Nova Cond | 36 | no | inherit |
| Results slide title | Arial Nova Cond | 20 | yes | theme TEXT_2 |
| Results commentary | Arial Nova Cond | 12 | no | inherit |
| Technical title | inherit | 12 | yes | #0070C0 |
| Technical outlook | inherit | 10 | levels bold, rest regular | #000000 |

Rules for the build: replace text **run by run**, keeping each run's properties. Never collapse a paragraph into one run where the template has two (events) or many (outlook). Never add a run without copying the properties of its neighbour.

## Timing

- Team final saves (file "modified" time) mostly fall between 10:10 and 10:50 PKT, after the build (about 9:50). The deck keeps being edited (results, TPs, heatmap) until the PDF is made.
