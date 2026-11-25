---
name: task-09-dmu
description: Task 9 of the ANB DMW cycle. Builds the internal Daily Market Update (DMU) deck from the previous DMU, DMW - Data.xlsx, the final DMW deck, the ANB Rough doc and the Technicals deck. Captures the two Bloomberg pictures, writes the commentary to the team's pattern, refreshes tables and charts, fills the Results slide on results days, and checks the deck against the house rules. Built after Task 8 and the Technicals (usually about 9:50 Karachi). Use for "DMU", "Daily Market Update", "build the DMU", "Task 9".
---

# Task 9: Daily Market Update (after Task 8 and the Technicals, usually about 9:50)

> Times are usual windows, not deadlines (per-task cut-offs dropped 4 Nov 2026; late-input deadlines in Task 0).

The DMU is an internal deck (FOR INTERNAL USE ONLY) for ANB's senior management, reviewed by the team.
File: `Daily Market Update - YYYYMMDD.pptx` in the day folder.

Rebuilt 21 Oct 2026 from a review of the team's 80 DMUs (1 Jul to 21 Oct 2026). **Where the review
conflicts with an earlier DMU instruction, the review wins** (user, 21 Oct 2026). The team's feedback
on the 20 and 21 Oct runs was: fonts off, charts and pictures not updated, and commentary that took
longer to fix than to write. Updated 29 Oct 2026 from the team's training video (sources, Bloomberg
terminal steps) and the slide-by-slide test against the team's 29 Oct final (picture capture, fixed frames),
1 Nov 2026 (slide 4 from the ANB Rough doc and its selection rule; Results slide order, footnote, layout)
and 2 Nov 2026 (first from-scratch practice run: no template leftovers, technical slide from the team,
Bloomberg pictures from Saad's Mac screenshots, stock futures from the Futures page, captions on one
line, shorter agency items, latest data, Zawya and Arab News allowed when the news is new, slide 4
corporate items in the DMW's order, consensus TPs from an EQS export).

**The DMU work rotates weekly between team members, so the team's decks are not a reliable standard on
their own** (Saad, 2 Nov 2026): the stock futures picture switched between two Bloomberg pages three
times in October. Where the decks disagree, this file sets the standard and Claude keeps to it.

SHARED = `../../shared/` relative to this skill's base directory. Read `SHARED/rules/config.md` first,
then `references/dmu-structure.md` (next to this file) for the slide map, charts and formatting spec,
and `references/dmu-writing.md` (next to this file) for the commentary rules. Both are required reading.

## Prerequisites

- [ ] Template: the DMU file the data team left in today's day folder if there is one, otherwise the
      previous edition's DMU. Keep the previous edition's DMU path for the checks as well. **Practice or
      "from scratch" runs** (Saad, 2 Nov 2026): use the previous edition's DMU as the template, never the
      team's same-day DMU, and do not read or copy anything from the team's same-day DMU (text, names,
      TPs, estimates) until the build is finished; it may be used afterwards only for the step 8 diff,
      with one exception: the technical slide (below).
- [ ] Today's `DMW - Data.xlsx`
- [ ] **The final DMW EN deck** (`ANB Daily Market Watch - EN.pptx`, after Task 8) for headline 1 on
      slide 2.
- [ ] **The ANB Rough doc** as it stands at build time, for the slide 4 items. The DMU goes out before
      the DMW, so slide 4 is picked from the doc, morning additions included (Saad, 29 Oct 2026).
- [ ] **The technical slide from the team** (Saad, 2 Nov 2026): the data/technical team makes the whole
      TASI TECHNICAL ANALYSIS slide, including the three right-hand images, and leaves it in the DMU file
      in today's day folder. Take that slide **whole**, as Task 8 takes slides 4-5 for the DMW decks.
      Today's `Technicals -ENG-YYYYMMDD.pptx` (about 9:30) is the fallback when the team's slide is not
      there yet; then the three images become TO UPDATE placeholders.
- [ ] Results days: the Results rows (`DMW - Results.xlsx` / Task 10) with consensus and anbc estimates
- [ ] Sundays: `DMU Commodity Data - <date>.xlsx` (weekly forecasts). Other days pass the week's latest
      file with `--forecasts` too, so the forecast columns are sourced, not carried from the template.
- [ ] The two Bloomberg pictures (step 1)

## Steps

1. **Get the two Bloomberg pictures.** **Preferred (Saad, 2 Nov 2026): ask Saad for two full-size Mac
   screenshots of bloomberg.com (US Edition) and crop them.** His screenshots are Retina, 4112 px wide,
   and gave much sharper snips than anything Claude captured: futures 1760 x 575 px (324 px per inch on
   the slide) against the team's usual 1098 x 359 (202) and Claude's browser capture 792 x 252 (146).
   Coordinates read off the image as shown in the chat must be scaled to the file's real width (shown at
   2000 px, file 4112 px: factor 2.056); check the file size first.
   - **Market recap**: homepage with the **Markets** menu open (hover it). Crop the Top Securities block,
     **first two rows only** (S&P 500 to Russell 2000, Nikkei 225 to Gold), no currency row or heading,
     with about 16 px of margin left and right and 8 px top and bottom. The menu is still there (Saad's
     Safari, 2 Nov 2026); Claude's browser tools could not open it on 2 Nov (hover and scripted events
     did nothing, no `MarketDataDropdown` element found), so this picture needs Saad's screenshot.
   - **Stock futures: the Futures page** (Saad, 2 Nov 2026): bloomberg.com → Markets → **Futures**
     (`bloomberg.com/markets/futures`, page title "Futures"), "Today in the Markets", **Popular Indices**
     tab: the 10 index-futures cards (2 Nov: KM1, ES1, A51, BC1, IFB1, NQ1, VG1, BZ1, NZ1, Z 1). Crop
     the cards with a 3 px margin, without the heading, tabs or heatmap. It matches the slide's "STOCK
     FUTURES" heading. **Not the Stocks page**: `bloomberg.com/markets/stocks` has a tab with the same
     name, "Popular Indices", showing index cards (SHSZ300 to SPTSX); the rota used that one from 11 to
     29 Oct and the Futures page on 1 to 8 Oct and 1 to 2 Nov. Check the cards read `xx1:IND` before
     cropping.
   - **Fallback when Saad is not at his Mac**: capture in Chrome (Claude in Chrome) as below, or in the
     desktop app's built-in browser.
     - Chrome: set the US Edition, decline any cookie banner. Region coordinates are in the screenshot's
       coordinate frame: take one screenshot first and check that its frame equals `innerWidth` x
       `innerHeight`. Chrome's default zoom should be 125% (`devicePixelRatio` 1.25). Move the mouse off
       the region first, or the cursor shows in the crop.
     - Built-in browser: its pane scales the page down (a plain crop came out 460 px wide). Pin the
       cards' grid container (`[class*=tabGridContainer]`) with `position:fixed; left:0; top:0;
       transform:scale(1.72); transform-origin:0 0` at a 1536 px emulated viewport and crop it: 792 px.
     - Hide Bloomberg's terms pop-up with a script; never click Accept. The built-in browser hit a
       Bloomberg bot check on the homepage: never try to solve it.
   - The build pads the pictures to the fixed frames (`dmu-structure.md`, Picture frames). Write the
     capture time in the source log: the tiles and the slide 2 headline must describe the same moment.
2. **Write `content.json`** (format in `build_dmu.py`) by the rules in `references/dmu-writing.md`:
   - `global_headlines`: [0] = the Global write-up copied word for word from the final DMW deck.
     [1] to [4] follow the slots: geopolitics, US macro, US/China/Japan, then Japan (China only when
     nothing relevant on Japan). Each is a tight lead of 2 sentences, **about 40 to 60 words**, every
     figure kept, written in house style, and always on the **latest data release** published before
     the cut-off (Saad, 2 Nov 2026). Reuters first; when Reuters has nothing that fits a slot, use
     another outlet (Bloomberg, CNBC and so on). Pages to check are in the "Where to look" table of
     `dmu-writing.md`. **No source tags on the slide**; put every link in the source log, with the
     original publisher when the article came through an aggregator.
   - `events`: `DD-Mon-YY: <event>: x (P), y (F)` from the calendar file. API/EIA rows without figures.
   - `commodity_headlines`: oil, then gold, each the story of the **last settlement**.
     `commodity_captions`: [oil, gold], the Reuters headline of the story, **on one line at 12 pt**:
     word for word when it fits, trimmed to the main clause when it does not, "on Friday" added on
     Mondays when it fits (`dmu-writing.md`, Captions). It describes the move on the chart's last date
     (on Mondays that is Friday's session).
   - `saudi_headlines`: from the **ANB Rough doc**: one official-data Economic item first, then the top
     Corporate items **in the same order as the day's DMW Corporate section** (Saad, 2 Nov 2026), as
     many as fit (usually 2 or 3), skipping interviews, MoUs, licences, housekeeping and unlisted names,
     **copied in full**, source tag dropped. Depart from the DMW order only for a stated reason that
     day. Zawya and Arab News items are allowed when the news itself is new (`dmu-writing.md`, Slide 4).
   - `results` on results days: rows `[Company, TASI code, BBG code, NP, NP prior, chg SAR mn, chg %,
     consensus, var %, anbc est, var %]` with `-` where nothing applies; names with an anbc estimate
     first, then by net profit, largest first; `footnote` (e.g. `*FY26`) when a row is starred;
     commentary only for names with an anbc estimate; `split_at` when a long list plus commentary needs
     two slides (render first: on 2 Nov 14 rows and 3 paragraphs fitted on one slide despite the FLAG).
     Consensus comes from Bloomberg `EEO <ticker>`; when Saad sends EEO screenshots, tie every
     consensus figure in the Results file to them (Net Income Adj+, quarter estimate column) and
     report any difference. Analyst bullets come from WhatsApp via Saad. On a practice run done after
     the team has sent, Saad may point to the Sales Email (`EMAIL MB YYYYMMDD.docx`): its coverage
     sentences carry the analysts' paragraphs (copied there from the team's DMU); take the paragraph
     and drop the email's own consensus sentence and tag (2 Nov 2026: Bupa, Mouwasat). An analyst
     bullet for a company that is not in today's Results file is asked about, not added on a guess
     (2 Nov 2026: the ELM and CARE bullets belonged to the 1 Nov edition).
   - `target_prices` {data-file name: TP}: from Saad's **EQS export** when he sends one (2 Nov 2026,
     `EQS TASI BEst TP - <date>.xlsx`: Tadawul All Share Index members, column "BEst Target Px:D-1",
     tickers in column A as `XXX AB Equity`). Map each leader and laggard to its Bloomberg ticker
     through `Mapping.xlsx`, round to one decimal (half up), `-` where the export has no value, and
     report any TP that differs from `ANR` or from the team's usual `-` (2 Nov: Sumou 50.5 in EQS,
     `-` in the team's deck; Saad to check `ANR SRE AB`). Without an export, leave TPs blank for the
     team (`ANR <ticker>`).
3. **Build**:
   ```
   python SHARED/scripts/build_dmu.py --template "<template>.pptx" --data "DMW - Data.xlsx" \
     --date YYYY-MM-DD --out "Daily Market Update - YYYYMMDD.pptx" --technicals "Technicals -ENG-YYYYMMDD.pptx" \
     --content content.json --names state/name_overrides_dmu.json --recap recap.png --futures futures.png \
     [--heatmap heatmap.png] [--team-technical "<team's DMU in the day folder>.pptx"] \
     [--forecasts "DMU Commodity Data - <date>.xlsx"] [--results-template "<latest DMU with a Results slide>.pptx"]
   ```
   It finds slides by title (decks run 6 to 8 slides) and refreshes the tables, all three charts
   (data cache rebuilt, ranges extended, links pointed at today's data file), the two pictures (fixed
   frames, padded to shape), the events (date bold), the technical slide (Market Brief joined into one
   paragraph, levels bold) and the Results slide (added, filled, split or removed; source note;
   commentary box moved under the table). It writes `_checks.md`. When a FLAG says "no short name", add the team's
   short name to `state/name_overrides_dmu.json` and rebuild. The DMU keeps its own list: the DMU team shortens more names than the DMW team does on slide 3 (Mouwasat, Naseej, Arab Sea), so the two lists must not be merged (DMW back-test, 21 Oct 2026).
   An INFO saying a caption wraps means it should have been trimmed (step 2): trim it and rebuild.
   **Then replace the technical slide with the team's slide, whole** (all shapes, pictures re-linked), when
   the team's DMU in the day folder has it. 2 Nov 2026 this was done with a small post-build script
   (`post.py`); since 10 Nov 2026 it is `build_dmu.py --team-technical <team's DMU>.pptx`, which copies
   the slide whole and ignores `--technicals`. The heatmap snip goes in with `--heatmap <snip>.png`
   (padded black to the frame's shape; replaced `heat.py`). Tested 10 Nov on the 9 Nov DMU with the
   team's 10 Nov technical slide: shapes and pictures identical to the team's.
   **No content from the template may stay in the deck** (Saad, 2 Nov 2026). Every picture still
   identical to the template's (market recap, heatmap, anything else except the cover logo) is replaced
   with a grey `TO UPDATE: <what>` placeholder of the same size and position, and the same goes for any
   text or table cell left over from the template. Check by hashing every picture against the template
   before handing over. The heatmap is then pasted over its placeholder by Saad.
4. **Keep the evidence**: copy the built file as `Daily Market Update - YYYYMMDD - Claude build.pptx`
   to `WORK/DMW/<date>/` before anyone edits it.
5. **Check**:
   ```
   python SHARED/scripts/check_dmu.py --deck "Daily Market Update - YYYYMMDD.pptx" --date YYYY-MM-DD \
     --data "DMW - Data.xlsx" --prev "<previous DMU>.pptx" --dmw "ANB Daily Market Watch - EN.pptx" \
     --out "WORK/DMW/<date>/DMU check - <date>.md"
   ```
   Fix every FAIL. Read every WARN, which covers lengths outside the team's range. TEAM lines are the
   team's open items. Also confirm in code that each chart's last category and last values equal the
   data file (oil `Sheet1!T:V`, gold `X:Y`, P/E `'Graph PE'!O:S`).
6. Render the deck and look at every slide (fonts from `WORK/DMW Tools/fonts/`); slides 2 and 4 in
   particular, whose last line must clear the FOR INTERNAL USE ONLY footer. When Saad needs a PDF, make
   it with LibreOffice from the deck: its bullets come out blue, as in the team's PowerPoint PDF.
   Exports from Keynote or the Mac's Preview get the bullet colours wrong (2 Nov 2026).
7. Open-file check, then save to the day folder. Hand the user the TEAM list, with the Bloomberg
   terminal steps spelled out (shared terminal over AnyDesk, one user at a time, so allow for a wait).
   When reporting, name every placeholder plainly as "still to update", never as "not updated":
   - **Heatmap (Saad)**: `SASEIDX Index`, then `IMAP`. Group by Sector (GICS), Market Cap All, Pie
     Chart & Table. **Previous session**, not today's live one: the DMU goes out before the Saudi open,
     so use Historical Performance for the last session when snipping later in the day; the "All
     Securities" line should match the TASI DoD on the slide. Move the mouse off the chart so no
     pop-up shows. Snip the pie and sector table, paste over the slide 4 picture at the same size and
     position.
   - **Consensus TPs**: an `EQS` export (Tadawul All Share Index members, "BEst Target Price" added as
     "No condition, display only", plus Ticker), copied to Excel and sent to Claude; or `ANR <ticker>`
     for each leader and laggard, `-` where there is none. In EQS, a new criterion replaces whichever
     line is highlighted: click empty space first (2 Nov 2026: the index line was replaced once).
   - **Results consensus** (results days, if not already in the Results rows): `EEO <ticker>`.
   - A Sukuk slide if one is due.
8. **Next edition**: diff yesterday's Claude build against the team's final DMU and write
   `DMU diff - <date>.md`. Every change the team made becomes a rule in `references/dmu-writing.md`,
   `references/dmu-structure.md` or the scripts, the same day, after checking it against several of
   the team's decks (the rota changes weekly, so one deck is not a rule). Ask Saad before taking
   anything from the team's DMU into a practice deck. A side-by-side page of our bullets against the
   team's (slot, both texts, word counts, our source) helps Saad review; give a one-line reason per
   difference.

## Hand-off check

- [ ] `check_dmu.py`: 0 FAIL; chart last points tie to the data file
- [ ] No picture, text or cell left over from the template (placeholders instead)
- [ ] Technical slide taken whole from the team's DMU (or the Technicals deck plus placeholders)
- [ ] Commentary follows `dmu-writing.md` (house text word for word; agency items 40 to 60 words on the latest data; no source tags; captions one line at 12 pt; slide 4 corporate items in the DMW order)
- [ ] Both Bloomberg pictures in (Saad's screenshots, or captures, or placeholders), stock futures from the Futures page, and the time logged
- [ ] Every slide 2 and 3 link in the source log
- [ ] Claude build copy saved to WORK; TEAM items, with the terminal steps, listed for the user
