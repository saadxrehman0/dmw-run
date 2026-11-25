# DMW deck: house rules (EN and AR)

## Inputs (in the day folder)

| File | Used for |
|---|---|
| Fixed templates `SHARED/templates/DMW Template - EN.pptx` / `- AR.pptx` (in the skill, not the day folder; since 29 Oct 2026) | Start of every build; their frames are the layout spec |
| `DMW - Data.xlsx` (about 9:00) | `PPT-ENG`: slide 1 cards and Index Performance, all slide 3 tables. `PPT-AR`: same for AR. `Face`: date check |
| `DMW Graphs - EN.pptx` / `- AR.pptx` | Cached chart images (slide 1 TASI Performance, P/E & DY; slide 3 Sector YTD). Their tables are stale; ignore them |
| `Manual File-Calendar.xlsx` | Event and economic calendars, EN and AR (Arabic wording from this file, not the glossary). The scripts read by header; check the tables if a column move is flagged |
| `Technicals -ENG-YYYYMMDD.pptx` / `Technicals -AR- YYYYMMDD.pptx` (about 9:30) | Replace slides 4 and 5 wholesale |
| ANB Rough Draft tab | News and write-ups for the commentary boxes (read only) |
| `DMW - Results.xlsx` (results days, after 10:00) | Results Announcements rows |

Confirm the report date from the Technicals title or the `Face` sheet. The previous edition's decks
are no longer read by the build.

**Calendar file missing at build time (5 Nov 2026):** the build takes every row in the file it is
given, whatever the date (it flags them but still places them), so never pass the previous
edition's file as a stand-in. Pass a header-only copy (data rows cleared) so both calendars come
out empty, say so, and rebuild when the file lands if nobody has saved the deck.

## House formatting

- Closes 1dp with commas; rates 2dp; headline closes 2dp. Round **half-up** (0.55 → 0.6).
- Percent 1dp. Negative red `FF0000`, keeping the sign even when it rounds to `-0.0%`. Positive
  green `00B050`, except `0.0%`, which stays default colour.
- bps: integer, negatives in brackets `(13) bps` red; `0 bps` default.
- **Headline DoD box (house rule, approved by Saad 5 Nov 2026):** on both cards in both decks the
  DoD figure sits **0.05 in** from the close, measured ink to ink (the close's last glyph to the
  DoD's first sign or digit, glyph side bearings from `font_widths.json`; the EN leading space on a
  positive change is not counted), and its top is **1.67 in**. The build sizes the DoD box to its
  text with wrap off and 0.03 in slack (a fitted AR box with wrap on dropped the % to a second line).
  `build_dmw.place_dod` (`DOD_GAP_IN`, `DOD_TOP_IN`) places it; `snap_spec.py` checks it and
  re-places it with `--fix` (Task 11).
  **Fonts the measure uses:** EN Segoe UI Bold (`segoe` tables). AR **anb Apercu** (`apercu`,
  `apercu_ink`, read from `anbApercu-Regular.otf`) plus PowerPoint's fake bold, about 0.023 em per
  glyph (`APERCU_FAKE_BOLD`), because the user's PC has anb Apercu and its PDFs draw the AR figures
  in it. Until 5 Nov the AR measure used Calibri Bold (a 4 Nov PDF exported on a PC without Apercu),
  which put the AR TASI % nearly against the close in PowerPoint; Saad saw it as "different from the
  other three". Checked on the 5 Nov PowerPoint PDF: EN exact, AR within 0.005 in on both cards.
  Previews made here render AR in Apercu with LibreOffice's own fake bold, close enough to judge;
  the PowerPoint PDF is the final check.
- Headline value traded: EN `SAR 3.4 bn` / `SAR 19.0 mn`; AR `3.4 مليار ريال` / `19.0 مليون ريال`
  (TASI in billions, NOMU in millions). `build_dmw_ar.py` writes the AR unit (4 Nov 2026).
- Calendar prior/estimate: one decimal, half-up, for numbers, percents and K/M/B text alike
  (`3.98M` → `4.0M`); blanks `-`; EN dates `15-Sep-26`.
- After a Saudi holiday, TASI/NOMU DoD compares against the holiday date (flat 0.0%). Convention,
  not an error.
- Calendar tables (EN and AR): the title bar and the column-label row are **bold** (user,
  13 Oct 2026). `build_dmw_ar.py` sets the AR title bars bold on every build.
- Commentary paragraphs are **justified** in both decks (EN `algn="just"`; AR `algn="just"` with
  `rtl="1"`), as in the published decks (user, 13 Oct 2026).
- AR fonts: anb Apercu. Pass `--arial-arabic` only on a machine without it.
- Client-mandated calendar wording: `US API Weekly Report on Crude Oil Inventories` (never
  `US API Report on Crude Oil Inventories - Weekly`); the Wednesday EIA row is
  `US EIA Crude Oil Inventories`, AR `مخزونات النفط الخام الأمريكية (EIA)` unless the calendar owner's Arabic already names the EIA (10:02 file: "وفقًا لإدارة معلومات الطاقة الأمريكية", kept) (Muryum via Saad,
  25 Nov 2026; Arabic written by Claude). `build_dmw.py` holds the EN forms in
  `CLIENT_ECON_WORDING`; add new client wordings there. `build_dmw_ar.py` sets the AR EIA form
  and writes any Latin inside an AR calendar event ("(EIA)", "(API)") in its own en-US run.
- **AR slide 5 header rule (user, 5 Nov 2026):** the Stock of the Day slide carries a navy line
  under the title (0.40 in, 0.72 in, 7.70 in wide, 1.25 pt, `002060`). The EN Technicals deck has
  it; the AR one keeps arriving without it and the team added it by hand (4 and 5 Nov finals).
  `build_dmw_ar.py` and `dmw_patch.py tech --lang ar` add it when it is missing
  (`ensure_s5_rule`) and log it.

## Fixed templates and layout spec (user, 29 Oct 2026)

The team's position mistakes used to carry over day to day because each build started from the
previous day's deck. Since 29 Oct 2026 Task 7 builds from two fixed templates in the live deck
structure (`SHARED/templates/`), so nothing the team moves in one day's deck reaches the next.

- **Spec:** `SHARED/data/template_spec.json`, read from the templates' own frames: page headers,
  rules, title, date, logo, headline cards (the DoD box's left edge and width are placed from the
  close by the build), the slide 1 and 3 chart frames, the slide 1 Markets header, the slide 2
  Corporate header, the Index Performance table and the slide 3 tables (left, top, width; table
  heights follow their rows). News bodies and the headers and calendars that flow with them are
  `fill_commentary.py`'s, not the spec's.
- **Snap-to-spec (backup):** `snap_spec.py` matches elements by slide and label, not shape id, and
  moves anything off spec back (drift is reverted, not just flagged; user, 29 Oct 2026). It also
  checks the headline DoD gap and top against the house rule above. The builds run it on the
  template and self-check the output; Task 11 runs it on the team's saved decks.
- **How the templates were set (29 Oct 2026):** the 29 Oct published decks, with every fixed element
  at the frame it had most often over 21 to 29 Oct; the slide 3 sector chart at 3.0 x 2.3 in; the
  two headline cards at one top and height; Results tables removed and the slide 1 Corporate box
  restored. The templates keep the smaller logo on each deck (user, 29 Oct).
- **Template changes, 4 Nov 2026:** EN slide 2: the stray empty "Text 22" box under the Corporate
  header deleted and the "(cont.) - paste" box moved to 0.96 in.
- **Template changes, 5 Nov 2026:** (1) DoD boxes at a 1.67 in top and placed by the house gap
  rule on both cards in both templates (re-placed at 11:50 with the corrected AR measure). (2) AR
  card frames from Saad's edited deck: TASI close top 1.57 in, NOMU YTD label and both YTD % boxes.
  (3) **AR slide 2 Corporate group fixed** (the defect logged 1 Nov 2026): `Group 6` ungrouped as
  Saad did in the 5 Nov deck; the Corporate header at 0.846 in (0.23 in tall); the news box at left
  0.456 in, width 7.673 in, top 1.144 in (header bottom plus the EN 0.068 in gap). On a results day
  the slide 2 table now starts at 1.08 in under the header and the news follows it, as Saad laid it
  out. Spec rewritten each time.
- **Template changes, 25 Nov 2026 (Saad's 24 Nov hand layout, made standing; he asked for even
  spacing where his values varied):**
  - EN slide 1: charts flush to the content edges, 0.40 and 4.42 in, both 3.68 in wide (right edge
    8.10). Markets header 6.29 in, 0.175 in tall like the other headers. Markets, Economic and
    Corporate text boxes 0.40 in, 7.70 in wide, each 0.016 in under its header.
  - EN slide 2: news box 0.40 in, 7.70 in wide, top 0.908 in. Both calendars 0.40 in, 7.70 in wide
    (columns scaled; they were 7.73 in wide at 0.397 in).
  - AR slide 1: the Markets and Economic groups ungrouped. All three headers and text boxes 0.40 in,
    7.70 in wide; each text box 0.02 in under its header (Saad had 0.018 / 0.038 / 0.019). Charts
    0.48 in w3.58 and 4.40 in w3.60, both 2.30 in tall (this also cut the P/E chart stretch); chart
    titles 0.48 in w3.67 and 4.40 in w3.69; "أداء المؤشرات" at 6.0 in.
  - AR slide 2: Corporate header and news box 0.40 in, 7.70 in wide, news top 1.09 in; both
    calendars 0.40 in, 7.70 in wide (columns scaled).
  - Slide 3, EN and AR (Saad's 24 Nov alignment, evened): two columns at 0.40 and 4.40 in, every
    table and column header 3.70 in wide (columns scaled), full-width headers 0.40 in, 7.70 in wide.
    Table tops as Saad left them (EN International and Commodity 6.47 in; AR Global and Commodity
    6.524, Regional 8.756, Interest rates 9.045 in). EN footnote and source line at 4.40 in.
    Checked on the 25 Nov data: no slide 3 name wraps at the narrower width.
  - Spec rewritten. Templates in `WORK/DMW Tools/templates/`, skill zip rebuilt.
- **A deliberate change** (the team resizes something on purpose): the team tells the user, the user
  tells Claude, Claude edits the template, runs `snap_spec.py --write-spec`, updates the spec in the
  master copy and rebuilds the skill zip. Templates are kept in `WORK/DMW Tools/templates/` (binary
  files cannot live in the project) and bundled by `make_skill.py`.
- **Results days:** no separate template. The build removes Results tables and Task 10 adds the
  day's with `place_results.py`.
- **Slide 3 standing labels (AR):** the template carries the team's wording ("العالم (MSCI)",
  "S&P 500", "مؤشر ناسداك"); `build_dmw_ar.py` keeps it when the data file has the same characters
  in another order or spacing.
- **Parked:** the 21 Oct table-cell redesign (Saad's drafts of 21 and 29 Oct in
  `Claude Outputs/Workflow Automation/`). Do it only after the team has seen it.

## House spacing rule (EN and AR)

- **Slide 1: one gap, 0.10 in**, from the last ink of the block above to the top of the next
  header bar (`SECTION_GAP` in `build_dmw.py`, the single source). Each text box sits a fixed
  offset under its header, set by the template (EN 0.016 in, AR 0.02 in; 25 Nov 2026).
- **Slide 2 calendars: 0.20 in** above the Event calendar (from the news ink) and above the
  Economic calendar (from the Event calendar ink), EN and AR (`CAL_GAP` in `fill_commentary.py`;
  Saad, 25 Nov 2026, from his 24 Nov hand layout of 0.21 to 0.31 in, evened to 0.20).
- **Slide 1:** Markets text, gap, Economic header, Economic text, gap, Corporate header, Corporate
  text. Corporate fills down to **10.60 in** (`PAGE_BOTTOM`; the page-number band starts at
  10.66 in). Never run into that band.
- **Slide 2:** Corporate continued, calendar gap, Event calendar, calendar gap, Economic calendar,
  top-down. Spare space stays at the bottom. Never pin the calendars to the page bottom.
- **EN and AR break slide 1 Corporate at the same item** (the smaller of the two fits).
- **Overflow:** items move to slide 2; anything still not fitting is left out of both decks
  (lowest ranked first) and listed as a FLAG. Measure the real bottom first: the largest top +
  height among shapes that start above 10.60 in (the page-number box is not content; a false
  10.90 in reading on 5 Nov 2026 nearly cost the RIYAL item).
- **Crude oil inventories rows are never cut for room** (API on Tuesdays, EIA on Wednesdays; user,
  4 Nov 2026). If slide 2 still runs past 10.60 in after the lowest-ranked Corporate items are out,
  keep the crude row, say how far past it runs and ask the user to check the PDF page bottom.
- **Slide 3:** no reflow. The build flags any company name wider than its column; fix with a short
  name in `name_overrides.json` (EN) or `name_overrides_ar.json` (AR) (the team's call: use the full
  name until they supply one). The DMW list is the DMW team's own; the DMU has a separate
  `name_overrides_dmu.json`. Additions: 1 Nov 2026 Gulf General, CGS, Gulf Union Alahlia, Saleh
  Alrashed; 5 Nov 2026 "National Gas & Industrialization" -> "GASCO Holding" (Saad) with the AR
  "شركة الغاز والتصنيع الأهلية" -> "الغاز القابضة" (names sheet Short Name - Arabic). When the user
  shortens an EN name, take the AR short name from the names sheet and add both overrides. Slide 3
  tables are set in Segoe UI: render checks need `segoeui.ttf` / `segoeuib.ttf` from
  `WORK/DMW Tools/fonts/`, or the fallback font shows wraps that are not there.
- **Slide 3 Arabic spelling:** the data file's Arabic names carry Persian letters (ک ی), "الإتصالات",
  "ارامكو" and a detached "و الشخصية". `build_dmw_ar.py` corrects them on every build (`ar_label`).
- **Results days** (settled from the 70 published decks, 15 Jul to 21 Oct 2026; 47 were results
  days): the Results table sits under the slide 1 Corporate header and runs to 10.60 in; slide 1
  carries **no** Corporate item. Run `fill_commentary.py --results`, which puts every Corporate item
  on slide 2. Rows that do not fit on slide 1 continue in a second table under the slide 2 Corporate
  header, above the Corporate news (Task 10 has the detail; `place_results.py` does the placing).
  Different-period rows go last, starred, with the period footnote under the source line, which
  `place_results.py` now writes (EN `*For 2Q27`, AR `* للربع الثاني من عام 2027`; 1 Nov 2026 form
  `*For 1H26` / `* للنصف الأول من عام 2026`). Never add the footnote by hand: the 5 Nov hand edit
  grew the table without moving the news, which then overlapped the table by 0.05 in.
  Results sentences are not added to the deck news unless the user asks (table only).
- **On days without results** slide 1 carried 2 Corporate items on 15 of 23 days and 3 on 5; the
  fit decides, not a fixed count.
- Never hand-tune gaps per day. A different look means changing `SECTION_GAP` / `CAL_GAP` /
  `PAGE_BOTTOM` or the template.
- If the team still moves a header after a fill, diff their saved deck against the build and
  recalibrate `SECTION_GAP` / `CAL_GAP` / `PAGE_BOTTOM`.

## Chart geometry

Frames come from the fixed template (the spec). The two slide 1 charts are the same size and level
(EN 3.68 x 2.341 in; AR 3.58 and 3.60 x 2.30 in, Saad's 24 Nov sizes). The slide 3 sector chart
frame fills its column, 3.705 x 2.3 in (`CHART3_IN`; Saad's 2 Nov call, kept on 25 Nov). The
data team's cached image has drifted to about 1.29 wide to 1 tall, so it stretches about 20%; the
fix is the image size in `DMW Graphs`, not a smaller frame (open item). A cached image with a different aspect ratio is stretched to fill its frame
and the stretch is reported past 2%.

## Technicals missing

If today's Technicals are not in the folder, do not leave the template's old slides 4 and 5 in.
Build slides 1 to 3, then delete slides 4 and 5 from both decks (after the build script, which needs
them to exist):

```python
from pptx import Presentation
p = Presentation(path); lst = p.slides._sldIdLst
for s in list(lst)[3:]:
    p.part.drop_rel(s.rId); lst.remove(s)
p.save(path)
```

Say so in the reply and the check report. When the Technicals land, paste them with Task 8 (rerun
the build for that language with `--technicals` only if nobody has saved the deck since). Task 7
re-lists the day folder right before saving, so Technicals that land during the build go in.

**Technicals with blank cells (1 Nov 2026):** the EN Technicals deck came with the slide 4 TASI
summary table empty while the AR deck had it in full. With the author unreachable, the user had the
EN cells filled from the AR table, then the close checked against the data file. Ask the user
before doing this; never fill from any other source.

## The calendar file in the morning

During the deck build, do not edit `Manual File-Calendar.xlsx`: it is built and approved on the
evening run, and changes to it go through the user. When the owner's file is missing in the
morning and the user says Claude covers it, Task 3 builds it from the sources and it is saved to the
day folder on the user's OK (5 Nov 2026).

**One exception (pre-approved by the user, 29 Oct 2026):** every morning, on anyone's week, Task 7
runs `fix_cal_ar_names.py` before the build. Where an event row's English company name matches the
names sheet exactly and its Arabic company name differs, the script writes the sheet's Arabic long
name into the file. Nothing else changes; rows with no exact match are listed. Log every change and
give the user one line for the calendar owner.

When the calendar file changes after the decks are built (5 Nov 2026: Saad dropped the Saudi Azm
EGM row at 10:09), remove the dropped rows from both decks' slide 2 tables (reduce the frame height
by the row height) before the next fill, so the calendars re-flow.

When slide 2 is too full (a busy results day), propose cutting the lowest-impact US releases (never
the crude rows) in both decks, and ask the user before cutting.

## anb capital's published version (reference only)

anb capital re-exports the EN deck before publishing and sends the user a copy of their PDF only
now and then. Compared on 27 Oct 2026: every figure matched; they cut the last, lowest-ranked slide 2
Corporate item when slide 2 was tight, and the slide 1 date line came out in a lighter font (font
swap). Compared on 1 Nov 2026 with their 29 Oct PDF: every figure matched; two editorial cuts of
Argaam interview material (management-commentary bullets from the middle of the Corporate list, and
a named-spokesperson attribution). Ask the user before treating that as a selection rule. Checked on
4 Nov 2026 with their 3 Nov PDF: same layout; two data-team carry-overs on slide 3 (Nikkei -0.3% at
an unchanged level on a Japan holiday; US natural gas +1.2% at an unchanged 2.9). Their Results table
shows "-" where there is no usable consensus. When a new copy of their PDF arrives, compare it the
same way and add what changed here.

## Conflicts and flags

Every CONFLICT line the build prints goes to the user as a question before delivery: never choose
between sources and never edit the deck to resolve it. FLAG lines are reported as they are.
On the user's Pasting weeks, team text that Claude's checked sources or house rules contradict is
overridden in the decks and listed in one line (Task 8, user 4 Nov 2026); a CONFLICT between two
data sources still goes to the user.

**Slide 3 carry-overs (10 Nov 2026, Saad's call on Claude's recommendation):** when a slide 3 row
shows a non-zero DoD at a level unchanged from the previous edition (a data-team carry-over, usually
a market holiday: 3 Nov Nikkei -0.3% on a Japan holiday, US natural gas +1.2% at 2.9), list it as a
FLAG in the build report with the market and the likely holiday. Do not change the figure; the data
team's file stands unless the user says otherwise.
