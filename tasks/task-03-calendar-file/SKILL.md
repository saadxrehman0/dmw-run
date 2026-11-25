---
name: task-03-calendar-file
description: Task 3 of the ANB DMW cycle. Builds the next edition's Manual File-Calendar.xlsx (event and economic calendars, English and Arabic) from the Saudi Exchange issuer financial calendars, Argaam and investing.com, and saves it to the next edition's day folder after approval. The calendars go to the team in the Task 4 Word file, not the Google Doc. Use for "calendar", "Manual File-Calendar", "event calendar", "economic calendar", "Task 3".
---

# Task 3: Manual File-Calendar (evening, next edition)

SHARED = `../../shared/` relative to this skill's base directory. Read
`SHARED/rules/config.md` first (day folder, `WORK` folder, approvals), then
`SHARED/rules/calendar.md` in full before building; it governs.

## Prerequisites

- [ ] The user has not said a teammate is building the calendar tonight (if so, skip this task)
- [ ] The previous edition's `Manual File-Calendar.xlsx` (copy source) and the latest final DMW
      decks (wording source) are reachable in the day folders
- [ ] Browser available (Saudi Exchange issuer financial calendars and announcements, Argaam
      calendar, investing.com)

## Morning cover (5 Nov 2026)

When the owner's file is missing from the day folder in the morning and the user says Claude covers
it, build it for the **same day's** edition with the steps below (about 10 minutes), show the rows,
and save it on the user's OK. Meanwhile Task 7 builds with a header-only stand-in and rebuilds
when the file is ready (`deck-house-rules.md`). Once the user edits the saved file, diff it against
Claude's version and patch the decks for dropped or changed rows (5 Nov: the Saudi Azm EGM row was
dropped at 10:09).

## Steps

1. Work out the edition date (Sunday's run → Monday's folder; Thursday's → Sunday's) and which days
   each block covers (empty-day rules, 1-star skip rule).
2. **Fetch with `cal_fetch.py` first (26 Nov 2026).** One call reads the Saudi Exchange dividends,
   general assembly and corporate actions tabs (Main and Nomu), the assembly announcements and the
   investing.com US and KSA calendar in a visible Chrome (Mac) or Edge (Windows) window, applies the
   rules below and writes `rows.json` for step 5 plus `cal.json` (`sources` = source log lines,
   `review` = what is left to settle, `dropped` = routine assemblies left out):
   `python SHARED/scripts/cal_fetch.py --edition <YYYY-MM-DD> --out cal.json --rows rows.json
   --names state/Mapping.xlsx --prev-cal "<previous Manual File-Calendar.xlsx>"`
   (`--after` defaults to the previous trading day; `--from-raw` rebuilds without a browser). It does
   not read Argaam: earnings calls, listing and index events and GASTAT/SAMA releases are still added
   by hand, as below. The steps below are what the script does, and the manual route if it fails.
   **Event rows:** follow the source order in `calendar.md` ("Event calendar: sources and
   inclusion"): Saudi Exchange Dividends Calendar for dividends; its General Assembly meetings
   Calendar for assemblies (`Extraordinary - Corporate Actions` type; plain extraordinary meetings
   get an agenda check); its Corporate Actions tab as a cross-check for bonus, capital and rights
   rows; Argaam for earnings calls, listing and index events. Compare Argaam with the exchange tabs
   for the edition date(s) and log any event only one side has. Resolve short names through the
   names sheet (state `Mapping.xlsx`, sheet `mapping`). Never fabricate; unsourced = `TK`.
   **What worked on 5 Nov 2026 (built-in browser):**
   - Dividends page: the table is `#issuerTable` (about 1,400 Main rows). Switch board by setting
     `#marketsList` to `M` or `S` (Nomu) and triggering `change` with jQuery, and set `#drp_end`
     past the edition date before clicking the search button; then keep only rows whose text holds
     the edition date (`2026-11-05`), eligibility or distribution.
   - General Assembly page: same `#marketsList`; each row links its announcement (`anId=`). Fetch
     the announcement HTML from the same origin and read the "General Assembly Meeting Agenda" text
     to decide a plain Extraordinary meeting (5 Nov: Saudi Azm item 2, buyback of up to 3.6 mn
     shares, counted as a corporate action).
   - **Argaam failed in the built-in browser** (ERR_CERT_COMMON_NAME_INVALID, 5 Nov 2026); use
     Chrome. The calendar rows read as `div.row` with `div.events`, `div.eventType`,
     `div.company`, `div.detail a` (id). The detail endpoint answered for one id and returned
     "[BLOCKED: Cookie/query string data]" for others in Chrome; the row list alone was enough.
3. **Economic rows:** investing.com calendar data service, US and Saudi Arabia, star pooling;
   standing API row on Tuesdays. Prior = Previous, Estimation = Forecast. From any
   `www.investing.com/economic-calendar/` page (built-in browser, 5 Nov 2026):
   ```js
   const fd=new URLSearchParams();['5','52'].forEach(c=>fd.append('country[]',c));
   fd.append('dateFrom','YYYY-MM-DD');fd.append('dateTo','YYYY-MM-DD');fd.append('timeZone','77');
   fd.append('timeFilter','timeOnly');fd.append('currentTab','custom');fd.append('limit_from','0');
   const j=await (await fetch('/economic-calendar/Service/getCalendarFilteredData',{method:'POST',
     headers:{'X-Requested-With':'XMLHttpRequest','Content-Type':'application/x-www-form-urlencoded'},
     body:fd.toString()})).json();
   // parse j.data rows: td.flagCur, count of td.sentiment i.grayFullBullishIcon = stars,
   // td.event, td.act, td.fore, td.prev; tr data-event-datetime (timeZone 77 = PKT)
   ```
   Return only the parsed lines, never the HTML. Also check the Argaam calendar for KSA releases
   investing.com lacks (GASTAT, SAMA).
4. **Wording:** copy how the latest final EN and AR decks worded the same rows (slide 2 tables),
   except the fixed Arabic dividend forms in `calendar.md`, which override the decks.
   Client-mandated rows keep the client wording. Open (5 Nov 2026): the team's files and decks this
   week used "Earnings Call" / مؤتمر الأرباح and "(0.27 ريال سعودي/سهم)"; calendar.md fixes
   "(ريال سعودي 0.27/للسهم)". Follow calendar.md until the user settles it.
   A buyback EGM with no house form yet (5 Nov 2026): EN `Extraordinary General Meeting (Share
   Buyback of up to 3.60 mn shares)`, AR (Claude's wording, name it to the user)
   `جمعية عمومية غير عادية (شراء ما يصل إلى 3.60 مليون سهم من أسهم الشركة)`.
5. Write `rows.json` and run
   `python SHARED/scripts/build_cal.py <previous file> <new file> rows.json`
   with `blank` set to the previous file's no-value marker. It keeps the file's images and charts.
   Then run `python SHARED/scripts/fix_cal_ar_names.py <new file> --names state/Mapping.xlsx --report <WORK>/DMW/<date>/cal_ar_names.md`.
   Include its changes and its no-match list in what you show the user. Run `check_ar_text.py` on
   the Arabic you wrote.
6. **Show the user the rows** (EN and AR) with the source of each figure, and keep the source log
   (`WORK/DMW/<date>/Source log - <date>.md`). Save a draft copy in `WORK/DMW/<date>/`; save into
   the edition's day folder only after approval (create the folder if needed; open-file check in
   `SHARED/rules/deck-safety.md`, content check after the save).
7. Do not paste the calendars into the ANB Rough doc. Task 4 fills the Word hand-off file's
   calendar tables from this approved Excel file.

## Hand-off check

- [ ] Right edition folder and dates; empty-day and 1-star rules applied
- [ ] Every figure traced to the exchange or investing.com; short names resolved to long names
- [ ] Argaam vs Saudi Exchange comparison done for the edition date(s); one-sided events flagged
- [ ] Wording matches the latest final decks; fixed dividend forms and client wording kept
- [ ] File built with `build_cal.py` (images and charts intact); user approved before saving
- [ ] Nothing pasted into the ANB Rough doc
