# Manual File-Calendar: process reference

## Where and when

- Built on the evening run for the **next edition**, starting from the previous edition's
  `Manual File-Calendar.xlsx`. Build it in the workspace; after approval save it into the next
  edition's day folder. A Sunday-evening run writes into Monday's folder; a Thursday-evening run
  into Sunday's. Create the folder if needed.
- **Show the user the file first**; save to the work folder only after approval.
- Do not paste the calendars into the ANB Rough doc (pasted tables broke its formatting,
  12 Oct 2026). Task 4 adds them to the Word hand-off file from the approved Excel file.
- Some nights a teammate builds the calendar; the user says so, and the task is skipped.

## What to touch

Only the data rows of the calendar sheet. Two layouts exist:

- **Current (from 10 Oct 2026):** one sheet (`Sheet1`), titles row 1, headers row 2, data from
  row 3. AR dates written as Arabic-digit text (`١٣ أكتوبر ٢٠٢٦`); no-value marker `_`.
- **Older:** sheet `Calender`, headers row 3, data rows 4 to 26, other sheets with charts.

Always copy the previous file's conventions (layout, date form, blank marker). The file carries an
image or charts, so never save it through openpyxl (it drops them): use
`shared/scripts/build_cal.py`, which edits the sheet XML directly, finds the header row, keeps the
previous row styles, and writes AR dates the way the previous file did.

| Block | Cols | Fields |
|---|---|---|
| Event Calendar (EN) | B, C, D | Date, Company, Event |
| Event Calendar (AR) | G, H, I | الحدث, الشركة, التاريخ |
| Economic Calendar (EN) | N, O, P, Q | Date, Event, Prior, Estimation |
| Economic Calendar (AR) | S, T, U, V | التوقعات, الرقم السابق, الحدث, التاريخ |

Percent figures stored as decimals with `0.0%`; K/M/B figures as text (`38.0K`, `7.27M`, `4.1B`).
The two blocks do not share rows: edit each block's columns only, never insert or delete sheet rows.

**Wording drift to watch:** the 12 Oct 2026 file worded the API row "Issuance of API weekly report
on the US crude oil inventories", while the client-mandated deck wording is
`US API Weekly Report on Crude Oil Inventories` (the deck build rewrites it). When the file's
wording and the client wording differ, keep the client wording in the deck and flag it.

## Which dates each block covers

- Both blocks cover the **edition date** of the folder they sit in. Friday and Saturday corporate
  events roll into Sunday's file.
- **Empty-day rule (both calendars):** if the edition day has any item, use only that day's. If
  none, pull in the next day that has items, keep their real dates, and repeat them on that day's
  own edition too.
- **Skip a borrowed day whose only items are 1-star events with no figures** (e.g. "IMF Meetings"
  on a US holiday); take the next day with a 3-star item. 1-star items with figures (bill
  auctions, Dallas Fed) still count.
- **Release dates: Bloomberg wins (client rule; user, 14 Oct 2026).** When Bloomberg, Argaam and
  investing.com disagree on the date of a release, use Bloomberg's date. Example: Saudi
  Arabia's September CPI, which Argaam listed on 14 Oct and Bloomberg on 15 Oct; the team left it
  out of the 14 Oct edition and ran it on 15 Oct. When checking a teammate's file, do not flag a
  row whose date follows Bloomberg.

## Event calendar: sources and inclusion

Sources, in this order (user, 19 Oct 2026). Base URL
`https://www.saudiexchange.sa/wps/portal/saudiexchange/newsandreports/issuer-financial-calendars/`
(`?locale=en`; `locale=ar` for Arabic names). Reachable in the built-in browser or Chrome on the
UAE VPN server (19 Oct 2026); `WebFetch` is blocked.

1. **Dividends: Saudi Exchange Dividends Calendar** (the base page). One row per payout: symbol,
   official long name, announcement date, eligibility date, distribution date, amount per share.
   Set Market Type to Main Market, then Nomu, and the period to cover the edition date(s). It
   lists future payouts too (e.g. Kingdom Holding's quarterly SAR 0.07 to 31 Dec 2026). A row with
   a blank eligibility or distribution date (Ataa, ACWA Power on 19 Oct 2026): take the dates from
   the announcement.
2. **General assemblies: Saudi Exchange General Assembly meetings Calendar**
   (`.../general-assembly-meetings`): long name, type, date, time, venue, convening status, link to
   the announcement. Type `Extraordinary - Corporate Actions` meets the inclusion rule as it
   stands; read the figure from the linked announcement. A plain `Extraordinary` meeting still
   needs its agenda checked (Argaam detail summary or the announcement); `Ordinary` is dropped
   unless its agenda is a corporate action.
3. **Bonus shares, capital changes, rights issues: cross-check against the Saudi Exchange
   Corporate Actions tab** (`.../corporate-actions`): recommendation date, issue type, eligibility
   date, new and previous capital. It is a log of approved actions (latest row on 19 Oct 2026 was
   dated 6 Oct), not a forward calendar, and it uses short names (SAIB, SOLUTIONS): use it to check
   an eligibility row's date and capital, not to find events.
4. **Earnings calls, listing and index events, and anything else: Argaam company calendar**
   (`/en/company/calendar/details/marketid/3`, `/ar/` for Arabic). The exchange tabs do not carry
   these.
5. **Completeness check:** compare the edition date(s) on Argaam with the exchange tabs and list in
   the source log any event only one of them carries. Flag it to the user rather than dropping it.
   After about two weeks of logs, review whether Argaam can be dropped for dividends and
   assemblies.

The Board of directors sessions tab is not used (not checked as of 19 Oct 2026).

Figures always come from the Saudi Exchange. For anything the calendar tabs do not give, **go to
the issuer announcement first** (dividend decision announcements state the per-share amount,
eligibility and distribution dates; the `locale=ar` version gives the exchange's Arabic company
name). The company profile page is a fallback: a bare
`.../company-profile-main/?companySymbol=<code>` link redirected to the exchange home page on
12 Oct 2026.

**Reading the Argaam calendar:** the page text comes back empty to plain page-text reads. Read
the rows from the page itself in the browser: each event is a `div.row` with `div.events`,
`div.eventType`, `div.company` and a `div.detail a[onclick="Detail(this)"]` whose `id` is the
Argaam detail id; the date blocks are headed `Today`, `Tomorrow`, then dated headings. Then
fetch the details from the same page with `fetch('/en/calendar/calendardetail/' + id)`.

Each Argaam row's Details link carries an id: `GET /en/calendar/calendardetail/<id>` returns a
plain-language summary that usually states the figure, the fastest way to see what a bare
"Extraordinary general meeting" is about (its agenda). It is not the source of record; the figure
comes from the Saudi Exchange (31 Aug 2026).

Include: cash dividend eligibility and distribution; bonus-share eligibility; general assemblies
**only when the agenda is a corporate action** (capital hike or reduction, buyback, bonus shares);
rights-issue rows; listing and index events; company earnings calls (not industry events); Nomu
alongside Main Market. A routine general assembly is dropped, not written with a bare label.

Company names: the Dividends and General Assembly tabs already give the official long name. Resolve
Argaam's and the Corporate Actions tab's short names through `TASI Mapping.xlsx` (Short Name
column) to the official long name. Watch brand-vs-legal traps: AVALON PHARMA = Middle East Pharmaceutical
Industries Co. (4016); BUDGET SAUDI = United International Transportation Co. (4260); SFICO =
Saudi Fisheries Co. (6050); APICO = Arabian Plastic Industrial Co. (9548); MUNAWLA = Munawla
Cargo Co. (9571); SABIC (2010) is not SABIC AGRI-NUTRIENTS (2020).

**Refreshing the mapping.** Use the exchange's symbol directory, not profile pages:
`GET /tadawul.eportal.theme.helper/ThemeSearchUtilityServlet` returns JSON for about 1,900
instruments (`symbol`, `companyNameEN`, `companyNameAR`, `tradingNameEn`, `tradingNameAr`, `isin`,
`market_type`: `M` Main Market, `S` Nomu). Cross-check against the live board
(`ourmarkets/main-market-watch`, `ourmarkets/nomuc-market-watch`, table `#marketWatchTable1`)
before adding rows: the servlet keeps companies that migrated between boards, and their old Nomu
rows would duplicate short names (23 such rows excluded on 2 Sep 2026).

**Profile pages** (only `companySymbol` changes):
`.../hidden/company-profile-main/...?companySymbol=<code>` for Main Market,
`.../hidden/company-profile-nomu-parallel/...?companySymbol=<code>` for Nomu. REITs (43xx) render
no `#company_table1`: take the name from the logo `alt` attribute. They carry `#companyDividends`
normally.

**Finding a company's announcements.** The listing page
(`newsandreports/issuer-news/issuer-announcements?locale=en`) is client-rendered, so URL
parameters do not filter it. Set the fields and call the page's own loader in the browser:

```js
$("#symbolId").val("6014");
$("#textSearch").val("capital");   // optional keyword; clear it between companies
ns_Z7_5A602H80O0HTC060SG6UT81DI1_getAnnouncementList(1);
```

then read the `a[href*="anId="]` links. If the site is rebuilt the portlet id changes; recover it
with `Object.keys(window).filter(k => /getAnnouncementList/.test(k))`. One page returns about ten
announcements, so an older item needs the `textSearch` keyword.

Capital figures: post-increase capital as the exchange states it (`mn` to one decimal, `bn` above
1,000 mn); if the EGM invitation only attaches its agenda, the figures are in the earlier board
recommendation announcement. Per-share dividends keep two decimals (`SAR 0.35/sh`).

## Economic calendar: sources

investing.com economic calendar, **United States and Saudi Arabia**, star ratings read from the
**English** site. Pool both countries: take the day's 3-star events; if none, 2-star; if none,
1-star. Prior = *Previous*, Estimation = *Forecast*. Standing weekly row: the API crude inventories
report every Tuesday (investing.com lists it at 01:30 Wednesday; date it Tuesday), worded exactly
`US API Weekly Report on Crude Oil Inventories`. The EIA report arrives as 3-star
`Crude Oil Inventories` on Wednesday. **Word it in the house form, not investing.com's name**
(user, 20 Oct 2026; as in the 14 Oct 2026 final decks): `US Weekly EIA Report on Crude Oil
Inventories` / `الولايات المتحدة – مخزونات النفط الخام الصادرة عن إدارة معلومات الطاقة (EIA) (أسبوعي)`.
If investing.com ever rates it below 3-star, add it by hand on the Wednesday entry the same way.

**Crude inventory rows carry no figures (user, 21 Oct 2026).** Both the API and the EIA rows take
Prior and Estimation as the blank marker (`_` in the file, `-` in the decks), whatever
investing.com shows. The 21 Oct 2026 file carried `-8.6M` / `-0.3M` on the EIA row and the team
removed them. In the DMU events calendar these rows are written without figures:
`21-Oct-26: US Weekly EIA Report on Crude Oil Inventories` (as the 20 Oct DMU wrote the API row).

**Dates by the releasing country's own day (26 Nov 2026, `cal_fetch.py`).** investing.com lists
times in PKT, so a US release after 2:00 PM New York time (FOMC minutes, Treasury auctions, the API
report) shows on the next PKT day. `cal_fetch.py` dates US rows by their US Eastern date and KSA rows
by Riyadh date, which matches the API rule below (dated Tuesday). Claude's default; Saad has not
ruled on it yet.

The release **date** follows Bloomberg when the sources disagree (see "Which dates each block
covers").

**Do not switch to the Bloomberg web calendar** (`bloomberg.com/markets/economic-calendar`;
checked 19 Oct 2026): in Chrome it returned no page text and the tab froze on screenshots (bot
blocking), and the star-tier rules here depend on investing.com's ratings. Read investing.com
through its calendar data service, filtered to US (5) and Saudi Arabia (52), not through the site's
page, which avoids the page clutter. Bloomberg dates come from the team's Terminal when the sources
disagree.

**API after a US Monday holiday:** investing.com moves the API report a day later (Columbus Day,
12 Oct 2026: shown Wednesday 14 Oct). Keep it on Tuesday anyway (user's decision, 12 Oct 2026);
Argaam still lists it on Tuesday. Do not ask again.

**Saudi releases listed only on Argaam (user, 18 Oct 2026):** before building the economic block,
check the Argaam calendar for the edition date for KSA data releases that investing.com does not
carry (GASTAT, SAMA), e.g. GASTAT's quarterly Real Estate Price Index (Argaam lists it under
"Reports", company "General Authority For Statistics"). Add each as the **top row(s)** of the
economic block, whatever the investing.com star tier that day, worded
`KSA Real Estate Price Index` / `السعودية – الرقم القياسي لأسعار العقارات` (EN as in the 20 Jul 2026
deck). Prior and Estimation `-` unless a figure is sourced. The 18 Oct 2026 run missed it.

Number format: **one decimal everywhere, millions included**, half-up, suffix kept: `3.98M` →
`4.0M`, `0.095M` → `0.1M`, `-0.25%` → `-0.3%` (user, 13 Oct 2026: the 13 Oct file went out at
`3.98M` / `3.96M` and nobody caught it). `build_cal.py` writes the file that way, and the deck build
rounds again on read, so an older file still comes out right.

## House wording: follow the final published decks

The wording that counts is what the latest final `ANB Daily Market Watch - EN/AR.pptx` calendar
tables (slide 2) used. Check them before writing a row. **Exception:** the two Arabic dividend
forms and the client-mandated rows below are fixed; they override whatever a deck used.

- EN economic: country prefix (`US CB Consumer Confidence`, `KSA SAMA Monthly Report`,
  `OPEC Meeting`, `US Existing Home Sales`). Drop investing.com's trailing month from the event
  name (`US ISM Manufacturing PMI`, not `... (Sep)`).
- AR economic: country first with an en dash: `الولايات المتحدة – مبيعات المنازل القائمة`,
  `السعودية – النشرة الإحصائية الشهرية لساما`.
- EN events: `Cash Dividend Distribution (SAR 0.50/sh)`, `Eligibility for Cash Dividend (SAR 4.00/sh)`,
  `Eligibility for Bonus Shares (1:5)`, `Extraordinary General Meeting (Capital Hike to SAR 15.00 bn)`,
  `Extraordinary General Meeting (Capital Reduction to SAR 76.52 mn)`. **Capital figures take two
  decimals, mn and bn alike** (every EGM row from 27 Sep 2026 on: 5.18 bn, 401.92 mn, 76.52 mn,
  15.00 bn); this replaces the one-decimal `mn` rule above for calendar rows.
  Earnings calls take the bare label `Earnings call`, no parenthetical; link the company's
  earnings-call announcement.
- AR events: `صرف أرباح نقدية (ريال سعودي 0.50/للسهم)`, `أحقية أرباح نقدية (ريال سعودي 4.00/للسهم)`,
  `أحقية أسهم المنحة (1:5)`, `جمعية عمومية غير عادية (زيادة رأس المال إلى 15.00 مليار ريال سعودي)`,
  `جمعية عمومية غير عادية (تخفيض رأس المال إلى 76.52 مليون ريال سعودي)`: always **ريال سعودي**
  (every EGM row since 27 Sep 2026; early Sep rows wrote ريال alone).
  The two dividend forms are fixed (user's decision, 12 Oct 2026, after a count of the Sep and Oct
  final AR decks: صرف on 32 of 46 distribution rows, أحقية أرباح نقدية (ريال سعودي X/للسهم) the
  most used eligibility form). Do not copy other forms from a deck (توزيع ...، استحقاق ...،
  الأحقية في توزيع ...، أحقية توزيع ...); they match `ar_glossary.json`.
- The file uses the full company name; the team may shorten it in the deck.

## Write-up Word file for the team (Task 4)

Name `DMW Writeup - <YYYYMMDD>` (next edition's date). The previous day's proofread file is the
template; its name varies (`DMW Writeups 20261012 copy.docx`, `DMW Writeup - 20261013.docx`), so
match any `.docx` starting with `DMW Writeup` in that day folder and take the most recently
saved one. Layout like that file: four tables after the TASI section, with no headings above them; each table's first row is
its merged title (`Event Calendar`, `تقويم الأحداث`, `Economic Calendar`, `التقويم الاقتصادي`),
then the header row, the data rows and one empty closing row. Order: EN events, AR events, EN
economic, AR economic. Copy the previous file's tables and replace only the data rows, adding or
removing rows when the count changes (plain style; AR dates as the previous file wrote them,
e.g. `١٣ أكتوبر ٢٠٢٦`; blank marker as in the Excel file). **AutoFit every table to window**
(table width 100%, `tblW type=pct w=5000`).
