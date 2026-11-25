# Argaam: the cheap route (open item 2, user 29 Oct 2026)

Order agreed with the user: (1) RSS feeds, (2) list-page extraction for items the feeds lack,
(3) a helper agent on a cheaper model only if Argaam's layout breaks the script. Never pull a
whole Argaam page into the chat: a list page is about 530 KB of HTML and an article page about
560 KB, while the article body is about 600 characters.

## Lock and bar-chart rows are always skipped (user, 7 Nov 2026)

Skip every Argaam row with a **lock icon** (paywalled, usually labelled "Argaam Special") or a
**bar-chart icon** (financial-report link), whichever route found it. The RSS feeds carry no
icon, so a locked article can come through the feed looking normal: the script now takes the
flags from the list pages and moves flagged items, feed ones included, into `skipped`. A feed
item not seen on any scanned list page comes back `unchecked: true`; check its icon before using
it. Bar-chart rows link to `financial-reports/...` pages, not `articledetail`, so they never
reach `listOnly`. This overrides the earlier note that Argaam's monthly cement sales report is in
scope: when it carries a lock (it did on 5 Nov 2026), it is skipped too.

**Lock is not "Argaam Special" (checked 25 Nov 2026).** The lock icon (`i.ico.lock`, row class
`locked-article`) marks paywalled rows; "Argaam Special" is only Argaam's byline. On the 25 Nov
list page 16 rows were Argaam Special and 8 of them were locked. A locked article page shows a
teaser of about 130 characters inside `.restricted-article` / `.subscription-package-article`;
`argaam_fetch.py` flags those `lock`. Its old `special` flag was taken as a paywall flag and hid
free items (24 Nov replay: foreign ownership, Arab Sea).

## How to run

1. Open any `www.argaam.com` page in the built-in browser (it is signed in as the user).
2. First call in that tab: set `CFG` at the top of `SHARED/browser/argaam_news.js` and paste the
   whole file as the `javascript_tool` text. It returns compact JSON:
   - `feed`: items from the main-news and company-disclosures RSS feeds (`f` M or D), with the
     article ID, headline, feed time and text length (`n` 1 = a report link with no text);
     lock/chart items removed
   - `listOnly`: items on the main-news list page(s) and disclosures page 1 that no feed carries
     (`d` date only, `src`); lock/chart items removed
   - `skipped`: the lock and chart items, for the source log
   - `text`: body text for the IDs in `CFG.text` (from the feed when it has it, else from the
     article page, with its publish time `pub`)
   - `maxId`: log it in the source log; the next run passes it as `sinceId`
3. Later calls in the same tab (no navigation in between) are one line:
   `await argaamNews({sinceId: <last maxId>, text: ["<id>", ...]})`.
4. Triage on headlines first; ask for text only for the items worth reading.
5. Store the result in `window.__r` and return only id + headline lines to the chat.

## Quirks

- Feed `pubDate` says GMT but is not. Main-news times match the article page's own publish time
  (Riyadh). The disclosures feed showed the same article one hour later (08:27 vs 09:27). Use the
  article ID (rises with time) for "new since", and the article page's `pub` when the time matters.
- Feed titles carry a hidden left-to-right mark; the script strips it.
- Feed size on 29 Oct morning: main news 6 items, disclosures 15. The main-news list page 1 held
  50 rows (28 and 29 Oct). From the 28 Oct morning to 29 Oct 11:30 PKT the feeds carried 14
  items and the list pages 15 more: the feeds alone are not enough.
- Main-news list pages 2 and up return no rows to a script fetch (7 Nov 2026); page 1 holds about
  1.5 to 2 days. After a weekend, page 1 plus the disclosures feed reached back to Thursday.
- Report links (dividends, the international market daily report) have no text in the feed.

## Tadawul list endpoint (updated 25 Nov 2026)

The issuer-announcements list page's own AJAX call (`NJgetAnnouncementListData`, POST from the
list page itself, `pageSize` up to 50) **works** again: it returns JSON with the anId, symbol,
trading name, title and filing time (Riyadh time) for each row. `scripts/tadawul_fetch.py` calls
it from a visible browser window and reads each detail page's Element/Explanation table. (On 7 Nov
2026 it returned 404 and the detail pages were walked by `anId`; that is the fallback if it breaks
again.) Hidden (headless) browsers get "Access Denied", and the site needs a Saudi route.

## Still to measure (deep analysis, fresh chat)

Feed coverage of the main-news list over several days (share of the items the DMW used that the
feeds carried, by time of day), and the token cost of each route per run (feeds only, feeds plus
list, list plus text for N items) against the 29 Oct baseline (Argaam page fetches were about 12%
of the morning chat's tokens).
