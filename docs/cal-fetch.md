# cal_fetch.py: the next edition's calendar rows

Built 26 Nov 2026 to replace the one-off scripts the S6 agents wrote in every Mac replay
(`fetch_sx.py`, `fetch_all.py`, `fetch_inv.py`, `parse_inv.py`, `make_rows.py` under
`Claude Outputs/DMW Tools/runs/replay/2026-11-*/`). It folds their page reads and the rules they
applied by hand into one script; `scripts/cal_fetch.py --help` has the full rules.

    python scripts/cal_fetch.py --edition 2026-11-29 --out cal.json --rows rows.json \
           --names state/Mapping.xlsx --prev-cal "Manual File-Calendar.xlsx"
    python scripts/build_cal.py "Manual File-Calendar.xlsx" new.xlsx rows.json

The fetch is saved next to `--out` as `cal.raw.json`; `--from-raw cal.raw.json` rebuilds the rows
with no browser.

## Tests

`python3 -m pytest tests/calendar -q`. The fixtures are what the replays actually read:

| Fixture | From | Checks |
|---|---|---|
| `raw_20261125.json` + `names_20261125.json` | 24 Nov replay S6 (edition 25 Nov) | Rows equal the agent's hand-made `rows.json`, apart from Derayah (on Argaam only) |
| `raw_20261105.json` + `inv_service_20261105.html` | 4 Nov replay S6 (edition 5 Nov) | East Pipes, Float Glass, Miral rows; jobless claims prior 194K; failed board switch flagged |
| `inv_next_20261125.json` | investing.com `__NEXT_DATA__`, 25 Nov | the fallback parser |

`test_page_scripts_on_stand_in_pages` runs the in-page JavaScript against stand-in pages with
Playwright's own Chromium (skipped when Python Playwright is missing; set `DMW_TEST_CHROMIUM` to a
Chromium binary if Playwright's bundled one is not installed).

## Still to check live (Mac, home router VPN)

Nothing here could reach saudiexchange.sa from the cloud, so these are untested against the live
sites:

1. The board switch and date range on all three tabs (`SET_JS`, the 23 Nov replay's method). The
   4 Nov and 24 Nov reads got Nomu rows for both boards on the assemblies and corporate-actions
   tabs; the script now checks the codes and retries once, then flags it under `review`.
2. Whether the general assembly table is paged (every read returned exactly 6 rows). The script
   reads all rows through the DataTables API when the table is one, and flags a pager otherwise.
3. The dividends range: the script filters the announcement date from half a year back to the
   edition. Check that a payout announced months ago still shows.
4. Assembly announcements fetched from inside the page (`ANN_JS`): the capital and buyback phrases
   were tested on the ACIG invitation (98897) and on made-up text in the 5 Nov Saudi Azm wording.
5. investing.com from Edge on the Windows server (only Chrome on the Mac has been used), and whether
   the data service pages a 10-day US+KSA window (`bind_scroll_handler`; flagged if so).
6. Timing: about 2 minutes expected (six tab reads with 9-second waits).

## Defaults to confirm with Saad

- US rows take their US Eastern date (an FOMC-minutes evening shows after midnight PKT). Written
  into calendar.md as Claude's default.
- Distribution rows first, then eligibility, then assemblies (as the 24 Nov replay ordered them).
- Prior is the figure as first published: investing.com's "revised from" value when the page shows
  one (a replay, or a late build), else "previous".
- The API row's Arabic is the deck form from `ar_glossary.json`; calendar.md gives no calendar-file
  Arabic for it.
