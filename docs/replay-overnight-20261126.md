# Overnight Mac replays: 4, 5, 18 and 23 Nov 2026

Run on Saad's Mac, 25 to 26 Nov 2026, in practice mode, from branch `claude/mac-replays` (PR #7).
Each stage ran as a background agent (Opus, medium effort) from the run folder's own stage prompt,
because the `claude` CLI on the Mac is signed out (**Saad: run `claude`, then `/login`**, before
`run_edition.py` can drive stages itself). Run folders and the full comparisons (`compare.md`) are under
`Claude Outputs/DMW Tools/runs/replay/2026-11-DD/` in the personal Drive. Nothing was written to the work
Drive, no gate was answered, and nothing was sent.

## Coverage against the published editions

| Edition | Economic news | Corporate news | Results rows | TASI bullet | PDFs | Next-day calendar |
|---|---|---|---|---|---|---|
| 4 Nov (results day) | 2 of 2 | 2 of 6 | 16 of 20 | match | 5 + 5 pages | same events; AR date format differed (fixed) |
| 5 Nov (results day) | 1 of 2 (matcher shows 0) | 2 of 5 | 17 of 18 | match | 5 + 5 | 9 Nov rows only; team also listed 8 Nov items |
| 18 Nov | 1 of 2 | 6 of 8 | none due | identical | 5 + 5 | same events and US rows |
| 23 Nov | 2 of 2 | 6 of 8 | none due | match | 5 + 5 | same event and US rows; 8 of 56 cells differ |
| **Total** | **6 of 8** | **16 of 27** | **33 of 38** | | | |

Corporate coverage is much better on 18 and 23 Nov (12 of 16) than on 4 and 5 Nov (4 of 11), because
the Tadawul fetcher and the late-news round came in after the pilot. The Global bullet is left as TK
on every replay: a replay has no in-window Dow or Asia-morning source.

## Mismatches worth a look

- **Capital increases (your call to make).** Triage keeps capital increases out under your 22 Oct
  rule, but the published 5 Nov deck ran SPPC's capital increase and AYYAN's withdrawn capital cut.
  Either the rule is narrower than the code (for example, only bonus-share increases out) or the team
  overrode it that day. I left the rule as it is; tell me which reading is right.
- **Missed Corporate items, by cause:**
  - Tadawul filings with no rule match: GAS contract (18 Nov), ACWA DMTT update and WSM closure (23 Nov),
    ALSAIF tariff (5 Nov).
  - Not in either source by the cut-off: PIF US holdings (18 Nov, a 13F), Aramco East-West, PIF facility,
    Retal/Ajdan, Armah (4 Nov).
- **Missed Economic items:** MISA's 1,200 foreign enterprises (5 Nov), RCU AlUla SAR 10 bn (18 Nov).
- **Results:** 4230 Red Sea (5 Nov) was filed at 12:09 PKT, after the 10:10 cut-off; the live team took
  results later than the run can. Avalon, Nofoth and Salama (4 Nov) were in neither source by the
  cut-off. Consensus is "-" where the team had a figure (YCC 16.8, TCC 0.3): the run has no consensus
  source.
- **Ranks** differ on every day (for example 23 Nov GAS: published 2, run 6). The order is still the
  run's own pick.
- **Calendar:** the 23 Nov run read the US Conference Board prior after its release (91.9 vs the team's
  91.7), a replay artifact. The 5 Nov file missed the 8 Nov items because Argaam's calendar has no past
  view. Arabic wording of US rows differs from the team's.
- **Deck tables (23 Nov):** a dozen numbers differ between the run's v2 tables and the published deck
  (TASI levels on one side, chart labels on the other). Not traced yet.

## Stage times (from each run's own stamps)

| Stage | 4 Nov | 5 Nov | 18 Nov | 23 Nov |
|---|---|---|---|---|
| S0 | done | done | done | done |
| S1 | 15m 43s | 20m 07s | 13m 29s | 10m 32s |
| S2 | 1m 13s | 2m 31s | 1m 23s | 2m 25s |
| S3 | 1m 35s | 1m 11s | 1m 07s | 1m 18s |
| S4 | 14m 00s | 14m 20s | 9m 14s | 7m 32s |
| S5 | 2m 43s | 3m 08s | 3m 15s | 2m 25s |
| S6 | 6m 16s | 8m 16s | 6m 30s | running |

S1 and S4 are the slow stages, and most of their time is the Argaam ID walk. With `--reuse` the 23 Nov
window took 308 s (search 141 s, widen 58 s, bodies 101 s) against 356 to 540 s before; bodies will
drop further once S1 writes the new `no_article` list.

## Fixes in PR #7

- **Mac port:** `export_measure.py` exports PDFs and PowerPoint's own layout through PowerPoint for
  Mac (no file-access prompt; both decks exported first try on all four days). Mac font folders,
  Chrome for fetchers, a Mac how-to. The Windows path still calls `export_measure.ps1` unchanged.
- **Argaam search:** an unreadable probe no longer ends the search (5 Nov S4 returned 0 items twice);
  galloping ID search; `--reuse` of S1's file.
- **Results:** `place_results` reads Bloomberg codes written as `=SUBSTITUTE(...)` (FLYNAS was dropped
  silently) and workbooks without cached values.
- **Checks:** `check_final` and the email builder flag a TK or a bracketed note (TK had reached the
  5 Nov PDFs and email unflagged); the Results header is no longer read as an Economic bullet;
  `check_ar_text` allows TK and credit ratings.
- **Triage:** "fine" only as a word (it matched "Refineries"), credit rating actions, GASTAT's long name.
- **Calendar:** `build_cal.py` no longer wipes the team's wording list (D14:G21) and dates new rows.
- **Shared state stays read-only in replays.** One replay overwrote `state/chart_state`; I restored it
  from the Drive revision before any live use.
- **Comparison tool:** reads v1 Results tables, late news and the Arabic Results header.
- Tests: 55 pass; the new ones fail on the old code.

## Still open

1. A calendar fetcher (`cal_fetch.py`): S6 agents wrote ad-hoc Playwright scripts each time.
2. A source for the Global bullet's Asia and Dow lines, and a consensus source for Results.
3. Triage rules for the missed Tadawul filing types (contracts with government bodies, tax updates,
   branch closures).
4. Replay leak: a later edition's day folder (for example `runs/replay/2026-11-24/day/`) is readable
   by an earlier replay's S6. No agent opened it, but the sandbox should hide it.
5. Smaller: `write_news_doc` has no title option, the Docs export returns base64 only,
   `doc_calls.json` lacks the tab id, number parity flags Arabic ordinals ("Zone 3") and "Tier 2" in
   words.
