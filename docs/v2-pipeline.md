# DMW Template v2 pipeline (build step 2, 25 Nov 2026)

Template v2 (EN 5.14, AR 5.26, CEO-approved 25 Nov 2026) is all native tables with fixed shape
names, so every step below fills named cells instead of moving text boxes. The v1 scripts are kept
as `scripts/legacy_*` for the old decks only.

## Order of a run

| Step | Command | Replaces (v1) |
|---|---|---|
| 1. Build EN and AR | `build_dmw.py --lang en ...` and `build_dmw_ar.py ...` (same flags plus `--ar-names`, `--glossary`) | build_dmw.py, build_dmw_ar.py |
| 2. News, and Results on results days | `fill_commentary.py --news NEWS.json --en EN --ar AR --outdir OUT [--results "DMW - Results.xlsx" --quarter 3Q26 --mapping Mapping.xlsx --universe ... --skip ...]` | fill_commentary.py, place_results.py, insert_ar_news.py, dmw_patch.py arfill |
| 3. Final check | `check_final.py --en --ar --calendar --date [--results --quarter] [--layout ppt_layout.csv] --outdir` | check_final.py, check_results.py |
| 4. PDFs and real layout | `powershell -ExecutionPolicy Bypass -File export_measure.ps1 -Folder <day folder>` | export_pdf.ps1 (still works) |
| 5. Sales email | `build_email_mb.py DAY TEMPLATE Mapping.xlsx OUT [DMU]` (reads v2 decks automatically) | same |

Technicals are appended as slides 4-5 by step 1; `dmw_patch.py tech` is no longer needed for a
fresh build. Step 3 exit code 1 means a FLAG to read before the PDFs.

## News input (NEWS.json)

```json
{"markets":   {"en": ["<Global>", "<TASI>"], "ar": ["<Global>", "<TASI>"]},
 "economic":  [{"en": "...", "ar": "...", "short_en": "(optional)", "short_ar": "(optional)"}],
 "corporate": [{"en": "...", "ar": "..."}]}
```
Corporate items are in rank order. The run writes this file from the edition's Claude Doc (the
doc reader is the next step of the plan). `tests/replay/news_from_v1_deck.py` writes it from a
published v1 edition for replays.

## v2 layout rules the scripts enforce

- Slide 1: Markets (Global, TASI), Economic items, ONE Corporate bullet.
- Slide 2: at most 7 Corporate items at the top of `s2.flow`; the calendars follow in the same table.
- Nothing below 10.60 in on slides 1-2. Slide 1 overflow: the short text of the last Economic item,
  then of both; `fill_commentary.py` exits 3 with the lines to cut when a short text is missing.
  Slide 2 overflow: the lowest-ranked Corporate items leave both decks.
- Results days: the Results table sits under "Major Corporate News" on slide 1 instead of the
  Corporate bullet; the rows that do not fit continue at the top of slide 2 under the same header,
  then the Corporate news, then the calendars. EN and AR split at the same row.
- Index Performance values over 6 characters are flagged. Company names: Mapping long name if it
  fits one line in its cell (HarfBuzz, real font), else the short name; a FLAG if the short wraps.
- The EN date is plain text (the master's date field would show the PC's date).

## Height model

`news_text.py`: 11.0 pt per line plus 4.0 pt between paragraphs (AR 0.5 pt less), with HarfBuzz
line breaks in Arial Nova Cond 9.5 pt (EN) and anb Apercu 9 pt (AR). Fitted on 16 cells of the
CEO finals; on the 5 replayed editions PowerPoint's own layout matched it on all 20 slide bottoms.
