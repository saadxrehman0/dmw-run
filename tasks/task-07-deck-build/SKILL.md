---
name: task-07-deck-build
description: Task 7 of the ANB DMW cycle. Builds the ANB Daily Market Watch EN and AR decks (slides 1 to 3 numbers, tables, charts and calendars) from DMW - Data.xlsx, the Graphs decks and the calendar file, starting from the fixed templates. Use for "build the DMW", "deck build", "data is in", "Task 7".
---
# Task 7: DMW deck build (from about 9:00, when the data lands)

Numbers, tables, calendars and charts only. Commentary is Task 8. Late data: deadline 9:30
(Task 0 table).

SHARED = `../../shared/` relative to this skill's base directory. Read
`SHARED/rules/config.md` first (day folder, `WORK` folder, fonts), then
`SHARED/rules/deck-house-rules.md` and `SHARED/rules/deck-safety.md`.

## Prerequisites

- [ ] In today's day folder: `DMW - Data.xlsx`, `DMW Graphs - EN.pptx`, `DMW Graphs - AR.pptx`,
      `Manual File-Calendar.xlsx`. The data file usually lands about 9:00
- [ ] Fixed templates `SHARED/templates/DMW Template - EN.pptx` / `- AR.pptx` and the spec
      `SHARED/data/template_spec.json` (in the skill; since 29 Oct 2026 the previous edition's
      decks are not used)
- [ ] Names sheet: the **state copy** `WORK/DMW Tools/state/Mapping.xlsx` (the live one; the
      skill's `SHARED/data/Mapping.xlsx` can be a rebuild behind; user, 4 Nov 2026)
- [ ] State files in `WORK/DMW Tools/state/` (seed them from `SHARED/data/` on the first run)
- [ ] Technicals decks: optional now (they land about 9:30)

## Steps

1. Stage the inputs, and run the scripts from the plugin's `SHARED/scripts/` folder (or copy the
   whole `SHARED/` folder, so `scripts/` keeps its siblings `data/` and `templates/`: build_dmw.py
   reads `data/font_widths.json`, `data/template_spec.json` and `templates/`).
   Confirm the report date from the `Face` sheet.
   **Calendar Arabic company names (every morning, on anyone's week; pre-approved by the user
   29 Oct 2026):** open `Manual File-Calendar.xlsx` with the deck open-file routine
   (`deck-safety.md`: lock file and recent-save check, latest copy, mtime guard, re-stage to
   confirm), run `python fix_cal_ar_names.py "Manual File-Calendar.xlsx" --names state/Mapping.xlsx
   --report <WORK>/DMW/<date>/cal_ar_names.md` (local files only, no internet), commit the file back with the mtime guard if it changed, and build from the
   corrected file. It changes only Arabic company names that differ from the names sheet on an
   exact English match; rows without a match are listed, not changed. Put its lines in the check
   report and give the user its one-line note for the calendar owner. Any other calendar change
   still goes to the user as a question.
2. EN build (the template defaults to `SHARED/templates/DMW Template - EN.pptx`):
   ```
   python build_dmw.py --data "DMW - Data.xlsx" --graphs "DMW Graphs - EN.pptx" \
     --calendar "Manual File-Calendar.xlsx" --names state/name_overrides.json --chart-state state/chart_state.json \
     [--technicals "Technicals -ENG-YYYYMMDD.pptx"] [--writeup "DMW Writeup - YYYYMMDD.docx"] \
     --date YYYY-MM-DD --out "ANB Daily Market Watch - EN.pptx"
   ```
   The build snaps the template to the layout spec first (`snap_spec.py`; any move is listed in the
   checks) and self-checks the output against the spec before saving. `--writeup` (optional, EN
   only): the Task 4 hand-off docx; the build cross-checks its TASI bullet and calendar tables
   against the data and calendar files and reports CONFLICTs.
   The build removes the template's Results tables on slides 1 and 2 and, in AR, fixes the
   Technicals typos the team always corrects (thousands written 10.132.38, "التقرير.هذا",
   "هاتف: .800", الإتصالات) and the data file's Arabic misspellings on slide 3 (see
   `deck-house-rules.md`).
   AR build, same shape (template defaults to `SHARED/templates/DMW Template - AR.pptx`), with the AR
   graphs, AR technicals, `--glossary state/ar_glossary.json --ar-names state/name_overrides_ar.json`
   and **`--names state/Mapping.xlsx`** (always pass the state copy; the script's default is the
   skill's bundled copy; user, 4 Nov 2026), no `--chart-state`. The AR headline cards get their
   units from the build ("3.4 مليار ريال", "19.0 مليون ريال"; fixed 4 Nov 2026: every AR build
   before that left them out and the team typed them in).
3. Read both `_checks.md` reports. **Every CONFLICT goes to the user as a question**; never pick
   a source or edit the deck to resolve it. FLAG lines are reported as they are. A "Layout ...
   moved back to spec" line means the template itself was off spec: say so, and fix the template
   (`deck-house-rules.md`, Fixed templates). Calendar rows not dated to the report date are
   expected under the empty-day rule; say so.
   With Technicals in both builds, also run `check_final.py --technicals-only` on the two new
   decks (its other checks need a filled deck and stop on the placeholders; 20 Oct 2026):
   it compares slides 4-5 between EN and AR (13 Oct 2026: the two Technicals decks gave
   different intraday highs), so a mismatch reaches the Technicals author before the team
   starts pasting.
4. If the Technicals are not in yet, delete slides 4 and 5 from both decks (snippet in
   `deck-house-rules.md`; they are the template's old pages) and say slides 4-5 are pending.
5. Render slides 1 to 3 of both decks (fonts per `SHARED/rules/config.md`) and look: charts present
   and equal in size, calendars not overlapping, no stale values, slide 3 names not wrapping (a
   wrap FLAG needs a short name from the team; add it to `state/name_overrides.json`, or for an
   Arabic name to `state/name_overrides_ar.json`, then rebuild if nobody has saved the deck). If the team
   edits a short name into a deck, diff their saved deck against the build to spot it and add it
   to `state/name_overrides.json` too.
6. **Re-list the day folder (`device_list_dir`) right before saving** (user, 4 Nov 2026: the
   Technicals landed at 9:30 and the 9:36 save missed them because the folder was listed once at
   the start). If `Technicals -ENG-` / `-AR-` are there now and the build ran without them, stage
   them and either rerun the build with `--technicals` (nobody has saved the deck yet) or patch
   them in with `dmw_patch.py tech` (Task 8 step 5) before the save. Also pick up a newer data or
   calendar file this way (rebuild if the data file changed).
   Then the open-file check, and write both decks to the day folder under the final names. Check
   reports to `WORK/DMW/<date>/`; later changes in the day are appended to them under a
   timestamped "Update" heading. Update `state/chart_state.json`.

## When the team changes the layout on purpose

The template, not the day's deck, carries the layout. If the team resizes or moves something on
purpose, they tell the user and the user tells Claude: edit the template in
`WORK/DMW Tools/templates/`, run `snap_spec.py --write-spec`, write the new
`template_spec.json` to the master copy and rebuild the skill zip. Until then the build and Task 11
put the element back.

## Hand-off check

- [ ] Report date confirmed; all CONFLICT lines asked, FLAG lines reported
- [ ] Layout lines read (template on spec; output self-check OK)
- [ ] AR build run with `--names` pointing at the state `Mapping.xlsx`
- [ ] Day folder re-listed just before saving; Technicals in if they had landed, else old slides 4-5 removed
- [ ] Renders looked at; slide 3 wraps resolved or flagged
- [ ] Decks written only after the open-file check; reports and state saved in `WORK`

Next: Task 8 (on the user's Pasting weeks straight away; on other weeks when the user says the
sequence is final).
