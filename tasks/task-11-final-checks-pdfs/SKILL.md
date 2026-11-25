---
name: task-11-final-checks-pdfs
description: Task 11 of the ANB DMW cycle. Runs the pre-PDF check on the team's saved EN and AR DMW decks (EN/AR figure parity, calendar wording, slide 1 spacing, layout drift against the template spec), refreshes calendar numbers changed in the morning, and exports the PDFs with PowerPoint on the user's PC or Mac usually 10:10 to 10:20 Karachi, after the last deck change. Use for "final check", "pre-PDF check", "check the decks", "PDFs", "Task 11".
---

# Task 11: Final checks and PDFs (after the last deck change, usually 10:10 to 10:20)

> Times are usual windows, not deadlines (per-task cut-offs dropped 4 Nov 2026; late-input deadlines in Task 0).

SHARED = `../../shared/` relative to this skill's base directory. Read
`SHARED/rules/config.md` first (day folder, `WORK` folder, fonts), then `SHARED/rules/deck-safety.md`.

## Prerequisites

- [ ] Both decks saved by the team after their review
- [ ] Current `Manual File-Calendar.xlsx` (the team updates some numbers in the morning)
- [ ] On results days, the day's `DMW - Results.xlsx`
- [ ] Fonts staged and the Arial Nova Cond measuring copy installed (`SHARED/rules/config.md`);
      without them run with `--no-render` and say the spacing was not checked

## Steps

1. Stage the two saved decks, the calendar file and (results days) `DMW - Results.xlsx` (note mtimes).
2. **Calendar refresh:** re-read the event and economic calendar numbers from the file; report any
   that changed since the build as FLAGs, and patch them into both decks only when the user asks
   (patch the team's latest copy, never rebuild). Rows the user or a reviewer has removed from the
   decks by hand (25 Nov 2026: three economic rows Muryum cut) stay removed even though the file
   still carries them; do not flag them back in.
3. Run (about 30 seconds, read only):
   `python SHARED/scripts/check_final.py --en "ANB Daily Market Watch - EN.pptx" --ar "ANB Daily Market Watch - AR.pptx" --calendar "Manual File-Calendar.xlsx" --date YYYY-MM-DD --outdir <out>`
   It pairs EN and AR commentary paragraphs and compares every number and ticker, checks every
   calendar row against the file, compares the Technicals figures on slides 4-5, and measures
   the slide 1 gaps (0.10 in) and where the content ends on slide 1 and slide 2 (10.60 in limit;
   slide 2 matters most on results days, when all Corporate news sits there).
   When it reports a different paragraph count for EN and AR, name the unpaired paragraph (29 Oct
   2026: an Alqemam contract item sat in the AR deck only).
   Run it again after any Task 10 patch that lands after this check (late news or Results), and
   after a teammate finishes pasting the English on a Task 8 teammate route.
4. **Results days (29 Oct 2026):** the team fills the Results tables by hand, so check them every
   time, read only:
   `python SHARED/scripts/check_results.py "ANB Daily Market Watch - EN.pptx" "ANB Daily Market Watch - AR.pptx" "DMW - Results.xlsx" SHARED/data/Mapping.xlsx`
   It matches every row by Bloomberg ticker across the EN deck, the AR deck and the Results file,
   and flags rows on one side only, TASI code differences and any figure that differs after 1dp
   rounding (misses so far: BAHRI and AMERICANA consensus, 28 Oct; SAUDI CABLE missing from the
   AR deck and ARMAH under its old Nomu code 9590 instead of 6022, 29 Oct).
5. **Layout drift (29 Oct 2026):** on each saved deck run
   `python SHARED/scripts/snap_spec.py "ANB Daily Market Watch - EN.pptx" --lang EN` (and `--lang AR`).
   It lists every fixed element (headers, logo, cards, charts, slide 3 tables) the team moved away
   from the template spec. The user approved reverting drift: after the open-file check
   (`deck-safety.md`), rerun with `--fix` on the team's latest copy and say which elements went
   back. If a deck is open, send the list instead so the team undoes it before export. A move the
   team says is deliberate goes to the user (template change, Task 7).
6. **Send every FLAG to the user at once**, in a few lines, so the team fixes them before export.
   Apart from the drift revert in step 5, do not edit the team's decks unless asked. The checks
   compare figures and tickers, not wording, so a dropped word still needs a read.
7. **PDF export (user, 1 Nov 2026): after every deck update Claude saves to the day folder, export
   both PDFs, without asking.** PowerPoint does it, never LibreOffice, so the PDFs match a manual
   export. Use whichever machine the session is linked to:
   - **Windows PC:** Desktop Commander `start_process` (PowerShell):
     `powershell -NoProfile -ExecutionPolicy Bypass -File "H:\My Drive\Claude Outputs\DMW Tools\export_pdf.ps1" -Folder "<day folder>"`.
     The script opens `ANB Daily Market Watch - EN.pptx` and `- AR.pptx` read-only with no window,
     saves each as PDF over the old one in the day folder (`SaveAs`, format 32) and prints the slide
     count, size and save time. Producer "Microsoft PowerPoint 2021".
   - **Mac (25 Nov 2026, first run 10:00, about 30 s for both):** Desktop Commander
     `start_process` (zsh), PowerPoint for Mac via AppleScript, one call per deck:
     `osascript -e 'on run argv' -e 'tell application "Microsoft PowerPoint"' -e 'open (POSIX file (item 1 of argv)) with read only' -e 'set p to active presentation' -e 'save p in (POSIX file (item 2 of argv)) as save as PDF' -e 'close p saving no' -e 'end tell' -e 'end run' "<deck>.pptx" "<deck>.pdf"`
     with the day folder under `~/Library/CloudStorage/GoogleDrive-<work account>/.shortcut-targets-by-id/...`.
     Producer "Quartz PDFContext". The Mac has Arial Nova Cond, Segoe UI, anb Apercu and Apercu
     Arabic installed; its PDFs carry the real anb Apercu (the 24 Nov PC PDFs showed Calibri in its
     place, so that PC may lack the font). Arabic text renders in Arial on both machines.
   - Either way it exports the **saved** file: anyone's unsaved edits are not in it.
   - Then stage both PDFs and check: newer than their decks, 5 pages each, the expected fonts
     (`pdffonts`), the latest change is in the text (e.g. the newest Results row or calendar edit),
     and look at a render of slides 1-2.
   - If the machine cannot be reached or PowerPoint fails, say so and ask the team to export by hand.
   Save the check output as `WORK/DMW/<date>/DMW pre-PDF check - <date>.md`.

## Hand-off check

- [ ] Calendar numbers re-read; changes flagged (patched only on request)
- [ ] Results days: `check_results.py` run on both decks against `DMW - Results.xlsx`
- [ ] Layout drift checked on both decks; reverted after the open-file check, or listed for the team
- [ ] Every FLAG sent before export
- [ ] PDFs exported by PowerPoint (PC script or Mac AppleScript) after the last deck update, newer than the decks, opened and looked at
