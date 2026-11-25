---
name: task-04-writeup-handoff
description: Task 4 of the ANB DMW cycle. Exports the approved ANB Rough content to the team's "DMW Writeup - YYYYMMDD" Word file with the event and economic calendars (EN and AR) at the bottom, laid out like the previous proofread file, for the proofreader and the team's WhatsApp. Use for "writeup file", "hand-off docx", "export ANB Rough", "Task 4".
---

# Task 4: Write-up hand-off (evening)

SHARED = `../../shared/` relative to this skill's base directory. Read
`SHARED/rules/config.md` first (day folder, `WORK` folder, approvals), then the "Write-up Word
file" section of `SHARED/rules/calendar.md`.

## Prerequisites

- [ ] Tasks 1 to 3 approved (news and TASI write-up posted; the next edition's
      `Manual File-Calendar.xlsx` saved in its day folder)
- [ ] The previous day's proofread write-up in its day folder as the layout template. Its name
      varies with who saved it: match any `.docx` whose name starts with `DMW Writeup`
      (`DMW Writeups 20261012 copy.docx` and `DMW Writeup - 20261013.docx` both count). If
      several match, take the most recently saved one and tell the user which file was used

## Steps

1. Read the ANB Rough doc with the Docs connector (`read_doc`; the result is saved to a file,
   parse it with a script). Do not rely on the Drive .docx export: it comes back as inline base64
   that cannot be saved to disk (12 Oct 2026). Build the Word file on the previous day's
   proofread write-up as template: remove everything above its first calendar table and write
   the sections back in its style (Heading 2 section titles; EN paragraph, blank, link as a
   hyperlink, blank, AR paragraph right-to-left with `w:bidi` and Latin tickers in their own
   non-RTL runs).
   - **Economic News / Corporate News:** from the Draft tab as it stands at build time.
   - **Global:** leave the section empty. On the evening run the Draft tab's Global paragraph is
     still the current edition's; the next edition's Global is written in the morning.
   - **TASI:** from the Draft tab. On the user's own TASI turns the pair they approved should
     already be there (Task 2 replaces any other version); if the Draft tab still holds a
     different pair, use the approved one and say so.
2. Keep the previous file's four calendar tables after the TASI section, laid out as the
   "Write-up Word file" section of `SHARED/rules/calendar.md` describes (no headings; each
   table's first row is its merged title). Fill the data rows from the approved
   `Manual File-Calendar.xlsx` in the next edition's day folder (the calendars are no longer
   pasted into the Google Doc), adding or removing rows when the count changes. Plain style, no
   added shading; AR dates and the blank marker as in the Excel file. AutoFit every table to
   window (`tblW type=pct w=5000`).
   Then justify every paragraph above the first calendar table (news, Global and TASI sections,
   headings and blank lines included, EN and AR: `w:jc w:val="both"`; user, 27 Oct 2026). The
   template's EN paragraphs carry no alignment, so this has to be set, not inherited.
3. Name it `DMW Writeup - <YYYYMMDD>` (the next edition's date). Open it and check it renders.
4. Send it to the user (the user posts it to the team's WhatsApp). Keep a copy in
   `WORK/DMW/<date>/`.

## Hand-off check

- [ ] News matches the Draft tab exactly (EN and AR); Global empty; TASI is the approved pair
- [ ] Calendar rows match the approved Excel file; previous file's layout; tables fit the page width
- [ ] Every paragraph above the calendar tables justified
- [ ] Named with the next edition's date
