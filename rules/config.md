# Team settings and file locations

Every task skill reads this file first. Change a setting here, not in the task skills.

## Settings

| Setting | Value |
|---|---|
| Time zone | Karachi (PKT, UTC+5). Riyadh is PKT minus 2 hours |
| Day folder (work drive, shared) | `G:\.shortcut-targets-by-id\18rkHxDXGFRnp_kbzgyFhbwCjVH4-XOpc\2026\03. Submission\<MM. Mon>\<YYYYMMDD>\` (e.g. `...\10. Oct\20261012\`). `G:\My Drive\2026.lnk` is a shortcut to it. This is the Windows path; on a Mac the same shared folder sits under Google Drive for desktop, so ask the user for its path once |
| Data team's morning file | **Find it by its contents, not its name** (user, 1 Nov 2026: the team renames it). Any `DMW - *.xlsx` in the day folder, other than `DMW - Results.xlsx` and anything with "format" in its name, whose sheets include `Face`, `PPT-ENG` and `PPT-AR`. Seen as `DMW - Data.xlsx` (to 29 Oct 2026) and `DMW - Main.xlsx` (1 Nov 2026). Wherever a task says `DMW - Data.xlsx`, it means this file. If two files match, take the newest and say which; if none, say so and fall back as Task 0 says. Pass its real path to `--data` |
| ANB Rough Google Doc | id `1g6xUFKRcqnUITM-5oVeDwXx8Esf8i0vTSVzifq9kGQM` (**Draft** tab `t.0`; the **English - Paste** tab belongs to the sidebar) |
| Names sheet (team) | `Mapping.xlsx`, sheet `mapping`: the one names file since 29 Oct 2026 (names EN/AR, code, RIC, Bloomberg, market, sectors EN/AR, ISIN). Re-laid out by the user on 1 Nov 2026 (new headers, columns reordered, sector lookup on an `appendix` sheet); scripts find columns by header through `mapcols.py`. The user's working copy (the live one Claude reads and the monthly check updates) is `WORK/DMW Tools/state/Mapping.xlsx` (`H:\My Drive\Claude Outputs\DMW Tools\state\Mapping.xlsx` on the PC). **The team's copy: `G:\.shortcut-targets-by-id\18rkHxDXGFRnp_kbzgyFhbwCjVH4-XOpc\2026\Mapping.xlsx`** (user, 1 Nov 2026; not yet placed there on that date: it goes there, and to the project file, once the user calls the file final, with the open-file check). The skill carries a copy in `SHARED/data/` and the project holds one as a file; refresh both from the state copy after a change |
| ANB coverage universe | `G:\.shortcut-targets-by-id\18rkHxDXGFRnp_kbzgyFhbwCjVH4-XOpc\2026\ANB Coverage Universe\`: read the **newest** `ANB Coverage Universe Market Cap - <date>.xlsx` (sheets "ANB Universe" and "Values"; 13 Oct 2026 file as of 1 Nov). Used for the Results row order (Task 10) and the DMU |
| Deck templates | `WORK/DMW Tools/templates/DMW Template - EN.pptx` / `- AR.pptx`: the fixed Task 7 templates (source copies; the skill carries them in `SHARED/templates/`). Layout changes go here, then `snap_spec.py --write-spec` and a skill rebuild (`deck-house-rules.md`) |
| PDF export | `WORK/DMW Tools/export_pdf.ps1`, run on the user's Windows PC with PowerPoint after every deck update Claude saves (Task 11, user 1 Nov 2026) |
| Working-files folder | Each user's own personal folder, set once per user (see below). Never the work drive |
| Saudi trading week | Sunday to Thursday |

## Working-files folder (per user)

The firm wants Claude-built working material (drafts, check reports, logs, backups, scratch decks)
kept **off the work drive**, because it clutters it and confuses people. Each user keeps it in
their own folder, called `WORK` in the task skills:

- Ask the user once, at the first run, where their working-files folder is (for example a
  personal Google Drive folder `Claude Outputs`). Remember it for the session.
- Layout: `WORK/DMW/<YYYY-MM-DD>/` for the day's reports and logs, **dated by the edition the work
  is for** (the evening run of 12 Oct writes into `2026-10-13`; user, 12 Oct 2026); `WORK/DMW Tools/state/` for
  the files the build updates (see below).
- If `WORK` cannot be reached, keep the files in the session, send them in the reply, and say
  they still need filing.
- When `WORK` is on a connected drive, save files with `device_commit_files` rather than pushing
  them through the Drive connector as encoded text. If a write on the Mac fails with "Resource
  deadlock avoided" (seen 6 Oct 2026 on a folder first created through the Drive connector), save
  small text files with Drive `create_file` into the same folder instead and say so.

## Finding files in the day folder on a Mac (cloud-only files)

Google Drive for desktop streams the shared drive, so most day-folder files are **cloud-only**
until opened. The `device_bash` shell does not list them at all (`ls` showed an empty `20261028`
folder while `Manual File-Calendar.xlsx` sat in it, 27 Oct 2026), and python-pptx or openpyxl
fails on them in the shell ("Package not found").

- Check whether a file exists with `device_list_dir` (it lists cloud-only files and marks them
  `cloudOnly`), never with `ls`, `find` or a script in `device_bash`.
- Read a cloud-only file by staging it with `device_stage_files` (it downloads the file) and
  working on the staged copy in the workspace.
- Never tell the user a file is missing, or ask them about it, on the strength of a shell listing.

## Files the build updates (state)

The plugin ships seed copies in `shared/data/`. The build changes some of them day to day, so the
live copies sit in `WORK/DMW Tools/state/`:

| File | What changes it |
|---|---|
| `chart_state.json` | Task 7, every build (chart image fingerprints for the stale-chart check) |
| `name_overrides.json` | Task 7, when the team supplies a short company name (EN slide 3) |
| `name_overrides_ar.json` | Task 7, when the team supplies a short Arabic name (AR slide 3, full name -> short name; added 13 Oct 2026) |
| `name_overrides_dmu.json` | Task 9, when the DMU team uses a short company name. Kept apart from the DMW list because the two teams shorten different names (21 Oct 2026) |
| `ar_glossary.json` | Tasks 1 and 6, when a proofread Arabic name is confirmed |
| `schedule.json` | Task 0, when the user gives a new week's schedule (their DMW/DMU jobs per week; added 29 Oct 2026) |
| `names_check.json` | Task 0, after each monthly Saudi Exchange check of `Mapping.xlsx` (`{"last": "YYYY-MM-DD"}`; first check 29 Oct 2026). The November 2026 check was skipped by the user (1 Nov); the next is due on the first daily run of December |
| `Mapping.xlsx` | Task 0 monthly names check (after the user approves the proposed copy), or a names fix the user approves. When it changes, also refresh the project file and the skill's `SHARED/data/` copy (rebuild), and tell the user the team's shared copy needs the same change |

On the first run, copy `shared/data/chart_state.seed.json` to `state/chart_state.json`, and copy
the others from `shared/data/`. Later runs use the `state/` copies. Send additions back to
whoever maintains the plugin so the next version carries them for the whole team.

## Where finished files go

| What | Where |
|---|---|
| The two DMW decks under the final names `ANB Daily Market Watch - EN.pptx` / `- AR.pptx` and their PDFs, `Daily Market Update - YYYYMMDD.pptx`, `EMAIL MB YYYYMMDD.docx`, `Manual File-Calendar.xlsx` (next edition) | The day folder, because the team edits them there |
| Check reports, post logs, doc backups, drafts, any deck under a non-final name | `WORK/DMW/<YYYY-MM-DD>/` |

The team's working files (the two DMW decks, the DMU, the EMAIL MB) are written to the day folder
as each task finishes, after its checks, because the team edits them there. Anything else final
(including the next edition's calendar file and the team's `Mapping.xlsx`) goes there only after
the user approves it.

## Fonts (not shipped in the plugin)

The deck fonts are licensed, so the plugin does not carry them. Stage them from the user's PC
when a task needs to render a deck:

- **First place to look: `WORK/DMW Tools/fonts/`**, if the user keeps a copy there (set up
  12 Oct 2026; holds `ArialNovaCond.ttf`, `ArialNovaCond-Bold.ttf`, `anbApercu-Regular.otf`,
  `ApercuArabic-Regular-Pro.otf`). This works from any machine, including the Mac when the PC is
  not at hand. If the folder is missing or a file is absent, fall back to the PC locations below.
- EN deck: Arial Nova Cond (`ArialNovaCond.ttf`, usually `C:\Windows\Fonts\` or the user's
  `AppData\Local\Microsoft\Windows\Fonts\`). Install only the **measuring copy** built by
  `make_measure_font.py` (kerning off, +0.25% width); it matched PowerPoint's line breaks on 43 of
  43 published boxes. Never install the unmodified file alongside it, and never use a proxy font.
- AR deck: anb Apercu and Apercu Arabic (`anbApercu-Regular.otf`, `ApercuArabic-Regular-Pro.otf`).
  Copy them to `~/.fonts` as they are, and also next to `build_dmw_ar.py` in a `fonts/` folder
  (the AR name-width check reads `fonts/ApercuArabic-Regular-Pro.otf` beside the script).

Then `fc-cache -f` and confirm with `fc-match -f "%{file}\n" "Arial Nova Cond"` that the
measuring copy is the one in use. If the fonts cannot be staged, say the layout cannot be measured
and stop before any step that relies on it.

## Approvals

Nothing goes into the ANB Rough doc, the day folder or an email without the user's approval of the
content. When two sources disagree (write-up vs data file vs calendar file), ask the user; never
pick one silently.

## Keeping runs cheap

- Ask for Argaam links, not PDFs: a fetched link returns compact article text, a printed PDF
  carries menus and ads. Read page text, not screenshots. Go straight to known URLs.
- **Argaam news (29 Oct 2026):** use `SHARED/browser/argaam_news.js` in the built-in browser
  (RSS feeds first, then list pages for what the feeds lack, body text only for chosen items),
  as set out in `SHARED/rules/argaam.md`. Never `get_page_text` an Argaam list or article page.
- **Saudi Exchange names (29 Oct 2026):** `SHARED/browser/sx_names.js` takes a short digest of
  `Mapping.xlsx` and returns only the differences (Task 0). Never return the whole directory.
- **ANB Rough:** write with the position map (`rough_map.py`, `anb-rough-doc.md`), not a full
  read before every write.

## Writing in every task

- Never write em-dashes in anything you say or produce.
- Never fabricate a figure. An unsourced number is written `TK`.
- **Read figures from the source's own cells, then cross-check (user, 3 Nov 2026; every task,
  every source).** When a figure comes from a filing, web page, PDF or screenshot, read it from
  the table's cells (the page's `td` elements, a PDF table extractor, or the image cell by cell),
  never by splitting run-together page text by eye. Then check it against a figure the source
  itself gives (its % change, total, subtotal or prior period) before it goes into any output; a
  mismatch stops it. 3 Nov 2026: BCI's Tadawul page text read "15.489.857.959..." and was split as
  15.48 / 9.85; the cells were 15.48 / 9.8 / 57.959%, and 15.48 / 9.85 - 1 = 57.2% would have
  failed against the filing's 57.959%.
- **Times come from a clock, never an estimate (3 Nov 2026).** State a time only from the clock
  tool or from a file's save time converted to PKT; on 3 Nov 2026 estimated times ran 5 to 13
  minutes fast.
- Keep a source log for anything taken from the web: `WORK/DMW/<date>/Source log - <date>.md`.
