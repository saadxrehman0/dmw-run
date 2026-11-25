# The ANB Rough Google Doc: layout and posting rules

Doc id `1g6xUFKRcqnUITM-5oVeDwXx8Esf8i0vTSVzifq9kGQM`, editable by anyone with the link. Edit it
with the Google Docs connector. If the connector is not on, ask the user to turn it on; never
recreate the doc.

## Layout

- **Draft** tab (`tabId` t.0): Heading 2 sections **Economic News**, **Corporate News**,
  **Global**, **TASI**, in that order. **No calendar tables in the doc** (pasted tables broke its
  formatting, 12 Oct 2026); the calendars go to the team in the Task 4 Word file. There is no
  personal drafts tab (the "Saad's Drafts" tab was deleted on 13 Oct 2026): approved text
  lives in Draft only.
  Each news section has an empty Normal paragraph under its heading: insert items into it so the
  heading keeps its style. One blank paragraph between items.
- Each news item is three paragraphs: English paragraph ending with its source tag, the source
  link on its own line, the Arabic paragraph (right-to-left). The write-ups have no link line.
- **Items are normal text** (10.5 pt EN, link underlined). Only section titles are headings.
  Locate sections by heading text, not by style.
- There is no **Final - Sequenced** tab now. If the team ever brings one back, it is theirs:
  never write into it; ask the user before clearing it.
- The **English - Paste** tab (added 13 Oct 2026) belongs to the sidebar's **Copy for decks**
  button, which rewrites it in full from Draft. **Claude never posts to it, edits it or clears
  it**, and never reads it as a source: Draft is the only source for the deck fill. There is no
  Arabic paste tab any more: the Arabic deck is always filled by Claude (Task 8). If an old
  **Arabic - Paste** tab is still in the doc, the button empties it and marks it as not used;
  the team can delete it.

## Posting rules

- **Post only after the user approves the content.**
- **Guard every write** with `writeControl.requiredRevisionId`, using the position map below
  instead of a fresh read. The team edits live in the morning; if a write is refused, re-read,
  rebuild the map and the requests, and retry.
- If writes keep being refused (the team is typing), send the batch with `targetRevisionId` set
  to the revision you read, then read back and verify every inserted paragraph.
- Every write request carries the Draft tab's `tabId`; every `replaceAllText` is scoped with
  `tabsCriteria` to the Draft tab.
- **Morning additions** go in at their ranked position and are **highlighted yellow, EN and AR**.
  The global write-up is never highlighted. If the team is already sequencing in Draft, still post
  to Draft (the deck fill reads only the Draft tab) and tell the user in one line
  which items are new and where they rank, so the sequencer is not confused.
- **Never clear or overwrite the TASI or Global write-ups** when refreshing news; a teammate may
  have posted the TASI write-up first. **Exception (user, 12 Oct 2026):** on the user's own TASI
  turns, the TASI pair they approved replaces whatever TASI pair is under the Draft tab's TASI
  heading (EN and AR, Arabic right-to-left), because the deck fill reads only the Draft tab.
  An approved pair must never sit anywhere but the Draft tab.
- **Small corrections:** `replaceAllText` with the exact old EN and AR strings, scoped to Draft,
  match strings long enough not to hit the wrong paragraph. Show rewritten wording to the user
  first; plain deletions they ask for can go straight in.
- For corrections, produce **only the changed lines**; never regenerate a full write-up to fix
  one bullet.

## Evening refresh (before posting the evening news)

1. **Back up** the doc with Drive `copy_file` to `WORK/DMW/<YYYY-MM-DD>/`, named
   `ANB Rough backup <date> <time> PKT`.
2. **Compare the Draft tab with the post logs** (the previous run's and any later one for the same
   edition). Anything Claude did not post is someone else's: **ask the user before changing it.**
   With no post log (first run), confirm with the user before clearing.
3. Clear the previous day's Economic and Corporate items, keeping headings and the write-ups.
4. Post the approved items.

## Post log

After every post or edit, write `ANB Rough post log - <date>.md` to `WORK/DMW/<YYYY-MM-DD>/`:
exactly what was posted, changed or cleared, in which section, whether highlighted, and the Arabic
names Claude wrote itself. The next refresh compares against it.

## Highlight convention in the doc

Yellow = added that morning. No highlight = posted the evening before. Red = dropped.

## Position map (29 Oct 2026): write without re-reading

A full `read_doc` of ANB Rough is one of the most expensive steps in a run (every paragraph comes
with its whole style block). `rough_map.py` keeps the Draft tab's paragraphs, their indices and
the doc's revision in `WORK/DMW/<YYYY-MM-DD>/rough_map.json`, so Claude reads the doc only when
someone else has edited it.

1. After any full read (saved to disk by the harness): `python rough_map.py build READ_DOC.json
   --out rough_map.json`.
2. Before a write: `python rough_map.py show rough_map.json "Economic News" "Corporate News"`
   prints each paragraph's start-end index and kind (H heading, E English, L link, A Arabic,
   - blank). Build the requests from those indices.
3. Send `update_doc` with `writeControl.requiredRevisionId` = the map's `revisionId`.
   - **Accepted:** the response carries `writeControl.requiredRevisionId`, the new revision. Save
     the requests to a file and run `python rough_map.py apply rough_map.json requests.json
     --revision <new revision>`. The write landed on exactly the text the map holds, so no
     read-back is needed. Commit the map to the day's `WORK` folder.
   - **Refused** ("does not match the latest revision"): nothing was applied. Someone else edited:
     re-read, `build` the map again, rebuild the requests, retry.
4. A map marked stale (a request it cannot replay, such as a table) needs a fresh read before the
   next write. When a full read happens anyway (Task 4, Task 8), `python rough_map.py verify
   rough_map.json READ_DOC.json` first: MATCH confirms the map, DIFFERENT means someone edited.

Tested on a scratch doc on 29 Oct 2026 (insert, delete, heading style, Arabic paragraph): the map
matched the live doc index for index. A private Claude-only doc was considered instead and kept
as an option if morning refusals turn out frequent.

## Reading the doc cheaply

Read content with Drive `read_file_content` for a quick look. Use the Docs connector's full read
(`read_doc`) when positions or structure are needed: always for posting, for Task 4's Word file and
for Task 8's `fill_commentary.py` (`read_file_content` can also come back blank for tabs).

## Sequencing sidebar

`shared/sidebar/` holds an Apps Script drag-to-reorder sidebar for sequencing news items in the
doc (one person sequences each morning). Its **Copy for decks** button fills the
**English - Paste** tab (top 2 Economic and every Corporate item as English bullets; Global and
TASI are not copied; red-dropped items left out) for a teammate who pastes the English deck
news by hand. There is no Arabic paste route. Install steps are in its README.
