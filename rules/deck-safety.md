# Writing into the day folder safely (open-deck check)

The team starts editing decks and the calendar file within minutes of a build. PowerPoint holds an
unsaved copy in memory, so a file written to disk while someone has it open is silently replaced
when they save. Google Drive does not always sync the `~$` lock file, so a missing lock file does
not prove the file is closed.

**Before every write into the day folder:**

1. List the folder. Look for a `~$` lock file for the file. `device_stage_files` also reports
   `saveBackBlocked: open_in_another_app` when the file is open on the PC: then do not commit;
   prepare the change and ask the user to close the file (5 Nov 2026).
2. Compare the file's modified time with the time you staged it or last wrote it.

Then:

- **Saved since your write, and no change was asked for:** do not overwrite and do not rebuild.
  Stage their copy, diff it against yours if needed, report the differences, and stop.
- **A change was asked for** (news swap, Results table, a bullet): stage the team's latest copy,
  diff it against your last version so their edits are known, apply **only** the requested change
  on top of their copy, and commit with `expectedMtimeMs` set to the staged mtime. Re-list right
  before committing; if it was saved again, re-stage and reapply. Warn the user that anyone with
  the file open must not save over it from an older copy.
  When the team's only edits are ones Claude can reproduce on its own base copies (a slide 3 short
  name, a moved box), apply them to the base copies and refill from there, so a refill never drops
  them (5 Nov 2026: Saad's "GASCO Holding" slide 3 edit carried into the 10:43 decks this way).
- **Never rebuild a deck the team has edited; patch it** (`shared/scripts/dmw_patch.py`).
- If an input (calendar file, Results workbook, data file) has a newer modified time than when
  you staged it, re-stage it before trusting it, and say so.
- If the user says a file is off limits, it stays off limits for the session.

**After every write into the day folder: verify it held, by content.** A commit can report success
and still be undone seconds later by an open copy or a Drive sync, and the modified time can keep
your write's timestamp while the content is the older version (20 Oct 2026: the Results workbook
lost the Mobily row at 10:16 and the Yamama row at 10:20 this way; 5 Nov 2026 10:20: both decks
showed the new mtime but held the 10:15 content, with no lock file visible). So after each commit:

1. Stage the file again from the day folder (wait a second or two first).
2. Compare its sha256 with the file you wrote. A match is the check. If they differ, check that
   your change is in it (the new row, bullet, table or text), never the size or modified time.
3. If it is missing, re-stage the latest copy, reapply only your change, commit again with the
   mtime guard, and check again. Tell the user someone likely has the file open, and ask them to
   have it closed and reopened before anyone saves.
4. Only report the file as updated once the check passes.

Files on a user's PC usually cannot be deleted from a session; ask the user to delete.
