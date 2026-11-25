# Future update ideas (queue)

Ideas Saad has asked to keep for later. Newest last; each with the date and why.

1. **Real sandbox for the day folder** (25 Nov 2026). The stages' deny list blocks shell deletes
   and moves, but a stage's own Python script could still move or delete files. Today the day
   folder is protected only because the guarded writer (deck-safety.md) is the one script that
   writes there, and only in live mode. Harden it: run the stage sessions as a Windows user with
   read-only rights on the day folder, and let only the finals writer (a separate account or a
   small service) write there.
