# Future update ideas (queue)

Ideas Saad has asked to keep for later. Newest last; each with the date and why.

1. **Real sandbox for the day folder** (25 Nov 2026). The stages' deny list blocks shell deletes
   and moves, but a stage's own Python script could still move or delete files. Today the day
   folder is protected only because the guarded writer (deck-safety.md) is the one script that
   writes there, and only in live mode. Harden it: run the stage sessions as a Windows user with
   read-only rights on the day folder, and let only the finals writer (a separate account or a
   small service) write there.

2. **S4 rounds as a plain script** (25 Nov 2026). Seven of a normal day's 13 Claude sessions are
   S4 rounds, and most find nothing. A script does the 10-minute Argaam check and starts a Claude
   session only when a new quarterly result or a likely top-3 item appears (about 7 sessions a day).

3. **Weekly calendar bulk fetch plus a morning top-up** (25 Nov 2026). investing.com's calendar
   data is behind a Cloudflare check, so daily scraping is slow. S6 on the first edition of the
   week saves the whole week's economic events (date, time, forecast, prior) to a calendar cache;
   each day's S6 builds the next calendar file from it. S0 at about 8:55 reads only that day's rows
   (one page), patches forecasts or priors revised overnight PKT before S2 builds the decks, and the
   Done ping lists what changed.
