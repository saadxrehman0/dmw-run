## Replay of a past edition

This session replays the edition of {EDITION}; the real date is later. Work as if it were the
morning of {EDITION} at this stage's time. Only what was public by then counts:

- News (S1): items published from 9:00 AM PKT on {PREV_EDITION} (the previous edition) up to
  9:00 AM PKT on {EDITION}. argaam_fetch.py range returns exactly that window; for WebSearch
  results, check each article's own date and time, and leave out any whose time cannot be read.
- Results and late news (S4): announcements up to 10:10 AM PKT on {EDITION}, in one round.
- Next edition prep (S6): what was public by 10:20 AM PKT on {EDITION}.
- Never use a later article, a later price or a later announcement, even to check a figure.
  Market figures come only from the data file in {DAY}.

{DAY} holds only the morning's inputs (data file, Graphs, Technicals, Manual File-Calendar.xlsx)
and, in {DAY}/prev/, the previous edition's published EN and AR decks (to drop news already used)
and the most recent DMW - Results.xlsx (S4's starting workbook).
The edition's own published files are not available and must not be looked for.

Gates: nobody answers in a replay. The pick gate takes the run's own pick at 9:25; a conflict gate
stays open and its cell stays a yellow TK. Name the news Doc "REPLAY · DMW news · <edition date>".
Every output goes under {RUN}/practice/ (practice mode).
- Shared state is read-only in a replay. Never write to the DMW Tools state/ folder
  (chart_state.json, name overrides, Results, schedule): copy what a step needs into
  {RUN}/practice/state/ and point the step there (for build_dmw.py:
  --chart-state {RUN}/practice/state/chart_state.json). A 5 Nov replay overwrote the live
  chart_state on 26 Nov 2026, which the next live build would have read as a stale-chart baseline.
