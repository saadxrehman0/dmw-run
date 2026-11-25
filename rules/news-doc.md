# The edition's news Doc (replaces ANB Rough)

Decided with Saad on 25 Nov 2026; mock-up of the 24 Nov edition:
https://claude.ai/artifact/8EEXKGnoshg7bzf3JUW2wg. ANB Rough (`anb-rough-doc.md`) is retired for
the automated run: the team that sequenced it has been offboarded.

## One Doc per edition

- Title `DMW news · <Ddd> <d> <Mon> <yyyy>` (the edition date). The run makes it in S1 with
  `write_news_doc.py` and saves its link to the day folder.
- Layout, top to bottom:
  1. Lead: pick by 9:25 AM PKT.
  2. **Pick**: one table for Economic, one for Corporate. Columns Rank, Item (`E1 · headline`),
     Pick (In / Out dropdown), Run says (Include / Maybe), Why (one line).
  3. **Markets**: Global and TASI, each EN then AR.
  4. **Economic** and **Corporate**: one `### <id> · <headline>` per item, then the EN paragraph,
     the AR paragraph, and `Source: [<source>](<link>)` (order set by Saad, 25 Nov 2026).
     Economic items add `**Short version** (used only if slide 1 overflows)`, then a short EN and
     a short AR. The top two Economic items always carry one (Saad, 25 Nov 2026).
  5. **Left out**: ID (X1...), Headline, Source, Reason, Pick, Rank. Saad switches a row to In to
     bring it back; the run drafts its EN and AR and adds it as a new E or C item.
- Item ids never change. Order lives only in Rank: Saad reorders by typing a number (1.5 puts an
  item between 1 and 2). Ties keep table order.
- The run's own pick is the Doc as written: Include items In, Maybe items Out.

## The pick

- About 9:08 PKT: the pick card goes to DMW Ops (`pick_card.py card`), with Go with proposed and
  I edited the doc. Saad can also reply in words ("drop C7", "put C6 second"); the run makes that
  change in the Doc, so the Doc stays the one record of the pick.
- 9:18: one reminder. 9:25: the run reads the Doc as it stands. If Saad has not touched it, that
  is the run's own Include tier in its proposed order, and the Done ping says the pick was the
  run's (decided 25 Nov 2026).
- Saad's edits to item text win: the reader takes the text as it stands. Edits up to 10:00 refill
  the decks.

## Reading the Doc

Export the Doc's one tab with the Claude Docs connector, format **markdown** (the `text` format
drops links and headings), decode it, and run `read_news_doc.py`. It stops (exit 1, nothing
written) on a missing Markets bullet, a picked item with no EN or AR, an unknown or doubled id.
It flags (and carries on) a non-number Rank, a Pick cell that is neither In nor Out (read as Out),
split paragraphs (joined), a short version with only one language, and an Economic count other
than 2.

## Headless access (tested 25 Nov 2026 on Saad's PC, DESKTOP-3S8HFHF, Windows)

`claude -p` signed in with the claude.ai account sees the Claude Docs connector as
`mcp__claude_ai_Claude_Docs__*` (`claude mcp list` shows "claude.ai Claude Docs ... Connected").
A headless run created a Doc, exported it as markdown and returned its link, with
`--allowedTools "mcp__claude_ai_Claude_Docs"`. Install notes for the server: PowerShell's script
policy blocks `npm.ps1`, so install from cmd with
`npm.cmd install -g --allow-scripts=@anthropic-ai/claude-code @anthropic-ai/claude-code` (without
`--allow-scripts` the native binary is not fetched), then `claude auth login` once.
