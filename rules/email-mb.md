# Sales Email (EMAIL MB): rules

Built from the **final EN DMW deck** by `shared/scripts/build_email_mb.py`. Derived from 532
results sentences in published EMAIL MBs (1Q and 2Q 2026 seasons) plus the desk's rulings.

## Sections

Markets Performance/News · Major Economic/Sector News · Major Corporate News (results sentences
first) · Corporate Calendar.

## Corporate Calendar (confirmed 12 Oct 2026)

**Every event row from the final EN deck except extraordinary general meetings.** Listing
upgrades, bonus-share eligibility, dividends and rights-issue rows all go in. The script drops only
rows matching "extraordinary general meeting / assembly" or "EGM" and lists each one in its checks.
(The older cash-dividend-only rule is withdrawn.)

## Results sentence

`<Full name> (<TICKER AB>) posted a net profit|loss of SAR X mn for <period>, <change clause>.  <consensus clause> (Tadawul, Bloomberg)`

- Two spaces before the consensus clause (house habit).
- Tag always `(Tadawul, Bloomberg)`, with or without consensus.
- Coverage names (26 Oct 2026, team edit): a company with an anbc estimate in the DMU results
  table takes the DMU analyst paragraph in place of the change clause, led by the full name and
  the Bloomberg ticker from the table (the DMU paragraph wrote `SOLUTIONS AB`; the ticker is
  `SOLUTION AB`), figures at 1dp (`SAR 417.0 mn`, estimate `SAR 468.7 mn`), then two spaces, the
  consensus clause and the tag `(Tadawul, Bloomberg, anbc research)`. This replaces the older
  `(ANBC Estimates, Tadawul, Bloomberg)` form (4 Oct 2026).
- One decimal, halves up. Figures from the deck's Results table, but change and variance are
  **recomputed** because the table does not auto-update; a gap over 0.15 against the table is
  flagged (within it, the table's figure stands). The table's change % and variance come from
  unrounded figures, so a table figure that the 1dp inputs can produce within their rounding
  (each input +/-0.05) also stands and is not flagged (AZM 23.0%, 26 Oct 2026).
- Period: the table header (e.g. 3Q26), except rows marked `*` / `**`, which take the period from
  the table's footnote (`*For 1Q27`, `**For FY26`) (26 Oct 2026).
- Current and prior-year profit equal (change rounds to 0.0): "reflecting no change from the net
  profit of SAR Z mn in the same period last year." (no %). The script also flags it so the
  prior-year figure is confirmed (Solutions 3Q26, 26 Oct 2026).
- **Always the full change clause**, never the short "from the net profit of" form.
- Estimated (unaudited management) results never appear.

| Last year → this year | Wording |
|---|---|
| profit → profit | reflecting an increase / a decline of SAR X mn (Y%) from the net profit of SAR Z mn in the same period last year. Decline % carries a minus sign: (-9.1%) |
| loss → loss | reflecting an increase in loss / a decrease in loss of SAR X mn from the net loss of SAR Z mn ... (no %) |
| loss → profit | reflecting a swing to profit of SAR X mn from the net loss of SAR Z mn ... |
| profit → loss | reflecting a decrease of SAR X mn from the net profit of SAR Z mn ... |

## DMU cross-check (26 Oct 2026)

Every row of the DMU results table is compared with the DMW Results table; any difference is
listed in the checks. The DMW table comes from `DMW - Results.xlsx`; settle the right figure there
(26 Oct: the DMU showed EPCCO consensus 47.7 and variance 4.9%; the source is 47.567, 5.1%).

- TASI codes are compared too (29 Oct: Armah 6022 on the DMU, 9590 on the DMW; 6022 is right).
- Rows on the DMW only are expected and not flagged: the DMU's cut-off (9:50) is earlier than the
  DMW's (10:15), so results filed in between miss the DMU (user, 29 Oct 2026). Rows on the DMU
  only are flagged.
- Names: a code missing from the names sheet falls back to the Bloomberg ticker, for companies
  that moved from Nomu to the main market (Armah 9590 to 6022, 29 Oct 2026).

## Consensus clause

Variance = actual ÷ consensus − 1.

| abs(variance) | Wording (profit vs profit) |
|---|---|
| up to 2% (inclusive) | The result was in line with the consensus net profit of SAR X mn, up/down Y%. (no "by", team edit 26 Oct 2026) |
| over 2% to 5% (inclusive) | The result was broadly in line with ... |
| over 5% | The result exceeded / fell short of ... |

- Loss vs loss consensus, same bands on the size of the loss: "The loss came in line with / came
  broadly in line with / came lower than / came higher than the consensus net loss of SAR X mn." (no %)
- Profit vs loss consensus: "The result exceeded the consensus net loss of SAR X mn." (no %)
- Loss vs profit consensus: "The result fell short of the consensus net profit of SAR X mn." (no %)
- Consensus losses are written as positive numbers.

## When the user edits the generated email

Diff their version against yours, explain what went wrong, fix the script, and rerun it to confirm
it reproduces their version exactly. Then send the script change to the plugin maintainer.
