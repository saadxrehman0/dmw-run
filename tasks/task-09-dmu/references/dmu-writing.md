# DMU writing guide (commentary on slides 2 to 5)

Built 21 Oct 2026 from the team's 80 published DMUs (1 Jul to 21 Oct 2026). Each one was compared with the same day's final DMW EN deck. The counts come from `commentary.py`, `source_map.py` and `saudi_pick.py` in the review folder. Where a rule here conflicts with an earlier DMU instruction, this guide wins (Saad, 21 Oct 2026).

Updated 29 Oct 2026 from the team's DMU training video (22 Oct deck) and Saad's answers: source pages, the fallback when Reuters has no story, slot 5 priority, and no source tags. Walkthrough: `claude/DMU video walkthrough - 2026-10-29.md`. Updated again 29 Oct 2026 from the slide-by-slide test against the team's 29 Oct final (story choice for slots 3 to 5, events lines), 1 Nov 2026 (slide 4 selection rule from all 86 decks) and 2 Nov 2026 (captions trimmed to one line, from the team's 12 DMUs of 15 Oct to 2 Nov; always the latest data release; agency items kept short; Zawya and Arab News allowed when the news is new; economic item first confirmed; slide 4 corporate items in the DMW's order; approved by Saad). Updated 3 Nov 2026 from Areeb's correction session with Saad (audio and screen recording, `claude/DMU corrections - Areeb - 2026-11-03.md`): thin-day fallback, Japan/China order, settle prices from our data, Monday captions, and the slide 4 order (economic item only when strong, contracts by size; mix chosen by Saad). Areeb's stock futures instruction was wrong and is not followed (`dmu-structure.md`).

## The one big rule

**The DMU reuses text; it does not condense it.** In the team's decks every paragraph is either lifted from our own house text or is a news-agency lead (Reuters by default) with light edits. My 20 and 21 Oct runs summarised and shortened instead, and the team had to replace that text. Areeb confirmed this on 3 Nov 2026: the team copy-pastes the agency lead, so the deck reads the same whoever is on the rota.

- **House text (headline 1 on slide 2 from the final DMW deck; slide 4 from the ANB Rough doc): copy it word for word.** It is our own house text.
- **Agency items (slide 2 items 2 to 5, slide 3): match the lead in facts and figures** (2 sentences, about 40 to 60 words, every number and comparison kept). Write it in house style rather than pasting the agency's sentences. I don't reproduce third-party articles verbatim. The team may still paste the original if they prefer. **Keep them tight** (Saad, 2 Nov 2026: "make our bullets a bit shorter"): one sentence for the news with its figures, one for the context; drop side details before figures. On 2 Nov the Logan item went from 76 to 60 words. (Open with Saad, 3 Nov 2026: whether to follow the DMW Global write-up's approach here, see the corrections file.)
- Never merge two stories into one item or drop the figures to save space.
- **Nothing critical of Saudi Arabia**, its government, leadership or policies, anywhere in the deck: the red line in `shared/rules/core-style.md` (team, 29 Oct 2026). Check every agency item against it.
- **No source tags on the slides** ("(Reuters)", "(Bloomberg)", "- CNBC" and the like). Management does not want them (Saad, 29 Oct 2026). Put every link in the source log instead, so the item can be traced later.

## Source map

| Box | Source | Evidence (80 decks) |
|---|---|---|
| Slide 2 headline 1 | Global write-up **exactly as it appears in the final DMW deck** (after Task 8), not the draft or the ANB Rough doc. Drop the trailing "(Reuters)" tag | 78 of 80 match the DMW text; median similarity 1.00 |
| Slide 2 headlines 2 to 5 | Reuters first. If Reuters has nothing relevant for a slot, any other outlet (Bloomberg, CNBC, Yahoo Finance and so on), no fixed order: keep searching until a relevant story turns up. Areeb, 3 Nov 2026: Reuters is the main source; others "only if we don't have anything" | 0 of 80 match the DMW. 22 Oct slot 3 was a Bloomberg story (video, 3:19) |
| Slide 3 oil and gold paragraphs | Reuters, not in the DMW | 3 of 80 match the DMW |
| Slide 3 captions | The Reuters headline of the story behind each paragraph, **kept to one line at 12 pt**: word for word when it fits, otherwise trimmed (see Captions) | 22 Oct: both captions match the Reuters headlines exactly (video, 13:04 and 14:42); 1 and 2 Nov: trimmed |
| Slide 4 headlines | **ANB Rough doc** (Economic and Corporate sections) as it stands at build time, morning additions included, copied word for word. The DMU goes out before the DMW, so not the final DMW deck (Saad, 29 Oct 2026) | 29 Oct: 3 of the team's 4 items were not in the final DMW deck |
| Results slide | Results table plus our-estimate commentary (Task 10 inputs) | 46 decks have it |

**Headline 1 comes from the final DMW deck in the day folder; slide 4 comes from the ANB Rough doc** (read it with `rough_map.py`; strip the doc's "[...]" markers, trailing URLs and source tags before copying).

### Where to look (from the team's 22 Oct video)

| Need | Page |
|---|---|
| Slot 2, geopolitics | Reuters Iran War hub, `reuters.com/world/iran/`: the top story's first paragraph. Skip a pinned story already run in an earlier edition; take the newest story published since the last edition |
| Slots 3 to 5, macro | Reuters Econ World, `reuters.com/markets/econ-world/`. Then Google News ("us economic news", "japan economy", "china economy"). Then the Bloomberg terminal: `TOP` <GO>, click **More »** on the **Eco & Gov** heading, set the date to **Last Day** (Saad does the terminal step; it is shared over AnyDesk) |
| Slide 3 oil and gold | Reuters Commodities, `reuters.com/markets/commodities/`, or a Google search "oil prices reuters" / "gold prices reuters" |

Aggregators (Yahoo Finance and similar) republish other outlets. Log the original publisher shown on the article, not the aggregator.

**Reuters in the browser (29 Oct 2026):** WebFetch is blocked for reuters.com, and the built-in browser was shown "Access temporarily restricted" after three quick page loads. Chrome loaded every page. Use Chrome for Reuters, and read a story's headline, time and paragraphs with a short script (`[data-testid^="paragraph-"]`), not the whole page. 2 Nov 2026: in the built-in browser, one page load plus `fetch()` calls from that page to the section API (`/pf/api/v3/content/fetch/articles-by-section-alias-or-id-v1`) and to story pages worked without being blocked.

**Reuters briefs** (2 Nov 2026): a fresh data story is often a one-paragraph brief at first (the Japan PMI story at 00:37 GMT had one paragraph). The figures then appear in a fuller story, often Reuters' Asia factory roundup. When the brief has no figures, take the figures and the second sentence from the fuller Reuters story and log both links.

## Slide 2: Global Markets, headline updates

- **5 items** (4 to 6 seen; median 5), 11 pt body. The box holds about **204 to 347 words** (median 302). On 21 Oct mine came to 232 words and the team's to 306. **Since the 29 Oct layout keep it to about 310 words or fewer**: at 323 words the last line ran into the FOR INTERNAL USE ONLY footer in the render, at 309 it fitted. Render slide 2 and check the last line before handing over; when it overflows, cut the Iran item to the story's first paragraph (the team does this) before cutting any figures. 2 Nov 2026 practice: 259 words with the tighter items, with room to spare.
- Fixed slot order (since September it held every day):

| Slot | Topic | Sep to Oct (36 decks) |
|---|---|---|
| 1 | DMW Global write-up (US close plus the Asia morning) | 36 of 36 |
| 2 | Geopolitics: the leading conflict story (Iran, Hormuz, Houthis) | 36 of 36 |
| 3 | US macro: data, Fed speakers, Treasuries, housing, mortgages | 36 of 36 |
| 4 | A second US item, or China or Japan | US 16, China 8, Japan 7, other 5 |
| 5 | Asia macro: **Japan first** (BOJ, data, JGBs). China (data, PBOC, trade) only when there is nothing relevant on Japan (Saad, 29 Oct 2026) | Japan 20, China 14 |

- **Japan and China (Areeb, 3 Nov 2026):** one or two US items, then one China or Japan item; when there is good Japan news, prefer Japan. When both run, **order them by importance, not by country**: a bigger China story sits above a Japan one.
- Each item is a lead of 2 sentences, **about 40 to 60 words**. Keep the figures, the comparison with the forecast and the attribution inside the sentence ("government data showed", "a Reuters poll"). A one-line summary is too thin. My 21 Oct Japan item had 26 words against the team's 55.
- **Always the latest data** (Saad, 2 Nov 2026). When two releases cover the same thing, use the newest one published before the cut-off. On 2 Nov, China's private RatingDog PMI (Monday, 52.0) replaced the official NBS PMI (Saturday, 49.3), which is what the team ran too. The older release can be mentioned in a clause only if the item stays within length.
- Pick the stories Reuters is leading on that morning, in the slot order above. When Reuters has nothing that fits a slot, use another outlet (see the source map). Do not use a story that is already in headline 1.
- **Story choice for slots 3 to 5: what happened, not what is coming** (29 Oct 2026 test). Take the main news story on the last US session's events and data releases: the Fed decision story itself on a Fed day (vote split, rate range, market pricing for the next meeting), and the data released that session (29 Oct: the advance goods trade deficit, with the figure, the change and the Reuters poll). Do not use a preview of data due later today (I used the Q3 GDP preview) or a side story about the main event (I used Warsh on politics while the team used the decision story). 2 Nov: with no US data on Friday, I used Logan's Friday call for two more hikes, and the team kept the Fed decision story from Wednesday. **When the last US session had no data** (settled 10 Nov 2026, Saad's call on Claude's recommendation, from Areeb's 3 Nov thin-day rule): take the biggest US macro story published since the previous edition that has not run in a DMU (a Fed speaker counts); if nothing new fits, the main story of the most recent data or Fed day that has not run yet, as the team did on 2 Nov.
- **Japan slot: the same morning's hard numbers first** (JGB auction results with bid-to-cover, data releases), over a preview piece from earlier in the week. On 29 Oct the team used the 2-year JGB auction (bid-to-cover 4.03 vs 3.89, the 12-month average 3.84); I used a 3-day-old BOJ preview.
- Only stories published before the DMU build (about 9:50) count. Check the time on the article.
- **Search window: everything published since the previous edition's cut-off** (about 10:00 PKT the day before), not just the top of the page. The team's 29 Oct picks were published at 15:05 and 17:41 PKT on 28 Oct (Fed decision, goods trade deficit) and had scrolled down Econ World by the next morning. Use the section's "Load more" or the Reuters US economy pages until the list reaches the previous cut-off.
- **Thin days (Areeb, 3 Nov 2026).** News is thin on Monday mornings and sometimes on Sunday. When nothing new fits a slot, an older relevant story that has not run yet is fine (on 2 Nov Areeb used a Thursday story because Wednesday's was old). Check the earlier DMUs so it has not already run, and still prefer a newer story when there is one. The new-news test for slide 4 does not change.
- **Read headlines, not URL slugs.** Reuters keeps a preview's URL for the main story: on 29 Oct `fed-likely-go-into-wait-and-see-mode-rates-2026-10-28` carried "Fed leaves rates unchanged as it awaits 'accumulation of more information'".
- **Japan slot: Bloomberg is the second stop.** On 29 Oct Bloomberg ran "Japan 2-Year Bond Sale Sees Firm Demand Before BOJ Decision" at 8:38 PKT; Reuters' auction story came at 10:41, after the cut-off. When Reuters has no same-morning Japan story by about 9:15, check bloomberg.com (Markets, Asia / Japan headlines) in Chrome.
- A short item is acceptable when the lead is one long sentence: on 29 Oct the team's Iran item was the story's first paragraph alone (43 words).
- On 21 Oct the team dropped my South Korea exports and BOJ poll items and used Iran/Pakistan, US pending home sales and the Reuters Fed poll. Other international items (Korea, India, IMF and similar) are rare: 2 of 71 items in slots 4 to 6 from September to October.

## Slide 3: Commodity Markets

- **Paragraph 1 = oil, paragraph 2 = gold.** A third paragraph appears in 16 of 80 decks for a big oil-market story (an OPEC output survey, for example). About 80 to 210 words in total (median 116). The slide has room below the two paragraphs (2 Nov 2026 render: 139 words, the box half empty), so a third context sentence may stay when it adds something.
- Each is a Reuters lead, about 45 to 90 words. Openings seen: "Oil prices settled ..." (21 decks), "Oil prices rose / fell / were ..." and "Gold prices ..." / "Gold rose / edged ...".
- **Take the story about the last settlement**, which is the previous US session for an edition from Tuesday to Thursday. The 21 Oct team gold paragraph covered Tuesday's close ("Gold edged lower on Tuesday ..."). Mine covered the Wednesday 0319 GMT move and was replaced.
- On Monday editions the paragraphs carry the Monday morning story ("Oil prices slipped on Monday ...") in 15 of 16 decks. On 24 Aug they kept Friday's story.
- **Check the direction against the chart before choosing the story** (29 Oct 2026 test). Read the last-day change from `DMW - Data.xlsx` `Sheet1` (oil T:V, gold X:Y) for the chart's last date. The paragraph and the caption must describe that move. Reuters' story is often filed before the close: on 28 Oct it said spot gold "lost 0.5% ... by 3:12 p.m. EDT" while the data file closed gold **up** 0.18% at 4,067.35.
- **Settle prices must match our data** (Areeb, 3 Nov 2026). The Brent and WTI settle prices in the oil paragraph must match the slide's Performance recap table (Last Close, from `DMW - Data.xlsx`). They can differ around a new month or a contract roll, when Reuters quotes the new front-month contract: then use our data, not Reuters' figure. Apply the same check to spot gold.
- **When Reuters and the data disagree, stop and tell Saad** (team, 29 Oct 2026). The team then uses a story from the Bloomberg terminal as a one-off (29 Oct caption: "Gold erased losses on Wednesday after Fed officials left rates unchanged"). Saad pastes the terminal story into the chat; draft the paragraph from it in house style, take its headline as the caption, and give the closing level and change from the data file. Do not hunt for another outlet on the web, and do not write the item from the Reuters story.
- Trim the Reuters lead to the move, the reason and the settle prices.
- **Oil wording: `Brent` and `WTI`** (management rule, 29 Oct 2026; `core-style.md`). Reuters writes "Brent futures rose ... while US West Texas Intermediate crude rose ..."; the team writes "Brent rose ... while WTI rose ..." (78 of 96 oil paragraphs, 1 Jul to 21 Oct, use bare Brent; 77 use WTI; "Brent futures" 3, "West Texas Intermediate" 2).
- **Gold wording** (80 gold paragraphs, 1 Jul to 21 Oct 2026): open with "Gold prices rose / fell / edged ..." or "Gold rose / edged ... on <day>" with the reason; then **"Spot gold rose / fell x% to $x,xxx.xx per ounce"** (77 of 80; "per ounce" 73, "an ounce" 7); optionally "US gold futures for December delivery settled x% higher / lower at $x,xxx.xx" (33 of 80). Dollar figures always with the thousands comma and two decimals (80 of 80). Drop the time stamp ("by 3:12 p.m. EDT", 1 of 80). "Closing at" never appears; the 29 Oct team text ("closing at $4067.4") is the exception, not the pattern. On 21 Oct the team cut "after reaching $111.99 earlier in the session".
- **Monday gold story and the cut-off** (2 Nov 2026): Reuters' Monday gold story came out at 10:04 PKT, after the cut-off. The team used Bloomberg's Monday gold story instead ("Gold rose on Monday as lower oil prices offered some respite ... Bullion gained as much as 1.1% to near $3,915 an ounce"). When Reuters has no Monday gold story by about 9:15, check bloomberg.com, as for the Japan slot.

### Captions (above each chart)

- Top caption = oil, bottom caption = gold. Start from the headline of the story used for the paragraph (Reuters; the Bloomberg terminal story Saad pastes when Reuters points the wrong way, see above), no full stop.
- **One line at 12 pt bold, always** (Saad, 2 Nov 2026, from the team's 12 DMUs of 15 Oct to 2 Nov: every caption one line, every one 12 pt; the font is never shrunk). The line holds about 5.3 in of Arial Nova Cond Bold, roughly 14 words. When the headline fits, copy it word for word (22 Oct: "Oil prices steady as investors weigh geopolitics, lower US stockpiles"). When it is longer, **trim it to the main clause** and drop the trailing clause, keeping the words in their original order (1 Nov team: "Tech rout drives gold to one-year low" from "Tech rout drives gold to one-year low, white metals sink deep into oversold zone"). `build_dmu.py` measures each caption and, as a last resort, moves the chart below a wrapping caption down one line; that INFO line means the caption should have been trimmed.
- **The caption describes the move on the chart's last date.** On every Monday edition (16 of 16) the captions describe Friday's session ("Oil touched $110 a barrel on Friday ...") while the paragraphs below describe Monday, because the chart ends on Friday's close. On Mondays the caption is therefore the headline of Friday's story, and the paragraph comes from Monday's story. This is deliberate. The 12 Oct check that flagged "Friday captions over Monday headlines" as an error was wrong and is withdrawn.
- **Monday caption = Sunday's caption, in the past tense with "on Friday"** (Areeb, 3 Nov 2026: "on Monday ... we will add Friday", "same for the gold part"; the charts stay as on Sunday because both editions end on Friday's close). The team did this in 5 of 5 Sunday/Monday pairs from 4 Oct to 2 Nov. Examples: Sunday "Oil jumps 2%, settles at 6-month highs on fears of escalation in Iran", Monday "Oil jumped 2% on Friday, settling at 6-month highs"; Sunday "Tech rout drives gold to one-year low", Monday "Tech rout drives gold to one-year low on Friday". Put the verb in the past tense where the headline has one, add "on Friday" after the main clause, and drop the trailing clause to keep it on one line.
- From Tuesday to Thursday, caption and paragraph come from the same story.

## Slide 2: Events calendar

- One line per event: `DD-Mon-YY: <Country> <Event>: <x> (P), <y> (F)`. The format moved from `DD-Mon-YYYY` to `DD-Mon-YY` on 6 Oct 2026, and every deck since uses `DD-Mon-YY`.
- **Drop the period tag** the calendar file carries after the event name ("(Sep)", "(Q3)"): `29-Oct-26: US GDP (QoQ): 2.2% (P), 3.2% (F)`, not `US GDP (QoQ) (Q3)` (29 Oct 2026 team final). Keep "(YoY)", "(QoQ)" and "(MoM)".
- **SAMA monthly rows** (KSA M3 Money Supply, SAMA Net Foreign Assets and the like) become **one line, first, without figures**: `29-Oct-26: KSA SAMA Monthly` (29 Oct 2026 team final).
- (P) = previous and (F) = forecast are included where they exist (114 of 142 lines). **API and EIA crude inventory rows carry no figures** (user, 21 Oct 2026; `shared/rules/calendar.md`): `21-Oct-26: US Weekly EIA Report on Crude Oil Inventories`.
- 0 to 6 lines. 29 of 80 decks had none. KSA lines first, then today's US releases in the calendar file's order. KSA releases include CPI, the real estate price index and the SAMA monthly.
- Source: the day's Manual File-Calendar / investing.com.

## Slide 4: Saudi Market, headline updates

Rule from all 86 team DMUs, 1 Jul to 29 Oct 2026 (`s4_hist.py`, 29 Oct 2026 test; approved by Saad 1 Nov 2026), revised 3 Nov 2026 with Areeb's correction (Saad chose a mix of the two). The style is meant to be the same every day. A deck that breaks it is a slip, an undertrained hand or a one-off from the line manager, and is logged below, not copied.

- **3 or 4 items, about 180 to 275 words** (median 220). Render slide 4 and check the last line clears the FOR INTERNAL USE ONLY footer: on 2 Nov, 4 items (302 words) overflowed and 3 items (235 words) fitted.
- **Item 1: one official-data economic item, when there is a strong one.** 75 of 86 decks led with an economic item. Prefer official data: GASTAT and other authority statistics, SAMA, the PMI, the budget, banks' monthly profits, sukuk and savings-product data (57 of 86 economic picks). Sector figures (cement sales, private-market inflows) come next. **When the day's economic items are weak** (ministry approvals, executive or official comments and project tenders only), **slide 4 can be all corporate** (Areeb, 3 Nov 2026: "you can also keep all the corporate ones if the eco one is not that good"). The line manager (Miriam) has the final say, so say in the hand-off when the economic item was left out and why.
- **Corporate items: contracts first, largest contract value first** (Areeb, 3 Nov 2026; Saad, 3 Nov 2026). A contract announced by a company in our coverage universe (the ANB Coverage Universe file in `config.md`) is kept in even when it is smaller. Then the other corporate items in the same order as the day's DMW Corporate section (the sequence Claude ranks in the ANB Rough doc and the team finalizes; take it from the doc's Corporate section as sequenced at build time, or from the final DMW EN deck when it is out). Fill as space allows (usually 2 or 3 corporate items, 4 on an all-corporate day). Compare contract sizes in SAR, from the doc paragraph. This replaces the 2 Nov rule "corporate items in the DMW's order".
- **Skip**: executive interviews and comments (about 6 of 195 ever ran; 29 Oct skipped Petro Rabigh, ADES, Bahri and NCLE), MoUs, licences and certificates, housekeeping items (renamings, lawsuit updates, expropriations) and unlisted companies.
- **Zawya and Arab News items are allowed** (Saad, 2 Nov 2026; the earlier blanket skip of Zawya corporate items is removed), on one condition: **the news must be new.** Check that the event itself (the announcement, the release, the deal) happened since the previous edition, not just that the article is recent, by finding the underlying announcement or release date; drop it if it is older or already ran. Same rule as `shared/rules/news-selection.md`, Sources. 2 Nov example: Aramco/HUMAIN (Zawya) was first announced in October 2025, but its closing was reported on 1 Nov after the last edition, so the closing is new.
- **Copy the doc paragraph in full**, as it stands at build time. Drop only the trailing source tag ("(Argaam)", "(Tadawul)", "(Zawya)", "(Arab News)") and any URL. On 21 Oct I cut Tanmiah to one sentence and the team restored the second one.

Deviation log (the all-corporate days below are allowed under the 3 Nov rule when the economic items were weak):
- 28 Oct 2026: foreign flows first, then three corporate items, no economic item.
- 29 Oct 2026: the team ran four corporate contracts (Al Moammar, ACWA Power, Alqemam, Al Ashghal Al Moysra) and no economic item, although the GASTAT mining licences were in the doc. This matches the 3 Nov rule (contracts first; mining licences are not strong official data).
- 2 Nov 2026: the team again ran no economic item (ACADEMY, EAT, then Aramco/HUMAIN) although the GASTAT 3Q26 flash GDP was in the doc. Flash GDP is strong official data, so under the 3 Nov rule it would still lead.

## Results Announcements slide (when present)

- On days with results only. 46 of 80 decks, nearly every day in results season. Two slides on 2, 4, 5, 6 and 9 Aug, when the list ran long.
- Table: Company, TASI code, Bloomberg code, net profit (this quarter, same quarter last year), change YoY (SAR mn, %), consensus and variance, anbc estimate and variance. `-` where there is no anbc estimate.
- **Consensus net profit comes from Bloomberg `EEO <ticker>`** (Saad, 29 Oct 2026). The terminal step is done on the shared Bloomberg terminal, not by me. Do not confuse the two columns: Consensus is Bloomberg's, anbc estimates are ours (Areeb, 3 Nov 2026).
- Commentary only for companies with an anbc estimate, one paragraph each, about 30 to 80 words: "<Company> (<TICKER> AB) reported a net profit of SAR <x> mn in <Q>, <in line with / above / below> our estimate of SAR <y> mn, <up/down> by <z>% (<+a% YoY>, <+b% QoQ>). The earnings variance was driven by ..."
- Inputs come from Task 10 (`DMW - Results.xlsx`) and the anbc estimates.
- **Row order** (Saad, 1 Nov 2026): **companies in our coverage universe** (ANB Capital names covered by the Akseer team) first, then the rest by this quarter's net profit, largest first. Coverage list: `state/coverage.json` once Saad sends it; until then, names with an anbc estimate stand in for it. On 29 Oct the team put ARMAH third; logged as a one-off.
- **Which rows: every result in `DMW - Results.xlsx` at build time** (Saad, 1 Nov 2026). The team leaves out late results only because it has to send by a fixed time (29 Oct: ARDCO and Saudi Cable missing from the team's deck); I am faster, so I include them all.
- **anbc estimates and commentary come from the covering analyst**, who sends the bullet over WhatsApp (Saad, 1 Nov 2026). Saad pastes it into the chat: take the estimate into the table (variance = actual / estimate - 1), and put the bullet in the commentary box in house style, keeping its figures. Until it arrives, leave the anbc columns `-`, the commentary box empty, and list the covered names as a TEAM item. Never write the commentary or invent an estimate myself.
- **Pending (3 Nov 2026): a fallback source for anbc estimates.** Areeb says that when the analyst's bullet has not come in, the team takes the estimate from the quarter's results preview file (3 Nov example: Leejam, a covered name, had no anbc estimate in the results file). The team is sending Saad the file; until it arrives and Saad confirms it, keep the rule above.
- Every row in `DMW - Results.xlsx` with a TASI code counts, quarter in column A. Number formats as in the file's DMW table: SAR mn with the thousands comma and one decimal, % one decimal without a comma (`1025.0`), `-` for a swing to profit or loss and for a missing consensus. A negative consensus is shown (`-35.8`) with `-` variance.
- **Non-quarter rows** (column A not the quarter, e.g. `FY26`): star the short name (`PURITY*`) and put the note on a second line of the source row: `Source: Tadawul, Bloomberg` / `*FY26` (`results.footnote`).
- **Commentary box** sits just under the table (the build moves it); about 0.48 in per paragraph.

## What changes nothing

- Fonts: commentary 11 pt Arial Nova Cond; captions 12 pt bold, colour #0070C0. The formatting spec is in `dmu-structure.md`, and `check_dmu.py` enforces it.
- No em-dashes, house number style, as in `core-style.md`.
