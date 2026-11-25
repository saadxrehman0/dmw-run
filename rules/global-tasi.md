# Global and TASI write-ups: process reference

## Shape of a daily entry

Each edition carries **that morning's** global bullet and the **previous session's** TASI close.
The TASI bullet written on the evening of 31 Aug appears in the 1 Sep edition.

- **Sunday:** US Friday close plus Asia's Friday close.
- **Monday:** Asia only (the US close was covered on Sunday).
- **Tue to Thu:** prior US close plus that morning's Asia session.
- Friday's global write-up goes into Sunday's edition (worked Saturday evening, finished Sunday).

## Global bullet (morning, cut-off 9:15)

Source: Reuters. The US close article for the US sentence; that morning's **Asia-session story**
(not the later London rewrite) for the Asia sentence. Reuters blocks plain fetching: use the
browser, and if that fails ask the user to paste the article.

Format: one sentence of narrative lifted from the article's lede, then the index moves, largest
move first. Tag `(Reuters)`, or `(Reuters, Bloomberg)` when the user folds in Bloomberg text.

- **Stick to the article's own wording.** Add no clauses, context or commentary. The client
  objects to additions, even true ones.
- **Keep the lede's clauses, too (user, 13 Oct 2026).** Do not drop a cause or context clause
  from the lede, including geopolitical ones that touch Saudi Arabia or the Middle East (e.g.
  "after attacks on a Saudi airport", "while a widening conflict in the Middle East soured risk
  sentiment"). Stay true to the Reuters article. The only trims are the ones listed here: other
  asset classes out of a compound subject, the figurative verb swapped for a plain one.
- **Use the latest update of the Asia-session story (user, 19 Oct 2026).** Reuters re-files the
  Sydney story during the session and the lede changes (19 Oct: the 05:38 PKT lede ended on the
  euro; the 07:48 PKT update ended on Chinese data). Take the lede from the latest Asia-session
  version, not the first one a search returns.
- **The Asia sentence opens "Asian stocks"** and carries **"morning"**: "Asian stocks slumped on
  Wednesday morning as…". Trim other asset classes out of a compound subject.
- **Index names, house form (user, 19 Oct 2026):** "Nikkei 225", "Hang Seng", "CSI 300",
  "MSCI Asia Pacific". **Never "Japan's Nikkei"** or any other country possessive, even when
  Reuters writes it.
- **Three Asia indices only: Nikkei 225, Hang Seng, CSI 300 (user, 20 Oct 2026).** MSCI Asia
  Pacific goes in **only when needed**: as a swap for an index that rounds to 0.0%, or for one
  whose market is closed. Never add it as a fourth index. Still read the article's MSCI
  Asia-Pacific (ex-Japan) move each day so it is ready for a swap; it comes from the article (no
  Yahoo quote), and the source log says which update it is from. (Replaces the 19 Oct "include
  it every day" rule.)
- **Plain verb**: rose / increased / fell / declined, not the article's figurative verb.
- **Indices moving in opposite directions → "mixed"**, whatever the lede says. Same for the US
  sentence when the three split ("U.S. stocks closed mixed on Wednesday as…").
- **Mixed line structure (user, 20 Oct 2026):** group the indices by direction, the larger group
  first, then "while" for the other: "The Nasdaq and S&P 500 rose 1.1% and 0.4%, respectively,
  while the Dow Jones fell 0.1%." Not "followed by the S&P 500, which increased 0.4%, while…".
  Arabic: "وارتفع مؤشرا ناسداك وستاندرد آند بورز 500 بنسبة 1.1% و0.4% على التوالي، فيما انخفض مؤشر
  داو جونز بنسبة 0.1%."
- **One decimal on every percentage**; a move of 0.05% to 0.1% rounds to 0.1%.
- **Never print "0.0%"** (user, 14 Oct 2026; no published bullet has). An index that rounds to
  0.0% is swapped for MSCI Asia Pacific on the Asia side (29 Sep, 30 Sep, 11 Oct), or written
  "was flat" (Nikkei 225, 28 Sep). On the US side follow the next rule; never "fell 0.2% and 0.0%".
- **US line when one index rounds to 0.0% (client, 21 Oct 2026):** if the other two moved the
  same way, write "U.S. stocks closed mostly up" (or "mostly down") and drop the flat index from
  the line: "U.S. stocks closed mostly up on Tuesday, as … The Dow Jones and S&P 500 rose 1.0% and
  0.3%, respectively." Not "mixed", not "while the Nasdaq was flat".
  Arabic: "أغلقت الأسهم الأمريكية على ارتفاع في معظمها يوم الثلاثاء، …" ("mostly down": على انخفاض
  في معظمها).
- **Round from the underlying figures**, half-up: change ÷ (price − change), not the page's
  displayed two decimals. `shared/scripts/global_calc.py` does this.
- **Check the prior close against the daily series** (Yahoo `interval=1d&range=1mo`, last
  completed session). The 1-day chart API has returned a two-day-old close. For the CSI 300, check
  the quote page's Previous Close against its displayed % change. After a mainland holiday the
  prior close is the last pre-holiday session.
- **Market holidays:** a closed market drops out (China and Hong Kong on 1 Oct; Japan on Sports
  Day, the second Monday of October; Hong Kong on Chung Yeung). Run the line on the indices that
  traded, and note the closure in the source log. Never report a closed market as a move.
- **Missing index:** if the Reuters story omits one the format needs (often CSI 300), take the
  latest level from Yahoo (`000300.SS`, `%5EN225`, `%5EHSI`). When one quote is looked up,
  refresh the others from the same source so the three are consistent, and log the timestamps.
- **An index sitting on 0.0% can be swapped for MSCI Asia Pacific** (figure from the article).
- Ties after rounding: "followed by the Nikkei 225 and CSI 300, which declined 0.1%, each."
- **Refresh the quotes just before sign-off**; moves of 0.1 to 0.3 points in minutes are common.
- In the ANB Rough doc the global write-up is **never highlighted yellow**.

**Settled 12 Oct 2026 (user):**
- **Sunday Asia sentence:** it reports Asia's completed Friday close, so it opens "Asian stocks
  ... on Friday", **without "morning"** (Arabic: يوم الجمعة, not صباح الجمعة).
- **No full stop after the source tag**, EN or AR: `(Reuters)`, `(رويترز)` (the decks used the
  bare tag 70 times in Sep and Oct, with a full stop once).

## TASI bullet (evening, after the 3:00 PM Riyadh close, about 5:30 PM Karachi)

Sources, and only these, all on `saudiexchange.sa` (browser):

1. **Market summary** (`ourmarkets/main-market-watch/marketsummary?locale=en`): TASI level and
   1D return, Top Gainers / Top Losers, sector heatmap.
2. **Indices Performance**, entity TASI, Key Parameters: the session's **Value Traded in SAR**.
3. **Historical Reports** (Indices, TASI): the **previous** session's value traded (the Value
   Traded and No. of Trades columns come back in the page text). Today's row does not post by
   6:00 PM Karachi; do not wait for it.

Rules:
- Main Market only. Exactly three companies from the page's gainers or losers table: TASI up →
  advancers; down → decliners; flat → advancers.
- Sector: top mover among the 22 true sectors. Exclude TASI50, Tadawul Large/Medium/Small Cap,
  Tadawul IPO and MSCI Tadawul 30.
- Use the exchange's label; Household & Personal Products is an "Index", not a "sector".
- Company form `Long name English (TICKER AB)`, looked up on the 4-digit code.
- **Listing transfers since the mapping file (`TASI Mapping.xlsx`, dated Aug/Sep 2026):** the
  mapping's Market column and tickers can be stale. A company moved from Nomu to the Main Market
  is a valid TASI pick; the exchange's gainers/losers table (and the data file) decide the market,
  not the mapping. Known transfers (user, 14 Oct 2026):

  | Code | Company | Now | Ticker |
  |---|---|---|---|
  | 9513 | Watani Iron Steel Co. (شركة وطني للحديد والصلب) | Main Market (TASI) | WATANI AB |

  Add new transfers here as the user confirms them.
- **Company moves round from the exchange's displayed two-decimal %, half-up** (EXTRA -8.05% →
  8.1%). Ties come only from displayed figures that round alike.
- **Sector ties:** two sectors on the same heatmap % are a tie, both named ("sectors", plural);
  never break a tie at further decimals.
- "each" trails the percentage, never before the verb: "surged 9.9% each, followed by…";
  "declining 2.1%, each." where it ends the sentence.
- **Points in brackets: no "+" on an up day** (`TASI closed 1.1% higher (112.9 points)`, 5 Oct 2026
  deck); a down day keeps the minus (`TASI fell 0.7% (-76.8 points)`, 13 Oct 2026 deck). Same in
  the Arabic (user, 13 Oct 2026).
- Tag `(Tadawul)`.

### Wording: copy the published pattern, vary the verbs (user, 13 Oct 2026)

Take the sentence frame from the last published deck for the same direction; do not repeat one
verb through the bullet ("rose" four times was flagged). Published frames:

- **Up day (5 and 7 Oct 2026 decks):** "TASI closed 1.1% higher (112.9 points) to close at
  10,505.84. Trading value decreased by 51.6% compared with the previous session. The Food &
  Beverages sector was the top gainer, rising 2.7%. Among the top gainers, Raydan Food Co.
  (RAYDAN AB) increased 10.0%, followed by … and … which closed higher by 7.7% and 6.7%,
  respectively. (Tadawul)"
  AR: "ارتفع المؤشر العام بنسبة 1.1% (112.9 نقطة) ليغلق عند مستوى 10,505.84 نقطة. وانخفضت قيمة
  التداول بنسبة 51.6% مقارنة بالجلسة السابقة. وكان قطاع … الأكثر ارتفاعًا، مسجلًا مكاسب بنسبة 2.7%.
  ومن بين الأسهم الأكثر ارتفاعًا، صعد سهم … بنسبة 10.0%، تلاه سهما … و…، اللذان ارتفعا بنسبة 7.7%
  و6.7% على التوالي. (تداول)"
- **Down day (8 and 13 Oct 2026):** "TASI fell 0.7% (-76.8 points) to close at 10,283.96. Trading
  value increased by 29.8% compared to the previous session. The Media and Entertainment sector was
  the top laggard, declining 2.1%. Among the top decliners, … declined 4.8%, followed by … and …,
  which fell 4.7% and 4.3%, respectively. (Tadawul)"
  AR: "انخفض المؤشر العام بنسبة 0.7% (-76.8 نقطة) ليغلق عند مستوى 10,283.96 نقطة. وارتفعت قيمة
  التداولات بنسبة 29.8% مقارنة بالجلسة السابقة. وكان قطاع … الأكثر تراجعاً بانخفاض 2.1%. ومن بين
  الأسهم الأكثر تراجعاً، انخفض سهم … بنسبة 4.8%، تلاه سهما … و… بانخفاض 4.7% و4.3% على التوالي.
  (تداول)"
- Before drafting, read the TASI bullet on slide 1 of the latest published EN and AR decks in the
  day folder and match it.

## Delivery and logs

English then Arabic, into the ANB Rough Draft tab under the Global / TASI headings, after the user
approves. Source links for every figure go in the day's source log, never in the client text.

## Sector map (exchange code → EN → AR)

| Code | English | Arabic |
|---|---|---|
| TENI | Energy | الطاقة |
| TMTI | Materials | المواد الأساسية |
| TCGI | Capital Goods | السلع الرأسمالية |
| TCPI | Commercial & Professional Svc | الخدمات التجارية والمهنية |
| TTNI | Transportation | النقل |
| TDAI | Consumer Durables & Apparel | السلع طويلة الاجل |
| TCSI | Consumer Services | الخدمات الإستهلاكية |
| TMDI | Media and Entertainment | الإعلام والترفيه |
| TRLI | Consumer Discretionary Distribution & Retail | تجزئة وتوزيع السلع الكمالية |
| TFSI | Consumer Staples Distribution & Retail | تجزئة وتوزيع السلع الاستهلاكية |
| TFBI | Food & Beverages | إنتاج الأغذية |
| THPI | Household & Personal Products (Index) | المنتجات المنزلية و الشخصية |
| THEI | Health Care Equipment & Svc | الرعاية الصحية |
| TPBI | Pharma, Biotech & Life Science | الادوية |
| TBNI | Banks | البنوك |
| TDFI | Financial Services | الخدمات المالية |
| TISI | Insurance | التأمين |
| TSSI | Software & Services | التطبيقات وخدمات التقنية |
| TTSI | Telecommunication Services | الإتصالات |
| TUTI | Utilities | المرافق العامة |
| TRTI | REITs | الصناديق العقارية المتداولة |
| TRMI | Real Estate Mgmt & Dev't | إدارة وتطوير العقارات |

The exchange's hamza use is inconsistent; the spellings above are the exchange's, and they are
used everywhere, including AR deck bullets (user, 12 Oct 2026: follow Tadawul). Household &
Personal Products is `المنتجات المنزلية و الشخصية`, with the space, as on the exchange and in 20
of 21 Sep and Oct deck uses.

**Do not take Arabic sector names from `DMW  Results.xlsx`.** Its Sector - AR column is
misaligned against Sector - EN (204 distinct pairings where there should be 22). Use the map above.
