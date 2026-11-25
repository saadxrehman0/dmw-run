# Arabic rules (all DMW, DMU and Global / TASI Arabic)

Modern Standard Arabic, Gulf financial press conventions, mirroring the English sentence for
sentence. **The translator proofreads; they do not retranslate**, so the drafted Arabic is the
text that ships. Nothing downstream will catch a wrong entity or a stale sentence in fluent Arabic.

## Matching the English

- The Arabic must match the shipped English sentence for sentence. When the English is trimmed or
  edited after the Arabic was drafted, cut or change the same material in the Arabic. If there is
  no time, give the user the list of sentences the English lost.
- When the user edits English only, hold the Arabic until the English is final, then translate the
  change and say which sentence you translated.
- End every Arabic bullet with the same source tag in Arabic, with no full stop after it: `(أرقام)`,
  `(تداول)`, `(زاوية)`, `(رويترز)`, `(عرب نيوز)`, `(بلومبرغ)`, `(رويترز، بلومبرغ)` (forms the Sep and
  Oct decks used; Bloomberg appeared once each as بلومبرغ and بلومبيرغ, so بلومبرغ is set as the
  house form, 12 Oct 2026).

## Names and Latin script

- **No Latin script for names inside an Arabic paragraph.** Foreign and unlisted names are written
  in Arabic, legal suffixes dropped, شركة prefix for a company: `Yara International Ltd.` →
  `يارا الدولية`.
- **Only two kinds of Latin text stay:** the Bloomberg ticker after the company's Arabic name on
  first mention (`الشركة العربية للأنابيب (APCO AB)`, never omitted), and the acronym indices and
  data names the house keeps in Latin: CSI 300, MSCI, PMI, S&P Global, ISM, API, EIA, ADP, JOLTS,
  FTSE Russell.
- **Listed companies use the exchange-official Arabic name** (from `TASI Mapping.xlsx` /
  `names_codes.xlsx`, keyed on the code) with no added diacritics (ملان, not مُلان). No trading
  brand suffixes (`الشركة السعودية للصناعات الأساسية`, not `... (سابك)`).
- **The Saudi Exchange is the sole authority for Arabic company names**, over the mapping files
  and any wording in earlier calendar sheets (31 Aug 2026). When a file and the exchange differ,
  the exchange wins and the file is corrected.
- 7205 Dar Albalad for Business Solutions: the exchange writes `شرکة` with a Persian kaf (U+06A9).
  Normalize it to the Arabic kaf (U+0643), `شركة دار البلد لحلول الأعمال`, the only knowing
  departure from the exchange's bytes.
- Open flag: 8270 Buruj Cooperative Insurance is marked Listed in the mapping but was absent from
  the Main Market board on 31 Aug 2026. Confirm on the board before relying on it.
- For Tadawul filings, take counterparty names from the Arabic filing (`locale=ar`).
- When an unlisted company's name is close to a listed one's, transliterate it (ريد سي جلوبال) so
  it cannot be read as the listed company, and flag it.
- Check `ar_glossary.json` first. **Tell the user which Arabic names you wrote yourself** and list
  them in the post log; after the proofreader has seen them, add the final form to the glossary.

## Period labels in Arabic (3Q26, 1H27, 9M26)

- **Written out in Arabic, never in Latin** (team decision, 25 Oct 2026; replaces the 13 Oct rule
  that kept `2Q26` in Latin). A digit next to Q or H flips in right-to-left text and broke the
  formatting in the ANB Rough doc and the deck (`Q27.1`). The English keeps `XQYY` / `XHYY`.
- Short form, full four-digit year, no من عام and no م:

| English | Arabic |
|---|---|
| 1Q26 / 2Q26 / 3Q26 / 4Q26 | الربع الأول 2026 / الربع الثاني 2026 / الربع الثالث 2026 / الربع الرابع 2026 |
| 1H27 / 2H27 | النصف الأول 2027 / النصف الثاني 2027 |
| 9M26 | التسعة أشهر الأولى 2026 |
| 3Q26 to 3Q29 | من الربع الثالث 2026 إلى الربع الثالث 2029 |
| the same quarter of 2025 (2Q25) | الربع الثاني 2025 |

- `check_ar_text.py` now flags any Latin period label left in the Arabic and prints the Arabic
  form to use.

## Proofreader edits

House rules win over proofreader edits that break them. Adopt pure wording fixes. Do not adopt a
name put back into Latin, a figure losing its one decimal (800.0 → 800), or wording that no longer
matches the English. List the edits not adopted to the user.

## TASI bullet

- TASI is **المؤشر العام**, never مؤشر تاسي: "ارتفع المؤشر العام بنسبة 1.1% (112.9 نقطة) ليغلق عند
  مستوى 10,505.84 نقطة" (as published in the 5 Oct 2026 deck).
- **No "+" on the points in brackets on an up day**, EN or AR: `(112.9 نقطة)`, `(112.9 points)`. A
  down day keeps the minus: `(-76.8 نقطة)`, `(-76.8 points)` (13 Oct 2026 deck). The earlier example
  here wrongly showed `(+112.9 نقطة)` and was copied into a draft (user, 13 Oct 2026).
- Sector names from the exchange's Arabic sector filter (`global-tasi.md`). Two tied sectors take
  the dual: قطاعا + the two names + بانخفاض X% لكل منهما.

## Global bullet

| English | Arabic |
|---|---|
| Dow Jones | داو جونز |
| S&P 500 | ستاندرد آند بورز 500 |
| Nasdaq | ناسداك |
| Nikkei 225 | نيكي 225 |
| CSI 300 | CSI 300 (Latin) |
| MSCI Asia Pacific | MSCI آسيا والمحيط الهادئ |
| Federal Reserve | الاحتياطي الفيدرالي |

Hang Seng has not been through a proofread pass: confirm it with the user before publishing.

- Percentages take بنسبة: "بنسبة 0.6%".
- Each index is introduced as مؤشر; dual مؤشرا for a tied pair.
- "each" on a tie → لكل منهما after the percentage; or كل من when the pair is the subject:
  "ارتفع كل من مؤشري ستاندرد آند بورز 500 وناسداك بنسبة 0.5%".
- "followed by" → تلاهما (dual antecedent), يليه (single).
- "morning" → صباح + day: صباح الخميس.
- "mixed": Asia وتباين أداء الأسهم الآسيوية صباح + day; US أغلقت الأسهم الأمريكية على أداء متباين.

## Calendars

Arabic calendar wording comes from the Arabic columns of `Manual File-Calendar.xlsx` and the final
published decks, not from the glossary (glossary is a fallback).

## Checks before posting

Run `shared/scripts/check_ar_text.py` on the Arabic, saved as a UTF-8 text file with one paragraph per line. It flags Latin other
than tickers and allowed acronyms, and double spaces. It must print OK. Then read the Arabic against
the English as control: no sentence dropped, same actor named.
