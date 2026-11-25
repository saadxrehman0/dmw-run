"""Checks Arabic news text before it is posted to ANB Rough or inserted into the AR deck.
Flags Latin script other than Bloomberg tickers and the allowed acronym indices, Latin period labels
(3Q26, 1H27, 9M26: written in Arabic since 25 Oct 2026, the flag prints the Arabic form), double spaces
and em-dashes. Also flags two TASI-bullet grammar slips (user, 21 Oct 2026): "يغلق عند" without
the لـ (house frame: ليغلق عند مستوى), and two companies after تلاه with no و between them or with
the singular سهم instead of the dual سهما.
Usage:  python check_ar_text.py file.txt     (one paragraph per line; prints flags, then OK or a count)
"""
import re, sys

# Bloomberg exchange codes: AB (Tadawul) plus foreign primary listings (core-style.md, 27 Oct 2026)
EXCH = r"(?:AB|US|LN|HK|JP|CH|GR|FP|NA|SW|IM|SM|KS|IN|SP|AU|CN|UH|DH|QD|KK|OM|BI|EY)"
TICKER = re.compile(r"\b[A-Z0-9&]{1,}(?: [A-Z0-9&]+)? " + EXCH + r"\b")
PERIOD = re.compile(r"(?<![A-Za-z0-9])(?:([1-4])Q(\d{2})|([12])H(\d{2})|9M(\d{2}))(?![A-Za-z0-9])")  # written out in Arabic (team, 25 Oct 2026)
Q_AR = {"1": "الربع الأول", "2": "الربع الثاني", "3": "الربع الثالث", "4": "الربع الرابع"}
H_AR = {"1": "النصف الأول", "2": "النصف الثاني"}


def period_ar(m):
    """Arabic form of a Latin period label: 3Q26 -> الربع الثالث 2026."""
    if m.group(1):
        return f"{Q_AR[m.group(1)]} 20{m.group(2)}"
    if m.group(3):
        return f"{H_AR[m.group(3)]} 20{m.group(4)}"
    return f"التسعة أشهر الأولى 20{m.group(5)}"
ALLOWED_LATIN = ["CSI 300", "MSCI", "S&P Global", "PMI", "ISM", "API", "EIA", "ADP", "JOLTS", "FTSE Russell"]
NO_LAM_CLOSE = re.compile(r"(?<![ء-ي])يغلق عند")
TICKER_PAREN = re.compile(r"\([A-Z0-9&]{1,}(?: [A-Z0-9&]+)? " + EXCH + r"\)")
NAME_START = re.compile(r"\s*(?:شركة|سهم|مصرف|بنك|مجموعة)")
LATIN_WORD = re.compile(r"[A-Za-z][A-Za-z0-9&.\-’']*(?: [A-Za-z0-9&.\-’']+)*")


def normalise_spaces(text):
    """Collapse runs of spaces/tabs to one space and trim. Newlines are kept."""
    return re.sub(r"[ \t ]{2,}", " ", text).strip()


def latin_flags(text):
    t = TICKER.sub(" ", text)
    t = PERIOD.sub(" ", t)  # reported separately by check_paragraphs
    for a in ALLOWED_LATIN:
        t = t.replace(a, " ")
    return [m.group().strip() for m in LATIN_WORD.finditer(t) if m.group().strip()]


def followed_by_grammar(p):
    """Flags in each sentence that has تلاه / تلاهما followed by two or more tickers."""
    out = []
    for sent in re.split(r"(?<=[.؟!])\s+", p):
        k = re.search(r"تلاه(?:ما)?\s", sent)
        if not k:
            continue
        rest = sent[k.start():]
        ticks = list(TICKER_PAREN.finditer(rest))
        if len(ticks) < 2:
            continue
        for a, b in zip(ticks, ticks[1:]):
            gap = rest[a.end():b.start()]
            if NAME_START.match(gap) and not re.match(r"\s*و", gap):
                out.append(f"missing و before the next company after {a.group()} (تلاه ... {a.group()} و...)")
        if re.match(r"تلاه\s+سهم\s", rest):
            out.append("singular سهم after تلاه with two companies: use the dual سهما")
    return out


def check_paragraphs(paras):
    issues = []
    for i, p in enumerate(paras):
        if re.search(r"[ \t ]{2,}", p):
            issues.append((i, "double space"))
        if "—" in p:
            issues.append((i, "em-dash: rephrase with a comma or full stop (en dash in calendar wording is fine)"))
        if re.search("[کیۀە]", p):
            # Persian kaf/yeh look the same but break search and font shaping; the team replaced
            # "شرکة وطني للحدید" with Arabic letters on 15 Oct 2026
            issues.append((i, "Persian letter (ک or ی): use the Arabic ك / ي"))
        if NO_LAM_CLOSE.search(p):
            issues.append((i, "'يغلق عند' without لـ: use 'ليغلق عند مستوى' (house TASI frame)"))
        for m in PERIOD.finditer(p):
            issues.append((i, f"Latin period label '{m.group()}' in Arabic: write '{period_ar(m)}'"))
        for msg in followed_by_grammar(p):
            issues.append((i, msg))
        for span in latin_flags(p):
            issues.append((i, f"Latin text in Arabic: '{span}' (write the name in Arabic, drop Ltd./Co./Inc.)"))
    return issues


if __name__ == "__main__":
    paras = [l for l in open(sys.argv[1], encoding="utf-8").read().splitlines() if l.strip()]
    issues = check_paragraphs(paras)
    for i, msg in issues:
        print(f"para {i + 1}: {msg}")
    print("OK" if not issues else f"{len(issues)} issue(s)")
