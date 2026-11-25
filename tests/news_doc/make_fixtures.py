#!/usr/bin/env python3
"""Turn a published edition's news (published_<day>.json, read back from the v1 decks by
tests/replay/news_from_v1_deck.py) into run candidates, as the S1 sweep would write them.

Headlines and the Maybe tiers for 24 Nov follow the mock-up Saad approved on 25 Nov 2026
(https://claude.ai/artifact/8EEXKGnoshg7bzf3JUW2wg). The short Economic versions for 24 Nov are the
ones written for that mock-up. 18 Nov gets generated headlines and no Maybe items.
"""
import json, os, re, sys

HERE = os.path.dirname(os.path.abspath(__file__))
FIX = os.path.join(HERE, "fixtures")

MOCK_24 = {
    "economic": [("GASTAT construction cost index, Oct: +2.3% YoY", "Include", "Official data (GASTAT), leads"),
                 ("Certificates of origin, Oct: 55,965", "Include", "Ministry sector statistic")],
    "corporate": [("Foreign ownership, week to 19 Nov", "Include", "Weekly ownership item opens Corporate"),
                  ("Naseej: SAR 46.7 mn Oman contract", "Include", "Contract with value"),
                  ("Jarir: DMTT about 11% of profit from 2027", "Include", "Large-cap impact disclosure"),
                  ("NADEC CEO on red meat and feed stocks", "Include", "Management commentary"),
                  ("SFICO shrimp project timeline", "Include", "Project update with timing"),
                  ("Nahdi: tariff impact 0.5% of revenue", "Include", "Impact disclosure"),
                  ("Riyadh Cables: DMTT impact", "Include", "Impact disclosure, no figure"),
                  ("Zain KSA: DMTT about 2% of profit", "Include", "Impact disclosure"),
                  ("Meyar: non-binding MoU with Najran Cement", "Maybe", "MoU without a value (rule 20)"),
                  ("Arab Sea: fair value model, +SAR 15 mn", "Maybe", "Accounting change, small"),
                  ("SHL: SAR 250 mn facility renewed", "Include", "Credit facility, bottom filler")],
    "short": [("Saudi Arabia’s construction cost index (CCI) rose 2.3% YoY to 104.1 points in October 2026, "
               "according to GASTAT. Residential costs rose 2.2% YoY and non-residential costs 2.7%. (Argaam)",
               "ارتفع مؤشر تكاليف البناء في المملكة العربية السعودية بنسبة 2.3% على أساس سنوي إلى 104.1 نقطة في "
               "أكتوبر 2026، وفقاً للهيئة العامة للإحصاء. وارتفعت تكاليف البناء السكني بنسبة 2.2% وغير السكني "
               "بنسبة 2.7%. (أرقام)"),
              ("Saudi Arabia’s Ministry of Industry and Mineral Resources issued 55,965 certificates of origin in "
               "October, down from September’s yearly high of about 58,000 but above the 2026 monthly average of "
               "about 49,300. (Arab News)",
               "أصدرت وزارة الصناعة والثروة المعدنية في المملكة العربية السعودية 55,965 شهادة منشأ خلال شهر أكتوبر "
               "2026، بتراجع طفيف عن أعلى مستوى سنوي سُجّل في سبتمبر عند نحو 58,000 شهادة، لكنها ظلت أعلى من "
               "المتوسط الشهري لعام 2026 البالغ نحو 49,300 شهادة. (عرب نيوز)")],
    "left_out": [("A listed company declares a cash dividend", "Tadawul", "Dividend declaration (rule 9)"),
                 ("A board appoints a new CEO", "Tadawul", "Appointment (rule 8)"),
                 ("A company sets its 3Q26 results date", "Tadawul", "Results-date notice (rule 13)"),
                 ("Weekly negotiated deals screen", "Argaam", "Market report (rule 2)")],
}


def source_of(en):
    m = re.search(r"\(([A-Za-z][A-Za-z ]*)\)\s*$", en)
    return m.group(1) if m else ""


def auto_headline(en):
    t = re.search(r"\(([A-Z0-9]+) AB\)", en)
    words = " ".join(en.split()[:6]).rstrip(",.")
    return f"{t.group(1)}: {words}" if t else words


def candidates(day):
    pub = json.load(open(os.path.join(FIX, f"published_{day}.json"), encoding="utf-8"))
    mock = MOCK_24 if day == "20261124" else None
    c = {"edition": f"{day[:4]}-{day[4:6]}-{day[6:]}",
         "markets": {"global": {"en": pub["markets"]["en"][0], "ar": pub["markets"]["ar"][0]},
                     "tasi": {"en": pub["markets"]["en"][1], "ar": pub["markets"]["ar"][1]}},
         "economic": [], "corporate": [], "left_out": []}
    for sec, pre in (("economic", "E"), ("corporate", "C")):
        for i, x in enumerate(pub[sec]):
            h, tier, why = mock[sec][i] if mock else (auto_headline(x["en"]), "Include", "Published that day")
            it = {"id": f"{pre}{i + 1}", "headline": h, "tier": tier, "why": why, "en": x["en"], "ar": x["ar"],
                  "source": source_of(x["en"]), "url": ""}
            if sec == "economic" and i < 2:
                it["short_en"], it["short_ar"] = mock["short"][i] if mock else (x["en"], x["ar"])
            c[sec].append(it)
    if mock:
        c["left_out"] = [{"id": f"X{i + 1}", "headline": h, "source": s, "reason": r}
                         for i, (h, s, r) in enumerate(mock["left_out"])]
    return c


if __name__ == "__main__":
    for day in sys.argv[1:] or ("20261124", "20261118"):
        p = os.path.join(FIX, f"candidates_{day}.json")
        json.dump(candidates(day), open(p, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
        print(p)
