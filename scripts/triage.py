"""One-pass triage of an Argaam / Tadawul headline list into include / maybe / exclude.
Rules derived from the Jun to Oct 2026 published DMW decks (88 editions). First match wins.
Usage: python triage.py headlines.txt   (one headline per line)
"""
import re, sys

RULES = [  # (verdict, section, reason, regex)
 ('exclude', None, 'market report / price move', r'^(TASI|NomuC|Nomu)\b.*(pts|falls|rises|adds|sheds)|52-week|all-time low|historical low|negotiated deals|what you need to know|weekly market cap|weekly traded value|major ownership changes|largest weighted on NomuC|Petrochemical Index|^Gold |^Brent |^Crypto|ex-dividend'),
 ('exclude', None, 'estimated financial results: never in the DMW (8 Oct 2026)', r'estimated (financial )?results|estimated (net )?profit|estimated earnings'),
 ('exclude', None, 'fund rankings / screens', r'equity funds|funds above'),
 ('exclude', None, 'foreign news, no Saudi link', r'^US |^UK |^China '),
 ('include', 'corp', 'weekly flows/ownership opens Corporate', r'foreign ownership|net (buyers|sellers)|foreign institutions'),
 ('include', 'eco', 'official macro data (Economic rank 1)', r'\bPMI\b|GDP|inflation|consumer spending|bank loans|deposits|GASTAT|General Authority for Statistics|budget|reserve assets|remittances|money supply'),
 ('exclude', None, 'board / committee / management change', r'\bCEO\b|chairman|appoint|resign|audit committee|board member'),
 ('exclude', None, 'dividend declaration, any company incl. large-cap banks (none ran 26 Aug to 21 Oct 2026; user, 22 Oct 2026)', r'dividend'),
 ('exclude', None, 'sukuk procedure / default notice', r'sukuk'),
 ('exclude', None, 'lawsuit update / appeal', r'lawsuit|appeals? ruling|court|legal action'),
 ('include', 'corp', 'credit facility: bottom-of-Corporate filler', r'credit facilit|banking facilit|renews? .*facilit|murabaha'),
 ('exclude', None, 'capital increase / reduction / restructuring / buyback (capital reductions never ran; user, 22 Oct 2026)', r'accumulated losses|restructuring|buyback|capital (increase|reduction|hike)'),
 ('include', 'corp', 'rebrand / new trading name of a listed company: bottom of Corporate', r'brand identity|rebrand|new (trading|commercial) name|change (its|of) (trading|commercial) name'),
 ('exclude', None, 'interview / outlook piece', r'sees promising|growth opportunities'),
 ('include', 'corp', 'contract / award / work order', r'contract|awards?|work order|project'),
 ('include', 'corp', 'M&A / stake', r'acquir|stake|merger'),
 ('include', 'corp', 'listed heavyweight / market structure story', r'heavyweight|overtakes'),
 ('include', 'corp', 'large-cap strategic commentary', r"Aramco's .* sees|Aramco CEO|Nasser"),
 ('maybe', 'corp', 'tax / regulatory penalty on a listed name', r'tax assessment|\bfined?\b|\bfines\b|penalt|zakat'),
 ('maybe', 'corp', 'credit rating action on a listed name (23 Nov replay)', r"Moody's|Fitch|S&P|credit rating|\brating\b"),
 ('maybe', 'corp', 'MoU / cooperation without a value', r'MoU|cooperation|agreement'),
 ('exclude', None, 'results-date notice (goes to event calendar)', r'to disclose .* results|results on'),
 ('include', 'corp', 'bottom filler: insurance product approval / showroom / branch / opening', r'Insurance Authority|insurance product|showroom|opens|branch'),
]


def triage(h):
    for v, s, why, rx in RULES:
        if re.search(rx, h, re.I):
            return v, s, why
    return 'maybe', None, 'no rule matched'


if __name__ == "__main__":
    for line in open(sys.argv[1], encoding="utf-8"):
        if line.strip():
            v, s, why = triage(line.strip())
            print(f"{v:8s} {s or '-':5s} {line.strip()}  [{why}]")
