"""Build 'EMAIL MB YYYYMMDD.docx' from the finalized ANB Daily Market Watch - EN deck.
Usage: python build_email_mb.py DAY_FOLDER TEMPLATE_EMAIL_DOCX NAMES_XLSX OUT_DOCX [DMU_PPTX]
"""
import sys, re, copy, openpyxl
from docx import Document
from pptx import Presentation

HEAD_MAP = {'MARKETS PERFORMANCE & NEWS': 'Markets Performance/News',
            'MAJOR ECONOMIC/SECTOR NEWS': 'Major Economic/Sector News',
            'MAJOR CORPORATE NEWS': 'Major Corporate News'}
EMAIL_HEADS = ['Markets Performance/News', 'Major Economic/Sector News', 'Major Corporate News', 'Corporate Calendar']
# 12 Oct 2026 (confirmed final): the Corporate Calendar carries every event row from the
# final EN deck EXCEPT extraordinary general meetings. Listing upgrades, bonus shares, dividends
# and rights rows all go in. The cash-dividend-only rule (29 Sep / 6 Oct) is withdrawn.
DROP_EVENT = r'\bextraordinary\s+general\s+(meeting|assembly)\b|\bEGM\b'
INLINE_BAND = 2.0          # |% variance vs consensus| up to this: 'in line' (inclusive)
TABLE_TOL = 0.15           # tolerance for rounding noise when checking the table's derived columns
BROADLY_INLINE_BAND = 5.0  # up to this: 'broadly in line'; beyond: exceeded / fell short of

def clean(t):
    t = t.replace('\xa0', ' ').replace('​', '').replace('\x0b', ' ')
    t = re.sub(r'[ \t]{2,}', ' ', t)
    return t.strip()

V2_HEADS = {'markets performance & news': 'Markets Performance/News',
            'major economic/sector news': 'Major Economic/Sector News',
            'major corporate news': 'Major Corporate News'}


def read_deck_v2(prs):
    """Template v2 (25 Nov 2026): news are cells of s1.news and s2.flow under their header rows;
    the Results tables are s1.results / s2.results; the Event Calendar is a block of s2.flow."""
    news = {h: [] for h in HEAD_MAP.values()}
    tables, checks = [], []
    cal_rows = []
    for si, s in enumerate(prs.slides, 1):
        if si > 3: break
        for sh in sorted([x for x in s.shapes if x.has_table], key=lambda x: x.top):
            rows = [[c.text for c in r.cells] for r in sh.table.rows]
            if sh.name in ('s1.news', 's2.flow'):
                # results days: Results has its own band (s1.news row 8, s2.results row 0);
                # s2.flow keeps its Corporate band
                sec = 'Major Corporate News' if sh.name == 's2.flow' else None
                for ri, r in enumerate(sh.table.rows):
                    t = clean(r.cells[0].text)
                    if t.lower() in V2_HEADS:
                        sec = V2_HEADS[t.lower()]; continue
                    if t.lower() in ('event calendar', 'economic calendar'):
                        sec = t.lower(); continue
                    if sec in news and t:
                        news[sec] += [clean(p.text) for p in r.cells[0].text_frame.paragraphs if clean(p.text)]
                    elif sec == 'event calendar':
                        cal_rows.append([c.text for c in r.cells])
            elif sh.name in ('s1.results', 's2.results'):
                tables.append(rows)
    # the calendar in the v1 table shape, so calendar_from_tables reads it unchanged
    tables.append([['EVENT CALENDAR'], ['Date', 'Company', 'Event']] + cal_rows)
    return news, tables, checks


def read_deck(path):
    prs = Presentation(path)
    if any(sh.name == 's1.news' for sh in prs.slides[0].shapes):
        return read_deck_v2(prs)
    news = {h: [] for h in HEAD_MAP.values()}
    tables = []
    checks = []
    for si, s in enumerate(prs.slides, 1):
        if si > 3: break
        heads, bodies = [], []
        for sh in s.shapes:
            if sh.has_table:
                tables.append([[c.text for c in r.cells] for r in sh.table.rows])
            if not sh.has_text_frame: continue
            t = sh.text_frame.text.strip()
            if sh.name.startswith('TextBox') and t in HEAD_MAP:
                heads.append((sh.top, HEAD_MAP[t]))
            elif sh.name.startswith('Text 22') and len(t) > 60:
                bodies.append(sh)
        if si > 2: continue
        heads.sort()
        for b in sorted(bodies, key=lambda x: x.top):
            above = [h for top, h in heads if top <= b.top]
            if not above:
                checks.append(f'slide {si}: body "{b.text_frame.text[:40]}" has no header above it'); continue
            sec = above[-1]
            for p in b.text_frame.paragraphs:
                t = clean(p.text)
                if t: news[sec].append(t)
    return news, tables, checks

def load_names(path):
    # 9 Nov 2026: find columns by header (mapcols.py). The fixed indexes (code r[3], name r[1],
    # ticker r[8]) predated the 1 Nov 2026 re-layout, so no names loaded and every result fell
    # back to the table's short name.
    import mapcols
    ws = mapcols.sheet(openpyxl.load_workbook(path, read_only=True, data_only=True))
    ix = mapcols.idx(ws)
    by_code = {}
    for r in ws.iter_rows(min_row=2, values_only=True):
        c = mapcols.get(r, ix, 'code')
        if c is None: continue
        try: code = str(int(c))
        except (TypeError, ValueError): continue
        full = mapcols.get(r, ix, 'en')
        tick = mapcols.get(r, ix, 'bbg2') or mapcols.get(r, ix, 'bbg')
        tick = str(tick).strip().removesuffix(' Equity').strip() if tick else None
        by_code[code] = (str(full).strip() if full else None, tick)
    return by_code

def lookup_name(names, code, bbg):
    # 29 Oct 2026: a company that moved from Nomu to the main market changes TASI code (Armah
    # 9590 -> 6022) before the names sheet catches up; fall back to the Bloomberg ticker
    if code in names: return names[code]
    return next((v for v in names.values() if v[1] == bbg), (None, None))

def num(x):
    x = x.strip().replace(',', '')
    if x in ('', '-', 'NM', 'N/A', 'n/a', 'nm'): return None
    return float(x)

from decimal import Decimal, ROUND_HALF_UP
def hu(v):  # house rounding: 1dp, halves up (0.55 -> 0.6)
    return Decimal(repr(abs(v))).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
def f1(v): return f'{hu(v):,}'

def rounding_range(f, *ins):
    # 26 Oct 2026: the deck shows inputs at 1dp, the table's derived columns come from unrounded
    # figures (AZM 17.299/14.060 -> 23.0%, not 22.7% from 17.3/14.1). Return the min/max of f over
    # every input within +/-0.05 of its shown value, so a table figure inside that range is kept.
    import itertools
    vals = []
    for c in itertools.product(*[(x - 0.05, x, x + 0.05) for x in ins]):
        try: vals.append(f(*c))
        except ZeroDivisionError: pass
    return (min(vals), max(vals)) if vals else None

def result_sentence(row, period, names, checks, foot=None):
    stars = len(row[0].strip()) - len(row[0].strip().rstrip('*'))
    if stars and foot and stars in foot:
        period = foot[stars]   # 26 Oct 2026: '*For 1Q27' / '**For FY26' footnotes set the period
    elif stars:
        checks.append(f'Result {row[0].strip()}: footnote marker with no matching footnote, period left as {period}')
    short, code, bbg = row[0].strip().rstrip('*'), row[1].strip(), row[2].strip()
    cur, pri, chg, pct, cons, var = [num(c) for c in row[3:9]]
    full, bbg2 = lookup_name(names, code, bbg)
    if not full:
        full = short; checks.append(f'Result {short} ({code}): not in names sheet, used table name')
    bbg = bbg or bbg2
    if cur is None or pri is None:
        checks.append(f'Result {short}: missing current/prior figure, sentence skipped'); return None
    # 4 Oct 2026: the Results table does not auto-update, so do not trust its derived columns.
    # Recompute change, change % and variance from the profit and consensus figures; flag any mismatch.
    chg_c = cur - pri
    pct_c = (cur / pri - 1) * 100 if pri and pri > 0 and cur >= 0 else None
    var_c = (cur / cons - 1) * 100 if cons not in (None, 0) else None
    rng = {'change SAR mn': rounding_range(lambda a, b: a - b, cur, pri),
           'change %': rounding_range(lambda a, b: (a / b - 1) * 100, cur, pri) if pct_c is not None else None,
           'variance %': rounding_range(lambda a, c: (a / c - 1) * 100, cur, cons) if var_c is not None else None}
    def ok(lab, tbl, calc):
        r = rng.get(lab)
        return abs(tbl - calc) <= TABLE_TOL or (r is not None and r[0] - TABLE_TOL <= tbl <= r[1] + TABLE_TOL)
    for lab, tbl, calc in (('change SAR mn', chg, chg_c), ('change %', pct, pct_c), ('variance %', var, var_c)):
        if tbl is not None and calc is not None and not ok(lab, tbl, calc):
            checks.append(f'Result {short}: table {lab} {tbl:.1f} vs recomputed {calc:.1f}; sentence uses recomputed, check the table and deck')
    if cur == pri:
        checks.append(f'Result {short}: current and prior-year profit are both SAR {f1(cur)} mn; written as "no change", confirm the prior-year figure')
    # within tolerance or rounding range the table's own figure stands (team rounding calls, e.g.
    # Almarai 4.5 on 4 Oct; unrounded inputs, AZM 26 Oct); beyond it the table is treated as stale
    pick_l = lambda lab, tbl, calc: calc if (calc is not None and (tbl is None or not ok(lab, tbl, calc))) else tbl
    chg, pct = pick_l('change SAR mn', chg, chg_c), pick_l('change %', pct, pct_c)
    band_var = var_c if var_c is not None else var   # wording band always from the recomputed variance
    var = pick_l('variance %', var, var_c)
    kind = 'profit' if cur >= 0 else 'loss'
    s = f'{full} ({bbg}) posted a net {kind} of SAR {f1(cur)} mn for {period}, '
    d = cur - pri if chg is None else chg
    if cur >= 0 and pri >= 0:
        pct_s = f' ({"-" if pct < 0 else ""}{hu(pct)}%)' if pct is not None else ''
        if hu(d) == 0:
            # 26 Oct 2026: flat year on year (Solutions 3Q26 417.0 vs 417.0, Argaam) takes
            # "reflecting no change", not "an increase of SAR 0.0 mn (0.0%)"
            s += f'reflecting no change from the net profit of SAR {f1(pri)} mn in the same period last year.'
        else:
            s += f'reflecting {"an increase" if d >= 0 else "a decline"} of SAR {f1(d)} mn{pct_s} from the net profit of SAR {f1(pri)} mn in the same period last year.'
    elif cur < 0 and pri < 0:
        s += f'reflecting {"an increase in loss" if cur < pri else "a decrease in loss"} of SAR {f1(d)} mn from the net loss of SAR {f1(pri)} mn in the same period last year.'
    elif pri < 0 <= cur:
        # 12 Oct 2026 (email-mb.md): loss -> profit takes the full change clause
        s += f'reflecting a swing to profit of SAR {f1(d)} mn from the net loss of SAR {f1(pri)} mn in the same period last year.'
    else:
        # 12 Oct 2026 (email-mb.md): profit -> loss takes the full change clause
        s += f'reflecting a decrease of SAR {f1(d)} mn from the net profit of SAR {f1(pri)} mn in the same period last year.'
    if cons is not None:
        # consensus losses are written as positive numbers (f1 takes the absolute value)
        if cur >= 0 and cons >= 0:
            # 4 Oct 2026: |variance| <= 2% "was in line with"; <= 5% "was broadly in line with";
            # beyond 5% "exceeded" / "fell short of" (e.g. Almarai 3Q26, -0.6% -> in line)
            v = band_var
            if v is not None and abs(v) <= INLINE_BAND:
                verb = "was in line with"
            elif v is not None and abs(v) <= BROADLY_INLINE_BAND:
                verb = "was broadly in line with"
            else:
                verb = "exceeded" if cur > cons else "fell short of"
            s += f'  The result {verb} the consensus net profit of SAR {f1(cons)} mn'
            if var is not None:
                # 26 Oct 2026 (team edit): "up 5.1%" / "down 0.3%", no "by"
                s += f', {"up" if var >= 0 else "down"} {hu(var)}%'
            s += '.'
        elif cur < 0 and cons < 0:
            # 12 Oct 2026 (email-mb.md): loss vs loss, same bands on the size of the loss, no %
            v = (abs(cur) / abs(cons) - 1) * 100
            if abs(v) <= INLINE_BAND:
                verb = "came in line with"
            elif abs(v) <= BROADLY_INLINE_BAND:
                verb = "came broadly in line with"
            else:
                verb = "came lower than" if abs(cur) < abs(cons) else "came higher than"
            s += f'  The loss {verb} the consensus net loss of SAR {f1(cons)} mn.'
        elif cur >= 0:
            s += f'  The result exceeded the consensus net loss of SAR {f1(cons)} mn.'
        else:
            s += f'  The result fell short of the consensus net profit of SAR {f1(cons)} mn.'
        s += ' (Tadawul, Bloomberg)'
    else:
        # 12 Oct 2026 (email-mb.md): the tag is always (Tadawul, Bloomberg), never (Tadawul) alone
        s += '  (Tadawul, Bloomberg)'
    return s

def read_dmu_coverage(dmu_path, checks, names=None):
    """26 Oct 2026: companies in our coverage carry an anbc estimate in the DMU results table and
    an analyst paragraph on the same slide. Return {bbg: (paragraph, estimate)}."""
    out = {}
    if not dmu_path: return out
    prs = Presentation(dmu_path)
    for s in prs.slides:
        tbl = next((sh for sh in s.shapes if sh.has_table and 'anbc' in ' '.join(c.text.lower() for c in sh.table.rows[0].cells)), None)
        if tbl is None: continue
        rows = [[c.text.strip() for c in r.cells] for r in tbl.table.rows]
        # 29 Oct 2026: one text box can hold several analyst paragraphs (Riyadh Cables and Lumi
        # shared one box, and both email sentences took the whole box). Split into paragraphs and
        # match each company by its Bloomberg ticker in brackets first, then by its table name.
        paras = [p.text.strip() for sh in s.shapes if sh.has_text_frame
                 for p in sh.text_frame.paragraphs if re.search(r'our (?:[a-z-]+ ){0,3}estimate', p.text, re.I)]
        # 9 Nov 2026: Cenomi Centers' paragraph said "our net profit estimate" and named the
        # company only by its full name (Arabian Centres Co.), so it was missed; accept up to three
        # words between "our" and "estimate" and also match the full name from the names sheet.
        for r in rows:
            if len(r) < 11 or not r[1].isdigit(): continue
            est = num(r[9]) if r[9] else None
            if est is None: continue
            para = next((t for t in paras if f'({r[2].strip()})' in t), None) or \
                   next((t for t in paras if r[0].rstrip('*').lower() in t.lower() or r[2].split()[0].lower() in t.lower()), None)
            if not para and names:
                full = (names.get(r[1].strip()) or (None,))[0]
                if full: para = next((t for t in paras if full.lower().rstrip('.') in t.lower()), None)
            if not para:
                checks.append(f'Coverage {r[0]}: anbc estimate {r[9]} on the DMU but no analyst paragraph found'); continue
            out[r[2]] = (para, est, r)
    return out

def coverage_sentence(full, bbg, para, est, cur, cons, var, checks):
    """Team version (26 Oct 2026): DMU paragraph with full name and Bloomberg ticker in front,
    house 1dp figures, then two spaces, the consensus clause and (Tadawul, Bloomberg, anbc research)."""
    body = re.sub(r'^.*?\([A-Z0-9_]+ AB\)', f'{full} ({bbg})', para, count=1)
    if f'({bbg})' not in body:
        # 28 Oct 2026: the DMU paragraph can open with the full name and no ticker
        # ("Nahdi Medical Co. reported ..."); put the ticker after it instead of repeating the name.
        base = bbg.replace(' AB', '').strip()
        if para.lower().startswith(full.lower()):
            body = f'{full} ({bbg})' + para[len(full):]
        elif re.match(re.escape(base) + r'\b', para, re.I):
            # 10 Nov 2026: the DMU paragraph opened with the short name ("MIS reported ..."), which
            # came out as "Al Moammar Information Systems Co. (MIS AB) MIS reported"; replace it.
            body = f'{full} ({bbg})' + para[len(base):]
        else:
            body = f'{full} ({bbg}) ' + para
    # 29 Oct 2026: the DMU writes "SAR 325.1mn" as well as "SAR 11.0 mn"; normalize the spacing
    body = re.sub(r'SAR ([\d,.]+)mn\b', r'SAR \1 mn', body)
    body = re.sub(r'net profit of SAR ([\d,.]+) mn', lambda m: f'net profit of SAR {f1(cur)} mn', body, count=1)
    body = re.sub(r'our estimate of SAR ([\d,.]+) mn', lambda m: f'our estimate of SAR {f1(est)} mn', body, count=1)
    # 10 Nov 2026: read the earnings gap ("24.4% lower than our net profit estimate"), not a
    # revenue gap earlier in the paragraph ("revenue coming 6.9% below our estimate"); allow the
    # DMU's unrounded inputs (MIS 20.8 / 61.8 = -66.3%, DMU 66.4%).
    m = (re.search(r'([\d.]+)% (lower|higher) than our (?:net profit |earnings )?estimate', body)
         or re.search(r'([\d.]+)% (below|above) our (?:net profit |earnings )?estimate of SAR', body))
    v = (cur / est - 1) * 100
    lo, hi = rounding_range(lambda a, b: abs((a / b - 1) * 100), cur, est) or (abs(v), abs(v))
    if m and (not (lo - 0.05 <= float(m.group(1)) <= hi + 0.05) or (m.group(2) in ('below', 'lower')) != (v < 0)):
        checks.append(f'Coverage {bbg}: DMU says {m.group(1)}% {m.group(2)} estimate, recomputed {v:.1f}%; check the DMU paragraph')
    body = body.rstrip()
    if not body.endswith('.'): body += '.'
    tail = ''
    if cons is not None and cur >= 0 and cons >= 0:
        vb = (cur / cons - 1) * 100
        verb = "was in line with" if abs(vb) <= INLINE_BAND else ("was broadly in line with" if abs(vb) <= BROADLY_INLINE_BAND else ("exceeded" if cur > cons else "fell short of"))
        tail = f'  The result {verb} the consensus net profit of SAR {f1(cons)} mn, {"up" if var >= 0 else "down"} {hu(var)}%.'
    return body + tail + ' (Tadawul, Bloomberg, anbc research)'

def period_label(h):
    h = h.strip()
    return 'FY' + h[2:] if re.fullmatch(r'20\d\d', h) else h

def results_from_tables(tables, names, checks, cov=None):
    cov = cov or {}
    out = []
    for t in tables:
        hi = next((i for i, r in enumerate(t[:8]) if len(r) > 3 and 'tasi code' in r[1].lower()), None)
        if hi is None: continue
        period = period_label(t[hi][3])
        foot = {}
        for r in t[hi + 1:]:
            for line in r[0].splitlines():
                m = re.match(r'\s*(\*+)\s*For\s+(\S+)', line, re.I)
                if m: foot[len(m.group(1))] = period_label(m.group(2))
        for r in t[hi + 1:]:
            if not r[1].strip().isdigit(): continue
            s = result_sentence(r, period, names, checks, foot)
            bbg = r[2].strip()
            if s and bbg in cov:
                para, est, drow = cov[bbg]
                cur, cons = num(r[3]), num(r[7])
                full = lookup_name(names, r[1].strip(), bbg)[0] or r[0].strip()
                var = (cur / cons - 1) * 100 if cons else None
                for i, lab in ((3, 'net profit'), (4, 'prior year'), (7, 'consensus'), (8, 'variance')):
                    if num(drow[i]) != num(r[i]):
                        checks.append(f'DMU vs DMW Results table, {r[0].strip()} {lab}: DMU {drow[i]} vs DMW {r[i]}')
                s = coverage_sentence(full, bbg, para, est, cur, cons, var, checks)
            if s: out.append(s)
    return out

def calendar_from_tables(tables, checks):
    lines = []
    for t in tables:
        if not t or not t[0][0].strip().upper().startswith('EVENT CALENDAR'): continue
        for r in t[2:]:
            date, co, ev = (clean(c) for c in r[:3])
            if not re.match(r'\d{2}-[A-Za-z]{3}-\d{2}', date): continue
            if re.search(DROP_EVENT, ev, re.I):
                checks.append(f'Calendar row dropped (extraordinary general meeting): {date} {co}, {ev}'); continue
            lines.append(f'{date}: {co}, {ev}')
    return lines

def write_email(template, sections, out):
    doc = Document(template)
    body = doc.element.body
    paras = list(doc.paragraphs)
    head_tpl = next(p for p in paras if p.text.strip() == 'Markets Performance/News')
    item_tpl = next(p for p in paras if p._p.pPr is not None and p._p.pPr.numPr is not None)
    head_x, item_x = copy.deepcopy(head_tpl._p), copy.deepcopy(item_tpl._p)
    # numIds per section from template so each section keeps its own bullet list
    numids, cur = {}, None
    for p in paras:
        if p.text.strip() in EMAIL_HEADS: cur = p.text.strip()
        elif cur and p._p.pPr is not None and p._p.pPr.numPr is not None and cur not in numids:
            numids[cur] = p._p.pPr.numPr.numId.val
    for p in paras: body.remove(p._p)
    sect = body.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sectPr')
    from docx.text.paragraph import Paragraph
    def add(x, text, numid=None):
        e = copy.deepcopy(x)
        for r in e.findall('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')[1:]: e.remove(r)
        sect.addprevious(e)
        P = Paragraph(e, doc._body)
        P.runs[0].text = text
        if numid is not None: e.pPr.numPr.numId.val = numid
    for h in EMAIL_HEADS:
        items = sections.get(h, [])
        if not items: continue
        add(head_x, h)
        for t in items: add(item_x, t, numids.get(h))
    doc.save(out)

def build(day_dir, template, names_xlsx, out, dmu=None, deck_name='ANB Daily Market Watch - EN.pptx'):
    news, tables, checks = read_deck(f'{day_dir}/{deck_name}')
    names = load_names(names_xlsx)
    cov = read_dmu_coverage(dmu, checks, names)
    res = results_from_tables(tables, names, checks, cov)
    # 26 Oct 2026: cross-check every DMU results row against the DMW Results table
    if dmu:
        dm = {}
        for s_ in Presentation(dmu).slides:
            for sh in s_.shapes:
                if sh.has_table and 'anbc' in ' '.join(c.text.lower() for c in sh.table.rows[0].cells):
                    for r in sh.table.rows:
                        c = [x.text.strip() for x in r.cells]
                        if len(c) > 8 and c[1].isdigit(): dm[c[2]] = c
        # 29 Oct 2026: also flag TASI code differences (Armah 6022 on the DMU vs 9590 on the DMW)
        # and rows on the DMU only. Rows on the DMW only are expected and not flagged: the DMU's
        # cut-off (9:50) is earlier than the DMW's (10:15), so later results miss the DMU (user,
        # 29 Oct 2026: ARDCO and Saudi Cable)
        dw = {r[2].strip(): r for t in tables for r in t if len(r) > 8 and r[1].strip().isdigit()}
        for b in dw:
            if b not in dm: continue
            elif dm[b][1] != dw[b][1].strip(): checks.append(f'DMU vs DMW Results table, {dw[b][0].strip()} TASI code: DMU {dm[b][1]} vs DMW {dw[b][1].strip()}')
        for b in dm:
            if b not in dw: checks.append(f'DMU vs DMW Results table: {dm[b][0]} ({b}) is on the DMU but not the DMW')
        for t in tables:
            for r in t:
                if len(r) > 8 and r[1].strip().isdigit() and r[2].strip() in dm and r[2].strip() not in cov:
                    d_ = dm[r[2].strip()]
                    for i, lab in ((3, 'net profit'), (4, 'prior year'), (5, 'change'), (6, 'change %'), (7, 'consensus'), (8, 'variance')):
                        if num(d_[i]) != num(r[i]):
                            checks.append(f'DMU vs DMW Results table, {r[0].strip()} {lab}: DMU {d_[i]} vs DMW {r[i]}')
    cal = calendar_from_tables(tables, checks)
    sections = {'Markets Performance/News': news['Markets Performance/News'],
                'Major Economic/Sector News': news['Major Economic/Sector News'],
                'Major Corporate News': res + news['Major Corporate News'],
                'Corporate Calendar': cal}
    for h in EMAIL_HEADS[:3]:
        if not sections[h]: checks.append(f'Section empty: {h}')
    if not cal: checks.append('No event calendar rows left after dropping EGMs: Corporate Calendar section omitted')
    write_email(template, sections, out)
    return sections, checks

if __name__ == '__main__':
    # optional 5th argument: the day's Daily Market Update pptx (coverage paragraphs + cross-check)
    s, c = build(*sys.argv[1:6])
    print('\n'.join(c) or 'no checks')
