#!/usr/bin/env python3
"""Shared DMW helpers (formatting, inputs, chart previews, cross-checks, Technicals copy).

Extracted from the v1 build_dmw.py on 25 Nov 2026 so the v2 build and the other scripts share one copy.
The v1 slide builders are kept below for the scripts not yet ported to v2 (check_final, dmw_patch).

Original v1 docstring: Build the ANB Daily Market Watch (EN) deck.

Slides 1-3: numbers, tables, calendars and charts refreshed from the data team's files.
Since 29 Oct 2026 the build starts from the fixed template (../templates/DMW Template - EN.pptx), not
the previous day's deck, so a nudge made in one day's deck never carries over; snap_spec.py puts any
fixed element that is off spec back in place and lists it in the checks (user, 29 Oct 2026).
Slides 4-5: replaced wholesale by the Technicals deck.
Commentary boxes get visible placeholders (text is pasted by hand).
Writes <out>.pptx and <out>_checks.md. Conflicts are reported, never resolved.
"""
import argparse, copy, io, json, re, sys, zipfile
from datetime import datetime, date
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

import openpyxl
from lxml import etree
from pptx import Presentation
from pptx.opc.packuri import PackURI
from pptx.parts.image import ImagePart
from pptx.util import Emu

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS = {"a": A, "r": R, "p": P}
GREEN, RED = "00B050", "FF0000"
HERE = Path(__file__).resolve().parent


def data_file(name):
    """Plugin data files live in shared/data/; a copy staged next to the script also works."""
    p = HERE / name
    return p if p.exists() else HERE.parent / "data" / name

checks = []   # (level, message)  level in OK / FLAG / CONFLICT


def log(level, msg):
    checks.append((level, msg))


# ---------------------------------------------------------------- formatting
def rnd(x, dp):
    q = Decimal(1).scaleb(-dp)
    # Excel-style: clean float noise at 15 significant digits first (so -0.0055*100 is -0.55,
    # not -0.54999999), then round half away from zero.
    return Decimal(f"{float(x):.15g}").quantize(q, rounding=ROUND_HALF_UP)


def num(x, dp=1):
    if x is None or (isinstance(x, str) and x.strip().upper() in ("N/A", "#N/A", "")):
        return "N/A"
    return f"{rnd(x, dp):,.{dp}f}"


def pct(x, dp=1):
    if x is None or isinstance(x, str):
        return "N/A"
    return f"{rnd(x * 100, dp):,.{dp}f}%"


def bps(x):
    if x is None or isinstance(x, str):
        return "N/A"
    v = rnd(x, 0)
    if x < 0:
        return f"({abs(int(v))}) bps"
    return f"{int(v)} bps"


def sign_color(x):
    """House rule seen in published decks: negative -> red (even -0.0%),
    positive that still shows as zero -> default text colour, positive -> green."""
    if x is None or isinstance(x, str):
        return None
    if x < 0:
        return RED
    if x == 0:
        return None
    return "zero_pos"  # resolved by caller against displayed text


# ---------------------------------------------------------------- xml helpers
def set_run_color(rPr, rgb):
    for tag in ("solidFill", "noFill", "gradFill"):
        for el in rPr.findall(f"{{{A}}}{tag}"):
            rPr.remove(el)
    fill = etree.SubElement(rPr, f"{{{A}}}solidFill")
    if rgb is None:
        c = etree.SubElement(fill, f"{{{A}}}schemeClr"); c.set("val", "tx1")
    else:
        c = etree.SubElement(fill, f"{{{A}}}srgbClr"); c.set("val", rgb)
    # solidFill must precede effectLst/latin/cs etc.
    rPr.remove(fill)
    insert_at = 0
    for i, ch in enumerate(rPr):
        if etree.QName(ch).localname in ("ln",):
            insert_at = i + 1
    rPr.insert(insert_at, fill)


def set_cell(cell, text, color="keep"):
    """Replace a cell's text keeping the first run's formatting."""
    tf = cell.text_frame
    p = tf.paragraphs[0]
    runs = p.runs
    if runs:
        runs[0].text = text
        for r in runs[1:]:
            r._r.getparent().remove(r._r)
        rPr = runs[0]._r.get_or_add_rPr()
    else:
        r = p.add_run(); r.text = text
        rPr = r._r.get_or_add_rPr()
        end = p._p.find(f"{{{A}}}endParaRPr")
        if end is not None:
            for k, v in end.attrib.items():
                rPr.set(k, v)
            for ch in end:
                rPr.append(copy.deepcopy(ch))
    for extra in tf.paragraphs[1:]:
        extra._p.getparent().remove(extra._p)
    if color != "keep":
        set_run_color(rPr, color)


def colored(cell, text, raw):
    c = sign_color(raw)
    if c == "zero_pos":
        c = GREEN if re.search(r"[1-9]", text) else None
    set_cell(cell, text, c if raw is not None and not isinstance(raw, str) else None)


def shape_by_id(slide, sid):
    for sh in slide.shapes:
        if sh.shape_id == sid:
            return sh
    return None


def find_table(slide, title):
    for sh in slide.shapes:
        if sh.has_table and sh.table.cell(0, 0).text.strip().upper().startswith(title.upper()):
            return sh
    raise KeyError(f"table '{title}' not found")


def find_text(slide, pattern):
    for sh in slide.shapes:
        if sh.has_text_frame and re.search(pattern, sh.text_frame.text):
            return sh
    return None


def add_image_part(slide, blob, ext):
    pkg = slide.part.package
    partname = pkg.next_partname(f"/ppt/media/image%d.{ext}")
    ctype = {"png": "image/png", "emf": "image/x-emf", "wmf": "image/x-wmf",
             "jpeg": "image/jpeg", "jpg": "image/jpeg", "gif": "image/gif"}[ext]
    part = ImagePart(PackURI(partname), ctype, pkg, blob)
    return slide.part.relate_to(part, f"{R}/image")


def add_picture(slide, blob, ext, x, y, cx, cy, name):
    rId = add_image_part(slide, blob, ext)
    sid = max([sh.shape_id for sh in slide.shapes] + [0]) + 1
    slide.shapes._spTree.add_pic(sid, name, "", rId, x, y, cx, cy)


def remove_shape(sh):
    el = sh._element
    el.getparent().remove(el)


DRIFT_POS = 45_720   # 0.05 in: how far a frame may move before it is reported
DRIFT_REL = 0.03     # 3%: how much a frame may grow or shrink before it is reported


def inherit_slots(old_shapes, fallback, label, reference=None):
    """Chart frame geometry, taken from the template deck's own pictures.

    Since 29 Oct 2026 the template is the fixed one (../templates), already snapped to spec, so
    its frames are the spec frames. `fallback` (the script's built-in slots) is used only when the
    template has no usable frame. `reference`, when given, is compared with the frames and drift
    is reported (kept for a run with --template pointing at an old deck).
    """
    got = sorted(((sh.left, sh.top, sh.width, sh.height) for sh in old_shapes),
                 key=lambda t: t[0])
    if len(got) != len(fallback):
        log("FLAG", "%s: template deck has %d chart frame(s), expected %d. "
                    "Fell back to the script's default geometry." % (label, len(got), len(fallback)))
        return list(fallback)
    for i, (cur, ref) in enumerate(zip(got, reference or []), 1):
        moved = max(abs(cur[0] - ref[0]), abs(cur[1] - ref[1]))
        grew = max(abs(cur[2] - ref[2]) / ref[2], abs(cur[3] - ref[3]) / ref[3]) if ref[2] and ref[3] else 0
        if moved > DRIFT_POS or grew > DRIFT_REL:
            log("FLAG", "%s frame %d has drifted from the reference: moved %.2f in, size off "
                        "%.1f%%. Template geometry kept; check it is deliberate." %
                        (label, i, moved / 914400, grew * 100))
    return got


# Slide 3 sector chart frame: one size on the EN and AR decks. 2 Nov 2026 (user): the chart fills
# its column, the same 3.705 in width as the section header and the slide 3 tables, 2.3 in tall
# (the 3.0 in frame of 29 Oct left the chart too small and Task 11's snap kept reverting the
# user's manual resize). The image stretches to fill it (about 8% on 2 Nov, accepted by the user),
# so the sector chart's stretch is reported only past CHART3_TOL.
CHART3_TOL = 0.15
CHART3_IN = (3.705, 2.3)


def chart3_slot(slot, outer, ref_edge, ref_top, label):
    """The slide 3 chart frame at CHART3_IN, keeping the template frame's top and its outer edge
    ("left" on the EN deck, "right" on the AR deck). Reports drift of that edge or the top
    against the reference (published geometry), never reverts it."""
    x, y, cx, cy = slot
    W, H = int(CHART3_IN[0] * 914400), int(CHART3_IN[1] * 914400)
    edge = x if outer == "left" else x + cx
    moved = max(abs(edge - ref_edge), abs(y - ref_top))
    if moved > DRIFT_POS:
        log("FLAG", "%s frame has drifted from the reference: %s edge or top moved %.2f in. Template "
                    "position kept; check it is deliberate." % (label, outer, moved / 914400))
    if (cx, cy) != (W, H):
        log("OK", "%s frame set to %.3f x %.1f in (was %.3f x %.3f in) so the image keeps its shape."
            % (label, CHART3_IN[0], CHART3_IN[1], cx / 914400, cy / 914400))
    return (x if outer == "left" else x + cx - W), y, W, H


def check_pair(slots, label):
    """The two slide 1 charts must be the same size and sit level. Reported, not changed."""
    (x1, y1, w1, h1), (x2, y2, w2, h2) = slots
    size = max(abs(w1 - w2) / max(w1, w2), abs(h1 - h2) / max(h1, h2))
    if size > 0.01 or abs(y1 - y2) > 18_288:   # 1% or 0.02 in
        log("FLAG", "%s are not the same size / level: sizes %.2fx%.2f vs %.2fx%.2f in, tops %.2f vs %.2f in."
            % (label, w1 / 914400, h1 / 914400, w2 / 914400, h2 / 914400, y1 / 914400, y2 / 914400))


def fit_slot(slot, c, label, tol=0.02):
    """Return the frame geometry, flagging a visibly stretched chart image.

    The cached preview the data team leaves in the Graphs deck is re-rendered at
    whatever aspect their session produced. The frame wins so the charts stay
    the same size on the page, but a stretch past 2% is reported.
    """
    x, y, cx, cy = slot
    native = (emf_aspect(c["blob"]) if c["ext"] == "emf" else None) or \
             (c["width"] / c["height"] if c["height"] > 0 else None)
    if cy > 0 and native:
        frame = cx / cy
        dev = abs(native - frame) / frame
        if dev > tol:
            log("FLAG", "%s: cached image aspect %.3f vs frame aspect %.3f; image "
                        "stretched %.1f%% to fill the frame." % (label, native, frame, dev * 100))
    return x, y, cx, cy



def emf_aspect(blob):
    """Intrinsic aspect of an EMF from its header frame (rclFrame, 0.01 mm).

    This is the shape the chart was drawn at. The OLE box size in the Graphs deck is
    not: the data team's own deck can show the chart squashed or stretched.
    """
    import struct
    try:
        if struct.unpack_from("<I", blob, 0)[0] != 1:
            return None
        l, t, r, b = struct.unpack_from("<iiii", blob, 24)
        return (r - l) / (b - t) if b > t and r > l else None
    except struct.error:
        return None


def emf_text(blob):
    """Text strings drawn in an EMF (EMR_EXTTEXTOUTW records)."""
    import struct
    out, i = [], 0
    while i + 8 <= len(blob):
        t, sz = struct.unpack_from("<II", blob, i)
        if sz < 8:
            break
        if t == 84:
            n, off = struct.unpack_from("<II", blob, i + 44)
            out.append(blob[i + off:i + off + 2 * n].decode("utf-16le", "ignore").strip())
        i += sz
    return [x for x in out if x]


def chart_staleness(charts, g, report_date, session_date, state_path):
    """Flag charts in the Graphs deck that look stale."""
    import hashlib
    state = {}
    if state_path and Path(state_path).exists():
        state = json.loads(Path(state_path).read_text(encoding="utf-8"))
    new_state = {"report_date": f"{report_date}", "hashes": {}}
    named = {}
    if len(charts.get(1, [])) == 2:
        named["TASI Performance"], named["Tadawul P/E & DY"] = charts[1]
    if len(charts.get(3, [])) == 1:
        named["Sector YTD"] = charts[3][0]
    stale = False
    for name, c in named.items():
        h = hashlib.sha1(c["blob"]).hexdigest()
        new_state["hashes"][name] = h
        txt = emf_text(c["blob"]) if c["ext"] == "emf" else []
        prev = state.get("hashes", {}).get(name)
        if name != "Tadawul P/E & DY" and prev == h and state.get("report_date") != f"{report_date}":
            log("STALE", f"{name} chart is identical to the one used for the {state['report_date']} build. Data team has not refreshed it.")
            stale = True
        if name == "TASI Performance" and txt:
            months = [x for x in txt if re.fullmatch(r"[A-Z][a-z]{2}-\d{2}", x)]
            want = f"{session_date:%b-%y}"
            if months and months[-1] != want:
                log("STALE", f"TASI Performance chart ends at {months[-1]}; session month is {want}.")
                stale = True
        if name == "Tadawul P/E & DY" and txt:
            yrs = [x for x in txt if re.fullmatch(r"Dec-\d{2}", x)]
            want = f"Dec-{(report_date.year - 1) % 100:02d}"
            if yrs and yrs[-1] != want:
                log("STALE", f"P/E & DY chart ends at {yrs[-1]}; expected {want}.")
                stale = True
        if name == "Sector YTD" and txt:
            vals, labels, k = [], [], 0
            while k < len(txt):
                x = txt[k]
                if x == "-" and k + 1 < len(txt) and re.fullmatch(r"[\d.]+", txt[k + 1]):
                    vals.append(-float(txt[k + 1])); k += 2
                elif re.fullmatch(r"[\d.]+", x):
                    vals.append(float(x)); k += 1
                else:
                    if x != "%" and not x.endswith("%"):
                        labels.append(x)
                    k += 1
            chart = dict(zip(labels, vals))
            data = {}
            for r in range(77, 82):
                data[str(g.get(f"D{r}")).strip()] = g.get(f"H{r}")
                data[str(g.get(f"K{r}")).strip()] = g.get(f"O{r}")
            checked, bad = 0, []
            for lab, v in chart.items():
                if isinstance(data.get(lab), (int, float)):
                    checked += 1
                    # compare as numbers: the chart prints a tiny negative as "0.0" while the
                    # deck shows "-0.0%" (18 Oct 2026, false STALE on Pharma, Bio & LS)
                    if abs(rnd(data[lab] * 100, 1)) != abs(Decimal(str(v)).quantize(Decimal('0.1'))) or \
                            (rnd(data[lab] * 100, 1) != 0 and rnd(data[lab] * 100, 1) != Decimal(str(v)).quantize(Decimal('0.1'))):
                        bad.append(f"{lab} chart {v}% vs data {pct(data[lab])}")
            if bad:
                log("STALE", "Sector YTD chart does not match today's data: " + "; ".join(bad))
                stale = True
            elif checked:
                log("OK", f"Sector YTD chart agrees with the data file on the {checked} sectors it shares with the slide 3 tables.")
    if not stale and named:
        log("OK", "No stale charts detected.")
    if state_path:
        Path(state_path).write_text(json.dumps(new_state, indent=1), encoding="utf-8")

# ---------------------------------------------------------------- inputs
def load_data(path):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb["PPT-ENG"]
    grid = {}
    for row in ws.iter_rows(min_row=1, max_row=140, max_col=20):
        for c in row:
            if c.value is not None:
                grid[c.coordinate] = c.value
    face_dates = []
    for row in wb["Face"].iter_rows(min_row=1, max_row=8, values_only=True):
        face_dates += [v for v in row if isinstance(v, datetime)]
    return grid, face_dates


def block(g, cols, r0, r1):
    return [[g.get(f"{c}{r}") for c in cols] for r in range(r0, r1 + 1)]



def cal_sheet(path):
    """The calendar sheet and its header row.

    The sheet was 'Calender' with headers in row 3 until Oct 2026; since 10 Oct 2026 the
    file has one sheet 'Sheet1' with headers in row 2. Take 'Calender' if present, else the
    sheet whose rows 1-6 hold two 'Date' headers.
    """
    wb = openpyxl.load_workbook(path, data_only=False)
    sheets = ([wb["Calender"]] if "Calender" in wb.sheetnames else []) + [w for w in wb.worksheets if w.title != "Calender"]
    for ws in sheets:
        for r in range(1, 7):
            if sum(1 for c in ws[r] if c.value is not None and str(c.value).strip().lower() == "date") >= 2:
                if ws.title != "Calender" or r != 3:
                    log("OK", f"Calendar read from sheet '{ws.title}', headers in row {r}.")
                return ws, r
    raise SystemExit("Calendar file: no sheet with two 'Date' headers in rows 1-6.")


def cal_columns(ws, header_row=3):
    """Column numbers of each calendar field, found by the row-3 headers.

    The team edits this sheet by hand and columns shift (on 28 Sep 2026 the event
    block moved from C:E to B:D), so fixed column numbers silently read nothing.
    """
    hdr = {}
    for c in ws[header_row]:
        if c.value is not None and str(c.value).strip():
            hdr.setdefault(str(c.value).strip().lower(), []).append(c.column)
    dates = sorted(hdr.get("date", []))
    if len(dates) < 2:
        raise SystemExit(f"Calendar sheet: expected two 'Date' headers in row {header_row}, found {len(dates)}.")
    d1, d2 = dates[0], dates[1]
    def after(name, lo, hi=10**6):
        cols = [c for c in hdr.get(name, []) if lo < c < hi]
        return min(cols) if cols else None
    cols = dict(ev_date=d1, ev_company=after("company", d1, d2), ev_event=after("event", d1, d2),
                ev_event_ar=after("الحدث", d1, d2), ev_company_ar=after("الشركة", d1, d2),
                ec_date=d2, ec_event=after("event", d2), ec_prior=after("prior", d2),
                ec_est=after("estimation", d2) or after("estimate", d2), ec_event_ar=after("الحدث", d2))
    missing = [k for k, v in cols.items() if v is None]
    if missing:
        raise SystemExit(f"Calendar sheet: headers not found for {missing}.")
    usual = dict(ev_date=2, ev_company=3, ev_event=4, ev_event_ar=7, ev_company_ar=8,
                 ec_date=14, ec_event=15, ec_prior=16, ec_est=17, ec_event_ar=21)
    moved = [k for k in usual if cols[k] != usual[k]]
    if moved:
        from openpyxl.utils import get_column_letter as L
        log("FLAG", "Calendar file columns have moved (" + ", ".join(
            f"{k} {L(usual[k])}->{L(cols[k])}" for k in moved) + "). Read by header, check the tables.")
    return cols

CLIENT_ECON_WORDING_ALL = {
        "US API Report on Crude Oil Inventories - Weekly": "US API Weekly Report on Crude Oil Inventories",
        "Issuance of API weekly report on the US crude oil inventories": "US API Weekly Report on Crude Oil Inventories",
        # Wednesday EIA row (Muryum via Saad, 25 Nov 2026): "US Crude Oil Inventories" -> client form
        "US Crude Oil Inventories": "US EIA Crude Oil Inventories",
    }


def load_calendar(path, target):
    ws, hr = cal_sheet(path)
    def fmt(cell):
        v = cell.value
        if v is None or (isinstance(v, str) and v.strip() in ("", "_", "-", "–")):
            return "-"  # blanks, and the '_' the calendar file uses since 10 Oct 2026, show as '-'
        if isinstance(v, (int, float)):
            if "%" in cell.number_format:
                return pct(v)
            return num(v, 1)
        return str(v).strip()

    def fmt_val(cell):
        """Prior / estimate: one decimal, half-up, suffix kept (3.98M -> 4.0M). User, 13 Oct 2026."""
        t = fmt(cell)
        m = re.fullmatch(r"(-?\d+(?:\.\d+)?)([KMBT%]?)", t)
        if not m:
            return t
        from decimal import Decimal, ROUND_HALF_UP
        return f"{Decimal(m.group(1)).quantize(Decimal('0.1'), ROUND_HALF_UP)}{m.group(2)}"
    k = cal_columns(ws, hr)
    events, econ = [], []
    # Client-mandated wording for standing economic rows. Applied on read, whatever the sheet says.
    CLIENT_ECON_WORDING = {
        "US API Report on Crude Oil Inventories - Weekly": "US API Weekly Report on Crude Oil Inventories",
        "Issuance of API weekly report on the US crude oil inventories": "US API Weekly Report on Crude Oil Inventories",
        # Wednesday EIA row (Muryum via Saad, 25 Nov 2026): "US Crude Oil Inventories" -> client form
        "US Crude Oil Inventories": "US EIA Crude Oil Inventories",
    }
    def econ_name(cell):
        t = fmt(cell)
        if t in CLIENT_ECON_WORDING:
            log("OK", f"Economic calendar wording set to client form: '{t}' -> '{CLIENT_ECON_WORDING[t]}'.")
            return CLIENT_ECON_WORDING[t]
        return t
    r = hr + 1
    while isinstance(ws.cell(r, k["ev_date"]).value, datetime):
        events.append((ws.cell(r, k["ev_date"]).value, fmt(ws.cell(r, k["ev_company"])), fmt(ws.cell(r, k["ev_event"])))); r += 1
    r = hr + 1
    while isinstance(ws.cell(r, k["ec_date"]).value, datetime):
        econ.append((ws.cell(r, k["ec_date"]).value, econ_name(ws.cell(r, k["ec_event"])), fmt_val(ws.cell(r, k["ec_prior"])), fmt_val(ws.cell(r, k["ec_est"])))); r += 1
    for label, rows in (("Event", events), ("Economic", econ)):
        if not rows:
            log("FLAG", f"{label} calendar: no dated rows found in the calendar file. Table left empty; confirm there really are none today.")
    for d, *_ in events + econ:
        if d.date() != target:
            log("FLAG", f"Calendar row dated {d:%d-%b-%y} does not match report date {target:%d-%b-%y}.")
    return events, econ


def ole_previews(pptx_path):
    """Cached chart images behind the linked Excel objects in the Graphs deck."""
    prs = Presentation(pptx_path)
    out = {}
    for si, slide in enumerate(prs.slides, 1):
        items = []
        for sh in slide.shapes:
            el = sh._element
            blips = el.findall(".//p:oleObj//a:blip", NS) or el.findall(".//a:blip", NS)
            if not el.findall(".//p:oleObj", NS) or not blips:
                continue
            rId = blips[0].get(f"{{{R}}}embed")
            part = slide.part.related_part(rId)
            ext = part.partname.ext
            items.append(dict(left=sh.left, top=sh.top, width=sh.width, height=sh.height,
                              blob=part.blob, ext=ext, name=sh.name))
        out[si] = sorted(items, key=lambda i: i["left"])
    date_txt = find_text(prs.slides[0], r"\b20\d\d\b$")
    return out, (date_txt.text_frame.text.strip() if date_txt else None)


# ---------------------------------------------------------------- slide builders
def build_slide1(slide, g, charts, report_date, results_keep):
    dt = find_text(slide, r"^(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday), ")
    set_text_keep(dt, f"{report_date:%A, %B} {report_date.day}, {report_date.year}")

    # headline boxes: two cards, TASI left, NOMU right
    cards = {"TASI": dict(close=g["C6"], vt=g["F6"], ytd=g["E8"], dod=g["C8"], unit="bn"),
             "NOMU": dict(close=g["L6"], vt=g["O6"], ytd=g["N8"], dod=g["L8"], unit="mn")}
    pairs = card_pairs(slide)   # EN: TASI card on the left, NOMU on the right
    nomu_gap = card_gap(*pairs["right"]) if "right" in pairs else None
    for key, left_max in (("TASI", 3_900_000), ("NOMU", 10**9)):
        d = cards[key]
        grp = [sh for sh in slide.shapes if sh.has_text_frame and 1_100_000 < sh.top < 2_150_000
               and (sh.left < 3_900_000) == (key == "TASI")]
        for sh in grp:
            t = sh.text_frame.text
            runs = sh.text_frame.paragraphs[0].runs
            if re.match(r"^\s*[\d,]+\.\d{2}", t):
                runs[0].text = f"{rnd(d['close'], 2):,.2f}" + re.sub(r"^[^\t]*", "", runs[0].text)
            elif t.startswith("Value Traded"):
                val = f"SAR {num(d['vt'], 1)} {d['unit']}"
                # the value is the first run with a digit or "SAR"; everything after it goes. Writing
                # the last run only gave "SAR 8.1 SAR 11.7 mn" when the template split the value into
                # 'SAR 8.1 ' + 'mn' (back-test, 4 Aug 2026).
                k = next((i for i, r in enumerate(runs) if re.search(r"\d|SAR", r.text)), len(runs) - 1)
                if k == 0:
                    runs[0].text = re.match(r"Value Traded\s*", runs[0].text).group(0) + val
                else:
                    runs[k].text = val
                for r in runs[k + 1:]:
                    r._r.getparent().remove(r._r)
            elif t.startswith("YTD"):
                txt = pct(d["ytd"]); runs[-1].text = txt
                set_run_color(runs[-1]._r.get_or_add_rPr(), RED if d["ytd"] < 0 else (GREEN if re.search(r"[1-9]", txt) else None))
            elif re.match(r"^\s*\(?-?[\d.]+%\)?\s*$", t):
                txt = pct(d["dod"]); txt = txt if d["dod"] < 0 else " " + txt
                runs[0].text = txt
                # no wrap: '-1.0%' wrapped in the 0.50 in NOMU box (20 Oct 2026); the box grows right instead
                sh.text_frame.word_wrap = False
                for r in runs[1:]:
                    r._r.getparent().remove(r._r)
                # house rule: a positive change that shows as 0.0% stays default colour (18 Oct 2026, NOMU +0.05%)
                set_run_color(runs[0]._r.get_or_add_rPr(), RED if d["dod"] < 0 else (GREEN if re.search(r"[1-9]", txt) else None))
    if len(pairs) == 2:
        place_dod(pairs, nomu_gap, "Slide 1", log)
    else:
        log("FLAG", "Slide 1: headline card close/DoD boxes not found on both cards; DoD boxes not placed.")

    # Index Performance table, row 4
    tbl = [sh for sh in slide.shapes if sh.has_table and sh.table.cell(0, 0).text.startswith("Main Market")][0].table
    vals = [g[f"{c}8"] for c in "CDEFGHIJKLMNOPQRST"]
    for i, v in enumerate(vals):
        colored(tbl.cell(3, i), pct(v), v)

    # charts: reuse the template's own frames, then swap in the Graphs-deck previews
    band = [sh for sh in slide.shapes if sh.shape_type == 13 and 3_300_000 < sh.top < 5_700_000]
    ch = charts.get(1, [])
    if len(ch) != 2:
        log("FLAG", f"Graphs deck slide 1: expected 2 charts, found {len(ch)}. Charts not updated.")
    else:
        fb = [(x, y, cx, int(cx * c["height"] / c["width"]))
              for (x, y, cx), c in zip([(382088, 3565638, 3367671), (4010959, 3559594, 3388312)], ch)]
        # reference = published 28 Sep 2026 geometry (both frames equal, level)
        slots = inherit_slots(band, fb, "Slide 1 charts")
        check_pair(slots, "Slide 1 charts")
        for sh in band:
            remove_shape(sh)
        for c, slot, side in zip(ch, slots, ("left", "right")):
            geom = fit_slot(slot, c, f"Slide 1 {side} chart")
            add_picture(slide, c["blob"], c["ext"], *geom, "DMW chart")
        log("OK", "Slide 1 charts (TASI Performance, P/E & DY) replaced from Graphs deck.")

    # results table: yesterday's is stale
    for sh in list(slide.shapes):
        if sh.has_table and sh.table.cell(0, 0).text.strip().lower().startswith("result"):
            if results_keep:
                log("FLAG", "Results Announcements table kept from template: confirm it is today's.")
            else:
                remove_shape(sh)
                log("OK", "Removed yesterday's Results Announcements table (add by hand on days with results).")


def set_text_keep(sh, text):
    """Set text of a one-line box, keeping formatting. Handles date fields (a:fld)."""
    ts = sh.text_frame._txBody.findall(".//a:t", NS)
    ts[0].text = text
    for t in ts[1:]:
        t.text = ""


def placeholder_paragraphs(sh, labels):
    """Replace a commentary box with highlighted placeholders, one bullet each."""
    txBody = sh.text_frame._txBody
    paras = txBody.findall(f"{{{A}}}p")
    proto = copy.deepcopy(paras[0])
    for p_ in paras:
        txBody.remove(p_)
    rpr_src = proto.find(f"{{{A}}}r/{{{A}}}rPr")
    if rpr_src is None:
        rpr_src = proto.find(f"{{{A}}}endParaRPr")
    for lab in labels:
        p_ = etree.SubElement(txBody, f"{{{A}}}p")
        ppr = proto.find(f"{{{A}}}pPr")
        if ppr is not None:
            p_.append(copy.deepcopy(ppr))
        r = etree.SubElement(p_, f"{{{A}}}r")
        rpr = etree.SubElement(r, f"{{{A}}}rPr")
        if rpr_src is not None:
            for k, v in rpr_src.attrib.items():
                rpr.set(k, v)
            for chd in rpr_src:
                if etree.QName(chd).localname in ("latin", "cs", "ea"):
                    rpr.append(copy.deepcopy(chd))
        hl = etree.Element(f"{{{A}}}highlight"); c = etree.SubElement(hl, f"{{{A}}}srgbClr"); c.set("val", "FFFF00")
        rpr.insert(0, hl)
        t = etree.SubElement(r, f"{{{A}}}t"); t.text = lab


def build_commentary(prs):
    s1, s2 = prs.slides[0], prs.slides[1]
    # slide 1: text boxes sitting under each section header
    headers = {sh.text_frame.text.strip(): sh for sh in s1.shapes
               if sh.shape_type == 17 and sh.has_text_frame}
    bodies = sorted([sh for sh in s1.shapes if sh.shape_type == 1 and sh.has_text_frame
                     and sh.width > 6_500_000 and sh.height > 300_000], key=lambda s: s.top)
    hdr_order = sorted(headers.values(), key=lambda h: h.top)
    def body_under(header_txt):
        h = headers.get(header_txt)
        if h is None:
            return None
        nxt = [x.top for x in hdr_order if x.top > h.top + 50_000]
        limit = nxt[0] if nxt else 9_700_000
        cands = [b for b in bodies if h.top - 400_000 <= b.top < limit]
        if not cands:
            return None
        keep = cands[0]
        for extra in cands[1:]:          # stale duplicate bodies in the same section
            remove_shape(extra); bodies.remove(extra)
        keep.top = Emu(h.top + 165_000)
        keep.height = Emu(max(limit - keep.top - 40_000, 300_000))
        return keep
    for hdr, labels in (("MARKETS PERFORMANCE & NEWS", ["[GLOBAL MARKETS BULLET – paste]", "[TASI BULLET – paste]"]),
                        ("MAJOR ECONOMIC/SECTOR NEWS", ["[ECONOMIC / SECTOR NEWS – paste]"]),
                        ("MAJOR CORPORATE NEWS", ["[CORPORATE NEWS – paste, continues on slide 2]"])):
        b = body_under(hdr)
        if b is None and bodies and hdr in headers:
            # section body missing in template (e.g. a Results table sat there): clone one
            src = bodies[0]._element
            el = copy.deepcopy(src); src.getparent().append(el)
            el.find('.//p:cNvPr', NS).set('id', str(max(sh.shape_id for sh in s1.shapes) + 1))
            b = [sh for sh in s1.shapes if sh._element is el][0]
            h = headers[hdr]
            nxt = [x.top for x in hdr_order if x.top > h.top + 50_000]
            limit = nxt[0] if nxt else 9_650_000
            b.top = Emu(h.top + 165_000); b.height = Emu(max(limit - b.top - 40_000, 300_000))
            log("FLAG", f"Slide 1: no text box under '{hdr}' in template; added one.")
        if b is None:
            log("FLAG", f"Slide 1: no text box found under '{hdr}'.")
            continue
        placeholder_paragraphs(b, labels)
    big = max((sh for sh in s2.shapes if sh.has_text_frame and not sh.has_table), key=lambda s: s.height)
    placeholder_paragraphs(big, ["[CORPORATE NEWS (cont.) – paste]"])
    log("OK", "Commentary boxes on slides 1–2 set to yellow placeholders.")


def resize_table(gf, n_data, first_data_row, fill_rows, last_is_source=True):
    """Make the table hold exactly n_data data rows by cloning/removing rows."""
    tbl = gf.table._tbl
    trs = tbl.findall(f"{{{A}}}tr")
    data = trs[first_data_row:len(trs) - (1 if last_is_source else 0)]
    proto = data[0]
    for tr in data[1:]:
        tbl.remove(tr)
    anchor = proto
    for _ in range(n_data - 1):
        new = copy.deepcopy(proto)
        for ext in new.findall(f"{{{A}}}extLst"):
            new.remove(ext)
        anchor.addnext(new); anchor = new
    if n_data == 0:
        tbl.remove(proto)
    for i, row in enumerate(fill_rows):
        r = first_data_row + i
        for j, v in enumerate(row):
            set_cell(gf.table.cell(r, j), v)


def drop_results_s2(slide, is_results, header, keep=False, after=0.03):
    """Slide 2 Results continuation table. On heavy results days the team runs the table under the
    slide 1 Corporate header and continues it in a second table under the slide 2 Corporate header,
    above the Corporate news (21 Jul to 21 Oct 2026 finals). The template's copy is yesterday's:
    remove it and bring the Corporate news back up under its header, as on a day without results.
    Task 10 adds today's tables after the commentary is in."""
    E = 914400
    for sh in list(slide.shapes):
        if sh.has_table and is_results(sh.table.cell(0, 0).text):
            if keep:
                log("FLAG", "Slide 2 Results table kept from template: confirm it is today's."); continue
            t_top = sh.top
            remove_shape(sh)
            hd = [x for x in flat_shapes(slide.shapes) if x.has_text_frame and not getattr(x, "has_table", False)
                  and x.text_frame.text.strip() == header]
            if hd:
                h = min(hd, key=lambda x: x.top)
                for b in flat_shapes(slide.shapes):
                    if (b.has_text_frame and not getattr(b, "has_table", False) and b.width > 6 * E
                            and b.top > t_top and b.top < 3 * E and len(b.text_frame.text.strip()) > 0
                            and b.text_frame.text.strip() != header):
                        b.top = Emu(int(h.top + h.height + after * E))
            log("OK", "Removed yesterday's slide 2 Results table (continuation); Corporate news moved up "
                      "under its header. Task 10 adds today's on results days.")


def build_slide2(slide, events, econ):
    ev = find_table(slide, "EVENT CALENDAR")
    ec = find_table(slide, "ECONOMIC CALENDAR")
    resize_table(ev, len(events), 2, [(f"{d:%d-%b-%y}", c, e) for d, c, e in events])
    resize_table(ec, len(econ), 2, [(f"{d:%d-%b-%y}", e, p, s) for d, e, p, s in econ])
    # morning placement: the event table keeps the template's top, the economic table
    # follows it at the house gap. fill_commentary.py re-flows both under the news later.
    ROW = 162000
    ev_h = 434340 + ROW * len(events) + sum(ROW // 2 for _, c, e in events if len(c) + len(e) > 95)
    ec_h = 472440 + ROW * len(econ) + sum(ROW // 2 for _, e, *_ in econ if len(e) > 70)
    stack_calendars(slide, ev, ec, ev_h, ec_h)
    log("OK", f"Slide 2 calendars: {len(events)} events, {len(econ)} economic releases. Stacked at the house gap; fill_commentary re-flows them under the news.")


# ---------------------------------------------------------------- house spacing (shared with build_dmw_ar.py)
SECTION_GAP = 0.10          # inches, visible gap between consecutive blocks (text, header bar, table)
PAGE_BOTTOM = 10.60         # inches, nothing may end below this (page number starts at 10.66)


def stack_calendars(slide, ev, ec, ev_h, ec_h):
    """Place the event table, then the economic table SECTION_GAP below it. A table is never
    taken as shorter than the sum of its own row heights: the AR template's rows (about
    0.19 in) are taller than the EN-based estimate, which made the AR event table overlap
    the economic calendar header on 6 Oct 2026."""
    E = 914400
    ev_h = max(ev_h, sum(r.height for r in ev.table.rows))
    ec_h = max(ec_h, sum(r.height for r in ec.table.rows))
    top = ev.top
    if top + ev_h + SECTION_GAP * E + ec_h > PAGE_BOTTOM * E:
        top = int(PAGE_BOTTOM * E - ec_h - SECTION_GAP * E - ev_h)
    ev.top = Emu(top); ev.height = Emu(ev_h)
    ec.top = Emu(int(top + ev_h + SECTION_GAP * E)); ec.height = Emu(ec_h)
    big = max((sh for sh in flat_shapes(slide.shapes) if sh.has_text_frame and not getattr(sh, "has_table", False)),
              key=lambda s: s.height)
    big.height = Emu(max(ev.top - big.top - int(SECTION_GAP * E), 500_000))


def flat_shapes(shapes):
    """Shapes with groups opened up. The AR team grouped each news header with its body
    (15 Oct 2026); without this the slide 2 'biggest text box' became the page header."""
    for sh in shapes:
        if sh.shape_type == 6:
            yield from flat_shapes(sh.shapes)
        else:
            yield sh


SECTION_TITLES = {"MARKETS PERFORMANCE & NEWS", "MAJOR ECONOMIC/SECTOR NEWS", "MAJOR CORPORATE NEWS",
                  "EVENT CALENDAR", "ECONOMIC CALENDAR"}


def clean_hidden_titles(prs, logf=None):
    """Delete leftover section titles from an old template: unfilled text boxes whose text is a
    section title written in white, so they are invisible today but would show if the bar or
    table covering them moved. Found on 5 Oct 2026: EN slide 1 and 2 and slide 3, AR slide 3."""
    n = 0
    for slide in list(prs.slides)[:3]:
        for sh in list(slide.shapes):
            if sh.has_table or not sh.has_text_frame or sh.text_frame.text.strip() not in SECTION_TITLES:
                continue
            sppr = sh._element.find(f"{{{P}}}spPr")
            if sppr is None or sppr.find(f"{{{A}}}solidFill") is not None or sppr.find(f"{{{A}}}gradFill") is not None:
                continue
            cols = [c.get("val") for c in sh._element.iter(f"{{{A}}}srgbClr")]
            if cols and all(c.upper() == "FFFFFF" for c in cols):
                remove_shape(sh); n += 1
    if n and logf:
        logf("OK", f"Removed {n} invisible leftover section title(s) from the template.")
    return n


_WIDTHS = None


def text_width_in(text, pt, face="segoe"):
    """Rendered width in inches of one line of text, from the real font widths read out of the
    team's PowerPoint PDFs (font_widths.json, 1/1000 em). Unknown characters count as 0.6 em."""
    global _WIDTHS
    if _WIDTHS is None:
        _WIDTHS = json.loads(data_file("font_widths.json").read_text(encoding="utf-8"))
    w = _WIDTHS[face]
    return sum(w.get(str(ord(ch)), 600) for ch in text) * pt / 1000 / 72


# ---------------------------------------------------------------- headline cards: DoD box placement
# TASI crossed 10,000 on 29 Oct 2026 and the close ran into the DoD box (fixed by hand: EN
# 1.35 -> 1.47 in, AR 6.485 -> 6.4 in). The DoD box is now placed from the measured width of the
# close (shared with build_dmw_ar.py).
# House rule (user, 5 Nov 2026): ONE gap for both cards in both decks, DOD_GAP_IN from the close's
# last ink to the DoD's first ink (sign or digit, glyph side bearings from font_widths.json; the EN
# leading space on a positive change is not counted), the DoD box sized to its text, and the DoD
# top at DOD_TOP_IN. Until 4 Nov the gap was read from each
# template's NOMU card (EN 0.042 in to the space slot, AR 0.052 in), so EN and AR differed and
# the EN gap widened on up days.
DOD_GAP_IN = 0.05
APERCU_FAKE_BOLD = 23.0     # 1/1000 em added per glyph by PowerPoint's synthetic bold (5 Nov 2026)
DOD_TOP_IN = 1.67           # user, 5 Nov 2026 (the 4 Nov 1.68 in was a slip)
_GENERIC_EM = {**{str(d): 550 for d in range(10)}, ",": 270, ".": 270, "%": 800, " ": 250,
               "-": 330, "(": 300, ")": 300}


def _run_width_in(text, run, default_pt=16.0):
    """Width of `text` set in `run`'s font. Segoe UI uses the measured widths; any other face
    (the AR deck's Apercu) uses generic Latin-figure widths. The error cancels in the placement,
    because the gap is measured on the NOMU card with the same function."""
    pt = run.font.size.pt if run is not None and run.font.size else default_pt
    rPr = run._r.find(f"{{{A}}}rPr") if run is not None else None
    lat = rPr.find(f"{{{A}}}latin") if rPr is not None else None
    face = (lat.get("typeface") if lat is not None else "") or "Segoe UI"
    bold = rPr is not None and rPr.get("b") == "1"
    if "segoe" in face.lower():
        return text_width_in(text, pt, "segoe_bold" if bold else "segoe")
    if "apercu" in face.lower():
        # the AR card's figures render in anb Apercu on the user's PC (5 Nov 2026 PDF: anbApercu-Regular
        # with PowerPoint's bold); the 29 Oct / 4 Nov PDFs came from a PC without it (Calibri Bold),
        # which put the AR TASI % 0.07 in off. Widths and bearings read from anbApercu-Regular.otf.
        w = text_width_in(text, pt, "apercu")
        if bold:
            # PowerPoint fakes the bold (no anb Apercu Bold file): each glyph is about 0.023 em wider
            # (5 Nov 2026 PDF: "10,038.89" 0.972 in drawn vs 0.925 in plain at 16 pt)
            w += APERCU_FAKE_BOLD * len(text) * pt / 1000 / 72
        return w
    if "calibri" in face.lower() and bold:
        return text_width_in(text, pt, "calibri_bold")
    return sum(_GENERIC_EM.get(ch, 600) for ch in text) * pt / 1000 / 72


def _ink_key(run):
    """font_widths.json ink table for a headline run: Segoe UI Bold (EN) or Calibri Bold (the AR
    card's anb Apercu figures render in Calibri Bold in PowerPoint's PDF), else None."""
    rPr = run._r.find(f"{{{A}}}rPr") if run is not None else None
    lat = rPr.find(f"{{{A}}}latin") if rPr is not None else None
    face = ((lat.get("typeface") if lat is not None else "") or "Segoe UI").lower()
    bold = rPr is not None and rPr.get("b") == "1"
    if "segoe" in face and bold:
        return "segoe_bold_ink"
    if "apercu" in face:
        return "apercu_ink"
    if "calibri" in face and bold:
        return "calibri_bold_ink"
    return None


def _bearing_in(run, ch, side):
    """Side bearing (left 0 / right 1) of `ch` in inches, from the glyph outlines (5 Nov 2026)."""
    global _WIDTHS
    if _WIDTHS is None:
        _WIDTHS = json.loads(data_file("font_widths.json").read_text(encoding="utf-8"))
    key = _ink_key(run)
    pt = run.font.size.pt if run is not None and run.font.size else 16.0
    if not key or key not in _WIDTHS:
        return 0.0
    return _WIDTHS[key].get(str(ord(ch)), [0, 0])[side] * pt / 1000 / 72


def _visible_span(sh, text, keep_lead=False, ink=False):
    """(start, end) in EMU of the visible (space-trimmed) text of a one-line box. keep_lead keeps
    leading spaces as part of the text: the DoD's " 2.3%" space is the slot a minus sign takes, so
    the house layout keeps the box still when the sign flips (29 Oct 2026 decks)."""
    if keep_lead:
        text = text.rstrip()
    runs = sh.text_frame.paragraphs[0].runs
    run = runs[0] if runs else None
    bp = sh.text_frame._txBody.find(f"{{{A}}}bodyPr")
    lI = int(bp.get("lIns", 91440)) if bp is not None else 91440
    rI = int(bp.get("rIns", 91440)) if bp is not None else 91440
    ppr = sh.text_frame.paragraphs[0]._p.find(f"{{{A}}}pPr")
    algn = (ppr.get("algn") if ppr is not None else None) or "l"
    E = 914400
    full = _run_width_in(text, run) * E
    lead = 0 if keep_lead else _run_width_in(text[:len(text) - len(text.lstrip())], run) * E
    core = _run_width_in(text if keep_lead else text.strip(), run) * E
    if algn == "r":
        start = sh.left + sh.width - rI - full
    elif algn == "ctr":
        start = sh.left + lI + (sh.width - lI - rI - full) / 2
    else:
        start = sh.left + lI
    s0, s1 = int(start + lead), int(start + lead + core)
    if ink and text.strip():
        vis = text.strip()
        s0 += int(_bearing_in(run, vis[0], 0) * E)
        s1 -= int(_bearing_in(run, vis[-1], 1) * E)
    return s0, s1


def _fit_width(sh):
    """Size a one-line DoD box to its text plus insets, so the renderer's alignment of a wide
    no-wrap box (LibreOffice centres it; 5 Nov 2026 the TASI box was 0.09 in wider than its
    text and the figure sat 0.05 in off) cannot move the figure."""
    E = 914400
    runs = sh.text_frame.paragraphs[0].runs
    run = runs[0] if runs else None
    bp = sh.text_frame._txBody.find(f"{{{A}}}bodyPr")
    lI = int(bp.get("lIns", 91440)) if bp is not None else 91440
    rI = int(bp.get("rIns", 91440)) if bp is not None else 91440
    ppr = sh.text_frame.paragraphs[0]._p.find(f"{{{A}}}pPr")
    algn = (ppr.get("algn") if ppr is not None else None) or "l"
    text = sh.text_frame.text
    # no wrap: a fitted AR box with wrap on dropped the % to a second line when the font rendered
    # wider than measured (5 Nov 2026, "0.4" with the % cut off); 0.03 in slack on top
    sh.text_frame.word_wrap = False
    w = lI + rI + int(_run_width_in(text, run) * E) + int(0.03 * E)
    if algn == "r":                       # keep the right edge where it is
        sh.left = Emu(sh.left + sh.width - w)
    sh.width = Emu(w)


def card_pairs(slide):
    """{"left"/"right": (close box, DoD box)} for the two slide 1 headline cards. The close is
    the number box (9,761.78), the DoD the percent-only box above the YTD line."""
    out = {}
    for sh in slide.shapes:
        if not sh.has_text_frame or not (1_100_000 < sh.top < 2_150_000):
            continue
        side = "left" if sh.left < 3_850_000 else "right"
        t = sh.text_frame.text
        if re.match(r"^\s*[\d,]+\.\d{2}", t):
            out.setdefault(side, [None, None])[0] = sh
        elif re.fullmatch(r"\s*\(?-?[\d.]+%\)?\s*", t) and sh.top < 1_850_000:
            out.setdefault(side, [None, None])[1] = sh
    return {k: tuple(v) for k, v in out.items() if all(v)}


def _close_text(sh):
    return re.match(r"\s*([\d,]+\.\d{2})", sh.text_frame.text).group(1)


def card_gap(close, dod):
    """Gap in EMU between the close and the DoD figure on one card, as the template has it."""
    cs, ce = _visible_span(close, _close_text(close))
    ds, de = _visible_span(dod, dod.text_frame.text, keep_lead=True)
    right = ds + de > cs + ce   # by the text, not the box: the EN close box is 3.5 in wide (tabs)
    return (ds - ce) if right else (cs - de)


def place_dod(pairs, gap=None, label="Slide 1", logf=None):
    """Place each card's DoD box by the house rule: its first visible character DOD_GAP_IN from
    the close's last ink, its top at DOD_TOP_IN. `gap` is ignored (kept for old callers)."""
    E = 914400
    gap = int(DOD_GAP_IN * E)
    for side, (close, dod) in pairs.items():
        _fit_width(dod)
        if abs(dod.top - int(DOD_TOP_IN * E)) >= 914:
            old_t = dod.top
            dod.top = Emu(int(DOD_TOP_IN * E))
            logf("OK", f"{label}: {side} card DoD box top {old_t / E:.3f} -> {DOD_TOP_IN:.2f} in.")
        cs, ce = _visible_span(close, _close_text(close), ink=True)
        ds, de = _visible_span(dod, dod.text_frame.text, ink=True)
        right = ds + de > cs + ce
        shift = int((ce + gap - ds) if right else (cs - gap - de))
        if abs(shift) >= 914:      # 0.001 in
            old = dod.left
            dod.left = Emu(dod.left + shift)
            logf("OK", f"{label}: {side} card DoD box moved {old / E:.3f} -> {dod.left / E:.3f} in to sit "
                       f"{DOD_GAP_IN:.3f} in from the close ({_close_text(close)}).")
        if abs(shift) > 0.35 * E:
            logf("FLAG", f"{label}: {side} card DoD box moved {shift / E:.2f} in; check the card by eye.")


def dod_gaps(pairs):
    """{side: (gap_in, top_in)} as the deck has them, by the house-rule measure (for checks)."""
    E = 914400; out = {}
    for side, (close, dod) in pairs.items():
        cs, ce = _visible_span(close, _close_text(close), ink=True)
        ds, de = _visible_span(dod, dod.text_frame.text, ink=True)
        g = (ds - ce) if ds + de > cs + ce else (cs - de)
        out[side] = (g / E, dod.top / E)
    return out


def fill_block(gf, start_row, rows, spec, label_check=True, names=None):
    t = gf.table
    for i, row in enumerate(rows):
        r = start_row + i
        label = row[0]
        if names is not None:
            label = names.get(str(label).strip(), str(label).strip())
            col_in = t.columns[0].width / 914400
            pt = 7.0
            run = t.cell(r, 0).text_frame.paragraphs[0].runs
            if run and run[0].font.size:
                pt = run[0].font.size.pt
            w = text_width_in(label, pt)
            if w > col_in:
                log("FLAG", f"Name '{label}' is {w:.2f} in wide and the column is {col_in:.2f} in: it will wrap "
                            "and push the table down. Needs a short name (add it to name_overrides.json).")
        elif label_check:
            tmpl = t.cell(r, 0).text.strip()
            if tmpl and tmpl != str(label).strip():
                log("CONFLICT", f"{t.cell(0,0).text.strip()}: template row '{tmpl}' vs data row '{label}'.")
        set_cell(t.cell(r, 0), str(label).strip())
        for j, (kind, v) in enumerate(zip(spec, row[1:]), start=1):
            if kind == "n1":
                set_cell(t.cell(r, j), num(v, 1))
            elif kind == "n2":
                set_cell(t.cell(r, j), num(v, 2))
            elif kind == "pct":
                colored(t.cell(r, j), pct(v), v)
            elif kind == "bps":
                colored(t.cell(r, j), bps(v), v)


def build_slide3(slide, g, charts, names):
    S = lambda cols, r0, r1: block(g, cols, r0, r1)
    sec = ["n1", "pct", "n1", "pct"]; co = ["n1", "pct", "n1", "n1"]
    fill_block(find_table(slide, "TOP PERFORMERS — SECTORS"), 2, S("DEFGH", 77, 81), sec, label_check=False)
    fill_block(find_table(slide, "TOP PERFORMERS — SECTORS"), 10, S("KLMNO", 77, 81), sec, label_check=False)
    fill_block(find_table(slide, "TOP PERFORMERS — COMPANIES"), 2, S("DEFGH", 85, 89), co, names=names)
    fill_block(find_table(slide, "TOP PERFORMERS — COMPANIES"), 10, S("KLMNO", 85, 89), co, names=names)
    fill_block(find_table(slide, "VALUE LEADERS"), 2, S("DEFGH", 93, 97), co, names=names)
    fill_block(find_table(slide, "VALUE LEADERS"), 10, S("KLMNO", 93, 97), co, names=names)
    fill_block(find_table(slide, "INTERNATIONAL MARKETS"), 2, S("DEFGH", 109, 118), sec)
    fill_block(find_table(slide, "REGIONAL MARKETS"), 2, S("JKLMN", 109, 118), sec)
    fill_block(find_table(slide, "COMMODITY PRICES"), 2, S("DEFGH", 122, 133), ["n1", "pct", "pct", "pct"])
    fill_block(find_table(slide, "INTEREST RATES"), 2, S("JKLMN", 122, 125), ["n2", "bps", "bps", "bps"])
    log("OK", "Slide 3 tables filled from PPT-ENG.")

    ch = charts.get(3, [])
    old = [sh for sh in slide.shapes if sh.shape_type == 13 and 3_400_000 < sh.top < 5_700_000 and sh.left < 3_900_000]
    if len(ch) != 1:
        log("FLAG", f"Graphs deck slide 3: expected 1 chart, found {len(ch)}. Sector chart not updated.")
        return
    c = ch[0]
    cx = 3271761
    fb = [(395458, 3537627, cx, int(cx * c["height"] / c["width"]))]
    s0 = inherit_slots(old, fb, "Slide 3 sector chart")[0]
    slot = chart3_slot(s0, "left", ref_edge=s0[0], ref_top=s0[1], label="Slide 3 sector chart")
    for sh in old:
        remove_shape(sh)
    geom = fit_slot(slot, c, "Slide 3 sector chart", tol=CHART3_TOL)
    add_picture(slide, c["blob"], c["ext"], *geom, "Sector YTD chart")
    log("OK", "Slide 3 sector YTD chart replaced from Graphs deck.")


def replace_with_technicals(dst_slide, src_slide):
    tree = dst_slide.shapes._spTree
    for el in list(tree):
        if etree.QName(el).localname not in ("nvGrpSpPr", "grpSpPr"):
            tree.remove(el)
    for el in src_slide.shapes._spTree:
        if etree.QName(el).localname in ("nvGrpSpPr", "grpSpPr"):
            continue
        new = copy.deepcopy(el)
        for node in new.iter():
            for attr in (f"{{{R}}}embed", f"{{{R}}}link", f"{{{R}}}id"):
                rid = node.get(attr)
                if not rid:
                    continue
                rel = src_slide.part.rels[rid]
                if rel.is_external:
                    node.set(attr, dst_slide.part.relate_to(rel.target_ref, rel.reltype, is_external=True))
                elif rel.reltype.endswith("/image"):
                    part = rel.target_part
                    node.set(attr, add_image_part(dst_slide, part.blob, part.partname.ext))
                else:
                    node.set(attr, dst_slide.part.relate_to(rel.target_part, rel.reltype))
        tree.append(new)


# ---------------------------------------------------------------- cross-checks
def _norm_long_date(t):
    """'Monday, October 05, 2026' and 'Monday, October 5, 2026' are the same date."""
    return re.sub(r"\b0(\d),", r"\1,", " ".join((t or "").split()))


def cross_checks(g, face_dates, report_date, graphs_date, tech, writeup, events, econ):
    if report_date not in [d.date() for d in face_dates]:
        log("CONFLICT", f"Data file 'Face' sheet dates {[f'{d:%d-%b}' for d in face_dates]} do not include report date {report_date:%d-%b-%y}.")
    else:
        log("OK", f"Data file 'Face' sheet carries report date {report_date:%d-%b-%y}.")
    if graphs_date and _norm_long_date(graphs_date) != f"{report_date:%A, %B} {report_date.day}, {report_date.year}":
        log("STALE", f"Graphs deck is dated '{graphs_date}', not the report date. Charts may be stale.")
    # Technicals market summary vs data (skipped when Technicals are missing; the writeup checks still run)
    if tech is not None:
        _tech_checks(tech, g, report_date)
    _writeup_checks(g, writeup, events, econ)


def _tech_checks(tech, g, report_date):
    tt = tech.slides[0]
    date_sh = find_text(tt, r"^(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday), ")
    want = f"{report_date:%A, %B} {report_date.day}, {report_date.year}"
    if date_sh is None or _norm_long_date(date_sh.text_frame.text) != want:
        log("CONFLICT", f"Technicals deck date '{date_sh.text_frame.text.strip() if date_sh else '?'}' vs report date '{want}'.")
    for sh in tt.shapes:
        if sh.has_table and sh.table.cell(0, 0).text.strip() == "Close":
            kv = {sh.table.cell(i, 0).text.strip(): sh.table.cell(i, 1).text.strip() for i in range(len(sh.table.rows))}
            tclose = float(kv["Close"].replace(",", ""))
            if abs(tclose - g["C6"]) > 0.005:
                log("CONFLICT", f"TASI close: Technicals {kv['Close']} vs data file {g['C6']:,.2f}.")
            else:
                log("OK", f"TASI close agrees across data file and Technicals ({kv['Close']}).")
            tval = float(kv["Value (SAR mn)"].replace(",", ""))
            if abs(tval / 1000 - g["F6"]) > 0.05:
                log("CONFLICT", f"Value traded: Technicals SAR {tval:,.1f} mn vs data file SAR {g['F6']*1000:,.1f} mn.")


def _writeup_checks(g, writeup, events, econ):
    if not writeup:
        return
    import docx
    d = docx.Document(writeup)
    paras = [p.text.strip() for p in d.paragraphs if p.text.strip()]
    tasi = next((p for p in paras if p.startswith("TASI ") and "close at" in p), None)
    if tasi:
        m = re.search(r"TASI (rose|fell|declined|increased|decreased|gained) ([\d.]+)% \(-?([\d.,]+) points\) to close at ([\d,]+\.\d+)", tasi)
        if m:
            up = m.group(1) in ("rose", "increased", "gained")
            if up != (g["C8"] >= 0) or m.group(2) != f"{rnd(abs(g['C8'])*100,1)}":
                log("CONFLICT", f"Writeup TASI move '{m.group(1)} {m.group(2)}%' vs data DoD {pct(g['C8'])}.")
            if float(m.group(4).replace(",", "")) != round(g["C6"], 2):
                log("CONFLICT", f"Writeup TASI close {m.group(4)} vs data {g['C6']:,.2f}.")
        v = re.search(r"Trading value (increased|decreased|rose|fell|declined) by ([\d.]+)%", tasi)
        if v:
            up = v.group(1) in ("increased", "rose")
            if up != (g["F8"] >= 0) or v.group(2) != f"{rnd(abs(g['F8'])*100,1)}":
                log("CONFLICT", f"Writeup says trading value {v.group(1)} {v.group(2)}%; data file value traded DoD is {pct(g['F8'])}.")
            else:
                log("OK", "Writeup TASI bullet trading-value change matches data file.")
    if not any(p.lower().startswith("global write") and len(p) > 25 for p in paras):
        idx = next((i for i, p in enumerate(paras) if p.lower().startswith("global write")), None)
        if idx is not None and idx == len(paras) - 1:
            log("FLAG", "Writeup 'Global Write up' section is empty.")
    # calendars: writeup tables vs calendar file
    def norm(s):
        s = re.sub(r"\s+", " ", s.replace("\xa0", " ")).strip()
        s = CLIENT_ECON_WORDING_ALL.get(s, s)
        return "-" if s in ("_", "", "–") else s.lower()  # the file's "_" blank is the deck's "-"
    for t in d.tables:
        head = norm(t.cell(0, 0).text)
        rows = [[c.text for c in r.cells] for r in t.rows[2:]]
        rows = [r for r in rows if re.match(r"\d{1,2}-[A-Za-z]{3}-\d{2}", r[0].strip())]
        if head == "event calendar":
            mine = [(f"{a:%d-%b-%y}", b, c) for a, b, c in events]
            compare_rows("Event calendar", rows, mine, norm)
        elif head == "economic calendar":
            mine = [(f"{a:%d-%b-%y}", b, c, e) for a, b, c, e in econ]
            compare_rows("Economic calendar", rows, mine, norm)


def compare_rows(label, theirs, mine, norm):
    tk = {norm(r[1]): r for r in theirs}; mk = {norm(r[1]): r for r in mine}
    for k in tk.keys() - mk.keys():
        log("CONFLICT", f"{label}: writeup has '{tk[k][1].strip()}', calendar file does not.")
    for k in mk.keys() - tk.keys():
        log("CONFLICT", f"{label}: calendar file has '{mk[k][1].strip()}', writeup does not.")
    for k in tk.keys() & mk.keys():
        a = [norm(x) for x in tk[k]]; b = [norm(x) for x in mk[k]]
        if a != b:
            log("CONFLICT", f"{label} '{mk[k][1].strip()}': writeup {[x.strip() for x in tk[k][2:]]} vs calendar file {list(mk[k][2:])}.")
    if tk.keys() == mk.keys() and all([norm(x) for x in tk[k]] == [norm(x) for x in mk[k]] for k in tk):
        log("OK", f"{label} in writeup matches calendar file.")


# ---------------------------------------------------------------- main
