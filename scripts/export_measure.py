#!/usr/bin/env python3
"""Export DMW v2 decks to PDF with PowerPoint and record PowerPoint's own table layout, on Windows or
a Mac (26 Nov 2026).

  python export_measure.py --folder "<folder>" [--recurse]

For every "ANB Daily Market Watch - EN/AR.pptx" under the folder: saves the PDF beside it and writes
each table row on slides 1-3 (top, height, row height) to "<folder>/ppt_layout.csv" (the columns of
export_measure.ps1), so check_final.py can check the 10.60 in rule on PowerPoint's real layout.

- Windows: runs export_measure.ps1 (PowerPoint through COM), unchanged.
- Mac: PowerPoint for Mac through AppleScript. PowerPoint is sandboxed, so each deck is copied into
  Office's own group container (it may write there without a Grant File Access prompt; Mac PDF
  export, 5 Nov 2026), opened and measured there, saved as PDF and the PDF moved back. The deck
  itself is never opened. PowerPoint is quit afterwards only if this script started it.
Prints OK <deck> | slides N | pdf N bytes per deck, then DONE <n> decks. Exit 1 if a deck failed.
"""
import argparse, csv, glob, os, shutil, subprocess, sys, tempfile, time

HERE = os.path.dirname(os.path.abspath(__file__))
PATTERN = "ANB Daily Market Watch - *.pptx"
STAGE = os.environ.get("DMW_PDF_STAGE") or os.path.expanduser(
    "~/Library/Group Containers/UBF8T346G9.Office/claude-work/pdf-stage")
FIELDS = ["file", "slide", "shape", "top_in", "height_in", "bottom_in", "row", "row_h_in"]

# "repeat with sh in shapes" fails on PowerPoint for Mac (-2763), and "out" is a reserved word in
# its dictionary: index loops and other names only (26 Nov 2026).
MEASURE_SAVE = r'''
on run argv
  set src to item 1 of argv
  set dst to item 2 of argv
  set acc to ""
  tell application "Microsoft PowerPoint"
    open (POSIX file src)
    set p to missing value
    repeat with i from 1 to count of presentations
      if full name of presentation i is src then set p to presentation i
    end repeat
    if p is missing value then error "deck did not open: " & src
    set ns to count of slides of p
    set lim to 3
    if ns < lim then set lim to ns
    repeat with si from 1 to lim
      set s to slide si of p
      repeat with k from 1 to count of shapes of s
        set sh to shape k of s
        if has table of sh then
          set t to table object of sh
          repeat with r from 1 to count of rows of t
            set acc to acc & si & tab & (name of sh) & tab & (top of sh) & tab & (height of sh) & tab & r & tab & (height of row r of t) & linefeed
          end repeat
        end if
      end repeat
    end repeat
    save p in (POSIX file dst) as save as PDF
    close p saving no
  end tell
  return (ns as text) & linefeed & acc
end run
'''


def _num(s):
    return float(str(s).replace(",", "."))   # AppleScript prints reals with the locale's decimal mark


def measure_mac(src, dst, rel):
    os.makedirs(STAGE, exist_ok=True)
    fd, s_src = tempfile.mkstemp(prefix=os.path.basename(src)[:-5] + " ", suffix=".pptx", dir=STAGE)
    os.close(fd)
    s_dst = s_src[:-5] + ".pdf"
    shutil.copy2(src, s_src)
    try:
        r = subprocess.run(["osascript", "-e", MEASURE_SAVE, s_src, s_dst], capture_output=True,
                           text=True, timeout=300)
        if r.returncode != 0:
            raise RuntimeError(r.stderr.strip())
        for _ in range(60):   # some versions write the PDF a moment after the call returns
            if os.path.exists(s_dst) and os.path.getsize(s_dst) > 0:
                break
            time.sleep(0.5)
        shutil.move(s_dst, dst)
        lines = r.stdout.strip("\n").split("\n")
        rows = []
        for ln in lines[1:]:
            if not ln.strip():
                continue
            si, name, top, h, row, rh = ln.split("\t")
            top, h = _num(top), _num(h)
            rows.append({"file": rel, "slide": int(si), "shape": name, "top_in": round(top / 72, 4),
                         "height_in": round(h / 72, 4), "bottom_in": round((top + h) / 72, 4),
                         "row": int(row), "row_h_in": round(_num(rh) / 72, 4)})
        return int(lines[0]), rows
    finally:
        for f in (s_src, s_dst):
            if os.path.exists(f):
                os.remove(f)


def ppt_running():
    return subprocess.run(["pgrep", "-xq", "Microsoft PowerPoint"]).returncode == 0


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--folder", required=True)
    ap.add_argument("--recurse", action="store_true")
    a = ap.parse_args(argv)
    folder = os.path.abspath(a.folder)
    if os.name == "nt":
        cmd = ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File",
               os.path.join(HERE, "export_measure.ps1"), "-Folder", folder] + (["-Recurse"] if a.recurse else [])
        return subprocess.run(cmd).returncode
    if sys.platform != "darwin":
        print("export_measure.py: needs PowerPoint (Windows or Mac)", file=sys.stderr)
        return 2
    pat = os.path.join(folder, "**", PATTERN) if a.recurse else os.path.join(folder, PATTERN)
    files = sorted(glob.glob(pat, recursive=a.recurse))
    was_running, failed, layout = ppt_running(), False, []
    for f in files:
        rel = os.path.relpath(f, folder)
        pdf = f[:-5] + ".pdf"
        try:
            n, rows = measure_mac(f, pdf, rel)
        except Exception as e:  # noqa: BLE001
            print(f"FAIL {f} | {e}")
            failed = True
            continue
        layout += rows
        print(f"OK {f} | slides {n} | pdf {os.path.getsize(pdf):,} bytes")
    with open(os.path.join(folder, "ppt_layout.csv"), "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.DictWriter(fh, fieldnames=FIELDS)
        w.writeheader()
        w.writerows(layout)
    if not was_running and ppt_running():
        subprocess.run(["osascript", "-e",
                        'tell application "Microsoft PowerPoint" to if (count of presentations) is 0 then quit'])
    print(f"DONE {len(files)} decks")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
