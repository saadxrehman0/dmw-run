#!/usr/bin/env bash
# Replay one past edition: build EN+AR from the day folder, refill the news read back from the
# published v1 decks, and compare figures. Usage: replay.sh <day folder> <YYYY-MM-DD> <out dir> <state dir>
set -e
D="$1"; DATE="$2"; O="$3"; S="$4"; YMD="${DATE//-/}"
HERE="$(cd "$(dirname "$0")" && pwd)"; SC="$HERE/../../scripts"
mkdir -p "$O"
TE=$(ls "$D"/Technicals\ -ENG*"$YMD".pptx | head -1); TA=$(ls "$D"/Technicals\ -AR*"$YMD".pptx | head -1)
python3 "$SC/build_dmw.py" --lang en --data "$D/DMW - Data.xlsx" --graphs "$D/DMW Graphs - EN.pptx" \
  --calendar "$D/Manual File-Calendar.xlsx" --technicals "$TE" --date "$DATE" --mapping "$S/Mapping.xlsx" \
  --names "$S/name_overrides.json" --out "$O/build-EN.pptx" > "$O/build-EN.log"
python3 "$SC/build_dmw.py" --lang ar --data "$D/DMW - Data.xlsx" --graphs "$D/DMW Graphs - AR.pptx" \
  --calendar "$D/Manual File-Calendar.xlsx" --technicals "$TA" --date "$DATE" --mapping "$S/Mapping.xlsx" \
  --names "$S/name_overrides.json" --ar-names "$S/name_overrides_ar.json" --glossary "$S/ar_glossary.json" \
  --out "$O/build-AR.pptx" > "$O/build-AR.log"
python3 "$HERE/news_from_v1_deck.py" --en "$D/ANB Daily Market Watch - EN.pptx" --ar "$D/ANB Daily Market Watch - AR.pptx" --out "$O/news.json"
RES=(); [ -f "$D/DMW - Results.xlsx" ] && RES=(--results "$D/DMW - Results.xlsx" --quarter "${QUARTER:-3Q26}" --mapping "$S/Mapping.xlsx")
python3 "$SC/fill_commentary.py" --news "$O/news.json" --en "$O/build-EN.pptx" --ar "$O/build-AR.pptx" --outdir "$O" "${RES[@]}" > "$O/fill.log" || true
