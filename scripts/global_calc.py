"""Percent moves for the Global write-up, rounded half-up to one decimal from the underlying figures.
Usage: python global_calc.py moves.json
moves.json: {"us":   {"Dow": [points_change, close], ...},          # from the Reuters US close article
             "asia": {"Nikkei 225": [last, previous_close], ...}}   # previous close from the daily series
Prints each move to 4 dp and the published 1 dp figure, largest move first per block.
"""
import json, sys
from decimal import Decimal as D, ROUND_HALF_UP
r1 = lambda x: D(str(x)).quantize(D("0.1"), ROUND_HALF_UP)
m = json.load(open(sys.argv[1], encoding="utf-8"))
for block, f in (("us", lambda c, l: c / (l - c) * 100), ("asia", lambda l, p: (l / p - 1) * 100)):
    out = sorted(((k, f(*v)) for k, v in m.get(block, {}).items()), key=lambda kv: -abs(kv[1]))
    for k, p in out:
        print(f"{block:4s} {k}: {p:.4f}% -> {r1(p)}%")
