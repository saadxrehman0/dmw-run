"""Offline checks for tadawul_fetch.py (no browser): time handling, kinds, window filter."""
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "..", "scripts"))
import tadawul_fetch as tf  # noqa: E402


def test_time_is_riyadh_plus_two_for_pkt():
    t = tf.parse_t("24/11/2026 08:07:54")
    assert (t + tf.PKT_MINUS_AST).strftime("%Y-%m-%d %H:%M") == "2026-11-24 10:07"


def test_kinds():
    assert tf.kind_of("Naseej for Technology Co. Announces Contract Sign Off") == "contract"
    assert tf.kind_of("Correction announcement regarding its Interim Financial Results") == "results"
    assert tf.kind_of("Board invites its shareholders to attend the Ordinary General Assembly Meeting") == "assembly"
    assert tf.kind_of("Albilad Saudi Sovereign Sukuk ETF changes") == "fund"
    assert tf.kind_of("Nahdi Medical Company Announces the Expected Financial Impact") == "other"


class FakeTadawul(tf.Tadawul):
    def __init__(self, pages):  # no browser
        self.pages = pages

    def page(self, n, size=50):
        return self.pages[n - 1] if n <= len(self.pages) else []


def test_walk_stops_at_window_start_and_dedupes():
    rows = [{"id": i, "pub_ast": f"2026-11-24 08:{i - 90:02d}:00"} for i in range(99, 90, -1)]
    g = FakeTadawul([rows[:5], rows[4:]])          # row 95 on both pages
    got = g.walk(lambda x: x["pub_ast"] < "2026-11-24 08:03:00")
    assert [x["id"] for x in got] == [99, 98, 97, 96, 95, 94, 93]


def test_summary_sorted_and_max_id():
    out = tf.summary([{"id": 5, "pub_ast": "b"}, {"id": 7, "pub_ast": "a"}], {"sinceId": 1})
    assert out["maxId"] == 7 and [x["id"] for x in out["items"]] == [7, 5] and out["count"] == 2
    assert tf.summary([], {"sinceId": 9})["maxId"] == 9
