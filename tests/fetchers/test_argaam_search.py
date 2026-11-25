import os, sys
from datetime import datetime, timedelta
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "..", "scripts"))
import argaam_fetch as af  # noqa: E402

T0 = datetime(2026, 11, 1)


class Fake(af.Argaam):
    """IDs 1000..200000, one minute apart; IDs in `dead` cannot be read (no browser needed)."""
    def __init__(self, dead=()):
        self.dead = set(dead)

    def _time_at(self, i, step=1, tries=12):
        if i in self.dead or i < 1000:
            return None, None
        return T0 + timedelta(minutes=i - 1000), i


def target(i):
    return T0 + timedelta(minutes=i - 1000)


def test_search_finds_window():
    lo = Fake().first_id_at_or_after(target(150000), 200000)
    assert 150000 - 16 <= lo < 150000


def test_unreadable_probe_is_not_before_window():
    # the first widening probe (hi - 400) is unreadable: it must not end the search there
    lo = Fake(dead={199600}).first_id_at_or_after(target(150000), 200000)
    assert 150000 - 16 <= lo < 150000
