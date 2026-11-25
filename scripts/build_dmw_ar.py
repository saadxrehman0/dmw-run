#!/usr/bin/env python3
"""AR entry point for the Template v2 build: same as `build_dmw.py --lang ar`."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import build_dmw  # noqa: E402

if __name__ == "__main__":
    build_dmw.main(default_lang="ar")
