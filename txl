#!/usr/bin/env python3
"""Launcher: run term_xl straight from the source tree with the system Python."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.realpath(__file__)))

from termxl.__main__ import main  # noqa: E402

if __name__ == "__main__":
    sys.exit(main())
