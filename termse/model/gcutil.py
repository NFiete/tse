"""Garbage-collector helper for bulk operations.

Building or rewriting hundreds of thousands of small objects triggers many full
cyclic-GC passes over the (large) heap. Pausing the collector during bulk work and
freezing the survivors afterwards keeps big workbooks responsive. Our own objects
(ASTs, tokens, formulas) are freed by reference counting, so freezing doesn't leak
them; xlsx.load() unfreezes before replacing a workbook.
"""
from __future__ import annotations

import gc
from contextlib import contextmanager


@contextmanager
def paused_gc():
    enabled = gc.isenabled()
    gc.disable()
    try:
        yield
    finally:
        if enabled:
            gc.enable()
            gc.freeze()
