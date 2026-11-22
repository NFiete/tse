"""Undo/redo stack."""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Callable


@dataclass
class UndoEntry:
    label: str
    undo: Callable[[], object]
    redo: Callable[[], object]
    # where to put the cursor afterwards: (sheet title, Rect) or None
    where: object = None


@dataclass
class UndoStack:
    limit: int = 500
    entries: list = field(default_factory=list)
    pos: int = 0  # number of applied entries

    def push(self, entry: UndoEntry):
        del self.entries[self.pos:]
        self.entries.append(entry)
        if len(self.entries) > self.limit:
            del self.entries[0]
        self.pos = len(self.entries)

    def can_undo(self) -> bool:
        return self.pos > 0

    def can_redo(self) -> bool:
        return self.pos < len(self.entries)

    def undo(self) -> UndoEntry | None:
        if not self.can_undo():
            return None
        self.pos -= 1
        e = self.entries[self.pos]
        e.undo()
        return e

    def redo(self) -> UndoEntry | None:
        if not self.can_redo():
            return None
        e = self.entries[self.pos]
        e.redo()
        self.pos += 1
        return e

    def clear(self):
        self.entries.clear()
        self.pos = 0
