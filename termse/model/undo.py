"""Undo history kept as a tree (like Vim's undo tree).

Every state of the workbook is a node; each node but the root holds the edit (`UndoEntry`) that led to
it from its parent. Undoing moves to the parent; an edit made after undoing starts a new branch instead
of throwing the undone edits away, so any earlier state can be reached again with `goto`. Redo follows
the branch that was last visited.
"""
from __future__ import annotations

import time
from dataclasses import dataclass, field
from typing import Callable


@dataclass
class UndoEntry:
    label: str
    undo: Callable[[], object]
    redo: Callable[[], object]
    # where to put the cursor afterwards: (sheet title, Rect) or None
    where: object = None


@dataclass(eq=False)
class UndoNode:
    entry: UndoEntry | None  # the edit from parent to this state; None for the root
    parent: UndoNode | None = None
    seq: int = 0  # creation order; the root is 0
    time: float = field(default_factory=time.time)
    children: list = field(default_factory=list)  # oldest first
    redo_child: UndoNode | None = None  # the branch redo follows

    @property
    def label(self) -> str:
        return self.entry.label if self.entry is not None else "Original"


class UndoTree:
    def __init__(self, limit: int = 500):
        self.limit = limit  # maximum number of edits kept
        self.clear()

    def clear(self):
        self.root = UndoNode(None)
        self.current = self.root
        self.saved: UndoNode | None = self.root  # the state last loaded or saved
        self.count = 0
        self._seq = 0

    # ------------------------------------------------------------------ linear undo/redo

    def push(self, entry: UndoEntry) -> UndoNode:
        self._seq += 1
        node = UndoNode(entry, self.current, self._seq)
        self.current.children.append(node)
        self.current.redo_child = node
        self.current = node
        self.count += 1
        self._prune()
        return node

    def can_undo(self) -> bool:
        return self.current.parent is not None

    def can_redo(self) -> bool:
        return self.current.redo_child is not None

    def top(self) -> UndoEntry | None:
        """The edit that produced the current state."""
        return self.current.entry

    def undo(self) -> UndoEntry | None:
        n = self.current
        if n.parent is None:
            return None
        n.entry.undo()
        n.parent.redo_child = n
        self.current = n.parent
        return n.entry

    def redo(self) -> UndoEntry | None:
        n = self.current.redo_child
        if n is None:
            return None
        n.entry.redo()
        self.current = n
        return n.entry

    # ------------------------------------------------------------------ tree navigation

    def nodes(self) -> list[UndoNode]:
        """Every state, oldest first."""
        out, todo = [], [self.root]
        while todo:
            n = todo.pop()
            out.append(n)
            todo.extend(n.children)
        out.sort(key=lambda n: n.seq)
        return out

    def path(self, target: UndoNode) -> tuple[list[UndoNode], list[UndoNode]]:
        """The nodes to undo (from the current state up) and then redo (down to `target`)."""
        anc = set()
        n = target
        while n is not None:
            anc.add(n)
            n = n.parent
        ups = []
        n = self.current
        while n not in anc:
            ups.append(n)
            n = n.parent
        downs = []
        m = target
        while m is not n:
            downs.append(m)
            m = m.parent
        downs.reverse()
        return ups, downs

    def goto(self, target: UndoNode) -> list[UndoEntry]:
        """Move to any state in the tree; returns the entries undone or redone, in order."""
        ups, downs = self.path(target)
        done = []
        for n in ups:
            n.entry.undo()
            n.parent.redo_child = n
            self.current = n.parent
            done.append(n.entry)
        for n in downs:
            n.entry.redo()
            n.parent.redo_child = n
            self.current = n
            done.append(n.entry)
        return done

    def step_time(self, delta: int) -> UndoNode | None:
        """The state `delta` steps away in creation order (Vim's g- / g+), clamped to the ends;
        None if that is the current state."""
        order = self.nodes()
        i = order.index(self.current)
        j = max(0, min(len(order) - 1, i + delta))
        return order[j] if j != i else None

    # ------------------------------------------------------------------ editing the history

    def forget(self, node: UndoNode):
        """Drop a node and its descendants (not the current state or one of its ancestors)."""
        p = node.parent
        if p is None:
            return
        p.children.remove(node)
        if p.redo_child is node:
            p.redo_child = p.children[-1] if p.children else None
        node.parent = None
        sub = self._subtree(node)
        self.count -= len(sub)
        if self.saved in sub:
            self.saved = None

    def squash(self, mark: UndoNode, label: str):
        """Merge the edits made since state `mark` (an ancestor of the current state) into one."""
        chain = []
        n = self.current
        while n is not None and n is not mark:
            chain.append(n)
            n = n.parent
        if n is None or len(chain) < 2:
            return
        chain.reverse()
        entries = [c.entry for c in chain]
        saved = self.saved
        self.forget(chain[0])
        self.current = mark

        def undo():
            for e in reversed(entries):
                e.undo()

        def redo():
            for e in entries:
                e.redo()

        node = self.push(UndoEntry(label, undo, redo, entries[0].where))
        if saved is chain[-1]:
            self.saved = node

    def mark_saved(self):
        self.saved = self.current

    @staticmethod
    def _subtree(node: UndoNode) -> set:
        out, todo = set(), [node]
        while todo:
            x = todo.pop()
            out.add(x)
            todo.extend(x.children)
        return out

    def _prune(self):
        """Keep at most `limit` edits by dropping the oldest ones (with any branch that splits off
        before the oldest edit still kept)."""
        while self.count > self.limit:
            root = self.root
            on_path = self.current
            while on_path is not None and on_path.parent is not root:
                on_path = on_path.parent
            others = [c for c in root.children if c is not on_path]
            if others:
                self.forget(min(others, key=lambda c: c.seq))
            elif on_path is not None:
                # the state before the oldest edit becomes unreachable
                if self.saved is root:
                    self.saved = None
                on_path.parent = None
                on_path.entry = None
                root.children.clear()
                self.root = on_path
                self.count -= 1
            else:
                break


UndoStack = UndoTree  # old name


def graph(tree: UndoTree) -> list[tuple[str, UndoNode | None]]:
    """Lay the tree out like `git log --graph`, newest state first: one (graph prefix, node) per
    state, plus (prefix, None) rows where branches join."""
    lanes: list = []  # per column: the node that column's line is heading to, or None
    rows = []

    def trim():
        while lanes and lanes[-1] is None:
            lanes.pop()

    for n in reversed(tree.nodes()):
        idx = [i for i, w in enumerate(lanes) if w is n]
        if not idx:
            col = lanes.index(None) if None in lanes else len(lanes)
            if col == len(lanes):
                lanes.append(None)
        else:
            col = idx[0]
            if len(idx) > 1:
                last = idx[-1]
                s = ""
                for i, w in enumerate(lanes):
                    if i == col:
                        ch = "├"
                    elif i in idx:
                        ch = "╯" if i == last else "┴"
                    elif col < i < last:
                        ch = "┼" if w is not None else "─"
                    else:
                        ch = "│" if w is not None else " "
                    s += ch + ("─" if col <= i < last else " ")
                rows.append((s.rstrip(), None))
                for i in idx[1:]:
                    lanes[i] = None
        s = ""
        for i, w in enumerate(lanes):
            if i == col:
                ch = "●" if n is tree.current else "○"
            else:
                ch = "│" if w is not None else " "
            s += ch + " "
        rows.append((s, n))
        lanes[col] = n.parent
        trim()
    width = max((len(p) for p, _ in rows), default=0)
    return [(p.ljust(width), n) for p, n in rows]
