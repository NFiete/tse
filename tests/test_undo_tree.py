"""Undo history as a tree: branches survive new edits and any state can be revisited."""
import pytest

from termse.model.undo import UndoEntry, UndoTree, graph
from termse.ui.app import TermSEApp
from termse.ui.undotree import UndoTreeScreen
from tests.test_ui import _ready, type_text, val


def _branchy(ws):
    """A1=1, A1=2, undo, A1=3: two branches below state 1."""
    ws["A1"] = 1
    ws["A1"] = 2
    ws.book.do_undo()
    ws["A1"] = 3
    return ws.book.undo


def test_new_edit_after_undo_keeps_the_old_branch(ws):
    tree = _branchy(ws)
    assert ws["A1"] == 3
    assert tree.count == 3
    one = tree.current.parent
    assert [c.label for c in one.children] == ["Edit cell", "Edit cell"]
    two, three = one.children
    ws.book.undo_goto(two)
    assert ws["A1"] == 2 and tree.current is two
    ws.book.undo_goto(tree.root)
    assert ws["A1"] is None
    # redo follows the branch visited last
    ws.book.do_redo()
    ws.book.do_redo()
    assert ws["A1"] == 2
    ws.book.undo_goto(three)
    assert ws["A1"] == 3


def test_step_time_walks_states_in_creation_order(ws):
    tree = _branchy(ws)  # states 0 (empty), 1 (1), 2 (2), 3 (3); currently 3
    seen = []
    while (n := tree.step_time(-1)) is not None:
        ws.book.undo_goto(n)
        seen.append(ws["A1"])
    assert seen == [2, 1, None]
    ws.book.undo_goto(tree.step_time(2))
    assert ws["A1"] == 2


def test_graph_layout(ws):
    tree = _branchy(ws)
    rows = [(p.rstrip(), n.seq if n else None) for p, n in graph(tree)]
    assert rows == [("●", 3), ("│ ○", 2), ("├─╯", None), ("○", 1), ("○", 0)]


def test_squash_merges_steps(ws):
    tree = ws.book.undo
    ws["A1"] = 1
    mark = tree.current
    ws["A2"] = 2
    ws["A3"] = 3
    tree.squash(mark, "Both")
    assert tree.count == 2 and tree.top().label == "Both"
    ws.book.do_undo()
    assert ws["A2"] is None and ws["A3"] is None and ws["A1"] == 1
    ws.book.do_redo()
    assert ws["A3"] == 3


def test_limit_drops_the_oldest_edits():
    log = []
    tree = UndoTree(limit=4)

    def push(name):
        tree.push(UndoEntry(name, lambda: log.append("-" + name), lambda: log.append("+" + name)))

    push("a")
    push("b")
    tree.undo()
    push("c")
    push("d")
    push("e")  # 5 edits: "a" goes; the state before it is no longer reachable
    assert tree.count == 4 and tree.root.entry is None
    assert [n.seq for n in tree.nodes()] == [1, 2, 3, 4, 5]
    push("f")  # then "b", a branch that split off before the oldest edit now kept
    assert tree.count == 4
    assert [n.label for n in tree.nodes()][1:] == ["c", "d", "e", "f"]
    while tree.undo():
        pass
    assert log[-4:] == ["-f", "-e", "-d", "-c"]


def test_saved_state_is_tracked(ws, tmp_path):
    from termse.io import xlsx

    ws["A1"] = 1
    xlsx.save(ws.book, str(tmp_path / "s.xlsx"))
    assert ws.book.undo.saved is ws.book.undo.current
    ws["A1"] = 2
    assert ws.book.undo.saved is ws.book.undo.current.parent


@pytest.mark.asyncio
async def test_undo_tree_panel_previews_and_cancels():
    app = TermSEApp()
    async with app.run_test(size=(120, 30)) as pilot:
        await _ready(app, pilot)
        for t in ("1", "2"):
            await pilot.press("ctrl+home")
            await type_text(pilot, t)
            await pilot.press("enter")
        await pilot.press("ctrl+z")
        await pilot.press("ctrl+home")
        await type_text(pilot, "3")
        await pilot.press("enter")
        assert val(app, "A1") == 3
        await pilot.press("alt+z")
        await pilot.pause()
        assert isinstance(app.screen, UndoTreeScreen)
        # newest first: 3, then the "2" branch (moving there previews it)
        await pilot.press("down")
        await pilot.pause()
        assert val(app, "A1") == 2
        await pilot.press("down")  # skips the join row
        await pilot.pause()
        assert val(app, "A1") == 1
        await pilot.press("escape")
        await pilot.pause()
        assert not isinstance(app.screen, UndoTreeScreen)
        assert val(app, "A1") == 3
        # Enter keeps the chosen state, and ordinary undo/redo continue from there
        await pilot.press("alt+z")
        await pilot.pause()
        await pilot.press("down", "enter")
        await pilot.pause()
        assert val(app, "A1") == 2 and "Undo state 2" in app.message
        await pilot.press("ctrl+z")
        assert val(app, "A1") == 1
        await pilot.press("ctrl+y")
        assert val(app, "A1") == 2


@pytest.mark.asyncio
async def test_vi_g_minus_g_plus():
    app = TermSEApp(vi=True)
    async with app.run_test(size=(100, 30)) as pilot:
        await _ready(app, pilot)
        for t in ("1", "2"):
            await pilot.press("s")
            await type_text(pilot, t)
            await pilot.press("escape")
        await pilot.press("u")
        await pilot.press("s")
        await type_text(pilot, "3")
        await pilot.press("escape")
        assert val(app, "A1") == 3
        await pilot.press("g", "minus")
        assert val(app, "A1") == 2  # u would have gone to 1
        await pilot.press("g", "minus")
        assert val(app, "A1") == 1
        await pilot.press("2", "g", "plus")
        assert val(app, "A1") == 3
        await pilot.press("g", "plus")
        assert "newest" in app.message
