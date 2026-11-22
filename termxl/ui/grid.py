"""The spreadsheet grid widget: virtualized rendering, navigation and selection."""
from __future__ import annotations

from dataclasses import dataclass, field

from rich.cells import cell_len, set_cell_size
from rich.segment import Segment
from rich.style import Style
from textual import events
from textual.binding import Binding
from textual.strip import Strip
from textual.widget import Widget

from ..format.numfmt import fit_general, format_value
from ..format.style import CellLook, StyleResolver, contrast
from ..formula.errors import ExcelError
from ..model import outline
from ..model.refs import MAX_COL, MAX_ROW, Rect, col_letter, parse_addr
from .theme import DARK, Palette


@dataclass
class Layout:
    cols: list  # [(col, x, width)] width includes the 1-char gutter
    rows: list  # [row]
    header_w: int
    frozen_rows: int
    frozen_cols: int
    n_frozen_cols_visible: int
    n_frozen_rows_visible: int
    # outline symbols: a gutter left of the row numbers (2 chars per row level) and a band
    # above the column letters (1 line per column level)
    ro_w: int = 0
    co_h: int = 0
    nrl: int = 0
    ncl: int = 0
    marks: dict = field(default_factory=dict)  # (axis, index, level) -> "bar" | "+" | "-" | None

    @property
    def cells_x(self) -> int:
        """x where the row numbers end and the cells begin."""
        return self.ro_w + self.header_w


def _blend(a: str, b: str, t: float) -> str:
    ar, ag, ab = int(a[0:2], 16), int(a[2:4], 16), int(a[4:6], 16)
    br, bg, bb = int(b[0:2], 16), int(b[2:4], 16), int(b[4:6], 16)
    return f"{round(ar + (br - ar) * t):02X}{round(ag + (bg - ag) * t):02X}{round(ab + (bb - ab) * t):02X}"


def _fit(text: str, width: int, align: str) -> str:
    if width <= 0:
        return ""
    n = cell_len(text)
    if n >= width:
        return set_cell_size(text, width)
    if align == "right":
        return " " * (width - n) + text
    if align == "center":
        left = (width - n) // 2
        return " " * left + text + " " * (width - n - left)
    return text + " " * (width - n)


class Grid(Widget, can_focus=True):
    DEFAULT_CSS = """
    Grid { height: 1fr; width: 1fr; }
    """

    BINDINGS = [
        Binding("up", "move(-1,0)", show=False),
        Binding("down", "move(1,0)", show=False),
        Binding("left", "move(0,-1)", show=False),
        Binding("right", "move(0,1)", show=False),
        Binding("shift+up", "move(-1,0,True)", show=False),
        Binding("shift+down", "move(1,0,True)", show=False),
        Binding("shift+left", "move(0,-1,True)", show=False),
        Binding("shift+right", "move(0,1,True)", show=False),
        Binding("ctrl+up", "jump(-1,0)", show=False),
        Binding("ctrl+down", "jump(1,0)", show=False),
        Binding("ctrl+left", "jump(0,-1)", show=False),
        Binding("ctrl+right", "jump(0,1)", show=False),
        Binding("ctrl+shift+up", "jump(-1,0,True)", show=False),
        Binding("ctrl+shift+down", "jump(1,0,True)", show=False),
        Binding("ctrl+shift+left", "jump(0,-1,True)", show=False),
        Binding("ctrl+shift+right", "jump(0,1,True)", show=False),
        Binding("pageup", "page(-1)", show=False),
        Binding("pagedown", "page(1)", show=False),
        Binding("alt+pageup", "page_h(-1)", show=False),
        Binding("alt+pagedown", "page_h(1)", show=False),
        Binding("home", "home", show=False),
        Binding("ctrl+home", "ctrl_home", show=False),
        Binding("ctrl+end", "ctrl_end", show=False),
        Binding("enter", "enter_move(1)", show=False),
        Binding("shift+enter", "enter_move(-1)", show=False),
        Binding("tab", "tab_move(1)", show=False),
        Binding("shift+tab", "tab_move(-1)", show=False),
        Binding("ctrl+a", "select_all", show=False),
        Binding("ctrl+space,ctrl+@", "select_col", show=False),
        Binding("shift+space", "select_row", show=False),
    ]

    def __init__(self, **kw):
        super().__init__(**kw)
        self.book = None
        self.sheet = None
        self.active = (1, 1)
        self.anchor = (1, 1)
        self.corner = (1, 1)
        self.top = 1
        self.left = 1
        self.view_states: dict = {}
        self.palette: Palette = DARK
        self.resolver: StyleResolver | None = None
        self.edit_cell = None  # (sheet, r, c) being edited
        self.edit_text = ""
        self.ref_highlights: list = []  # [(sheet, Rect, color)]
        self._spill_hl = None  # spilled range around the cursor, set per rendered line
        self.clip_rect = None  # (sheet, Rect)
        self._styles: dict = {}
        self._layout_key = None
        self._layout: Layout | None = None
        self._dragging = False
        self._drag_mode = None

    # ------------------------------------------------------------------ setup

    def set_book(self, book, sheet=None):
        self.book = book
        self.resolver = StyleResolver(book.wb)
        self.view_states.clear()
        self._styles.clear()
        self.show_sheet(sheet or book.sheets[0], initial=True)

    def show_sheet(self, sheet, initial: bool = False):
        if self.sheet is not None and not initial:
            self.view_states[id(self.sheet)] = (self.active, self.anchor, self.corner, self.top, self.left)
        self.sheet = sheet
        st = self.view_states.get(id(sheet))
        if st:
            self.active, self.anchor, self.corner, self.top, self.left = st
        else:
            pos = (1, 1)
            try:
                sel = sheet.ws.sheet_view.selection
                if sel and sel[0].activeCell:
                    p = parse_addr(sel[0].activeCell)
                    if p:
                        pos = p
            except Exception:
                pass
            fr, fc = self.frozen()
            self.active = self.anchor = self.corner = pos
            self.top, self.left = fr + 1, fc + 1
            try:
                tl = sheet.ws.sheet_view.topLeftCell
                if tl and not fr and not fc:
                    p = parse_addr(tl)
                    if p:
                        self.top, self.left = p
            except Exception:
                pass
        self._layout_key = None
        if self.size.height:
            self.ensure_visible(*self.active)
        self.refresh()

    def set_palette(self, palette: Palette):
        self.palette = palette
        self._styles.clear()
        self.refresh()

    def invalidate(self):
        self._layout_key = None
        self.refresh()

    # ------------------------------------------------------------------ geometry

    def frozen(self) -> tuple[int, int]:
        fp = self.sheet.ws.freeze_panes if self.sheet is not None else None
        if not fp:
            return 0, 0
        p = parse_addr(fp)
        if not p:
            return 0, 0
        return p[0] - 1, p[1] - 1

    def col_px(self, c: int) -> int:
        w, hidden = self.sheet.col_width(c)
        if hidden or w <= 0:
            return 0
        return max(2, int(w + 0.5)) + 1

    def row_visible(self, r: int) -> bool:
        return not self.sheet.row_hidden(r)

    def bands(self) -> tuple[int, int, int, int]:
        """-> (row outline levels, column outline levels, row gutter width, column band height)"""
        s = self.sheet
        if s is None or not outline.symbols_shown(s):
            return 0, 0, 0, 0
        nrl, ncl = outline.max_level(s, "row"), outline.max_level(s, "col")
        return nrl, ncl, (2 * nrl + 1 if nrl else 0), ncl

    @staticmethod
    def _header_w(last_row: int, ro_w: int, ncl: int) -> int:
        w = max(4, len(str(last_row)) + 2)
        if ncl:  # room for the column level buttons "1 2 …" in the corner
            w = max(w, 2 * ncl + 2 - ro_w)
        return w

    def layout(self) -> Layout:
        width, height = self.size.width, self.size.height
        if self.sheet._colw is None:
            self.sheet.col_width(1)
        nrl, ncl, ro_w, co_h = self.bands()
        key = (width, height, self.top, self.left, id(self.sheet), id(self.sheet._colw),
               self.sheet.ws.freeze_panes, len(self.sheet.ws.row_dimensions), self.sheet.outline_gen, nrl, ncl)
        if key == self._layout_key and self._layout is not None:
            return self._layout
        fr, fc = self.frozen()
        rows: list[int] = []
        avail = max(0, height - 1 - co_h)
        r = 1
        while r <= fr and len(rows) < avail:
            if self.row_visible(r):
                rows.append(r)
            r += 1
        nfr = len(rows)
        r = max(self.top, fr + 1)
        while len(rows) < avail and r <= MAX_ROW:
            if self.row_visible(r):
                rows.append(r)
            r += 1
        last_row = rows[-1] if rows else 1
        header_w = self._header_w(last_row, ro_w, ncl)
        cols: list[tuple] = []
        x = ro_w + header_w
        c = 1
        while c <= fc and x < width:
            w = self.col_px(c)
            if w:
                cols.append((c, x, w))
                x += w
            c += 1
        nfc = len(cols)
        c = max(self.left, fc + 1)
        while x < width and c <= MAX_COL:
            w = self.col_px(c)
            if w:
                cols.append((c, x, w))
                x += w
            c += 1
        self._layout = Layout(cols, rows, header_w, fr, fc, nfc, nfr, ro_w, co_h, nrl, ncl)
        self._layout_key = key
        return self._layout

    def visible_rows_count(self) -> int:
        L = self.layout()
        return max(1, len(L.rows) - L.n_frozen_rows_visible)

    def ensure_visible(self, r: int, c: int):
        fr, fc = self.frozen()
        _, ncl, ro_w, co_h = self.bands()
        height = max(1, self.size.height - 1 - co_h)
        width = self.size.width
        changed = False
        if r > fr:
            avail = max(1, height - sum(1 for rr in range(1, fr + 1) if self.row_visible(rr)))
            if r < self.top:
                self.top = r
                changed = True
            else:
                count, rr = 0, self.top
                while rr <= r and count <= avail:
                    if self.row_visible(rr):
                        count += 1
                    rr += 1
                if count > avail:
                    k, rr = 0, r
                    while rr > fr + 1:
                        if self.row_visible(rr):
                            k += 1
                            if k == avail:
                                break
                        rr -= 1
                    self.top = max(fr + 1, rr)
                    changed = True
        if c > fc:
            if c < self.left:
                self.left = c
                changed = True
            else:
                header_w = self._header_w(self.top + height, ro_w, ncl)
                fixed = ro_w + header_w + sum(self.col_px(cc) for cc in range(1, fc + 1))
                avail = max(1, width - fixed)
                total, cc = 0, self.left
                while cc <= c and total <= avail:
                    total += self.col_px(cc)
                    cc += 1
                if total > avail:
                    left, acc = c, self.col_px(c)
                    while left - 1 > fc and acc + self.col_px(left - 1) <= avail:
                        left -= 1
                        acc += self.col_px(left)
                    self.left = left
                    changed = True
        if changed:
            self._layout_key = None

    def hit(self, x: int, y: int):
        """-> ("cell", r, c) | ("row", r) | ("col", c) | ("corner",) | ("outline", Group) |
        ("level", axis, n) | None"""
        L = self.layout()

        def col_at():
            for c, cx, w in L.cols:
                if cx <= x < cx + w:
                    return c
            return None

        def level_button(n_levels: int):
            n = x // 2 + 1
            return n if x % 2 == 0 and n <= n_levels + 1 else None

        if y < L.co_h:
            if x < L.cells_x:
                n = level_button(L.ncl) if y == 0 else None
                return ("level", "col", n) if n else None
            c = col_at()
            g = self.outline_group("col", c, y + 1) if c is not None else None
            return ("outline", g) if g else None
        if y == L.co_h:
            if x < L.ro_w:
                n = level_button(L.nrl)
                return ("level", "row", n) if n else None
            if x < L.cells_x:
                return ("corner",)
            c = col_at()
            return ("col", c) if c is not None else None
        idx = y - L.co_h - 1
        if idx >= len(L.rows):
            return None
        r = L.rows[idx]
        if x < L.ro_w:
            lvl = x // 2 + 1
            g = self.outline_group("row", r, lvl) if lvl <= L.nrl else None
            return ("outline", g) if g else None
        if x < L.cells_x:
            return ("row", r)
        c = col_at()
        return ("cell", r, c) if c is not None else None

    # ------------------------------------------------------------------ outline symbols

    def outline_mark(self, L: Layout, axis: str, i: int, lvl: int):
        """What to draw for row/column i at outline level lvl: "bar" for a group member,
        "-" / "+" for the button of an expanded / collapsed group, or None."""
        key = (axis, i, lvl)
        m = L.marks.get(key, 0)
        if m != 0:
            return m
        s = self.sheet
        if outline.level(s, axis, i) >= lvl:
            m = "bar"
        else:
            nb = i - 1 if outline.summary_after(s, axis) else i + 1
            if nb >= 1 and outline.level(s, axis, nb) >= lvl:
                m = "+" if outline.is_collapsed(s, outline.group_at(s, axis, nb, lvl)) else "-"
            else:
                m = None
        L.marks[key] = m
        return m

    def outline_group(self, axis: str, i: int, lvl: int):
        """The group drawn at row/column i, level lvl (as a member or by its button), if any."""
        s = self.sheet
        if outline.level(s, axis, i) >= lvl:
            return outline.group_at(s, axis, i, lvl)
        nb = i - 1 if outline.summary_after(s, axis) else i + 1
        if nb >= 1 and outline.level(s, axis, nb) >= lvl:
            return outline.group_at(s, axis, nb, lvl)
        return None

    # ------------------------------------------------------------------ selection model

    @property
    def selection(self) -> Rect:
        return Rect.of(self.anchor[0], self.anchor[1], self.corner[0], self.corner[1])

    def set_cursor(self, r: int, c: int, extend: bool = False, keep_active: bool = False):
        r = min(max(r, 1), MAX_ROW)
        c = min(max(c, 1), MAX_COL)
        if extend:
            self.corner = (r, c)
            self._expand_merges()
            self.ensure_visible(r, c)
        else:
            mr = self.sheet.merged_at(r, c)
            if mr is not None:
                r, c = mr.min_row, mr.min_col
            if not keep_active:
                self.active = (r, c)
            self.anchor = self.corner = (r, c)
            if mr is not None:
                self.corner = (mr.max_row, mr.max_col)
            self.ensure_visible(r, c)
        self.refresh()
        self.app.grid_changed()

    def _expand_merges(self):
        """Grow the selection so it fully contains any merged ranges it touches."""
        rect = self.selection
        if rect.rows * rect.cols > 100_000:
            return
        changed = True
        while changed:
            changed = False
            for mr in self.sheet.ws.merged_cells.ranges:
                if mr.max_row < rect.r1 or mr.min_row > rect.r2 or mr.max_col < rect.c1 or mr.min_col > rect.c2:
                    continue
                new = Rect(min(rect.r1, mr.min_row), min(rect.c1, mr.min_col),
                           max(rect.r2, mr.max_row), max(rect.c2, mr.max_col))
                if new != rect:
                    rect = new
                    changed = True
        a, k = self.anchor, self.corner
        self.anchor = (rect.r1 if a[0] <= k[0] else rect.r2, rect.c1 if a[1] <= k[1] else rect.c2)
        self.corner = (rect.r2 if a[0] <= k[0] else rect.r1, rect.c2 if a[1] <= k[1] else rect.c1)

    def _step(self, r: int, c: int, dr: int, dc: int) -> tuple[int, int]:
        """One step in a direction, skipping hidden rows/cols and merged interiors."""
        mr = self.sheet.merged_at(r, c)
        if mr is not None:
            if dr > 0:
                r = mr.max_row
            if dc > 0:
                c = mr.max_col
            if dr < 0:
                r = mr.min_row
            if dc < 0:
                c = mr.min_col
        nr, nc = r + dr, c + dc
        while 1 <= nr <= MAX_ROW and dr and not self.row_visible(nr):
            nr += dr
        while 1 <= nc <= MAX_COL and dc and self.col_px(nc) == 0:
            nc += dc
        if not (1 <= nr <= MAX_ROW):
            nr = r
        if not (1 <= nc <= MAX_COL):
            nc = c
        return nr, nc

    def action_move(self, dr: int, dc: int, extend: bool = False):
        if extend:
            r, c = self._step(*self.corner, dr, dc)
            self.set_cursor(r, c, extend=True)
        else:
            r, c = self._step(*self.active, dr, dc)
            self.set_cursor(r, c)

    def _is_blank(self, r: int, c: int) -> bool:
        v = self.book.get_value(self.sheet, r, c)
        return v is None or v == ""

    def jump_target(self, r: int, c: int, dr: int, dc: int) -> tuple[int, int]:
        s = self.sheet
        limit_r = MAX_ROW if dr > 0 else 1
        limit_c = MAX_COL if dc > 0 else 1
        nr, nc = r + dr, c + dc
        if not (1 <= nr <= MAX_ROW and 1 <= nc <= MAX_COL):
            return r, c
        cur_blank = self._is_blank(r, c)
        next_blank = self._is_blank(nr, nc)
        if not cur_blank and not next_blank:
            # move to the end of the contiguous block
            while True:
                tr, tc = nr + dr, nc + dc
                if not (1 <= tr <= MAX_ROW and 1 <= tc <= MAX_COL) or self._is_blank(tr, tc):
                    return nr, nc
                nr, nc = tr, tc
        # find the next non-blank cell
        max_r, max_c = s.max_row, s.max_col
        while 1 <= nr <= MAX_ROW and 1 <= nc <= MAX_COL:
            if not self._is_blank(nr, nc):
                return nr, nc
            if (dr > 0 and nr > max_r) or (dc > 0 and nc > max_c):
                break
            nr += dr
            nc += dc
        return (limit_r if dr else r), (limit_c if dc else c)

    def action_jump(self, dr: int, dc: int, extend: bool = False):
        base = self.corner if extend else self.active
        r, c = self.jump_target(*base, dr, dc)
        self.set_cursor(r, c, extend=extend)

    def action_page(self, d: int):
        n = self.visible_rows_count()
        r, c = self.active
        self.top = max(self.frozen()[0] + 1, min(MAX_ROW, self.top + d * n))
        self._layout_key = None
        self.set_cursor(max(1, min(MAX_ROW, r + d * n)), c)

    def action_page_h(self, d: int):
        L = self.layout()
        n = max(1, len(L.cols) - L.n_frozen_cols_visible - 1)
        r, c = self.active
        self.left = max(self.frozen()[1] + 1, min(MAX_COL, self.left + d * n))
        self._layout_key = None
        self.set_cursor(r, max(1, min(MAX_COL, c + d * n)))

    def action_home(self):
        self.set_cursor(self.active[0], 1)

    def action_ctrl_home(self):
        fr, fc = self.frozen()
        self.set_cursor(fr + 1, fc + 1)

    def action_ctrl_end(self):
        self.set_cursor(max(1, self.sheet.max_row), max(1, self.sheet.max_col))

    def action_enter_move(self, d: int):
        rect = self.selection
        if rect.rows * rect.cols > 1 and rect.rows * rect.cols < 10_000_000:
            r, c = self.active
            r += d
            if r > rect.r2:
                r, c = rect.r1, c + 1 if c < rect.c2 else rect.c1
            elif r < rect.r1:
                r, c = rect.r2, c - 1 if c > rect.c1 else rect.c2
            self.active = (r, c)
            self.ensure_visible(r, c)
            self.refresh()
            self.app.grid_changed()
        else:
            self.action_move(d, 0)

    def action_tab_move(self, d: int):
        rect = self.selection
        if rect.rows * rect.cols > 1 and rect.rows * rect.cols < 10_000_000:
            r, c = self.active
            c += d
            if c > rect.c2:
                c, r = rect.c1, r + 1 if r < rect.r2 else rect.r1
            elif c < rect.c1:
                c, r = rect.c2, r - 1 if r > rect.r1 else rect.r2
            self.active = (r, c)
            self.ensure_visible(r, c)
            self.refresh()
            self.app.grid_changed()
        else:
            self.action_move(0, d)

    def action_select_all(self):
        self.anchor = (1, 1)
        self.corner = (MAX_ROW, MAX_COL)
        self.refresh()
        self.app.grid_changed()

    def action_select_col(self):
        r, c = self.active
        sel = self.selection
        self.anchor = (1, sel.c1)
        self.corner = (MAX_ROW, sel.c2)
        self.refresh()
        self.app.grid_changed()

    def action_select_row(self):
        sel = self.selection
        self.anchor = (sel.r1, 1)
        self.corner = (sel.r2, MAX_COL)
        self.refresh()
        self.app.grid_changed()

    # ------------------------------------------------------------------ input

    async def on_key(self, event: events.Key) -> None:
        if event.is_printable and event.character and not any(
                event.key.startswith(m) for m in ("ctrl+", "alt+", "super+", "meta+")):
            event.stop()
            event.prevent_default()
            self.app.start_edit(initial=event.character, mode="enter")
        elif event.key == "f2":
            event.stop()
            self.app.start_edit(mode="edit")
        elif event.key == "backspace":
            event.stop()
            self.app.start_edit(initial="", mode="enter")
        elif event.key == "delete":
            event.stop()
            self.app.action_clear()

    def on_mouse_down(self, event: events.MouseDown) -> None:
        h = self.hit(event.x, event.y)
        if h is None:
            return
        self.focus()
        app = self.app
        if h[0] == "cell" and app.editing is not None and app.can_point():
            app.point_at(h[1], h[2], extend=event.shift)
            self._dragging, self._drag_mode = True, "point"
            self.capture_mouse()
            return
        if app.editing is not None:
            if not app.commit_edit(move=None):
                return
        if h[0] == "outline":
            app.toggle_group(h[1])
            return
        if h[0] == "level":
            app.show_outline_level(h[1], h[2])
            return
        if h[0] == "cell":
            self.set_cursor(h[1], h[2], extend=event.shift)
            self._dragging, self._drag_mode = True, "cell"
        elif h[0] == "col":
            if event.shift:
                self.corner = (MAX_ROW, h[1])
            else:
                self.active = (max(1, self.top), h[1])
                self.anchor, self.corner = (1, h[1]), (MAX_ROW, h[1])
            self._dragging, self._drag_mode = True, "col"
            self.refresh()
            app.grid_changed()
        elif h[0] == "row":
            if event.shift:
                self.corner = (h[1], MAX_COL)
            else:
                self.active = (h[1], max(1, self.left))
                self.anchor, self.corner = (h[1], 1), (h[1], MAX_COL)
            self._dragging, self._drag_mode = True, "row"
            self.refresh()
            app.grid_changed()
        elif h[0] == "corner":
            self.action_select_all()
        self.capture_mouse()

    def on_mouse_move(self, event: events.MouseMove) -> None:
        if not self._dragging:
            return
        h = self.hit(max(0, event.x), max(0, event.y))
        if h is None:
            return
        if self._drag_mode == "point" and h[0] == "cell":
            self.app.point_at(h[1], h[2], extend=True)
        elif self._drag_mode == "cell" and h[0] == "cell":
            if (h[1], h[2]) != self.corner:
                self.set_cursor(h[1], h[2], extend=True)
        elif self._drag_mode == "col" and h[0] in ("col", "cell"):
            c = h[1] if h[0] == "col" else h[2]
            self.corner = (MAX_ROW, c)
            self.refresh()
            self.app.grid_changed()
        elif self._drag_mode == "row" and h[0] in ("row", "cell"):
            self.corner = (h[1], MAX_COL)
            self.refresh()
            self.app.grid_changed()

    def on_mouse_up(self, event: events.MouseUp) -> None:
        if self._dragging:
            self._dragging = False
            self.release_mouse()

    def on_click(self, event: events.Click) -> None:
        if event.chain == 2:
            h = self.hit(event.x, event.y)
            if h and h[0] == "cell" and self.app.editing is None:
                self.app.start_edit(mode="edit")

    def on_mouse_scroll_down(self, event: events.MouseScrollDown) -> None:
        self._scroll(3, event.shift)
        event.stop()

    def on_mouse_scroll_up(self, event: events.MouseScrollUp) -> None:
        self._scroll(-3, event.shift)
        event.stop()

    def _scroll(self, n: int, horizontal: bool):
        fr, fc = self.frozen()
        if horizontal:
            self.left = max(fc + 1, min(MAX_COL, self.left + (1 if n > 0 else -1)))
        else:
            self.top = max(fr + 1, min(MAX_ROW, self.top + n))
        self._layout_key = None
        self.refresh()

    # ------------------------------------------------------------------ rendering

    def _style(self, fg: str | None, bg: str, bold=False, italic=False, underline=False, strike=False,
               dim=False) -> Style:
        key = (fg, bg, bold, italic, underline, strike, dim)
        st = self._styles.get(key)
        if st is None:
            st = Style(color=f"#{fg}" if fg else None, bgcolor=f"#{bg}", bold=bold, italic=italic,
                       underline=underline, strike=strike, dim=dim)
            self._styles[key] = st
        return st

    def _base_colors(self, look: CellLook, fmt_color: str | None) -> tuple[str, str]:
        p = self.palette
        bg = look.bg or p.bg
        fg = fmt_color or look.fg
        if fg is None:
            fg = p.fg if look.bg is None else ("000000" if contrast("000000", bg) >= contrast("FFFFFF", bg) else "FFFFFF")
        elif contrast(fg, bg) < 2.2:
            alt = p.fg if look.bg is None else (
                "000000" if contrast("000000", bg) >= contrast("FFFFFF", bg) else "FFFFFF")
            fg = alt
        return fg, bg

    def _display(self, r: int, c: int, cw: int):
        """-> (text, align, look, fmt_color, is_text, editing)"""
        sheet = self.sheet
        if self.edit_cell is not None and self.edit_cell == (sheet, r, c):
            look = self.resolver.look(sheet.ws._cells.get((r, c)))
            return self.edit_text.replace("\n", " "), "left", look, None, True, True
        cell = sheet.ws._cells.get((r, c))
        look = self.resolver.look(cell)
        v = self.book.get_value(sheet, r, c)
        if v is None:
            return "", "left", look, None, False, False
        fmt = cell.number_format if cell is not None else "General"
        color = None
        t = type(v)
        if t is int or t is float:
            if fmt == "General":
                text = fit_general(v, cw)
            else:
                text, color = format_value(v, fmt, self.book.epoch1904)
                if cell_len(text) > cw:
                    text = "#" * cw
            default = "right"
            is_text = False
        elif t is str:
            text, color = format_value(v, fmt, self.book.epoch1904)
            text = text.replace("\n", " ")
            default = "left"
            is_text = True
        elif t is bool:
            text, default, is_text = ("TRUE" if v else "FALSE"), "center", False
        elif isinstance(v, ExcelError):
            text, default, is_text = v.code, "center", False
        else:
            text, default, is_text = str(v), "left", True
        h = look.halign
        if h in ("left", "right", "center"):
            align = h
        elif h in ("centerContinuous", "distributed"):
            align = "center"
        elif h in ("fill", "justify"):
            align = "left"
        else:
            align = default
        if look.indent and align in ("left", "right"):
            pad = " " * min(look.indent * 2, max(0, cw - 1))
            text = pad + text if align == "left" else text + pad
        if (r, c) in sheet.stale and not is_text:
            color = color or self.palette.stale_fg
        return text, align, look, color, is_text, False

    def _active_spill(self):
        """(r1, c1, r2, c2) of the spilled range containing the cursor, if any."""
        sp = self.sheet.spill
        if not sp.anchors:
            return None
        r, c = self.active
        anchor = (r, c) if (r, c) in sp.anchors else sp.owner(r, c)
        return sp.rect(anchor) if anchor is not None else None

    def _cell_state(self, r: int, c: int, sel: Rect):
        """-> (is_cursor, is_selected, ref_color, in_clip, in_spill)"""
        is_cursor = (r, c) == self.active
        is_sel = sel.contains(r, c) and (sel.rows * sel.cols > 1)
        ref = None
        for s, rect, color in self.ref_highlights:
            if s is self.sheet and rect.contains(r, c):
                ref = color
                break
        clip = self.clip_rect is not None and self.clip_rect[0] is self.sheet and self.clip_rect[1].contains(r, c)
        sp = self._spill_hl
        spill = sp is not None and sp[0] <= r <= sp[2] and sp[1] <= c <= sp[3]
        return is_cursor, is_sel, ref, clip, spill

    def _piece_styles(self, look, fmt_color, state, editing) -> tuple[Style, Style]:
        """Styles for the content and the gutter of one cell."""
        p = self.palette
        is_cursor, is_sel, ref, clip, spill = state
        fg, bg = self._base_colors(look, fmt_color)
        gbg = bg
        if ref:
            bg = _blend(bg, ref, 0.35)
            gbg = bg
        if is_sel:
            bg = _blend(bg, p.sel_bg, 0.75 if look.bg is None else 0.45)
            gbg = bg
        if editing:
            bg, fg = p.edit_bg, p.edit_fg
        elif is_cursor:
            bg = p.cursor_bg
            fg = p.cursor_fg
        content = self._style(fg, bg, look.bold, look.italic, look.underline or clip, look.strike)
        gutter_fg = p.copy_fg if clip else p.spill_fg if spill else p.grid
        gutter = self._style(gutter_fg, gbg)
        return content, gutter

    def render_line(self, y: int) -> Strip:
        if self.sheet is None or self.book is None:
            return Strip.blank(self.size.width)
        L = self.layout()
        width = self.size.width
        p = self.palette
        sel = self.selection
        self._spill_hl = self._active_spill()
        if y < L.co_h:
            return self._band_strip(L, y, width)
        if y == L.co_h:
            return self._header_strip(L, sel, width)
        idx = y - L.co_h - 1
        if idx >= len(L.rows):
            return Strip([Segment(" " * width, self._style(None, p.bg))], width)
        r = L.rows[idx]
        segs: list[Segment] = self._row_gutter(L, r) if L.ro_w else []
        row_sel = sel.r1 <= r <= sel.r2
        hstyle = self._style(p.header_sel_fg if row_sel else p.header_fg, p.header_sel_bg if row_sel else p.header_bg,
                             bold=row_sel)
        frozen_last = idx == L.n_frozen_rows_visible - 1 and L.frozen_rows
        segs.append(Segment(str(r).rjust(L.header_w - 1) + " ", hstyle))
        merges = [mr for mr in self.sheet.ws.merged_cells.ranges if mr.min_row <= r <= mr.max_row] \
            if self.sheet.ws.merged_cells.ranges else []
        cols = L.cols
        n = len(cols)
        i = 0
        while i < n:
            c, x, w = cols[i]
            pane_end = L.n_frozen_cols_visible if i < L.n_frozen_cols_visible else n
            mr = None
            for m in merges:
                if m.min_col <= c <= m.max_col:
                    mr = m
                    break
            if mr is not None:
                j = i
                while j < pane_end and mr.min_col <= cols[j][0] <= mr.max_col:
                    j += 1
                span_w = sum(cols[k][2] for k in range(i, j))
                ar, ac = mr.min_row, mr.min_col
                va = self.resolver.look(self.sheet.ws._cells.get((ar, ac))).valign
                show_row = mr.min_row if va == "top" else (
                    (mr.min_row + mr.max_row) // 2 if va == "center" else mr.max_row)
                text, align, look, color, is_text, editing = self._display(ar, ac, span_w - 1)
                if r != show_row and not editing:
                    text = ""
                any_sel = sel.rows * sel.cols > 1 and not (
                    mr.max_row < sel.r1 or mr.min_row > sel.r2 or mr.max_col < sel.c1 or mr.min_col > sel.c2)
                st = (self.active == (ar, ac), any_sel, *self._cell_state(ar, ac, sel)[2:])
                cs, gs = self._piece_styles(look, color, st, editing)
                # lay the text out over the whole merged width, then show the visible slice
                last_c = min(mr.max_col, mr.min_col + 500)
                total = sum(self.col_px(cc) for cc in range(mr.min_col, last_c + 1))
                offset = sum(self.col_px(cc) for cc in range(mr.min_col, c))
                full = _fit(text, max(1, total - 1), align) + "│"
                piece = full[offset:offset + span_w].ljust(span_w)
                segs.append(Segment(piece[:-1], cs))
                end_char = piece[-1]
                if cols[j - 1][0] < mr.max_col:
                    boundary = L.frozen_cols and j - 1 == L.n_frozen_cols_visible - 1
                    end_char = "┃" if boundary else end_char
                    segs.append(Segment(end_char, gs if boundary else cs))
                else:
                    segs.append(Segment("│", gs))
                i = j
                continue
            text, align, look, color, is_text, editing = self._display(r, c, w - 1)
            state = self._cell_state(r, c, sel)
            text_w = cell_len(text)
            if (is_text and align == "left" and text_w > w - 1) or (editing and text_w > w - 1):
                # overflow into following empty cells
                j = i + 1
                span = w
                while j < pane_end and text_w > span - 1:
                    nc = cols[j][0]
                    if not editing and (not self._is_blank(r, nc) or any(m.min_col <= nc <= m.max_col for m in merges)):
                        break
                    span += cols[j][2]
                    j += 1
                full = _fit(text, span - 1, "left") + "│"
                pos = 0
                for k in range(i, j):
                    ck, _, wk = cols[k]
                    stk = state if k == i else self._cell_state(r, ck, sel)
                    if k > i:
                        stk = (False, *stk[1:])
                        if (r, ck) == self.active:
                            stk = (False, True, *stk[2:])
                    cs, gs = self._piece_styles(look, color, stk, editing)
                    chunk = full[pos:pos + wk]
                    if k < j - 1:
                        segs.append(Segment(chunk, cs))
                    else:
                        segs.append(Segment(chunk[:-1], cs))
                        segs.append(Segment(chunk[-1:], gs))
                    pos += wk
                i = j
                continue
            cs, gs = self._piece_styles(look, color, state, editing)
            segs.append(Segment(_fit(text, w - 1, align), cs))
            boundary = L.frozen_cols and i == L.n_frozen_cols_visible - 1
            segs.append(Segment("┃" if boundary else "│", gs))
            i += 1
        strip = Strip(segs)
        used = strip.cell_length
        if used < width:
            strip = Strip(segs + [Segment(" " * (width - used), self._style(None, p.bg))])
        strip = strip.crop(0, width)
        if frozen_last:
            strip = Strip([Segment(s.text, s.style + Style(underline=True)) if s.style else s for s in strip])
        return strip

    def _outline_styles(self) -> tuple[Style, Style, Style]:
        """-> (background, bar, button)"""
        p = self.palette
        base = self._style(p.header_fg, p.header_bg)
        return base, base, self._style(p.header_sel_fg, p.header_sel_bg, bold=True)

    def _level_buttons(self, n_levels: int, width: int) -> list[Segment]:
        """Excel's outline level buttons "1 2 …", padded or cut to width."""
        base, _, button = self._outline_styles()
        segs = []
        for k in range(1, n_levels + 2):
            segs.append(Segment(str(k), button))
            segs.append(Segment(" ", base))
        used = 2 * (n_levels + 1)
        if used > width:
            return list(Strip(segs).crop(0, width))
        return segs + [Segment(" " * (width - used), base)]

    def _row_gutter(self, L: Layout, r: int) -> list[Segment]:
        base, bar, button = self._outline_styles()
        segs = []
        for lvl in range(1, L.nrl + 1):
            m = self.outline_mark(L, "row", r, lvl)
            if m == "bar":
                segs.append(Segment("│ ", bar))
            elif m:
                segs += [Segment(m, button), Segment(" ", base)]
            else:
                segs.append(Segment("  ", base))
        segs.append(Segment(" ", base))
        return segs

    def _band_strip(self, L: Layout, y: int, width: int) -> Strip:
        """One line of the column outline band (level y + 1)."""
        base, bar, button = self._outline_styles()
        lvl = y + 1
        if y == 0:
            segs = self._level_buttons(L.ncl, L.cells_x)
        else:
            segs = [Segment(" " * L.cells_x, base)]
        after = outline.summary_after(self.sheet, "col")
        for c, x, w in L.cols:
            m = self.outline_mark(L, "col", c, lvl)
            if m == "bar":
                segs.append(Segment("─" * w, bar))
            elif m:
                label = f"[{m}]" if w >= 4 else m
                pad = " " * (w - len(label))
                segs += [Segment(label, button), Segment(pad, base)] if after else \
                    [Segment(pad, base), Segment(label, button)]
            else:
                segs.append(Segment(" " * w, base))
        strip = Strip(segs)
        if strip.cell_length < width:
            strip = Strip(segs + [Segment(" " * (width - strip.cell_length), base)])
        return strip.crop(0, width)

    def _header_strip(self, L: Layout, sel: Rect, width: int) -> Strip:
        p = self.palette
        base = self._style(p.header_fg, p.header_bg)
        segs = self._level_buttons(L.nrl, L.ro_w) if L.ro_w else []
        segs.append(Segment(" " * L.header_w, base))
        for i, (c, x, w) in enumerate(L.cols):
            on = sel.c1 <= c <= sel.c2
            st = self._style(p.header_sel_fg if on else p.header_fg, p.header_sel_bg if on else p.header_bg, bold=on)
            segs.append(Segment(_fit(col_letter(c), w - 1, "center"), st))
            boundary = L.frozen_cols and i == L.n_frozen_cols_visible - 1
            segs.append(Segment("┃" if boundary else "│", self._style(p.grid, p.header_bg)))
        strip = Strip(segs)
        if strip.cell_length < width:
            strip = Strip(segs + [Segment(" " * (width - strip.cell_length), base)])
        return strip.crop(0, width)
