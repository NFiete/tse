"""Map openpyxl cell styles (fonts, fills, alignment) to terminal colors/attributes."""
from __future__ import annotations

import colorsys
from dataclasses import dataclass

from openpyxl.styles.colors import COLOR_INDEX

_DEFAULT_THEME = ["FFFFFF", "000000", "E7E6E6", "44546A", "4472C4", "ED7D31", "A5A5A5", "FFC000",
                  "5B9BD5", "70AD47", "0563C1", "954F72"]


def parse_theme_colors(theme_xml) -> list[str]:
    """Return theme colors in Excel's index order (lt1, dk1, lt2, dk2, accent1-6, hlink, folHlink)."""
    if not theme_xml:
        return list(_DEFAULT_THEME)
    try:
        from lxml import etree

        root = etree.fromstring(theme_xml if isinstance(theme_xml, bytes) else theme_xml.encode())
        ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
        scheme = root.find(".//a:clrScheme", ns)
        if scheme is None:
            return list(_DEFAULT_THEME)
        vals = {}
        for child in scheme:
            tag = etree.QName(child).localname
            sub = child[0] if len(child) else None
            if sub is None:
                continue
            v = sub.get("lastClr") or sub.get("val")
            if v:
                vals[tag] = v[-6:].upper()
        order = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
                 "hlink", "folHlink"]
        return [vals.get(k, d) for k, d in zip(order, _DEFAULT_THEME)]
    except Exception:
        return list(_DEFAULT_THEME)


def apply_tint(hex6: str, tint: float) -> str:
    if not tint:
        return hex6
    r, g, b = (int(hex6[i:i + 2], 16) / 255 for i in (0, 2, 4))
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    if tint < 0:
        l = l * (1 + tint)
    else:
        l = l * (1 - tint) + tint
    r, g, b = colorsys.hls_to_rgb(h, max(0, min(1, l)), s)
    return f"{round(r * 255):02X}{round(g * 255):02X}{round(b * 255):02X}"


def resolve_color(color, theme: list[str]) -> str | None:
    """openpyxl Color -> 'RRGGBB' or None for automatic/system colors."""
    if color is None:
        return None
    try:
        t = color.type
    except AttributeError:
        return None
    try:
        if t == "rgb":
            v = color.rgb
            if not isinstance(v, str) or len(v) < 6:
                return None
            return v[-6:].upper()
        if t == "theme":
            idx = color.theme
            if idx is None or idx >= len(theme):
                return None
            return apply_tint(theme[idx], color.tint or 0)
        if t == "indexed":
            idx = color.indexed
            if idx is None or idx >= 64:
                return None  # 64/65 = system foreground/background
            if idx < len(COLOR_INDEX):
                return apply_tint(COLOR_INDEX[idx][-6:].upper(), color.tint or 0)
        if t == "auto":
            return None
    except Exception:
        return None
    return None


def luminance(hex6: str) -> float:
    def ch(v):
        v = v / 255
        return v / 12.92 if v <= 0.03928 else ((v + 0.055) / 1.055) ** 2.4

    r, g, b = (int(hex6[i:i + 2], 16) for i in (0, 2, 4))
    return 0.2126 * ch(r) + 0.7152 * ch(g) + 0.0722 * ch(b)


def contrast(a: str, b: str) -> float:
    la, lb = luminance(a), luminance(b)
    if la < lb:
        la, lb = lb, la
    return (la + 0.05) / (lb + 0.05)


@dataclass(frozen=True)
class CellLook:
    fg: str | None  # RRGGBB or None = default
    bg: str | None
    bold: bool = False
    italic: bool = False
    underline: bool = False
    strike: bool = False
    halign: str | None = None  # left/right/center/general/...
    valign: str | None = None
    indent: int = 0
    wrap: bool = False


class StyleResolver:
    """Caches CellLook per openpyxl style id (the StyleArray contents)."""

    def __init__(self, wb):
        self.theme = parse_theme_colors(getattr(wb, "loaded_theme", None))
        self._cache: dict[tuple, CellLook] = {}

    def look(self, cell) -> CellLook:
        if cell is None or cell._style is None:
            return _PLAIN
        key = tuple(cell._style)
        lk = self._cache.get(key)
        if lk is None:
            lk = self._compute(cell)
            self._cache[key] = lk
        return lk

    def invalidate(self):
        self._cache.clear()

    def _compute(self, cell) -> CellLook:
        font = cell.font
        fill = cell.fill
        al = cell.alignment
        fg = resolve_color(font.color, self.theme) if font is not None else None
        bg = None
        if fill is not None:
            pt = getattr(fill, "patternType", None) or getattr(fill, "fill_type", None)
            if pt and pt != "none":
                c = fill.fgColor if pt == "solid" else (fill.fgColor if fill.fgColor is not None else fill.bgColor)
                bg = resolve_color(c, self.theme)
                if bg is None and pt == "solid":
                    bg = None
            elif getattr(fill, "stop", None):
                try:
                    bg = resolve_color(fill.stop[0].color, self.theme)
                except Exception:
                    bg = None
        u = font.underline if font is not None else None
        return CellLook(
            fg=fg, bg=bg,
            bold=bool(font.b) if font is not None else False,
            italic=bool(font.i) if font is not None else False,
            underline=bool(u) and u != "none",
            strike=bool(font.strike) if font is not None else False,
            halign=al.horizontal if al is not None else None,
            valign=al.vertical if al is not None else None,
            indent=int(al.indent or 0) if al is not None else 0,
            wrap=bool(al.wrap_text) if al is not None else False,
        )


_PLAIN = CellLook(None, None)
