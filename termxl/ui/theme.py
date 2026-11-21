"""Color palettes for the grid (dark terminal mode and light "paper" mode)."""
from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class Palette:
    name: str
    bg: str
    fg: str
    grid: str
    header_bg: str
    header_fg: str
    header_sel_bg: str
    header_sel_fg: str
    sel_bg: str
    cursor_bg: str
    cursor_fg: str
    edit_bg: str
    edit_fg: str
    copy_fg: str
    stale_fg: str
    ref_colors: tuple[str, ...]


DARK = Palette(
    name="dark",
    bg="1B1D23", fg="D4D4D4", grid="30333B",
    header_bg="262930", header_fg="8B909A", header_sel_bg="3A4150", header_sel_fg="FFFFFF",
    sel_bg="27405F", cursor_bg="3B7DD8", cursor_fg="FFFFFF",
    edit_bg="3D3720", edit_fg="FFFFFF", copy_fg="7EE787", stale_fg="D7A35A",
    ref_colors=("4F8CFF", "E5534B", "8E6CEF", "3FB950", "D29922", "DB61A2", "39C5CF"),
)

LIGHT = Palette(
    name="light",
    bg="FFFFFF", fg="000000", grid="E1E1E1",
    header_bg="F2F2F2", header_fg="555555", header_sel_bg="D2E3FC", header_sel_fg="0B3D91",
    sel_bg="D6E6FB", cursor_bg="8AB4F8", cursor_fg="000000",
    edit_bg="FFF4CC", edit_fg="000000", copy_fg="1A7F37", stale_fg="9A6700",
    ref_colors=("1F6FEB", "CF222E", "8250DF", "1A7F37", "9A6700", "BF3989", "0A7480"),
)
