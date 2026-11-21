"""Command-line entry point: txl [FILE] [--sheet NAME]"""
from __future__ import annotations

import argparse
import sys


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="txl",
        description="A terminal spreadsheet that reads and writes Excel (.xlsx/.xlsm) and CSV files.",
    )
    parser.add_argument("file", nargs="?", help="workbook to open (created on first save if it doesn't exist)")
    parser.add_argument("--sheet", help="sheet to show first")
    parser.add_argument("--version", action="version", version="term_xl 0.1.0")
    args = parser.parse_args(argv)

    from .ui.app import TermXLApp

    app = TermXLApp(args.file, sheet=args.sheet)
    app.run()
    return 0


if __name__ == "__main__":
    sys.exit(main())
