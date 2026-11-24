"""Backward-compatible entry point for Excel/CSV to Microsoft SQL Server."""

from __future__ import annotations

import sys

from excel_to_mssql.cli import main as cli_main
from excel_to_mssql.gui import main as gui_main


def main() -> int:
    """Launch the CLI when `--cli` is passed, otherwise launch the desktop GUI."""
    if "--cli" in sys.argv[1:]:
        remaining = [arg for arg in sys.argv[1:] if arg != "--cli"]
        return cli_main(remaining)
    return gui_main()


if __name__ == "__main__":
    raise SystemExit(main())
