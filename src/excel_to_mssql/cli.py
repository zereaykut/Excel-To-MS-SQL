"""Command-line interface for Excel/CSV to Microsoft SQL Server."""

from __future__ import annotations

import argparse
import sys
from collections.abc import Sequence

from . import __version__, configure_logging
from .config import VALID_TABLE_EXISTS_MODES, load_config
from .database import (
    build_create_table_sql,
    import_dataframe,
    map_column_sql_type,
    test_connection,
)
from .reader import get_data_info, load_tabular_file


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="excel-to-mssql-cli",
        description="Import Excel or CSV/TSV files into Microsoft SQL Server.",
    )
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    parser.add_argument(
        "-c",
        "--config",
        default="config.json",
        help="Path to JSON configuration file (default: config.json).",
    )
    parser.add_argument(
        "-f",
        "--file",
        help="Path to the Excel (.xlsx/.xlsm/.xls) or CSV/TSV file to import.",
    )
    parser.add_argument(
        "--sheet",
        default=None,
        help="Excel worksheet name or 0-based index (default: first sheet).",
    )
    parser.add_argument(
        "--delimiter",
        default=None,
        help="CSV delimiter (default: ',' for CSV, tab for TSV).",
    )
    parser.add_argument(
        "--encoding",
        default="utf-8",
        help="CSV text encoding (default: utf-8).",
    )
    parser.add_argument("--server", help="Override SQL Server host/IP.")
    parser.add_argument("--database", help="Override target database name.")
    parser.add_argument("--schema", help="Override target schema (default: dbo).")
    parser.add_argument("--table", help="Override target table name.")
    parser.add_argument("--username", help="Override SQL login username.")
    parser.add_argument("--password", help="Override SQL login password.")
    parser.add_argument("--driver", help="Override ODBC driver name or 'auto'.")
    parser.add_argument(
        "--trusted-connection",
        action="store_true",
        default=None,
        help="Use Windows Integrated Authentication.",
    )
    parser.add_argument(
        "--if-exists",
        choices=VALID_TABLE_EXISTS_MODES,
        default=None,
        help="Behavior if target table already exists.",
    )
    parser.add_argument(
        "--batch-size",
        type=int,
        default=None,
        help="Number of rows per executemany batch.",
    )
    parser.add_argument(
        "--test-connection",
        action="store_true",
        help="Test database connectivity and exit.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Inspect file and print generated SQL DDL without connecting to SQL Server.",
    )
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    configure_logging()
    parser = build_parser()
    args = parser.parse_args(argv)

    config = load_config(args.config)
    if args.server:
        config.server_ip = args.server
    if args.database:
        config.database_name = args.database
    if args.schema:
        config.schema_name = args.schema
    if args.table:
        config.table_name = args.table
    if args.username:
        config.username = args.username
    if args.password is not None:
        config.password = args.password
    if args.driver:
        config.driver = args.driver
    if args.trusted_connection is not None:
        config.trusted_connection = args.trusted_connection
    if args.if_exists:
        config.if_table_exists = args.if_exists
    if args.batch_size is not None:
        config.batch_size = args.batch_size

    if args.test_connection:
        try:
            version_info = test_connection(config)
        except Exception as exc:
            print(f"Connection failed: {exc}", file=sys.stderr)
            return 1
        print(f"Connected successfully: {version_info}")
        return 0

    if not args.file:
        parser.error("--file is required unless --test-connection is specified.")

    sheet: str | int | None = args.sheet
    if isinstance(sheet, str) and sheet.isdigit():
        sheet = int(sheet)

    try:
        df = load_tabular_file(
            args.file,
            sheet_name=sheet,
            delimiter=args.delimiter,
            encoding=args.encoding,
        )
    except Exception as exc:
        print(f"Failed to read input file: {exc}", file=sys.stderr)
        return 1

    info_df = get_data_info(df)
    info_df["sql_type"] = info_df["logical_type"].apply(
        lambda lt: map_column_sql_type(lt, config.data_types)
    )
    print(f"Loaded {len(df)} rows and {len(df.columns)} columns from '{args.file}'.")
    print(info_df.to_string(index=False))

    if args.dry_run:
        ddl, qualified = build_create_table_sql(df, config)
        print(f"\n-- Dry Run DDL for {qualified}\n{ddl}")
        return 0

    def _on_progress(done: int, total: int) -> None:
        pct = int((done / total) * 100) if total else 100
        print(f"Inserted {done}/{total} rows ({pct}%)...", end="\r", flush=True)

    try:
        inserted = import_dataframe(df, config, progress_callback=_on_progress)
    except Exception as exc:
        print(f"\nImport failed: {exc}", file=sys.stderr)
        return 1

    print(
        f"\nImport complete: {inserted} rows inserted into "
        f"{config.schema_name}.{config.table_name}."
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
