"""Excel/CSV to Microsoft SQL Server importer package."""

from __future__ import annotations

import logging
from pathlib import Path

from .config import (
    AppConfig,
    BoolTypeConfig,
    DataTypesConfig,
    DatetimeTypeConfig,
    FloatTypeConfig,
    IntTypeConfig,
    ObjectTypeConfig,
    discover_sql_drivers,
    load_config,
    save_config,
)
from .database import (
    build_connection_string,
    build_create_table_sql,
    create_table,
    import_dataframe,
    insert_data,
    map_column_sql_type,
    prepare_rows,
    quote_identifier,
    resolve_driver,
    test_connection,
)
from .reader import classify_dtype, get_data_info, list_excel_sheets, load_tabular_file

__version__ = "2.0.0"

__all__ = [
    "AppConfig",
    "BoolTypeConfig",
    "DataTypesConfig",
    "DatetimeTypeConfig",
    "FloatTypeConfig",
    "IntTypeConfig",
    "ObjectTypeConfig",
    "__version__",
    "build_connection_string",
    "build_create_table_sql",
    "classify_dtype",
    "configure_logging",
    "create_table",
    "discover_sql_drivers",
    "get_data_info",
    "import_dataframe",
    "insert_data",
    "list_excel_sheets",
    "load_config",
    "load_tabular_file",
    "map_column_sql_type",
    "prepare_rows",
    "quote_identifier",
    "resolve_driver",
    "save_config",
    "test_connection",
]


def configure_logging(
    log_file: str | Path = "app.log",
    level: int = logging.INFO,
) -> logging.Logger:
    """Configure root logger for the application without duplicate handlers."""
    root_logger = logging.getLogger("excel_to_mssql")
    root_logger.setLevel(level)

    if not root_logger.handlers:
        formatter = logging.Formatter(
            "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
        )
        file_handler = logging.FileHandler(log_file, encoding="utf-8")
        file_handler.setFormatter(formatter)
        root_logger.addHandler(file_handler)

    return root_logger
