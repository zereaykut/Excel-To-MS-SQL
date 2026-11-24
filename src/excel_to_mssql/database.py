"""Microsoft SQL Server connection, DDL generation, and batch insertion."""

from __future__ import annotations

import datetime as dt
import logging
from collections.abc import Callable, Sequence
from typing import Any

import numpy as np
import pandas as pd

from .config import AppConfig, DataTypesConfig, discover_sql_drivers
from .reader import classify_dtype

logger = logging.getLogger(__name__)


def _get_pyodbc() -> Any:
    """Lazily import pyodbc and raise a clear error if system ODBC libs are missing."""
    try:
        import pyodbc
    except ImportError as exc:
        raise RuntimeError(
            "Unable to load pyodbc. On Linux, ensure unixODBC (libodbc.so.2) and "
            "Microsoft ODBC Driver for SQL Server are installed."
        ) from exc
    return pyodbc


def quote_identifier(identifier: str) -> str:
    """Safely bracket-quote a SQL Server identifier, escaping closing brackets."""
    cleaned = identifier.strip()
    if not cleaned:
        raise ValueError("SQL identifier cannot be empty.")
    return f"[{cleaned.replace(']', ']]')}]"


def format_qualified_table(schema_name: str, table_name: str) -> str:
    """Return `[schema].[table]` for SQL statements."""
    return f"{quote_identifier(schema_name)}.{quote_identifier(table_name)}"


def resolve_driver(
    configured_driver: str,
    available_drivers: Sequence[str] | None = None,
) -> str:
    """Resolve ODBC driver name from config or auto-detected installed drivers."""
    if configured_driver and configured_driver.lower() != "auto":
        return configured_driver

    drivers = list(available_drivers) if available_drivers is not None else discover_sql_drivers()
    if drivers:
        return drivers[0]

    return "ODBC Driver 17 for SQL Server"


def build_connection_string(
    config: AppConfig,
    available_drivers: Sequence[str] | None = None,
) -> str:
    """Build a pyodbc connection string from AppConfig."""
    driver = resolve_driver(config.driver, available_drivers=available_drivers)
    parts = [
        f"DRIVER={{{driver}}}",
        f"SERVER={config.server_ip}",
        f"DATABASE={config.database_name}",
    ]

    if config.trusted_connection:
        parts.append("Trusted_Connection=yes")
    else:
        parts.append(f"UID={config.username}")
        parts.append(f"PWD={config.password}")

    if config.trust_server_certificate:
        parts.append("TrustServerCertificate=yes")

    return ";".join(parts)


def map_column_sql_type(
    logical_type: str,
    data_types: DataTypesConfig,
) -> str:
    """Return the SQL Server column type definition (including NULL/NOT NULL)."""
    if logical_type == "datetime":
        sql_type = data_types.datetime.selected_sql_type()
        null_clause = "NULL" if data_types.datetime.nullable else "NOT NULL"
        return f"[{sql_type}] {null_clause}"

    if logical_type == "int":
        sql_type = data_types.int.selected_sql_type()
        null_clause = "NULL" if data_types.int.nullable else "NOT NULL"
        return f"[{sql_type}] {null_clause}"

    if logical_type == "float":
        null_clause = "NULL" if data_types.float.nullable else "NOT NULL"
        if data_types.float.use_sql_float:
            return f"[float] {null_clause}"
        return (
            f"[{data_types.float.type}]({data_types.float.size}, "
            f"{data_types.float.precision}) {null_clause}"
        )

    if logical_type == "bool":
        null_clause = "NULL" if data_types.bool.nullable else "NOT NULL"
        return f"[{data_types.bool.type}] {null_clause}"

    # Default / string / object
    null_clause = "NULL" if data_types.object.nullable else "NOT NULL"
    size_str = "MAX" if data_types.object.size <= 0 else str(data_types.object.size)
    return f"[{data_types.object.type}]({size_str}) {null_clause}"


def build_create_table_sql(
    df: pd.DataFrame,
    config: AppConfig,
) -> tuple[str, str]:
    """Generate parameterized `CREATE TABLE` DDL and the qualified object name."""
    if df.empty and len(df.columns) == 0:
        raise ValueError("Cannot create a table from a DataFrame with no columns.")

    column_defs: list[str] = []
    for col in df.columns:
        logical_type = classify_dtype(df[col].dtype)
        sql_def = map_column_sql_type(logical_type, config.data_types)
        column_defs.append(f"    {quote_identifier(str(col))} {sql_def}")

    qualified_table = format_qualified_table(config.schema_name, config.table_name)
    cols_block = ",\n".join(column_defs)

    ddl = (
        "IF OBJECT_ID(?, 'U') IS NULL\n"
        "BEGIN\n"
        f"    CREATE TABLE {qualified_table} (\n"
        f"{cols_block}\n"
        "    );\n"
        "END;"
    )
    return ddl, qualified_table


def _convert_cell_value(
    value: Any,
    logical_type: str,
    data_types: DataTypesConfig,
) -> Any:
    """Convert a pandas/numpy cell value into a pyodbc-compatible Python object."""
    if value is None or pd.isna(value):
        return None

    if logical_type == "datetime":
        if isinstance(value, pd.Timestamp):
            py_dt = value.to_pydatetime()
        elif isinstance(value, np.datetime64):
            py_dt = pd.Timestamp(value).to_pydatetime()
        elif isinstance(value, dt.datetime):
            py_dt = value
        elif isinstance(value, dt.date):
            return value
        else:
            py_dt = pd.to_datetime(value).to_pydatetime()

        if data_types.datetime.selected_sql_type().lower() == "date":
            return py_dt.date()
        return py_dt

    if logical_type == "int":
        return int(value)

    if logical_type == "float":
        num = float(value)
        if not data_types.float.use_sql_float:
            num = round(num, data_types.float.precision)
        return num

    if logical_type == "bool":
        return bool(value)

    return str(value)


def prepare_rows(df: pd.DataFrame, data_types: DataTypesConfig) -> list[tuple[Any, ...]]:
    """Convert a DataFrame into clean row tuples without mutating the input DataFrame."""
    logical_types = [classify_dtype(df[col].dtype) for col in df.columns]
    prepared: list[tuple[Any, ...]] = []

    for row in df.itertuples(index=False, name=None):
        converted = tuple(
            _convert_cell_value(val, ltype, data_types)
            for val, ltype in zip(row, logical_types, strict=True)
        )
        prepared.append(converted)

    return prepared


def build_insert_sql(df: pd.DataFrame, config: AppConfig) -> str:
    """Generate parameterized `INSERT INTO` statement for the DataFrame columns."""
    if len(df.columns) == 0:
        raise ValueError("DataFrame has no columns to insert.")

    qualified_table = format_qualified_table(config.schema_name, config.table_name)
    quoted_cols = ", ".join(quote_identifier(str(col)) for col in df.columns)
    placeholders = ", ".join("?" for _ in df.columns)
    return f"INSERT INTO {qualified_table} ({quoted_cols}) VALUES ({placeholders});"


def test_connection(config: AppConfig) -> str:
    """Test SQL Server connectivity and return the server version string."""
    config.validate(require_target=False)
    conn_str = build_connection_string(config)
    pyodbc = _get_pyodbc()
    logger.info("Testing connection to %s / %s", config.server_ip, config.database_name)

    with pyodbc.connect(conn_str, timeout=10) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT @@VERSION;")
        row = cursor.fetchone()
        return str(row[0]).splitlines()[0] if row else "Connected"


def create_table(
    df: pd.DataFrame,
    config: AppConfig,
    drop_existing: bool = False,
) -> str:
    """Create the target table in SQL Server (optionally dropping an existing one first)."""
    config.validate(require_target=True)
    ddl, qualified_table = build_create_table_sql(df, config)
    conn_str = build_connection_string(config)
    pyodbc = _get_pyodbc()

    logger.info("Creating table %s on %s", qualified_table, config.database_name)
    logger.debug("DDL:\n%s", ddl)

    with pyodbc.connect(conn_str) as conn:
        cursor = conn.cursor()
        try:
            if drop_existing:
                cursor.execute(f"DROP TABLE IF EXISTS {qualified_table};")
            cursor.execute(ddl, qualified_table)
            conn.commit()
        except Exception:
            conn.rollback()
            raise

    return ddl


def insert_data(
    df: pd.DataFrame,
    config: AppConfig,
    progress_callback: Callable[[int, int], None] | None = None,
) -> int:
    """Insert DataFrame rows into SQL Server in fast parameterized batches."""
    config.validate(require_target=True)
    total_rows = len(df)
    if total_rows == 0:
        logger.info("DataFrame is empty; nothing to insert.")
        return 0

    insert_sql = build_insert_sql(df, config)
    rows = prepare_rows(df, config.data_types)
    conn_str = build_connection_string(config)
    pyodbc = _get_pyodbc()
    batch_size = max(1, config.batch_size)
    qualified_table = format_qualified_table(config.schema_name, config.table_name)

    logger.info(
        "Inserting %d rows into %s (batch_size=%d)",
        total_rows,
        qualified_table,
        batch_size,
    )

    inserted = 0
    with pyodbc.connect(conn_str) as conn:
        cursor = conn.cursor()
        try:
            cursor.fast_executemany = True
        except Exception:
            logger.debug("fast_executemany not supported by ODBC driver; using standard mode.")

        try:
            if config.if_table_exists == "truncate":
                cursor.execute(f"TRUNCATE TABLE {qualified_table};")

            for start in range(0, total_rows, batch_size):
                batch = rows[start : start + batch_size]
                cursor.executemany(insert_sql, batch)
                inserted += len(batch)
                if progress_callback is not None:
                    progress_callback(inserted, total_rows)

            conn.commit()
        except Exception:
            conn.rollback()
            raise

    logger.info("Successfully inserted %d rows into %s.", inserted, qualified_table)
    return inserted


def import_dataframe(
    df: pd.DataFrame,
    config: AppConfig,
    progress_callback: Callable[[int, int], None] | None = None,
) -> int:
    """Create/replace target table according to config and insert all rows."""
    config.validate(require_target=True)
    drop_first = config.if_table_exists == "replace"
    create_table(df, config, drop_existing=drop_first)
    return insert_data(df, config, progress_callback=progress_callback)
