"""Tests for SQL generation, row preparation, and mocked pyodbc execution."""

from __future__ import annotations

import datetime as dt
from unittest.mock import MagicMock, patch

import numpy as np
import pandas as pd

from excel_to_mssql.config import AppConfig
from excel_to_mssql.database import (
    build_connection_string,
    build_create_table_sql,
    build_insert_sql,
    create_table,
    import_dataframe,
    insert_data,
    map_column_sql_type,
    prepare_rows,
    quote_identifier,
    resolve_driver,
)


def test_quote_identifier_escapes_closing_brackets() -> None:
    assert quote_identifier("NormalName") == "[NormalName]"
    assert quote_identifier("weird]col") == "[weird]]col]"


def test_resolve_driver_and_connection_string() -> None:
    drivers = ["SQL Server", "ODBC Driver 18 for SQL Server"]
    auto_driver = resolve_driver(
        "auto",
        available_drivers=["ODBC Driver 18 for SQL Server", "SQL Server"],
    )
    assert auto_driver == "ODBC Driver 18 for SQL Server"
    assert resolve_driver(drivers[0], available_drivers=drivers) == "SQL Server"

    cfg = AppConfig(
        server_ip="sql.local",
        database_name="SalesDB",
        username="etl",
        password="pw",
        trusted_connection=False,
        trust_server_certificate=True,
    )
    conn_str = build_connection_string(cfg, available_drivers=["ODBC Driver 18 for SQL Server"])
    assert "DRIVER={ODBC Driver 18 for SQL Server}" in conn_str
    assert "UID=etl;PWD=pw" in conn_str
    assert "TrustServerCertificate=yes" in conn_str

    cfg.trusted_connection = True
    trusted_str = build_connection_string(cfg, available_drivers=["ODBC Driver 17 for SQL Server"])
    assert "Trusted_Connection=yes" in trusted_str
    assert "PWD=" not in trusted_str


def test_type_mapping_respects_int_and_datetime_indexes() -> None:
    cfg = AppConfig()
    cfg.data_types.int.type_index = 1  # bigint
    cfg.data_types.datetime.type_index = 1  # date
    cfg.data_types.object.size = -1  # MAX
    cfg.data_types.object.nullable = False

    assert map_column_sql_type("int", cfg.data_types) == "[bigint] NULL"
    assert map_column_sql_type("datetime", cfg.data_types) == "[date] NULL"
    assert map_column_sql_type("object", cfg.data_types) == "[nvarchar](MAX) NOT NULL"
    assert map_column_sql_type("bool", cfg.data_types) == "[bit] NULL"


def test_build_ddl_and_insert_sql() -> None:
    df = pd.DataFrame(
        {
            "id": [1, 2],
            "amount": [10.12345, 20.5],
            "created_at": pd.to_datetime(["2026-01-01", "2026-01-02"]),
            "active": [True, False],
            "label": ["a", "b"],
        }
    )
    cfg = AppConfig(schema_name="analytics", table_name="Events")
    ddl, qualified = build_create_table_sql(df, cfg)

    assert qualified == "[analytics].[Events]"
    assert "IF OBJECT_ID(?, 'U') IS NULL" in ddl
    assert "[id] [int] NULL" in ddl
    assert "[amount] [decimal](18, 4) NULL" in ddl
    assert "[created_at] [datetime] NULL" in ddl
    assert "[active] [bit] NULL" in ddl
    assert "[label] [nvarchar](255) NULL" in ddl

    insert_sql = build_insert_sql(df, cfg)
    assert insert_sql == (
        "INSERT INTO [analytics].[Events] "
        "([id], [amount], [created_at], [active], [label]) "
        "VALUES (?, ?, ?, ?, ?);"
    )


def test_prepare_rows_does_not_mutate_df_and_handles_nulls() -> None:
    df = pd.DataFrame(
        {
            "id": pd.Series([1, 2], dtype="int64"),
            "score": pd.Series([3.141592, np.nan], dtype="float64"),
            "when": pd.to_datetime(["2026-05-01 12:30:00", None]),
            "flag": pd.Series([True, False], dtype="bool"),
            "note": ["ok", None],
        }
    )
    cfg = AppConfig()
    cfg.data_types.float.precision = 2
    cfg.data_types.datetime.type_index = 1  # date

    rows = prepare_rows(df, cfg.data_types)
    assert rows[0] == (1, 3.14, dt.date(2026, 5, 1), True, "ok")
    assert rows[1] == (2, None, None, False, None)
    # Input DataFrame must remain unrounded
    assert df.loc[0, "score"] == 3.141592


def test_create_table_and_batched_insert_with_mock_pyodbc() -> None:
    df = pd.DataFrame({"id": [1, 2, 3, 4, 5], "name": ["a", "b", "c", "d", "e"]})
    cfg = AppConfig(batch_size=2, if_table_exists="replace")

    mock_conn = MagicMock()
    mock_cursor = MagicMock()
    mock_conn.cursor.return_value = mock_cursor
    mock_conn.__enter__.return_value = mock_conn
    mock_pyodbc = MagicMock()
    mock_pyodbc.connect.return_value = mock_conn

    progress_updates: list[tuple[int, int]] = []

    with patch("excel_to_mssql.database._get_pyodbc", return_value=mock_pyodbc):
        total = import_dataframe(
            df,
            cfg,
            progress_callback=lambda done, tot: progress_updates.append((done, tot)),
        )

    assert total == 5
    assert progress_updates == [(2, 5), (4, 5), (5, 5)]
    assert mock_cursor.fast_executemany is True
    assert mock_cursor.executemany.call_count == 3
    assert mock_conn.commit.call_count == 2


def test_insert_truncate_executes_truncate_first() -> None:
    df = pd.DataFrame({"id": [10]})
    cfg = AppConfig(if_table_exists="truncate")

    mock_conn = MagicMock()
    mock_cursor = MagicMock()
    mock_conn.cursor.return_value = mock_cursor
    mock_conn.__enter__.return_value = mock_conn
    mock_pyodbc = MagicMock()
    mock_pyodbc.connect.return_value = mock_conn

    with patch("excel_to_mssql.database._get_pyodbc", return_value=mock_pyodbc):
        create_table(df, cfg)
        insert_data(df, cfg)

    executed_sqls = [call.args[0] for call in mock_cursor.execute.call_args_list]
    assert any("TRUNCATE TABLE [dbo].[SampleData]" in sql for sql in executed_sqls)
