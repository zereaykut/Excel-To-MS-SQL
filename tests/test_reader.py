"""Tests for Excel and CSV loading and dtype classification."""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest

from excel_to_mssql.reader import (
    _normalize_columns,
    classify_dtype,
    get_data_info,
    list_excel_sheets,
    load_tabular_file,
)


def test_load_sample_excel_file() -> None:
    sample_path = Path(__file__).resolve().parents[1] / "examples" / "data.xlsx"
    assert sample_path.exists()

    sheets = list_excel_sheets(sample_path)
    assert len(sheets) >= 1

    df = load_tabular_file(sample_path, sheet_name=sheets[0])
    assert not df.empty
    assert "Project Name" in df.columns

    info = get_data_info(df)
    assert set(["column", "pandas_dtype", "logical_type", "non_null_count", "null_count"]).issubset(
        info.columns
    )


def test_load_csv_and_normalize_columns(tmp_path: Path) -> None:
    csv_file = tmp_path / "sample.csv"
    csv_file.write_text(
        " Name ,Age,Amount,Active,Created\n"
        "Alice,30,12.5,true,2026-01-15\n"
        "Bob,,40.25,false,2026-02-20\n",
        encoding="utf-8",
    )

    df = load_tabular_file(csv_file)
    assert list(df.columns) == ["Name", "Age", "Amount", "Active", "Created"]
    assert classify_dtype(df["Name"].dtype) == "object"
    assert classify_dtype(df["Amount"].dtype) == "float"
    assert classify_dtype(df["Active"].dtype) == "bool"
    assert classify_dtype(df["Created"].dtype) == "datetime"


def test_unsupported_extension_raises(tmp_path: Path) -> None:
    bad_file = tmp_path / "data.json"
    bad_file.write_text("{}", encoding="utf-8")
    with pytest.raises(ValueError, match="Unsupported file format"):
        load_tabular_file(bad_file)


def test_duplicate_and_blank_column_names() -> None:
    raw = pd.DataFrame([[1, 2, 3]], columns=["col", "col", None])
    normalized = _normalize_columns(raw)
    assert list(normalized.columns) == ["col", "col_1", "column_3"]
