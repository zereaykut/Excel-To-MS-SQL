"""Excel and CSV file loading, sheet inspection, and schema inference."""

from __future__ import annotations

from pathlib import Path

import pandas as pd
from pandas.api import types as ptypes

EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xls"}
CSV_EXTENSIONS = {".csv", ".tsv"}
SUPPORTED_EXTENSIONS = EXCEL_EXTENSIONS | CSV_EXTENSIONS


def is_excel_file(file_path: str | Path) -> bool:
    """Return True if the path has a supported Excel file extension."""
    return Path(file_path).suffix.lower() in EXCEL_EXTENSIONS


def list_excel_sheets(file_path: str | Path) -> list[str]:
    """Return the list of worksheet names in an Excel workbook."""
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"File not found: {path}")
    if not is_excel_file(path):
        raise ValueError(f"Not an Excel file: {path.name}")

    with pd.ExcelFile(path) as excel_file:
        return [str(name) for name in excel_file.sheet_names]


def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Ensure every column header is a non-empty, unique string."""
    normalized: list[str] = []
    seen: dict[str, int] = {}

    for idx, col in enumerate(df.columns):
        name = str(col).strip() if col is not None and not pd.isna(col) else ""
        if not name:
            name = f"column_{idx + 1}"
        count = seen.get(name, 0)
        seen[name] = count + 1
        if count > 0:
            name = f"{name}_{count}"
        normalized.append(name)

    result = df.copy()
    result.columns = normalized
    return result


def load_tabular_file(
    file_path: str | Path,
    sheet_name: str | int | None = 0,
    delimiter: str | None = None,
    encoding: str = "utf-8",
) -> pd.DataFrame:
    """Load an Excel or CSV/TSV file into a normalized DataFrame."""
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"File not found: {path}")

    suffix = path.suffix.lower()
    if suffix not in SUPPORTED_EXTENSIONS:
        supported = ", ".join(sorted(SUPPORTED_EXTENSIONS))
        raise ValueError(f"Unsupported file format '{suffix}'. Supported: {supported}")

    if suffix in EXCEL_EXTENSIONS:
        df = pd.read_excel(path, sheet_name=0 if sheet_name is None else sheet_name)
    else:
        sep = delimiter if delimiter is not None else ("\t" if suffix == ".tsv" else ",")
        df = pd.read_csv(path, sep=sep, encoding=encoding)
        # Attempt lightweight date parsing for string columns that look like ISO dates
        for col in df.columns:
            if ptypes.is_object_dtype(df[col]) or ptypes.is_string_dtype(df[col]):
                non_null = df[col].dropna()
                if len(non_null) > 0 and non_null.astype(str).str.match(
                    r"^\d{4}-\d{2}-\d{2}([ T]\d{2}:\d{2}(:\d{2})?)?$"
                ).all():
                    df[col] = pd.to_datetime(df[col], errors="coerce")

    return _normalize_columns(df)


def classify_dtype(dtype: object) -> str:
    """Map a pandas/numpy dtype into a canonical logical type name."""
    if ptypes.is_bool_dtype(dtype):
        return "bool"
    if ptypes.is_datetime64_any_dtype(dtype):
        return "datetime"
    if ptypes.is_integer_dtype(dtype):
        return "int"
    if ptypes.is_float_dtype(dtype):
        return "float"
    return "object"


def get_data_info(df: pd.DataFrame) -> pd.DataFrame:
    """Inspect a DataFrame and return per-column type and null statistics."""
    rows = []
    for col in df.columns:
        series = df[col]
        null_count = int(series.isna().sum())
        rows.append(
            {
                "column": str(col),
                "pandas_dtype": str(series.dtype),
                "logical_type": classify_dtype(series.dtype),
                "non_null_count": int(len(series) - null_count),
                "null_count": null_count,
            }
        )
    return pd.DataFrame(rows)
