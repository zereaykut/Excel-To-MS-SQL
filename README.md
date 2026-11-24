# Excel to MS SQL

A modern desktop GUI and CLI tool for inspecting Excel (`.xlsx`, `.xlsm`, `.xls`) and CSV/TSV files, mapping column data types, creating tables, and importing data into Microsoft SQL Server in fast parameterized batches.

## Features

- **Desktop GUI & CLI**: Use the `ttk` desktop application (`excel-to-mssql`) or automate imports from the terminal (`excel-to-mssql-cli`).
- **Multi-Format & Multi-Sheet Support**: Import `.xlsx`, `.xlsm`, `.xls`, `.csv`, and `.tsv` files with Excel worksheet selection and configurable CSV delimiters and encodings.
- **Live Data & Schema Preview**: Preview the first 100 rows and inspect inferred source dtypes alongside mapped SQL Server column definitions (`nvarchar`, `decimal`/`float`, `int`/`bigint`, `datetime`/`date`, `bit`) before writing to the database.
- **Fast, Safe Batch Inserts**: Uses `pyodbc` parameterized queries with `fast_executemany`, configurable `batch_size`, transactional commits/rollbacks, bracket-escaped identifiers, and background worker threads so the GUI stays responsive.
- **Flexible Authentication & Table Modes**: Supports SQL Server authentication, Windows Integrated Authentication (`Trusted_Connection`), `TrustServerCertificate`, automatic ODBC driver discovery, and `append` / `truncate` / `replace` table policies.
- **Modern `uv` Workflow**: Managed with `uv`, `pyproject.toml`, and `uv.lock` for reproducible environments.

## Project Structure

```text
Excel-To-MS-SQL/
├── app.py                      # Backward-compatible launcher (GUI by default, --cli for CLI)
├── config.example.json         # Safe template configuration (copy to config.json)
├── examples/
│   └── data.xlsx               # Sample Excel dataset
├── pyproject.toml              # Project metadata, dependencies, and tool settings
├── uv.lock                     # Locked dependency graph
├── src/
│   └── excel_to_mssql/
│       ├── __init__.py
│       ├── config.py           # Typed configuration, validation, ODBC driver discovery
│       ├── reader.py           # Excel/CSV loading, sheet listing, dtype classification
│       ├── database.py         # SQL Server DDL generation and fast batch insertion
│       ├── cli.py              # Command-line interface
│       └── gui.py              # Tkinter/ttk desktop application
└── tests/
    ├── test_config.py
    ├── test_database.py
    └── test_reader.py
```

## Prerequisites

- **Python**: `>= 3.10`
- **Package Manager**: [`uv`](https://docs.astral.sh/uv/)
- **ODBC Driver**: [Microsoft ODBC Driver for SQL Server](https://learn.microsoft.com/en-us/sql/connect/odbc/download-odbc-driver-for-sql-server) (e.g., ODBC Driver 17 or 18).
  - On **Linux**, also ensure `unixodbc` (`libodbc.so.2`) is installed.
  - To use the desktop GUI, ensure Tk (`tkinter`) is available in your Python installation.

## Installation

Clone the repository and sync dependencies with `uv`:

```shell
git clone https://github.com/zereaykut/Excel-To-MS-SQL.git
cd Excel-To-MS-SQL
uv sync
```

## Configuration

Copy the example configuration file (or configure and save directly from the GUI):

```shell
cp config.example.json config.json
```

> **Note:** `config.json` is listed in `.gitignore` so database credentials are never committed to version control. If `config.json` does not exist, the application automatically falls back to `config.example.json` or built-in defaults.

### Type Mapping Overview

| Source Data Type | SQL Server Target Options | Notes |
| --- | --- | --- |
| String / Object | `nvarchar(N)` or `nvarchar(MAX)` | Set `size` to `1..4000` or `-1` for `MAX` |
| Float | `decimal(size, precision)` or `float` | Configurable scale/precision or native `float` |
| Integer | `int` or `bigint` | Configurable via `type_index` (`0` = `int`, `1` = `bigint`) |
| Date / Datetime | `datetime` or `date` | Configurable via `type_index` (`0` = `datetime`, `1` = `date`) |
| Boolean | `bit` | Supports `NULL` / `NOT NULL` |

## Usage

### Desktop GUI

Launch the GUI via `uv`:

```shell
uv run excel-to-mssql
```

Or using the root launcher:

```shell
uv run python app.py
```

Inside the GUI you can:
1. Configure **Connection & Target** (`Server`, `Database`, `Schema`, `Table`, `ODBC Driver`, SQL or Windows Auth) and click **Test Connection**.
2. Adjust **Data Types & Import Options** (`append`, `truncate`, `replace`, `batch_size`).
3. Click **Choose File...** (or **Load Sample Data** to open `examples/data.xlsx`), inspect the **Data Preview** and **Column & SQL Type Mapping** tabs, and run **Create Table**, **Insert Data**, or **Create & Insert**.

### Command-Line Interface (CLI)

Inspect a file and preview the generated SQL DDL without connecting to SQL Server (`--dry-run`):

```shell
uv run excel-to-mssql-cli --file examples/data.xlsx --dry-run
```

Test database connectivity:

```shell
uv run excel-to-mssql-cli --test-connection
```

Import a file directly into SQL Server:

```shell
uv run excel-to-mssql-cli \
  --file examples/data.xlsx \
  --server localhost \
  --database Analytics \
  --schema dbo \
  --table ProjectTasks \
  --username sa \
  --password "YourPassword" \
  --if-exists replace \
  --batch-size 1000
```

You can also invoke the CLI via `app.py --cli`:

```shell
uv run python app.py --cli --file examples/data.xlsx --dry-run
```

## Development & Testing

Run the test suite and linter with `uv`:

```shell
uv run pytest -v
uv run ruff check src tests app.py
```

## License

This project is licensed under the GNU General Public License v3.0 — see the [LICENSE](LICENSE) file for details.
