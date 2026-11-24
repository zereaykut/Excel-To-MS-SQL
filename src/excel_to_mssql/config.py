"""Configuration models, validation, persistence, and ODBC driver discovery."""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

VALID_TABLE_EXISTS_MODES = ("append", "truncate", "replace")

DRIVER_PRIORITY = (
    "ODBC Driver 18 for SQL Server",
    "ODBC Driver 17 for SQL Server",
    "ODBC Driver 13 for SQL Server",
    "SQL Server Native Client 11.0",
    "SQL Server",
    "FreeTDS",
)


@dataclass
class ObjectTypeConfig:
    """SQL Server mapping settings for string/object columns."""

    type: str = "nvarchar"
    size: int = 255
    nullable: bool = True

    def validate(self) -> None:
        if self.size != -1 and not (1 <= self.size <= 4000):
            raise ValueError(
                "nvarchar size must be between 1 and 4000 (or -1 for MAX)."
            )


@dataclass
class FloatTypeConfig:
    """SQL Server mapping settings for floating-point columns."""

    type: str = "decimal"
    size: int = 18
    precision: int = 4
    nullable: bool = True
    use_sql_float: bool = False

    def validate(self) -> None:
        if not self.use_sql_float:
            if not (1 <= self.size <= 38):
                raise ValueError("decimal size (precision) must be between 1 and 38.")
            if not (0 <= self.precision <= self.size):
                raise ValueError(
                    "decimal scale (precision) must be between 0 and decimal size."
                )


@dataclass
class IntTypeConfig:
    """SQL Server mapping settings for integer columns."""

    type: list[str] = field(default_factory=lambda: ["int", "bigint"])
    type_index: int = 0
    nullable: bool = True

    def selected_sql_type(self) -> str:
        idx = self.type_index if 0 <= self.type_index < len(self.type) else 0
        return self.type[idx]


@dataclass
class DatetimeTypeConfig:
    """SQL Server mapping settings for date/datetime columns."""

    type: list[str] = field(default_factory=lambda: ["datetime", "date"])
    type_index: int = 0
    nullable: bool = True

    def selected_sql_type(self) -> str:
        idx = self.type_index if 0 <= self.type_index < len(self.type) else 0
        return self.type[idx]


@dataclass
class BoolTypeConfig:
    """SQL Server mapping settings for boolean columns."""

    type: str = "bit"
    nullable: bool = True


@dataclass
class DataTypesConfig:
    """Container for all pandas-to-SQL Server type mappings."""

    object: ObjectTypeConfig = field(default_factory=ObjectTypeConfig)
    float: FloatTypeConfig = field(default_factory=FloatTypeConfig)
    int: IntTypeConfig = field(default_factory=IntTypeConfig)
    datetime: DatetimeTypeConfig = field(default_factory=DatetimeTypeConfig)
    bool: BoolTypeConfig = field(default_factory=BoolTypeConfig)

    def validate(self) -> None:
        self.object.validate()
        self.float.validate()


@dataclass
class AppConfig:
    """Application and database connection configuration."""

    server_ip: str = "localhost"
    username: str = "sa"
    password: str = ""
    database_name: str = "master"
    table_name: str = "SampleData"
    schema_name: str = "dbo"
    driver: str = "auto"
    trusted_connection: bool = False
    trust_server_certificate: bool = True
    batch_size: int = 1000
    if_table_exists: str = "append"
    data_types: DataTypesConfig = field(default_factory=DataTypesConfig)

    def validate(self, require_target: bool = True) -> None:
        if not self.server_ip.strip():
            raise ValueError("Server IP / Host cannot be empty.")
        if not self.database_name.strip():
            raise ValueError("Database name cannot be empty.")
        if require_target and not self.table_name.strip():
            raise ValueError("Table name cannot be empty.")
        if not self.schema_name.strip():
            raise ValueError("Schema name cannot be empty.")
        if not self.trusted_connection and not self.username.strip():
            raise ValueError(
                "Username cannot be empty unless Trusted Connection is enabled."
            )
        if self.batch_size < 1:
            raise ValueError("Batch size must be at least 1.")
        if self.if_table_exists not in VALID_TABLE_EXISTS_MODES:
            raise ValueError(
                f"if_table_exists must be one of {VALID_TABLE_EXISTS_MODES}."
            )
        self.data_types.validate()

    def to_dict(self) -> dict[str, Any]:
        return {
            "server_ip": self.server_ip,
            "username": self.username,
            "password": self.password,
            "database_name": self.database_name,
            "table_name": self.table_name,
            "schema_name": self.schema_name,
            "driver": self.driver,
            "trusted_connection": self.trusted_connection,
            "trust_server_certificate": self.trust_server_certificate,
            "batch_size": self.batch_size,
            "if_table_exists": self.if_table_exists,
            "data_types": {
                "object": {
                    "type": self.data_types.object.type,
                    "size": self.data_types.object.size,
                    "nullable": int(self.data_types.object.nullable),
                },
                "float": {
                    "type": self.data_types.float.type,
                    "size": self.data_types.float.size,
                    "precision": self.data_types.float.precision,
                    "nullable": int(self.data_types.float.nullable),
                    "use_sql_float": int(self.data_types.float.use_sql_float),
                },
                "int": {
                    "type": list(self.data_types.int.type),
                    "type_index": int(self.data_types.int.type_index),
                    "nullable": int(self.data_types.int.nullable),
                },
                "datetime": {
                    "type": list(self.data_types.datetime.type),
                    "type_index": int(self.data_types.datetime.type_index),
                    "nullable": int(self.data_types.datetime.nullable),
                },
                "bool": {
                    "type": self.data_types.bool.type,
                    "nullable": int(self.data_types.bool.nullable),
                },
            },
        }

    @classmethod
    def from_dict(cls, raw: dict[str, Any]) -> AppConfig:
        dt_raw = raw.get("data_types", {})
        obj_raw = dt_raw.get("object", {})
        flt_raw = dt_raw.get("float", {})
        int_raw = dt_raw.get("int", {})
        dtm_raw = dt_raw.get("datetime", {})
        bool_raw = dt_raw.get("bool", {})

        data_types = DataTypesConfig(
            object=ObjectTypeConfig(
                type=str(obj_raw.get("type", "nvarchar")),
                size=int(obj_raw.get("size", 255)),
                nullable=bool(obj_raw.get("nullable", 1)),
            ),
            float=FloatTypeConfig(
                type=str(flt_raw.get("type", "decimal")),
                size=int(flt_raw.get("size", 18)),
                precision=int(flt_raw.get("precision", 4)),
                nullable=bool(flt_raw.get("nullable", 1)),
                use_sql_float=bool(flt_raw.get("use_sql_float", 0)),
            ),
            int=IntTypeConfig(
                type=list(int_raw.get("type", ["int", "bigint"])),
                type_index=int(int_raw.get("type_index", 0)),
                nullable=bool(int_raw.get("nullable", 1)),
            ),
            datetime=DatetimeTypeConfig(
                type=list(dtm_raw.get("type", ["datetime", "date"])),
                type_index=int(dtm_raw.get("type_index", 0)),
                nullable=bool(dtm_raw.get("nullable", 1)),
            ),
            bool=BoolTypeConfig(
                type=str(bool_raw.get("type", "bit")),
                nullable=bool(bool_raw.get("nullable", 1)),
            ),
        )

        return cls(
            server_ip=str(raw.get("server_ip", "localhost")),
            username=str(raw.get("username", "sa")),
            password=str(raw.get("password", "")),
            database_name=str(raw.get("database_name", "master")),
            table_name=str(raw.get("table_name", "SampleData")),
            schema_name=str(raw.get("schema_name", "dbo")),
            driver=str(raw.get("driver", "auto")),
            trusted_connection=bool(raw.get("trusted_connection", False)),
            trust_server_certificate=bool(raw.get("trust_server_certificate", True)),
            batch_size=int(raw.get("batch_size", 1000)),
            if_table_exists=str(raw.get("if_table_exists", "append")),
            data_types=data_types,
        )


def discover_sql_drivers() -> list[str]:
    """Return installed ODBC drivers that appear to support SQL Server."""
    try:
        import pyodbc

        installed = list(pyodbc.drivers())
    except Exception:
        return []

    sql_drivers = [
        drv for drv in installed if "sql server" in drv.lower() or "freetds" in drv.lower()
    ]
    priority_map = {name.lower(): idx for idx, name in enumerate(DRIVER_PRIORITY)}
    sql_drivers.sort(key=lambda d: (priority_map.get(d.lower(), 999), d.lower()))
    return sql_drivers


def load_config(
    path: str | Path = "config.json",
    fallback_path: str | Path | None = "config.example.json",
) -> AppConfig:
    """Load configuration from JSON, falling back to example config or defaults."""
    target = Path(path)
    if target.exists():
        with target.open("r", encoding="utf-8") as fh:
            return AppConfig.from_dict(json.load(fh))

    if fallback_path is not None:
        fallback = Path(fallback_path)
        if fallback.exists():
            with fallback.open("r", encoding="utf-8") as fh:
                return AppConfig.from_dict(json.load(fh))

    return AppConfig()


def save_config(config: AppConfig, path: str | Path = "config.json") -> Path:
    """Validate and persist configuration as formatted JSON."""
    config.validate(require_target=False)
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    with target.open("w", encoding="utf-8") as fh:
        json.dump(config.to_dict(), fh, indent=4)
        fh.write("\n")
    return target
