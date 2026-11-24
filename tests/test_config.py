"""Tests for configuration loading, saving, validation, and driver resolution."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from excel_to_mssql.config import AppConfig, load_config, save_config


def test_load_config_falls_back_to_example_and_defaults(tmp_path: Path) -> None:
    missing = tmp_path / "missing.json"
    example = tmp_path / "config.example.json"
    example.write_text(
        json.dumps({"server_ip": "10.0.0.5", "database_name": "Analytics"}),
        encoding="utf-8",
    )

    cfg = load_config(missing, fallback_path=example)
    assert cfg.server_ip == "10.0.0.5"
    assert cfg.database_name == "Analytics"
    assert cfg.schema_name == "dbo"

    defaults = load_config(missing, fallback_path=tmp_path / "also_missing.json")
    assert defaults.server_ip == "localhost"


def test_save_and_reload_round_trip(tmp_path: Path) -> None:
    cfg = AppConfig(
        server_ip="db.internal",
        username="etl_user",
        password="secret",
        database_name="Warehouse",
        table_name="Sales",
        schema_name="staging",
        trusted_connection=False,
        batch_size=500,
        if_table_exists="truncate",
    )
    cfg.data_types.int.type_index = 1
    cfg.data_types.datetime.type_index = 1
    cfg.data_types.float.use_sql_float = True

    out_path = save_config(cfg, tmp_path / "config.json")
    reloaded = load_config(out_path)

    assert reloaded.server_ip == "db.internal"
    assert reloaded.schema_name == "staging"
    assert reloaded.batch_size == 500
    assert reloaded.if_table_exists == "truncate"
    assert reloaded.data_types.int.selected_sql_type() == "bigint"
    assert reloaded.data_types.datetime.selected_sql_type() == "date"
    assert reloaded.data_types.float.use_sql_float is True


def test_config_validation_catches_invalid_values() -> None:
    cfg = AppConfig()
    cfg.data_types.object.size = 5000
    with pytest.raises(ValueError, match="nvarchar size"):
        cfg.validate()

    cfg = AppConfig()
    cfg.data_types.float.size = 10
    cfg.data_types.float.precision = 12
    with pytest.raises(ValueError, match="decimal scale"):
        cfg.validate()

    cfg = AppConfig(batch_size=0)
    with pytest.raises(ValueError, match="Batch size"):
        cfg.validate()
