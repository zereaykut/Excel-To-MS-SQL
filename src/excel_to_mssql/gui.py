"""Modern Tkinter/ttk desktop GUI for Excel/CSV to Microsoft SQL Server."""

from __future__ import annotations

import logging
import sys
import threading
from pathlib import Path
from typing import Any

import pandas as pd

from . import __version__, configure_logging
from .config import (
    VALID_TABLE_EXISTS_MODES,
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
    create_table,
    import_dataframe,
    insert_data,
    map_column_sql_type,
    test_connection,
)
from .reader import get_data_info, is_excel_file, list_excel_sheets, load_tabular_file

logger = logging.getLogger(__name__)

SAMPLE_DATA_PATH = Path("examples/data.xlsx")
PREVIEW_ROW_LIMIT = 100


def _import_tk():
    import tkinter as tk
    from tkinter import filedialog, messagebox, scrolledtext, ttk

    return tk, ttk, filedialog, messagebox, scrolledtext


class ExcelToMSSQLApp:
    """Main Tkinter desktop application controller."""

    def __init__(self, root: Any, config_path: Path = Path("config.json")) -> None:
        tk, ttk, filedialog, messagebox, scrolledtext = _import_tk()
        self.tk = tk
        self.ttk = ttk
        self.filedialog = filedialog
        self.messagebox = messagebox
        self.scrolledtext = scrolledtext

        self.root = root
        self.config_path = config_path
        self.config = load_config(config_path)

        self.df: pd.DataFrame | None = None
        self.current_file: Path | None = None
        self.is_busy = False

        self.root.title(f"Excel/CSV to MS SQL Server v{__version__}")
        self.root.geometry("1020x740")
        self.root.minsize(880, 620)

        self._configure_styles()
        self._init_variables()
        self._build_ui()
        self._apply_auth_state()
        self._log("Application initialized. Load an Excel or CSV file to begin.")

    def _configure_styles(self) -> None:
        style = self.ttk.Style(self.root)
        available_themes = style.theme_names()
        for preferred in ("clam", "vista", "aqua", "default"):
            if preferred in available_themes:
                style.theme_use(preferred)
                break

        style.configure("Header.TLabel", font=("Segoe UI", 14, "bold"))
        style.configure("Subheader.TLabel", font=("Segoe UI", 9), foreground="#555555")
        style.configure("Section.TLabelframe.Label", font=("Segoe UI", 10, "bold"))
        style.configure("Accent.TButton", font=("Segoe UI", 9, "bold"))

    def _init_variables(self) -> None:
        tk = self.tk
        cfg = self.config

        # Connection variables
        self.server_var = tk.StringVar(value=cfg.server_ip)
        self.database_var = tk.StringVar(value=cfg.database_name)
        self.schema_var = tk.StringVar(value=cfg.schema_name)
        self.table_var = tk.StringVar(value=cfg.table_name)
        self.driver_var = tk.StringVar(value=cfg.driver)
        self.username_var = tk.StringVar(value=cfg.username)
        self.password_var = tk.StringVar(value=cfg.password)
        self.show_password_var = tk.BooleanVar(value=False)
        self.trusted_conn_var = tk.BooleanVar(value=cfg.trusted_connection)
        self.trust_cert_var = tk.BooleanVar(value=cfg.trust_server_certificate)

        # Data type variables
        self.nvarchar_size_var = tk.StringVar(value=str(cfg.data_types.object.size))
        self.nvarchar_nullable_var = tk.BooleanVar(value=cfg.data_types.object.nullable)

        self.decimal_size_var = tk.StringVar(value=str(cfg.data_types.float.size))
        self.decimal_precision_var = tk.StringVar(value=str(cfg.data_types.float.precision))
        self.decimal_nullable_var = tk.BooleanVar(value=cfg.data_types.float.nullable)
        self.use_sql_float_var = tk.BooleanVar(value=cfg.data_types.float.use_sql_float)

        self.int_bigint_var = tk.BooleanVar(value=cfg.data_types.int.type_index == 1)
        self.int_nullable_var = tk.BooleanVar(value=cfg.data_types.int.nullable)

        self.date_only_var = tk.BooleanVar(value=cfg.data_types.datetime.type_index == 1)
        self.date_nullable_var = tk.BooleanVar(value=cfg.data_types.datetime.nullable)

        self.bool_nullable_var = tk.BooleanVar(value=cfg.data_types.bool.nullable)

        # Import options
        self.if_exists_var = tk.StringVar(value=cfg.if_table_exists)
        self.batch_size_var = tk.StringVar(value=str(cfg.batch_size))

        # File options
        self.file_path_var = tk.StringVar(value="No file selected")
        self.sheet_var = tk.StringVar(value="")
        self.delimiter_var = tk.StringVar(value=",")
        self.encoding_var = tk.StringVar(value="utf-8")
        self.file_summary_var = tk.StringVar(value="Rows: 0 | Columns: 0")
        self.status_var = tk.StringVar(value="Ready")

    def _build_ui(self) -> None:
        ttk = self.ttk
        tk = self.tk

        outer = ttk.Frame(self.root, padding=12)
        outer.pack(fill=tk.BOTH, expand=True)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(2, weight=1)

        # Header
        header = ttk.Frame(outer)
        header.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        ttk.Label(
            header,
            text="Excel / CSV to Microsoft SQL Server",
            style="Header.TLabel",
        ).pack(anchor="w")
        ttk.Label(
            header,
            text=(
                "Configure your SQL Server target, inspect data & SQL type mappings, "
                "and run batch imports."
            ),
            style="Subheader.TLabel",
        ).pack(anchor="w")

        # Top settings notebook
        settings_nb = ttk.Notebook(outer)
        settings_nb.grid(row=1, column=0, sticky="ew", pady=(0, 10))

        conn_tab = ttk.Frame(settings_nb, padding=10)
        types_tab = ttk.Frame(settings_nb, padding=10)
        settings_nb.add(conn_tab, text="Connection & Target")
        settings_nb.add(types_tab, text="Data Types & Import Options")

        self._build_connection_tab(conn_tab)
        self._build_types_tab(types_tab)

        # File & Preview workspace
        workspace = ttk.LabelFrame(
            outer,
            text="Source File & Preview",
            style="Section.TLabelframe",
            padding=10,
        )
        workspace.grid(row=2, column=0, sticky="nsew", pady=(0, 10))
        workspace.columnconfigure(0, weight=1)
        workspace.rowconfigure(2, weight=1)

        self._build_file_bar(workspace)
        self._build_preview_notebook(workspace)

        # Bottom action bar
        self._build_action_bar(outer)

    def _build_connection_tab(self, parent: Any) -> None:
        ttk = self.ttk
        for col in (1, 3, 5):
            parent.columnconfigure(col, weight=1)

        # Row 0: Server, Database, Schema
        ttk.Label(parent, text="Server / Host:").grid(
            row=0, column=0, sticky="w", padx=(0, 6), pady=4
        )
        ttk.Entry(parent, textvariable=self.server_var).grid(
            row=0, column=1, sticky="ew", padx=(0, 14), pady=4
        )

        ttk.Label(parent, text="Database:").grid(
            row=0, column=2, sticky="w", padx=(0, 6), pady=4
        )
        ttk.Entry(parent, textvariable=self.database_var).grid(
            row=0, column=3, sticky="ew", padx=(0, 14), pady=4
        )

        ttk.Label(parent, text="Schema:").grid(
            row=0, column=4, sticky="w", padx=(0, 6), pady=4
        )
        ttk.Entry(parent, textvariable=self.schema_var, width=14).grid(
            row=0, column=5, sticky="ew", pady=4
        )

        # Row 1: Table, Driver, Actions
        ttk.Label(parent, text="Table Name:").grid(
            row=1, column=0, sticky="w", padx=(0, 6), pady=4
        )
        ttk.Entry(parent, textvariable=self.table_var).grid(
            row=1, column=1, sticky="ew", padx=(0, 14), pady=4
        )

        ttk.Label(parent, text="ODBC Driver:").grid(
            row=1, column=2, sticky="w", padx=(0, 6), pady=4
        )
        drivers = ["auto", *discover_sql_drivers()]
        ttk.Combobox(
            parent,
            textvariable=self.driver_var,
            values=drivers,
            state="normal",
        ).grid(row=1, column=3, columnspan=3, sticky="ew", pady=4)

        # Row 2: Username, Password, Auth toggles
        ttk.Label(parent, text="Username:").grid(row=2, column=0, sticky="w", padx=(0, 6), pady=4)
        self.username_entry = ttk.Entry(parent, textvariable=self.username_var)
        self.username_entry.grid(row=2, column=1, sticky="ew", padx=(0, 14), pady=4)

        ttk.Label(parent, text="Password:").grid(row=2, column=2, sticky="w", padx=(0, 6), pady=4)
        pwd_frame = ttk.Frame(parent)
        pwd_frame.grid(row=2, column=3, sticky="ew", padx=(0, 14), pady=4)
        pwd_frame.columnconfigure(0, weight=1)
        self.password_entry = ttk.Entry(pwd_frame, textvariable=self.password_var, show="*")
        self.password_entry.grid(row=0, column=0, sticky="ew")
        ttk.Checkbutton(
            pwd_frame,
            text="Show",
            variable=self.show_password_var,
            command=self._toggle_password_visibility,
        ).grid(row=0, column=1, padx=(6, 0))

        auth_flags = ttk.Frame(parent)
        auth_flags.grid(row=2, column=4, columnspan=2, sticky="w", pady=4)
        ttk.Checkbutton(
            auth_flags,
            text="Windows Auth",
            variable=self.trusted_conn_var,
            command=self._apply_auth_state,
        ).pack(side="left", padx=(0, 10))
        ttk.Checkbutton(
            auth_flags,
            text="Trust Server Cert",
            variable=self.trust_cert_var,
        ).pack(side="left")

        # Row 3: Buttons
        btn_row = ttk.Frame(parent)
        btn_row.grid(row=3, column=0, columnspan=6, sticky="e", pady=(8, 0))
        self.test_conn_btn = ttk.Button(
            btn_row,
            text="Test Connection",
            command=self.on_test_connection,
        )
        self.test_conn_btn.pack(side="left", padx=(0, 8))
        ttk.Button(
            btn_row,
            text="Save Config",
            command=self.on_save_config,
        ).pack(side="left")

    def _build_types_tab(self, parent: Any) -> None:
        ttk = self.ttk
        parent.columnconfigure(1, weight=1)
        parent.columnconfigure(3, weight=1)

        # String
        ttk.Label(parent, text="String [nvarchar] size (-1 = MAX):").grid(
            row=0, column=0, sticky="w", padx=(0, 8), pady=4
        )
        ttk.Entry(parent, textvariable=self.nvarchar_size_var, width=12).grid(
            row=0, column=1, sticky="w", pady=4
        )
        ttk.Checkbutton(
            parent,
            text="Nullable",
            variable=self.nvarchar_nullable_var,
            command=self._refresh_schema_preview,
        ).grid(row=0, column=2, sticky="w", padx=12, pady=4)

        # Float
        ttk.Label(parent, text="Decimal size / precision:").grid(
            row=1, column=0, sticky="w", padx=(0, 8), pady=4
        )
        flt_frame = ttk.Frame(parent)
        flt_frame.grid(row=1, column=1, sticky="w", pady=4)
        ttk.Entry(flt_frame, textvariable=self.decimal_size_var, width=6).pack(side="left")
        ttk.Label(flt_frame, text=" , ").pack(side="left")
        ttk.Entry(flt_frame, textvariable=self.decimal_precision_var, width=6).pack(side="left")

        flt_opts = ttk.Frame(parent)
        flt_opts.grid(row=1, column=2, columnspan=2, sticky="w", padx=12, pady=4)
        ttk.Checkbutton(
            flt_opts,
            text="Nullable",
            variable=self.decimal_nullable_var,
            command=self._refresh_schema_preview,
        ).pack(side="left", padx=(0, 12))
        ttk.Checkbutton(
            flt_opts,
            text="Use SQL [float] instead of [decimal]",
            variable=self.use_sql_float_var,
            command=self._refresh_schema_preview,
        ).pack(side="left")

        # Int
        ttk.Label(parent, text="Integer mapping:").grid(
            row=2, column=0, sticky="w", padx=(0, 8), pady=4
        )
        ttk.Checkbutton(
            parent,
            text="Use [bigint] instead of [int]",
            variable=self.int_bigint_var,
            command=self._refresh_schema_preview,
        ).grid(row=2, column=1, sticky="w", pady=4)
        ttk.Checkbutton(
            parent,
            text="Nullable",
            variable=self.int_nullable_var,
            command=self._refresh_schema_preview,
        ).grid(row=2, column=2, sticky="w", padx=12, pady=4)

        # Datetime
        ttk.Label(parent, text="Date/Time mapping:").grid(
            row=3, column=0, sticky="w", padx=(0, 8), pady=4
        )
        ttk.Checkbutton(
            parent,
            text="Use [date] instead of [datetime]",
            variable=self.date_only_var,
            command=self._refresh_schema_preview,
        ).grid(row=3, column=1, sticky="w", pady=4)
        ttk.Checkbutton(
            parent,
            text="Nullable",
            variable=self.date_nullable_var,
            command=self._refresh_schema_preview,
        ).grid(row=3, column=2, sticky="w", padx=12, pady=4)

        # Bool + Import behavior
        ttk.Label(parent, text="Boolean [bit]:").grid(
            row=4, column=0, sticky="w", padx=(0, 8), pady=4
        )
        ttk.Checkbutton(
            parent,
            text="Nullable",
            variable=self.bool_nullable_var,
            command=self._refresh_schema_preview,
        ).grid(row=4, column=1, sticky="w", pady=4)

        import_frame = ttk.Frame(parent)
        import_frame.grid(row=4, column=2, columnspan=2, sticky="w", padx=12, pady=4)
        ttk.Label(import_frame, text="If table exists:").pack(side="left", padx=(0, 6))
        ttk.Combobox(
            import_frame,
            textvariable=self.if_exists_var,
            values=list(VALID_TABLE_EXISTS_MODES),
            state="readonly",
            width=10,
        ).pack(side="left", padx=(0, 14))
        ttk.Label(import_frame, text="Batch size:").pack(side="left", padx=(0, 6))
        ttk.Entry(import_frame, textvariable=self.batch_size_var, width=8).pack(side="left")

        ttk.Button(
            parent,
            text="Refresh Schema Preview",
            command=self._refresh_schema_preview,
        ).grid(row=5, column=3, sticky="e", pady=(6, 0))

    def _build_file_bar(self, parent: Any) -> None:
        ttk = self.ttk

        top_bar = ttk.Frame(parent)
        top_bar.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        top_bar.columnconfigure(1, weight=1)

        ttk.Button(top_bar, text="Choose File...", command=self.on_choose_file).grid(
            row=0, column=0, padx=(0, 8)
        )
        ttk.Label(top_bar, textvariable=self.file_path_var).grid(
            row=0, column=1, sticky="ew"
        )
        if SAMPLE_DATA_PATH.exists():
            ttk.Button(
                top_bar,
                text="Load Sample Data",
                command=lambda: self._load_file(SAMPLE_DATA_PATH),
            ).grid(row=0, column=2, padx=(8, 0))

        opts_bar = ttk.Frame(parent)
        opts_bar.grid(row=1, column=0, sticky="ew", pady=(0, 8))

        ttk.Label(opts_bar, text="Excel Sheet:").pack(side="left", padx=(0, 6))
        self.sheet_combo = ttk.Combobox(
            opts_bar,
            textvariable=self.sheet_var,
            state="disabled",
            width=22,
        )
        self.sheet_combo.pack(side="left", padx=(0, 14))
        self.sheet_combo.bind("<<ComboboxSelected>>", self._on_sheet_changed)

        ttk.Label(opts_bar, text="CSV Delimiter:").pack(side="left", padx=(0, 6))
        self.delimiter_entry = ttk.Entry(opts_bar, textvariable=self.delimiter_var, width=4)
        self.delimiter_entry.pack(side="left", padx=(0, 10))

        ttk.Label(opts_bar, text="Encoding:").pack(side="left", padx=(0, 6))
        self.encoding_entry = ttk.Entry(opts_bar, textvariable=self.encoding_var, width=10)
        self.encoding_entry.pack(side="left", padx=(0, 10))

        ttk.Button(opts_bar, text="Reload File", command=self.on_reload_file).pack(
            side="left", padx=(0, 12)
        )
        ttk.Label(opts_bar, textvariable=self.file_summary_var, style="Subheader.TLabel").pack(
            side="right"
        )

    def _build_preview_notebook(self, parent: Any) -> None:
        ttk = self.ttk

        preview_nb = ttk.Notebook(parent)
        preview_nb.grid(row=2, column=0, sticky="nsew")

        # Tab 1: Data Preview
        data_tab = ttk.Frame(preview_nb)
        data_tab.rowconfigure(0, weight=1)
        data_tab.columnconfigure(0, weight=1)
        preview_nb.add(data_tab, text="Data Preview (Top 100)")

        self.data_tree = ttk.Treeview(data_tab, show="headings")
        data_vsb = ttk.Scrollbar(data_tab, orient="vertical", command=self.data_tree.yview)
        data_hsb = ttk.Scrollbar(data_tab, orient="horizontal", command=self.data_tree.xview)
        self.data_tree.configure(yscrollcommand=data_vsb.set, xscrollcommand=data_hsb.set)
        self.data_tree.grid(row=0, column=0, sticky="nsew")
        data_vsb.grid(row=0, column=1, sticky="ns")
        data_hsb.grid(row=1, column=0, sticky="ew")

        # Tab 2: Schema Preview
        schema_tab = ttk.Frame(preview_nb)
        schema_tab.rowconfigure(0, weight=1)
        schema_tab.columnconfigure(0, weight=1)
        preview_nb.add(schema_tab, text="Column & SQL Type Mapping")

        schema_cols = ("column", "pandas_dtype", "logical_type", "sql_type", "non_null", "nulls")
        self.schema_tree = ttk.Treeview(schema_tab, columns=schema_cols, show="headings")
        for col_id, heading, width in (
            ("column", "Column Name", 220),
            ("pandas_dtype", "Source Dtype", 120),
            ("logical_type", "Logical Type", 110),
            ("sql_type", "SQL Server Type", 200),
            ("non_null", "Non-Null", 90),
            ("nulls", "Nulls", 90),
        ):
            self.schema_tree.heading(col_id, text=heading)
            self.schema_tree.column(col_id, width=width, anchor="w")

        schema_vsb = ttk.Scrollbar(schema_tab, orient="vertical", command=self.schema_tree.yview)
        self.schema_tree.configure(yscrollcommand=schema_vsb.set)
        self.schema_tree.grid(row=0, column=0, sticky="nsew")
        schema_vsb.grid(row=0, column=1, sticky="ns")

        # Tab 3: Activity Log
        log_tab = ttk.Frame(preview_nb)
        log_tab.rowconfigure(0, weight=1)
        log_tab.columnconfigure(0, weight=1)
        preview_nb.add(log_tab, text="Activity Log")

        self.log_text = self.scrolledtext.ScrolledText(
            log_tab, state="disabled", height=8, wrap="word", font=("Consolas", 9)
        )
        self.log_text.grid(row=0, column=0, sticky="nsew")

    def _build_action_bar(self, parent: Any) -> None:
        ttk = self.ttk
        bar = ttk.Frame(parent)
        bar.grid(row=3, column=0, sticky="ew")
        bar.columnconfigure(0, weight=1)

        status_frame = ttk.Frame(bar)
        status_frame.grid(row=0, column=0, sticky="ew", padx=(0, 12))
        status_frame.columnconfigure(0, weight=1)

        self.progress = ttk.Progressbar(status_frame, mode="determinate", maximum=100)
        self.progress.grid(row=0, column=0, sticky="ew", pady=(0, 4))
        ttk.Label(status_frame, textvariable=self.status_var, style="Subheader.TLabel").grid(
            row=1, column=0, sticky="w"
        )

        btns = ttk.Frame(bar)
        btns.grid(row=0, column=1, sticky="e")
        self.create_btn = ttk.Button(btns, text="Create Table", command=self.on_create_table)
        self.create_btn.pack(side="left", padx=(0, 8))
        self.insert_btn = ttk.Button(btns, text="Insert Data", command=self.on_insert_data)
        self.insert_btn.pack(side="left", padx=(0, 8))
        self.import_btn = ttk.Button(
            btns,
            text="Create & Insert",
            style="Accent.TButton",
            command=self.on_import_all,
        )
        self.import_btn.pack(side="left")

    def _toggle_password_visibility(self) -> None:
        self.password_entry.configure(show="" if self.show_password_var.get() else "*")

    def _apply_auth_state(self) -> None:
        state = "disabled" if self.trusted_conn_var.get() else "normal"
        self.username_entry.configure(state=state)
        self.password_entry.configure(state=state)

    def _log(self, message: str, level: int = logging.INFO) -> None:
        logger.log(level, message)
        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"{message}\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _build_config_from_ui(self, require_target: bool = True) -> AppConfig:
        try:
            nvarchar_size = int(self.nvarchar_size_var.get().strip())
        except ValueError as exc:
            raise ValueError("nvarchar size must be an integer (1-4000 or -1 for MAX).") from exc

        try:
            decimal_size = int(self.decimal_size_var.get().strip())
            decimal_precision = int(self.decimal_precision_var.get().strip())
        except ValueError as exc:
            raise ValueError("Decimal size and precision must be integers.") from exc

        try:
            batch_size = int(self.batch_size_var.get().strip())
        except ValueError as exc:
            raise ValueError("Batch size must be an integer.") from exc

        cfg = AppConfig(
            server_ip=self.server_var.get().strip(),
            username=self.username_var.get().strip(),
            password=self.password_var.get(),
            database_name=self.database_var.get().strip(),
            table_name=self.table_var.get().strip(),
            schema_name=self.schema_var.get().strip() or "dbo",
            driver=self.driver_var.get().strip() or "auto",
            trusted_connection=self.trusted_conn_var.get(),
            trust_server_certificate=self.trust_cert_var.get(),
            batch_size=batch_size,
            if_table_exists=self.if_exists_var.get(),
            data_types=DataTypesConfig(
                object=ObjectTypeConfig(
                    type="nvarchar",
                    size=nvarchar_size,
                    nullable=self.nvarchar_nullable_var.get(),
                ),
                float=FloatTypeConfig(
                    type="decimal",
                    size=decimal_size,
                    precision=decimal_precision,
                    nullable=self.decimal_nullable_var.get(),
                    use_sql_float=self.use_sql_float_var.get(),
                ),
                int=IntTypeConfig(
                    type=["int", "bigint"],
                    type_index=1 if self.int_bigint_var.get() else 0,
                    nullable=self.int_nullable_var.get(),
                ),
                datetime=DatetimeTypeConfig(
                    type=["datetime", "date"],
                    type_index=1 if self.date_only_var.get() else 0,
                    nullable=self.date_nullable_var.get(),
                ),
                bool=BoolTypeConfig(
                    type="bit",
                    nullable=self.bool_nullable_var.get(),
                ),
            ),
        )
        cfg.validate(require_target=require_target)
        return cfg

    def on_save_config(self) -> None:
        try:
            cfg = self._build_config_from_ui(require_target=False)
            saved = save_config(cfg, self.config_path)
            self.config = cfg
        except Exception as exc:
            self.messagebox.showerror("Invalid Configuration", str(exc))
            return

        self._log(f"Configuration saved to {saved}")
        self.messagebox.showinfo("Save Config", f"Configuration saved to {saved}")

    def on_choose_file(self) -> None:
        path_str = self.filedialog.askopenfilename(
            title="Select Excel or CSV File",
            filetypes=[
                ("Tabular files", "*.xlsx *.xlsm *.xls *.csv *.tsv"),
                ("Excel files", "*.xlsx *.xlsm *.xls"),
                ("CSV/TSV files", "*.csv *.tsv"),
                ("All files", "*.*"),
            ],
        )
        if path_str:
            self._load_file(Path(path_str))

    def on_reload_file(self) -> None:
        if self.current_file is None:
            self.messagebox.showwarning("No File", "Please select a file first.")
            return
        self._load_file(self.current_file, keep_sheet=True)

    def _on_sheet_changed(self, _event: Any) -> None:
        if self.current_file is not None:
            self._load_file(self.current_file, keep_sheet=True)

    def _load_file(self, path: Path, keep_sheet: bool = False) -> None:
        try:
            sheet_arg: str | int | None = None
            if is_excel_file(path):
                sheets = list_excel_sheets(path)
                self.sheet_combo.configure(values=sheets, state="readonly")
                if keep_sheet and self.sheet_var.get() in sheets:
                    sheet_arg = self.sheet_var.get()
                else:
                    sheet_arg = sheets[0] if sheets else 0
                    if sheets:
                        self.sheet_var.set(sheets[0])
            else:
                self.sheet_combo.configure(values=[], state="disabled")
                self.sheet_var.set("")

            delimiter = self.delimiter_var.get() or None
            encoding = self.encoding_var.get().strip() or "utf-8"
            df = load_tabular_file(
                path,
                sheet_name=sheet_arg,
                delimiter=delimiter if not is_excel_file(path) else None,
                encoding=encoding,
            )
        except Exception as exc:
            self._log(f"Failed to load '{path}': {exc}", level=logging.ERROR)
            self.messagebox.showerror("File Load Error", str(exc))
            return

        self.df = df
        self.current_file = path
        self.file_path_var.set(str(path))
        self.file_summary_var.set(f"Rows: {len(df):,} | Columns: {len(df.columns)}")
        self._populate_data_preview(df)
        self._refresh_schema_preview()
        self._log(f"Loaded '{path.name}' ({len(df)} rows, {len(df.columns)} columns).")

    def _populate_data_preview(self, df: pd.DataFrame) -> None:
        self.data_tree.delete(*self.data_tree.get_children())
        cols = [str(c) for c in df.columns]
        self.data_tree["columns"] = cols
        for col in cols:
            self.data_tree.heading(col, text=col)
            self.data_tree.column(col, width=130, minwidth=80, anchor="w")

        for row in df.head(PREVIEW_ROW_LIMIT).itertuples(index=False, name=None):
            display = ["" if pd.isna(v) else str(v) for v in row]
            self.data_tree.insert("", "end", values=display)

    def _refresh_schema_preview(self) -> None:
        self.schema_tree.delete(*self.schema_tree.get_children())
        if self.df is None:
            return

        try:
            cfg = self._build_config_from_ui(require_target=False)
            data_types = cfg.data_types
        except Exception:
            data_types = self.config.data_types

        info_df = get_data_info(self.df)
        for row in info_df.itertuples(index=False):
            sql_type = map_column_sql_type(row.logical_type, data_types)
            self.schema_tree.insert(
                "",
                "end",
                values=(
                    row.column,
                    row.pandas_dtype,
                    row.logical_type,
                    sql_type,
                    row.non_null_count,
                    row.null_count,
                ),
            )

    def _set_busy(self, busy: bool, status_text: str) -> None:
        self.is_busy = busy
        state = "disabled" if busy else "normal"
        for btn in (self.test_conn_btn, self.create_btn, self.insert_btn, self.import_btn):
            btn.configure(state=state)
        self.status_var.set(status_text)

    def _run_background(self, status_text: str, worker: Any, on_success: Any) -> None:
        if self.is_busy:
            return
        self._set_busy(True, status_text)
        self.progress["value"] = 0

        def _runner() -> None:
            try:
                result = worker()
            except Exception as exc:
                self.root.after(0, lambda err=exc: self._handle_worker_error(err))
                return
            self.root.after(0, lambda res=result: self._handle_worker_success(res, on_success))

        threading.Thread(target=_runner, daemon=True).start()

    def _handle_worker_error(self, exc: Exception) -> None:
        self._set_busy(False, f"Error: {exc}")
        self._log(f"Operation failed: {exc}", level=logging.ERROR)
        self.messagebox.showerror("Operation Failed", str(exc))

    def _handle_worker_success(self, result: Any, callback: Any) -> None:
        self._set_busy(False, "Ready")
        callback(result)

    def on_test_connection(self) -> None:
        try:
            cfg = self._build_config_from_ui(require_target=False)
        except Exception as exc:
            self.messagebox.showerror("Invalid Configuration", str(exc))
            return

        self._run_background(
            "Testing SQL Server connection...",
            lambda: test_connection(cfg),
            lambda version: (
                self._log(f"Connection succeeded: {version}"),
                self.messagebox.showinfo("Connection Successful", version),
            ),
        )

    def on_create_table(self) -> None:
        if self.df is None:
            self.messagebox.showwarning("No File Loaded", "Please select a data file first.")
            return
        try:
            cfg = self._build_config_from_ui(require_target=True)
        except Exception as exc:
            self.messagebox.showerror("Invalid Configuration", str(exc))
            return

        df_snapshot = self.df
        drop_first = cfg.if_table_exists == "replace"
        self._run_background(
            f"Creating table {cfg.schema_name}.{cfg.table_name}...",
            lambda: create_table(df_snapshot, cfg, drop_existing=drop_first),
            lambda _ddl: (
                self.progress.configure(value=100),
                self._log(f"Table {cfg.schema_name}.{cfg.table_name} is ready."),
                self.messagebox.showinfo("Create Table", "Table creation completed."),
            ),
        )

    def on_insert_data(self) -> None:
        if self.df is None:
            self.messagebox.showwarning("No File Loaded", "Please select a data file first.")
            return
        try:
            cfg = self._build_config_from_ui(require_target=True)
        except Exception as exc:
            self.messagebox.showerror("Invalid Configuration", str(exc))
            return

        df_snapshot = self.df
        self._run_background(
            f"Inserting {len(df_snapshot):,} rows...",
            lambda: insert_data(df_snapshot, cfg, progress_callback=self._update_progress),
            lambda count: (
                self.progress.configure(value=100),
                self._log(f"Inserted {count:,} rows into {cfg.schema_name}.{cfg.table_name}."),
                self.messagebox.showinfo("Insert Data", f"Inserted {count:,} rows successfully."),
            ),
        )

    def on_import_all(self) -> None:
        if self.df is None:
            self.messagebox.showwarning("No File Loaded", "Please select a data file first.")
            return
        try:
            cfg = self._build_config_from_ui(require_target=True)
        except Exception as exc:
            self.messagebox.showerror("Invalid Configuration", str(exc))
            return

        df_snapshot = self.df
        self._run_background(
            f"Importing {len(df_snapshot):,} rows into {cfg.schema_name}.{cfg.table_name}...",
            lambda: import_dataframe(df_snapshot, cfg, progress_callback=self._update_progress),
            lambda count: (
                self.progress.configure(value=100),
                self._log(
                    f"Import complete: {count:,} rows into "
                    f"{cfg.schema_name}.{cfg.table_name}."
                ),
                self.messagebox.showinfo(
                    "Import Complete", f"Imported {count:,} rows successfully."
                ),
            ),
        )

    def _update_progress(self, done: int, total: int) -> None:
        pct = (done / total) * 100 if total else 100
        self.root.after(
            0,
            lambda: (
                self.progress.configure(value=pct),
                self.status_var.set(f"Inserted {done:,} / {total:,} rows ({pct:.0f}%)"),
            ),
        )


def main() -> int:
    configure_logging()
    try:
        tk, *_ = _import_tk()
    except Exception as exc:
        print(
            f"Tkinter GUI is unavailable in this Python environment ({exc}).\n"
            "Use the CLI instead: uv run excel-to-mssql-cli --help",
            file=sys.stderr,
        )
        return 1

    root = tk.Tk()
    ExcelToMSSQLApp(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
