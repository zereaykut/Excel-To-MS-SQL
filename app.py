import tkinter as tk
from tkinter import filedialog, messagebox
import logging
import pandas as pd
import pyodbc
import json
import warnings
warnings.filterwarnings("ignore")

logging.basicConfig(
    level=logging.DEBUG, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[logging.FileHandler('app.log')]
    )
logger = logging.getLogger(__name__)


def get_config() -> json:
    """
    Get json config for database connection
    """
    with open("config.json", "r") as f:
        config = json.load(f)
    return config

def get_data_info(df: pd.DataFrame) -> pd.DataFrame:
    """
    Get data types of given dataframe
    """
    cols = df.columns.to_list()
    types = []
    for col in cols:
        types.append(str(df[col].dtype))
    df_info = pd.DataFrame({"column": cols, "type": types})
    return df_info

def save_update_config():
    nvarchar_size = nvarchar_size_input.get()
    try:
        nvarchar_size = int(nvarchar_size)
    except Exception as e:
         messagebox.showerror("nvarchar size", e)
    
    decimal_size = decimal_size_input.get()
    try:
        decimal_size = int(decimal_size)
    except Exception as e:
         messagebox.showerror("decimal size", e)

    decimal_precision = decimal_precision_input.get()
    try:
        decimal_precision = int(decimal_precision)
    except Exception as e:
         messagebox.showerror("decimal precision", e)

    config = {
                "server_ip": server_ip_input.get(),
                "username": username_input.get(),
                "password": password_input.get(),
                "database_name": database_input.get(),
                "table_name": table_name_input.get(),
                "data_types": {
                    "object": {"type": "nvarchar", "size": nvarchar_size, "nullable": nvarchar_nullable.get()},
                    "float": {"type": "decimal", "size": decimal_size, "precision": decimal_precision, "nullable": decimal_nullable.get(), "use_sql_float": decimal_use_sql_float.get()},
                    "int": {"type": ["int", "bigint"], "type_index": int_type.get(), "nullable": int_nullable.get()},
                    "datetime": {"type": ["datetime", "date"], "type_index": date_type.get(), "nullable": date_nullable.get()}
                }
            }
    # Control response data
    with open("config.json", "w") as outfile:
       outfile.write(json.dumps(config, indent=4))
    
    messagebox.showinfo("Save/Update config.json", "config.json Saved/Updated")

def choose_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls"), ("CSV files", "*.csv")])
    if file_path:
        if file_path.endswith(('.xlsx', '.xls', '.csv')):
            global df
            global df_info
            if file_path.endswith(('.xlsx', '.xls')):
                df = pd.read_excel(file_path)
                df_info = get_data_info(df)
            elif file_path.endswith(('.csv')):
                df = pd.read_csv(file_path)
                df_info = get_data_info(df)
            messagebox.showinfo("Choose File", f"Selected File: {file_path}")
        else:
            messagebox.showerror("Invalid file", "Please select a valid Excel or CSV file.")

def create_table(df_info: pd.DataFrame, table_name: str, server_ip: str, database: str, username: str, password: str, data_types: dict) -> None:
    """
    Create table according to given table name if given table not exits
    """
    cols = df_info["column"]
    types = df_info["type"]
    query_cols = ""

    nvarchar_size = nvarchar_size_input.get()
    try:
        nvarchar_size = int(nvarchar_size)
    except Exception as e:
         messagebox.showerror("nvarchar size", e)
    
    decimal_size = decimal_size_input.get()
    try:
        decimal_size = int(decimal_size)
    except Exception as e:
         messagebox.showerror("decimal size", e)

    decimal_precision = decimal_precision_input.get()
    try:
        decimal_precision = int(decimal_precision)
    except Exception as e:
         messagebox.showerror("decimal precision", e)

    datetime_type = data_types["datetime"]["type"][date_type.get()]
    datetime_null = "NULL" if date_nullable.get() == 1 else ""

    object_type = data_types["object"]["type"]
    object_size = nvarchar_size
    object_null = "NULL" if nvarchar_nullable.get() == 1 else ""

    int_type = data_types["int"]["type"][date_type.get()]
    int_null = "NULL" if int_nullable.get() == 1 else ""

    float_type = data_types["float"]["type"]
    float_size = decimal_size
    float_precision = decimal_precision
    float_null = "NULL" if decimal_nullable.get() == 1 else ""
    float_use_sql_float = decimal_use_sql_float.get()

    for col, type_ in zip(cols, types):
        if "datetime" in type_:
            query_cols = f"""{query_cols}, [{col}] [{datetime_type}] {datetime_null}"""
        elif "object" in type_:
            query_cols = f"""{query_cols}, [{col}] [{object_type}] ({object_size}) {object_null}"""
        elif "int" in type_:
            query_cols = f"""{query_cols}, [{col}] [{int_type}] {int_null}"""
        elif "float" in type_:
            if float_use_sql_float == 0:
                query_cols = f"""{query_cols}, [{col}] [{float_type}] ({float_size}, {float_precision}) {float_null}"""
            else:
                query_cols = f"""{query_cols}, [{col}] [float] {float_null}"""
    query = f"""IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = '{table_name}')
            BEGIN
                CREATE TABLE {table_name} (
                    {query_cols[1:]}
                );
            END;
            """
    
    logger.debug(f"Create Table Query: \n\t{query}")

    conn_query = f"DRIVER=SQL Server;SERVER={server_ip};DATABASE={database};UID={username};PWD={password}"
    conn = pyodbc.connect(conn_query)
    cursor = conn.cursor()
    cursor.execute(query)
    conn.commit()

    conn.close()

def create_table_tk():
    database = database_input.get()
    table_name = table_name_input.get()
    decimal_precision = decimal_precision_input.get()

    try:
        decimal_precision = int(decimal_precision)
    except Exception as e:
        messagebox.showerror("Not Integer", "Please enter valid integer value.")

    create_table(df_info, table_name, config["server_ip"], database, config["username"], config["password"], config["data_types"])
    
    messagebox.showinfo("Create Table", "Run Completed")

def insert_data(df: pd.DataFrame, df_info: pd.DataFrame, table_name: str, server_ip: str, database: str, username: str, password: str, precision: int , use_sql_float: str) -> None:
    """
    Insert excel data to created table
    """
    query_cols = ""
    for col in df.columns:
        query_cols = f"""{query_cols}, [{col}]"""
    
    cols = df_info["column"]
    types = df_info["type"]

    if use_sql_float == 0:
        for col, type_ in zip(cols, types):
            if "float" in type_:
                df[col] = df[col].round(precision)

    query_insert = ""
    for col, type_ in zip(cols, types):
        if "date" in type_:
            query_insert = f"""{query_insert}, CONVERT(DATETIME, ?)"""
        else:
            query_insert = f"""{query_insert}, ?"""

    query = f"""INSERT INTO [{table_name}] ({query_cols[1:]})
                    VALUES ({query_insert[1:]});"""

    conn = pyodbc.connect(
        f"DRIVER=SQL Server;SERVER={server_ip};DATABASE={database};UID={username};PWD={password}"
    )
    cursor = conn.cursor()

    for index, row in df.iterrows():
        # Control row if any empty values
        data_ = []
        for col in cols:
            val = row.loc[col]
            if str(val) == "nan":
                data_.append(None)
            else:
                data_.append(val)

        logger.info(f"Insert Data Query: \n{query}\nInserted Data: \n{data_}")
        cursor.execute(query, data_)
    conn.commit()
    conn.close()

def insert_data_tk():
    database = database_input.get()
    table_name = table_name_input.get()

    insert_data(df, df_info, table_name, config["server_ip"], database, config["username"], config["password"], config["data_types"]["float"]["precision"], config["data_types"]["float"]["use_sql_float"] )
    messagebox.showinfo("Insert Data", "Run Completed")


config = get_config()

# Create the main window
root = tk.Tk()
root.title("Excel/CSV to MS SQL")
root.geometry("600x400")

#%% Config
config_label = tk.Label(root, text="Configuration", font=("Helvetica", 16))
config_label.grid(row=0, column=0, padx=20, pady=10)

# Server IP
server_ip_label = tk.Label(root, text="Server IP")
server_ip_label.grid(row=1, column=0, padx=10, pady=(20, 5), sticky="w")

server_ip_input = tk.Entry(root, width=50)
server_ip_input.insert(0, config["server_ip"])
server_ip_input.grid(row=2, column=0, padx=10, pady=(5, 20), sticky="w")

# Username
username_label = tk.Label(root, text="Username")
username_label.grid(row=1, column=1, padx=10, pady=(20, 5), sticky="w")

username_input = tk.Entry(root, width=50)
username_input.insert(0, config["username"])
username_input.grid(row=2, column=1, padx=10, pady=(5, 20), sticky="w")

# Password
password_label = tk.Label(root, text="Password")
password_label.grid(row=1, column=2, padx=10, pady=(20, 5), sticky="w")

password_input = tk.Entry(root, width=50, show="*")
password_input.insert(0, config["password"])
password_input.grid(row=2, column=2, padx=10, pady=(5, 20), sticky="w")

# Database Name
database_label = tk.Label(root, text="Database Name")
database_label.grid(row=3, column=0, padx=10, pady=(20, 5), sticky="w")

database_input = tk.Entry(root, width=50)
database_input.insert(0, config["database_name"])
database_input.grid(row=4, column=0, padx=10, pady=(5, 20), sticky="w")

# Table Name
table_name_label = tk.Label(root, text="Table Name")
table_name_label.grid(row=3, column=1, padx=10, pady=(20, 5), sticky="w")

table_name_input = tk.Entry(root, width=50)
table_name_input.insert(0, config["table_name"])
table_name_input.grid(row=4, column=1, padx=10, pady=(5, 20), sticky="w")

# Object Data Type
object_data_type_label = tk.Label(root, text="[nvarchar] size", font=("Helvetica", 10))
object_data_type_label.grid(row=5, column=0, padx=10, pady=(20, 5), sticky="w")

nvarchar_size_input = tk.Entry(root, width=50)
nvarchar_size_input.insert(0, config["data_types"]["object"]["size"])
nvarchar_size_input.grid(row=6, column=0, padx=10, pady=(5, 20), sticky="w")

nvarchar_nullable = tk.IntVar(value=config["data_types"]["object"]["nullable"])
nvarchar_nullable_checkbox = tk.Checkbutton(root, text="Nullable", variable=nvarchar_nullable, onvalue=1, offvalue=0)
nvarchar_nullable_checkbox.grid(row=6, column=1, padx=10, pady=(5, 20), sticky="w")

# Decimal Data Type
decimal_data_type_label = tk.Label(root, text="[decimal] size", font=("Helvetica", 10))
decimal_data_type_label.grid(row=7, column=0, padx=10, pady=(20, 5), sticky="w")

decimal_size_input = tk.Entry(root, width=50)
decimal_size_input.insert(0, config["data_types"]["float"]["size"])
decimal_size_input.grid(row=8, column=0, padx=10, pady=(5, 20), sticky="w")

decimal_data_type_label = tk.Label(root, text="[decimal] precision", font=("Helvetica", 10))
decimal_data_type_label.grid(row=7, column=1, padx=10, pady=(20, 5), sticky="w")

decimal_precision_input = tk.Entry(root, width=50)
decimal_precision_input.insert(0, config["data_types"]["float"]["precision"])
decimal_precision_input.grid(row=8, column=1, padx=10, pady=(5, 20), sticky="w")

decimal_nullable = tk.IntVar(value=config["data_types"]["float"]["nullable"])
decimal_nullable_checkbox = tk.Checkbutton(root, text="Nullable", variable=decimal_nullable, onvalue=1, offvalue=0)
decimal_nullable_checkbox.grid(row=8, column=2, padx=10, pady=(5, 20), sticky="w")

decimal_use_sql_float = tk.IntVar(value=config["data_types"]["float"]["use_sql_float"])
decimal_use_sql_float_checkbox = tk.Checkbutton(root, text="Use SQL float data type", variable=decimal_use_sql_float, onvalue=1, offvalue=0)
decimal_use_sql_float_checkbox.grid(row=8, column=3, padx=10, pady=(5, 20), sticky="w")

# Int Data Type
int_data_type_label = tk.Label(root, text="Check to use [bigint], uncheck to use [int]", font=("Helvetica", 10))
int_data_type_label.grid(row=9, column=0, padx=10, pady=(20, 5), sticky="w")

int_type = tk.IntVar(value=config["data_types"]["int"]["type_index"])
int_type_checkbox = tk.Checkbutton(root, text="Check for [bigint]", variable=int_type, onvalue=1, offvalue=0)
int_type_checkbox.grid(row=10, column=0, padx=10, pady=(5, 20), sticky="w")

int_nullable = tk.IntVar(value=config["data_types"]["int"]["nullable"])
int_nullable_checkbox = tk.Checkbutton(root, text="Nullable", variable=int_nullable, onvalue=1, offvalue=0)
int_nullable_checkbox.grid(row=10, column=1, padx=10, pady=(5, 20), sticky="w")

# Date Data Type
date_data_type_label = tk.Label(root, text="Check to use [date], uncheck to use [datetime]", font=("Helvetica", 10))
date_data_type_label.grid(row=11, column=0, padx=10, pady=(20, 5), sticky="w")

date_type = tk.IntVar(value=config["data_types"]["datetime"]["type_index"])
date_type_checkbox = tk.Checkbutton(root, text="Check for [date]", variable=int_type, onvalue=1, offvalue=0)
date_type_checkbox.grid(row=12, column=0, padx=10, pady=(5, 20), sticky="w")

date_nullable = tk.IntVar(value=config["data_types"]["datetime"]["nullable"])
date_nullable_checkbox = tk.Checkbutton(root, text="Nullable", variable=date_nullable, onvalue=1, offvalue=0)
date_nullable_checkbox.grid(row=12, column=1, padx=10, pady=(5, 20), sticky="w")

#%% Save/Update config.json
save_update_config_button = tk.Button(root, text="Save/Update config.json", command=save_update_config)
save_update_config_button.grid(row=13, column=0, columnspan=2, padx=10, pady=20, sticky="w")

#%% Choose File
choose_file_button = tk.Button(root, text="Choose File", command=choose_file)
choose_file_button.grid(row=14, column=0, columnspan=2, padx=10, pady=20, sticky="w")

#%% Create Table
create_table_button = tk.Button(root, text="Create Table", command=create_table_tk)
create_table_button.grid(row=14, column=1, columnspan=2, padx=10, pady=20, sticky="w")

#%% Insert Data
insert_data_button = tk.Button(root, text="Insert Data", command=insert_data_tk)
insert_data_button.grid(row=14, column=2, columnspan=2, padx=10, pady=20, sticky="w")

# Run the application
root.mainloop()
