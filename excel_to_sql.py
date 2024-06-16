#!/usr/bin/env python

import json
import pandas as pd
import pyodbc
import warnings

warnings.filterwarnings("ignore")


def get_config() -> dict:
    """
    Get json config for database connection
    """
    with open("config.json", "r") as f:
        config = json.load(f)
    return config


def get_data() -> pd.DataFrame:
    df = pd.read_excel("data.xlsx")
    return df


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


def create_table(
    df_info: pd.DataFrame,
    table_name: str,
    server_ip: str,
    database: str,
    username: str,
    password: str,
) -> None:
    """
    Create table according to given table name if given table not exits
    """
    cols = df_info["column"]
    types = df_info["type"]
    query_cols = ""
    for col, type_ in zip(cols, types):
        if "date" in type_:
            query_cols = f"""{query_cols}, [{col}] [datetime] NULL"""
        elif "object" in type_:
            query_cols = f"""{query_cols}, [{col}] [nvarchar](255) NULL"""
        elif "int" in type_:
            query_cols = f"""{query_cols}, [{col}] [int] NULL"""
        elif "float" in type_:
            query_cols = f"""{query_cols}, [{col}] [decimal] (15, 4) NULL"""
    query = f"""IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = '{table_name}')
            BEGIN
                CREATE TABLE {table_name} (
                    {query_cols[1:]}
                );
            END;
            """

    conn = pyodbc.connect(
        f"DRIVER=SQL Server;SERVER={server_ip};DATABASE={database};UID={username};PWD={password}"
    )
    cursor = conn.cursor()
    cursor.execute(query)
    conn.commit()

    conn.close()


def insert_data(
    df: pd.DataFrame,
    df_info: pd.DataFrame,
    table_name: str,
    server_ip: str,
    database: str,
    username: str,
    password: str,
) -> None:
    """
    Insert excel data to created table
    """

    query_cols = ""
    for col in df.columns:
        query_cols = f"""{query_cols}, [{col}]"""

    cols = df_info["column"]
    types = df_info["type"]
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

        cursor.execute(query, data_)
    conn.commit()
    conn.close()


def main() -> None:
    """
    Main
    """
    config = get_config()
    database = input("Database name\n")
    table_name = input("Table name\n")

    df = get_data()

    df_info = get_data_info(df)

    create_table(
        df_info,
        table_name,
        config["server_ip"],
        database,
        config["username"],
        config["password"],
    )

    insert_data(
        df,
        df_info,
        table_name,
        config["server_ip"],
        database,
        config["username"],
        config["password"],
    )


if __name__ == "__main__":
    main()
