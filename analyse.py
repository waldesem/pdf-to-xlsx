import os
import sqlite3

import pandas as pd

from main import base_dir


"""Converts an Excel file to a SQLite database.

Args:
    xlsx_file_path (str): Path to the Excel file to convert.

The Excel file is read into a Pandas DataFrame. The column headers are renamed to 
generic names like 'column_0'. The first two columns are converted to integer and 
date types respectively. 

A SQLite database is created with the same name as the Excel file, but with a .db
extension. The DataFrame is written to a 'xlsx' table in the database.
"""


def xlsx_to_db(xlsx_file_path):
    df = pd.read_excel(
        os.path.join(base_dir, xlsx_file_path),
    )

    headers = [f"column_{index}" for index, _ in enumerate(df.columns)]
    df.columns = headers
    df = df.replace("\n", " ", regex=True)

    # df["column_0"] = df["column_0"].astype(int)
    # df["column_1"] = pd.to_datetime(
    #     df["column_1"], dayfirst=True, format="%d.%m.%y"
    # ).dt.date
    # df["column_4"] = pd.to_datetime(
    #     df["column_4"], dayfirst=True, format="%d.%m.%y"
    # ).dt.date

    conn = sqlite3.connect(
        os.path.join(base_dir, xlsx_file_path.replace(".xlsx", ".db"))
    )

    df.to_sql("xlsx", conn, if_exists="replace", index=False)


if __name__ == "__main__":
    for xlsx_file_path in os.listdir(base_dir):
        if xlsx_file_path.endswith(".xlsx"):
            xlsx_to_db(xlsx_file_path)
