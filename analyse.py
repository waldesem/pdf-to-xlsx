import os
import sqlite3

import pandas as pd

from pdf_to_excel import base_dir


def xlsx_to_db(xlsx_file_path):
    df = pd.read_excel(
        os.path.join(base_dir, xlsx_file_path),
    )

    headers = [f"column_{index}" for index, _ in enumerate(df.columns)]
    df.columns = headers

    df["column_0"] = df["column_0"].astype(int)
    df["column_1"] = pd.to_datetime(
        df["column_1"], dayfirst=True, format="%d.%m.%y"
    ).dt.date
    df["column_4"] = pd.to_datetime(
        df["column_4"], dayfirst=True, format="%d.%m.%y"
    ).dt.date

    conn = sqlite3.connect(
        os.path.join(base_dir, xlsx_file_path.replace(".xlsx", ".db"))
    )

    df.to_sql("xlsx", conn, if_exists="replace", index=False)


if __name__ == "__main__":
    for xlsx_file_path in os.listdir(base_dir):
        if xlsx_file_path.endswith(".xlsx"):
            xlsx_to_db(xlsx_file_path)
