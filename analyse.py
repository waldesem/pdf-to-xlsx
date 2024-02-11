import os
import sqlite3

import pandas as pd

base_dir = os.path.abspath(os.path.join(__file__, "..", ".."))

for xlsx_file_path in os.listdir(base_dir):
    if xlsx_file_path.endswith(".xlsx"):
        df = pd.read_excel(
            os.path.join(base_dir, xlsx_file_path),
        )
        
        headers = [str(col) for col in df.columns]
        df.columns = headers

        df['0'] = df['0'].astype(int)
        df['1'] = pd.to_datetime(df['1'], dayfirst=True, format='%d.%m.%y').dt.date
        df['4'] = pd.to_datetime(df['4'], dayfirst=True, format='%d.%m.%y').dt.date

        conn = sqlite3.connect(os.path.join(base_dir, xlsx_file_path.replace(".xlsx", ".db")))

        df.to_sql('xlsx', conn, if_exists='replace', index=False)
