import os
import sqlite3

from pdf_to_excel import base_dir


conn = sqlite3.connect(os.path.join(base_dir, "xlsx.db"))
conn.execute("ALTER TABLE xlsx ADD COLUMN address TEXT")
conn.commit()
